"""Stateless template-deck actions. The browser keeps the session and sends what each call needs; files live in
blob storage (deck/blob.py). Entry points: api/deck.py on Vercel, `python3 -m deck.service <action>` in dev."""
import csv
import io
import json
import os
import sys
import tempfile

from . import agent, blob, engine


def _tmp(data, suffix):
    f = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    f.write(data)
    f.close()
    return f.name


def to_text(name, raw):
    ext = os.path.splitext(name)[1].lower()
    if ext in (".txt", ".md", ".csv", ".tsv", ".json"):
        return raw.decode("utf-8", "replace")
    path = _tmp(raw, ext)
    try:
        if ext == ".xlsx":
            import openpyxl
            out = []
            for ws in openpyxl.load_workbook(path, data_only=True).worksheets:
                buf = io.StringIO()
                w = csv.writer(buf)
                for row in ws.iter_rows(values_only=True):
                    if any(v is not None for v in row):
                        w.writerow(["" if v is None else v for v in row])
                out.append(f"[sheet {ws.title}]\n{buf.getvalue()}")
            return "\n".join(out)
        if ext == ".docx":
            import docx
            d = docx.Document(path)
            lines = [p.text for p in d.paragraphs if p.text.strip()]
            for t in d.tables:
                lines += [" | ".join(c.text for c in r.cells) for r in t.rows]
            return "\n".join(lines)
        if ext == ".pptx":
            return "\n".join(f"[slide {s['n']}] " + " / ".join(x["text"] for x in s["slots"] if x["text"])
                             for s in engine.analyze(path))
        if ext == ".pdf":
            from pypdf import PdfReader
            # page markers: "page 2 of the PDF" is found by read_file, and the router knows how many pages there are
            return "\n".join(f"[page {i}]\n{p.extract_text() or ''}" for i, p in enumerate(PdfReader(path).pages, 1))
        raise ValueError(f"unsupported file type {ext}")
    finally:
        os.unlink(path)


def read(p):
    """{url, name} -> {name, text}: an uploaded content file as text (capped)."""
    text = to_text(p["name"], blob.get(p["url"]))
    return {"name": p["name"], "text": text[:200000], "chars": len(text)}


def learn(p):
    """{templateUrl} -> what the model learns once about a template."""
    path = _tmp(blob.get(p["templateUrl"]), ".pptx")
    try:
        analysis = engine.analyze(path)
        return {"templateUrl": p["templateUrl"], "analysis": analysis, "profile": agent.profile(analysis, []),
                "chrome": engine.chrome(path)}
    finally:
        os.unlink(path)


def turn(p):
    """{session:{templateUrl, analysis, profile, chrome, docs, history, spec}, message} -> {reply, spec, deckUrl}."""
    s, message = p["session"], p["message"]
    analysis, prof, chrome, docs = s["analysis"], s["profile"], s.get("chrome", []), s.get("docs", [])
    history, spec = s.get("history", []), s.get("spec")
    reply, raw = agent.turn(analysis, prof, docs, history, spec, message, (), chrome)
    source = "\n".join([d["text"] for d in docs] + [h["text"] for h in history if h["role"] == "user"] + [message])
    spec, unsourced, unfilled, sparse = agent.finalize(raw, analysis, prof, source)
    stale = lambda: [c["id"] for c in chrome if spec["chrome"].get(c["id"], c["text"]).strip() == c["text"]]
    if unfilled or sparse or stale():  # one repair pass: leftover sample text reads as filler, a sparse box as a hole
        asks = [f"slide {i} slot {sid} still shows the template's sample text: fill it with a short phrase from the "
                f"material, or remove it if the slide reads well without it" for i, sid in unfilled]
        asks += [f"slide {i} slot {sid} has one line where the template has {k}: give it the template's line structure "
                 f"using facts from the material that belong to that item" for i, sid, k in sparse if k > 0]
        box = {(sl["n"], x["id"]): x.get("lines") for sl in analysis for x in sl["slots"]}
        tpl_of = lambda i: spec["slides"][i - 1]["template"]
        asks += [f"slide {i} slot {sid} is over its box (max {-k} characters, {box.get((tpl_of(i), sid)) or 'few'} "
                 f"lines): shorten it, keeping the most important items, moving a comparison to its neighbouring "
                 f"cell or label if needed" for i, sid, k in sparse if k < 0]
        asks += [f"deck chrome {cid} still shows the template's sample text: replace it with facts from the material "
                 f"(this deck's subject • period, e.g. program • report • year)" for cid in stale()]
        ask = "Fix only these slots and change nothing else: " + "; ".join(asks)
        _, raw = agent.turn(analysis, prof, docs, [], spec, ask, (), chrome)
        spec, unsourced, unfilled, sparse = agent.finalize(raw, analysis, prof, source)
        for i, sid in unfilled:  # still sample text after the repair: blank it, the shape (e.g. a banner) stays
            spec["slides"][i - 1]["fills"][sid] = ""
        for cid in stale():
            spec["chrome"][cid] = ""
    if unsourced:
        reply += f"\n⚠ Not in your material, please check: {', '.join(unsourced)}"
    if not spec["slides"]:
        return {"reply": reply or "No slides produced.", "spec": s.get("spec"), "deckUrl": None}
    tpl = _tmp(blob.get(s["templateUrl"]), ".pptx")
    out = tpl.replace(".pptx", "-deck.pptx")
    try:
        engine.build(tpl, spec, out)
        url = blob.put("deck.pptx", open(out, "rb").read(),
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    finally:
        os.unlink(tpl)
        if os.path.exists(out):
            os.unlink(out)
    return {"reply": reply, "spec": spec, "deckUrl": url}


ACTIONS = {"read": read, "learn": learn, "turn": turn}


def drop(p):
    """{url} -> remove a one-time upload (a deck opened from ChatSlide) once it has been read."""
    return {"dropped": blob.drop(p["url"])}


ACTIONS["drop"] = drop


def run(action, payload):
    if action not in ACTIONS:
        raise ValueError(f"unknown action {action}")
    return ACTIONS[action](payload)


if __name__ == "__main__":  # dev: JSON on stdin -> JSON on stdout (the vite dev plugin calls this)
    try:
        print(json.dumps(run(sys.argv[1], json.load(sys.stdin)), ensure_ascii=False))
    except Exception as e:  # surface the real error to the panel
        print(json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False))
