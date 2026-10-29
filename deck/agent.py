"""LLM side: learn a template once (profile), then turn chat messages into a deck spec for engine.build."""
import base64
import re
import json
import os

from openai import OpenAI

MODEL = os.environ.get("DECK_MODEL", "gpt-5.5")
_client = None


def _llm(system, parts, effort="medium"):
    global _client
    _client = _client or OpenAI(timeout=300, max_retries=3)  # gpt-5.5 calls take ~50-75s and occasionally time out
    content = []
    for p in parts:
        if not (p.endswith((".png", ".jpg")) and os.path.isfile(p)):  # paths are str too: test before treating as text
            content.append({"type": "input_text", "text": p})
        else:  # image path
            b64 = base64.b64encode(open(p, "rb").read()).decode()
            content.append({"type": "input_image", "image_url": f"data:image/png;base64,{b64}"})
    resp = _client.responses.create(
        model=MODEL, reasoning={"effort": effort},
        input=[{"role": "developer", "content": system}, {"role": "user", "content": content}],
        text={"format": {"type": "json_object"}},
    )
    return json.loads(resp.output_text)


PROFILE_SYSTEM = """You study a company's PowerPoint template so an automated system can reuse it.
For every slide you get a picture and the list of its text slots (id, current text, position in % of the slide).
Return JSON:
{"slides":[{"n":1,"purpose":"what this slide is for, in one line",
  "use_when":"what kind of content belongs here",
  "slots":{"<id>":{"role":"title|label|value|body|note|fixed","meaning":"short"}}}]}
Roles:
- title: the slide heading.
- label: a caption naming a metric/column/section (e.g. "Member base", "Frequency"). Keep unless content renames it.
- value: a number/data placeholder ("x", "x%", "xMn") or any spot that must hold the user's data.
- body: free text area (bullets, findings, descriptions).
- note: footnotes like data period/source.
- fixed: decoration/brand text that should never change (e.g. "Thank you !").
Describe slot meaning precisely using the picture (which metric, which row/column, which group)."""

TURN_SYSTEM = """You build business slide decks on the client's own PowerPoint template. You never redesign; you choose
template slides, write into their text slots, and may duplicate a template slide to make more pages.

The client's standard (non-negotiable):
- NO filler. Never write generic prose, slogans or padding. Every word is either taken from the user's material,
  or is a terse fact derived from their numbers (e.g. "Renewal 38% vs 31% LY"). If you cannot ground it, leave it out.
- Numbers are copied exactly from the material. Never invent, estimate, round differently or calculate new ones
  (no differences, sums or ratios): show a comparison as "A vs B" with both source numbers. A missing number -> "—".
  "—" is only for numbers; a text banner, heading or body area always gets a short phrase taken from the material.
- Find the story in the data: compare with prior period / target / benchmark when the material has them, and mark
  increases "up", decreases "down", the 1-3 numbers that matter most on a slide "key". The emphasis styling shows
  the direction (arrow + colour), so never write the words "up"/"down"/"上升"/"下降" into the text itself.
  In a slot holding several items with different directions, leave its emphasis out and start each line with
  "▲ " or "▼ " instead; those lines get coloured individually.
- Respect each slot's size (max = approx. CJK characters; Latin text fits ~1.8x as many; "1 line" = no line breaks).
  Shorten, never overflow.
- A number goes only into a slot whose meaning matches it (a 2024 penetration rate never sits in a member-count slot).
  If no slot matches, put it in the nearest matching label as "A vs B" or leave it out.
- A slot designed for a calculated gap (e.g. "+47%" between two bars) gets the compared source number instead
  (left bar = the subject's value, right bar = what it is compared with). "—" means the material lacks the data.
- Write in the language of the user's material unless told otherwise. Labels keep the template's meaning, translated
  ("Launch city active members" -> "试点城市活跃会员"); rename one only when the content in its row/column changes.
- Slots marked "has highlight style" mix a highlight colour into their text in the template: wrap the one key
  phrase or number per line in **double asterisks** there, the way the template highlights its key words.
- A slot that is empty in the template (now: '') is deliberate spacing: leave it empty.
- Never write page markers, slot names or category names as a slot's content ("第二页", "情感权益" in a text slot) to
  fill space; a slot with nothing from the material is removed if the rules allow, otherwise it takes "—".
- Match each slot's text density and line structure in the template (e.g. a card with lead line / key line / detail
  line gets three lines) as far as the material supports it; every line still comes from the material. Supporting
  lines can be the facts that belong to the item (under a goal: its mechanism, its metric, e.g. "年均到店 7.2次").
- A small cell (max under 10) holds one number; put its comparison in the neighbouring cell or the label, not inside.
- Extend smartly. A list only slightly over a slide's capacity (e.g. 5 short agenda sections for 4 slots) is condensed
  by merging closely related items ("渠道 & 权益"), but only if the merged text still fits each slot's size; never cram
  two items' full content into one panel. Otherwise repeat that template slide and give each copy a balanced share
  (3 cities on a 2-panel slide -> 2 slides, panels 北京 | 上海 and 广州 | removed/"—"). Parallel items (cities,
  regions, years, products) always get identical treatment: never put one of them in a banner or footer; continue a title with " (2/2)" only if the copies would otherwise look identical.
  If a repeating list (agenda entries, list rows) has more entries than items, delete the unused entries via
  "remove" (label and value together). Never remove part of a grid, table, bar row, ranking row or any region whose
  removal would leave an empty hole in the design: fill it from the material (a ranking is the source numbers sorted,
  with their names) or write "—". The finished slide must look as full and balanced as the template slide.
- Choose only template slides whose structure matches the content; skip the rest. Keep a logical order
  (cover/agenda -> overview -> performance -> diagnosis -> direction -> closing) when such slides exist.

When a current deck exists, apply the user's message as an edit: change only what they asked, keep everything else
identical, and return the COMPLETE updated deck. Slide numbers the user mentions refer to the current deck (1-based).

Deck chrome is text the template prints on every slide (footer, presentation name, year). Replace its template sample
text (team names, "Name of presentation", an old year) with facts from the material, or "" when the material has none;
keep brand names. In edits keep the current chrome unless asked.

Return JSON:
{"reply":"1-3 short sentences to the user: what you did, and which data was missing (if any)",
 "chrome":{"<chrome id>":"text"},
 "slides":[{"template":<template slide n>,"fills":{"<slot id>":"text, \\n for line breaks"},
            "emphasis":{"<slot id>":"up|down|key"},"remove":["<slot id>"]}]}
Slots you do not put in "fills" keep the template's text, except value slots, which become "—"."""


def profile(analysis, thumbs):
    parts = []
    for i, s in enumerate(analysis):  # thumbnails are optional (none on Vercel); the slot list always goes
        parts.append(f"Template slide {s['n']} slots: " + json.dumps(
            [{k: x[k] for k in ("id", "text", "x", "y", "w", "h")} for x in s["slots"]], ensure_ascii=False))
        if i < len(thumbs):
            parts.append(thumbs[i])
    return _llm(PROFILE_SYSTEM, parts, effort="low")


def _max_lines(x, role):
    """Titles, labels and values keep the template text's own line count: the design is built around that
    extent, not around the (often larger) box. Body/note slots may grow to fill their box."""
    if role in ("title", "label", "value") and x["text"]:
        return min(x["text"].count("\n") + 1, x.get("lines", 99))
    if role == "body" and x["text"].count("\n") >= 1:  # a multi-line body (lead/key/detail card) is a designed
        return min(x["text"].count("\n") + 1, x.get("lines", 99))  # shape: its box is often taller than the card
    return None


def _catalog(analysis, prof):
    by_n = {p["n"]: p for p in prof.get("slides", [])}
    lines = []
    for s in analysis:
        p = by_n.get(s["n"], {})
        lines.append(f"\nTemplate slide {s['n']}: {p.get('purpose', '')} | use when: {p.get('use_when', '')}")
        roles = p.get("slots", {})
        for x in s["slots"]:
            r = roles.get(x["id"], {})
            lines.append(f"  {x['id']} [{r.get('role', '?')}] {r.get('meaning', '')} | now: {x['text'][:60]!r} | max {x['max_chars']}" + (f", max {m} line(s)" if (m := _max_lines(x, r.get("role"))) else "")
                         + (", has highlight style" if x.get("accent") else ""))
    return "\n".join(lines)


def turn(analysis, prof, docs, history, spec, message, previews=(), chrome=()):
    parts = ["TEMPLATE:" + _catalog(analysis, prof)]
    if chrome:
        parts.append("DECK CHROME (master/layout text printed on every slide):\n" +
                     "\n".join(f"  {c['id']} | now: {c['text'][:80]!r}" for c in chrome))
    if docs:
        parts.append("USER MATERIAL:\n" + "\n\n".join(f"--- {d['name']} ---\n{d['text']}" for d in docs))
    if history:
        parts.append("CONVERSATION SO FAR:\n" + "\n".join(f"{h['role']}: {h['text']}" for h in history[-12:]))
    if spec:
        parts.append("CURRENT DECK:\n" + json.dumps(spec["slides"], ensure_ascii=False))
        parts.append("Current deck rendered, slide by slide:")
        parts.extend(previews)
    parts.append("USER MESSAGE:\n" + message)
    out = _llm(TURN_SYSTEM, parts)
    return out.get("reply", ""), {"slides": out.get("slides", []), "chrome": out.get("chrome", {})}


def _nums(text):
    # numbers are a fixed grammar (digits, separators), so a regex is the right tool here
    return {n.replace(",", "").rstrip(".") for n in re.findall(r"\d[\d,]*\.?\d*", text)}


def _arrows(text):
    """A line the model ended with the word up/down gets the arrow instead (the word must never show)."""
    out = []
    for line in str(text).split("\n"):
        word = line.rstrip().rsplit(" ", 1)[-1].lower()
        if word in ("up", "down") and " " in line.strip():
            line = ("▲ " if word == "up" else "▼ ") + line.rstrip()[: -len(word)].rstrip()
        out.append(line)
    return "\n".join(out)


def _fold(text, m):
    """Merge lines beyond the slot's max into the last allowed line."""
    parts = text.split("\n")
    return text if not m or len(parts) <= m else "\n".join(parts[:m - 1] + [" ".join(parts[m - 1:])])


def finalize(spec, analysis, prof, source=""):
    """Deterministic clean-up of the model's spec: valid slides only, unfilled value slots -> '—'.
    Returns (spec, numbers in neither the material nor the template, body slots still holding template text)."""
    by_n = {p["n"]: p.get("slots", {}) for p in prof.get("slides", [])}
    ids = {s["n"]: {x["id"] for x in s["slots"]} for s in analysis}
    tpl_text = {(s["n"], x["id"]): x["text"] for s in analysis for x in s["slots"]}
    cap = {(s["n"], x["id"]): x.get("max_chars", 999) for s in analysis for x in s["slots"]}
    roles = {p["n"]: p.get("slots", {}) for p in prof.get("slides", [])}
    max_lines = {(s["n"], x["id"]): m for s in analysis for x in s["slots"]
                 if (m := _max_lines(x, roles.get(s["n"], {}).get(x["id"], {}).get("role")))}
    known = _nums(source) | _nums(" ".join(x["text"] for s in analysis for x in s["slots"]))
    unsourced, unfilled, sparse = [], [], []
    out = []
    for item in spec["slides"]:
        n = int(item.get("template", 0))
        if n not in ids:
            continue
        fixed = {k for k, r in by_n.get(n, {}).items() if r.get("role") == "fixed"}  # brand text keeps its styling
        fills = {k: _arrows(v) for k, v in item.get("fills", {}).items() if k in ids[n] and k not in fixed}
        fills = {k: _fold(v, max_lines.get((n, k))) for k, v in fills.items()}  # extra lines would overlap the next shape
        for v in fills.values():
            unsourced += [x for x in _nums(v) if x not in known and x not in unsourced and not (x.isdigit() and int(x) < 10)]
        remove = [k for k in item.get("remove", []) if k in ids[n]]
        for sid, r in by_n.get(n, {}).items():
            if r.get("role") == "value" and sid in ids[n] and sid not in fills and sid not in remove:
                fills[sid] = "—"
            elif r.get("role") == "body" and sid in ids[n] and sid not in remove and tpl_text[(n, sid)] \
                    and str(fills.get(sid, tpl_text[(n, sid)])).strip() == tpl_text[(n, sid)]:
                unfilled.append((len(out) + 1, sid))  # template sample text would show as filler
        emph = {k: v for k, v in item.get("emphasis", {}).items() if k in fills and v in ("up", "down", "key")}
        lines = {k: max_lines[(n, k)] for k in fills if (n, k) in max_lines}
        out.append({"template": n, "fills": fills, "emphasis": emph, "remove": remove, "lines": lines})
        for sid, r in by_n.get(n, {}).items():  # far sparser than the template's own text in that box
            t = tpl_text.get((n, sid), "")
            if r.get("role") == "body" and t.count("\n") >= 2 and str(fills.get(sid, "")).count("\n") < 2 \
                    and sid in fills and sid not in remove:
                sparse.append((len(out), sid, t.count("\n") + 1))
        for sid, v in fills.items():  # far over capacity: the engine would shrink it to unreadable
            if len(str(v).replace("**", "")) > 1.5 * cap.get((n, sid), 999) + 2:
                sparse.append((len(out), sid, -cap[(n, sid)]))
    return {"slides": out, "chrome": {k: str(v) for k, v in spec.get("chrome", {}).items()}}, unsourced, unfilled, sparse


if __name__ == "__main__":  # self-check of the deterministic parts (no API call)
    a = [{"n": 1, "slots": [{"id": "2", "text": "x%"}, {"id": "3", "text": "body"}, {"id": "4", "text": "Renewal Rate", "lines": 2}]}]
    p = {"slides": [{"n": 1, "slots": {"2": {"role": "value"}, "4": {"role": "label"}, "3": {"role": "body"}}}]}
    s, bad, _, _ = finalize({"slides": [{"template": 1, "fills": {"3": "免费礼品 61% vs 55% up\n高 15 元", "4": "续费率\n2024 38%"}}]}, a, p, "61% 55% 2024 38%")
    assert s["slides"][0]["fills"] == {"3": "▲ 免费礼品 61% vs 55%\n高 15 元", "4": "续费率 2024 38%", "2": "—"}, s
    assert bad == ["15"], bad
    assert s["slides"][0]["lines"] == {"4": 1, "2": 1}, s["slides"][0]["lines"]
    p["slides"][0]["slots"]["3"]["role"] = "fixed"
    assert "3" not in finalize({"slides": [{"template": 1, "fills": {"3": "Thanks"}}]}, a, p)[0]["slides"][0]["fills"]
    print("ok")
