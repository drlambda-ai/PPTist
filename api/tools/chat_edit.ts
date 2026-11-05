import { chat, readBody } from '../_llm.js'

// Edit chatbot: given the current slide's elements + a natural-language instruction,
// return the updated elements array. The client merges the result onto the live slide.

const SYSTEM = `You are a slide-editing assistant for a PowerPoint-like editor.
You edit ONE slide. The canvas is {W}px wide and {H}px tall, origin at the top-left, units in px.

Each element is a JSON object with at least: id, type, left, top, width, height, rotate.
- type "text": has "content" (an HTML string, e.g. "<p>Hello</p>"), "defaultColor", "defaultFontName", optional "fill".
- type "image": has "src" (a placeholder token: keep it exactly as given).
- type "chart" / "table": keep them; you may move or resize them.
Text styling lives in two places: "defaultFontName"/"defaultColor" on the element, and inline CSS in "content"
(e.g. <span style="font-family: Microsoft YaHei; font-size: 24px; color: #1a2b8c;"><strong>...</strong></span>).
To change a font, size, colour, bold/italic/underline or alignment (<p style="text-align: center;">), change BOTH
the element default and any inline style that would override it. Font sizes are px.
"Move left/right/up/down" means change left/top by a visible amount (about 5-15% of the canvas) unless a size is given;
"bigger/smaller" scales width and height together around the element's centre.
- type "shape": has "fill", "path", "viewBox" ("path" may be a placeholder token: keep it exactly as given). A shape
  may also hold words: then it has "content", "defaultColor", "defaultFontName" exactly like a text element, and you edit
  them the same way (most text boxes of imported decks are shapes).
To MOVE an element, change its left/top; never add a copy of it.

Rules:
- Apply the user's instruction by modifying the elements, and change NOTHING else: every element and property the
  instruction does not ask about (position, size, colour, font, text) stays exactly as given.
- You MAY change text content, colors, sizes, positions; move/resize/rotate; add new TEXT or SHAPE elements; or delete elements.
- NEVER create a new "image" element, and never invent or alter an image "src". You have no real images. You may reposition, resize, or delete an existing image, but you cannot add one. If the user asks for a new picture, do not add an image element; say in "reply" that pictures are changed with a separate step.
- KEEP element "id" stable for elements you keep. Only omit an element if the user wants it deleted.
- For a NEW text/shape element, invent a short unique "id" and include all required fields for its type.
- Text must stay inside its own box: when wording gets longer (translation, rewrite, more points), shorten it or
  lower its font size so it fits the element's width and height; never let it spill onto other elements.
- Keep every element fully on-canvas (0..{W} horizontally, 0..{H} vertically).
- Preserve fields you are not changing.
- Text "content" must stay valid HTML.

Respond ONLY with JSON: {"reply": "<one short sentence to the user, in the language of the instruction>", "elements": [ ...full updated array... ]}.`

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const { elements = [], instruction = '', history = [], slideSize, patch = false, model: asked = '', advanced = false } = await readBody(req)
    const W = Math.round(slideSize?.width || 1000)
    const H = Math.round(slideSize?.height || 562)

    // patch mode (AI panel): the model writes only what changes, a fraction of the tokens of the whole slide
    const sys = SYSTEM.replaceAll('{W}', String(W)).replaceAll('{H}', String(H)).replace(/Respond ONLY with JSON:[\s\S]*$/, patch
      ? 'Respond ONLY with JSON: {"reply": "<one short sentence to the user>", "changes": [{"id": "<existing id>", <only the fields you change, each with its full new value>}], "added": [<new text/shape elements, complete>], "removed": ["<id>", ...]}. Never repeat an unchanged element or field.'
      : 'Respond ONLY with JSON: {"reply": "<one short sentence to the user, in the language of the instruction>", "elements": [ ...full updated array... ]}.')
    const messages = [
      { role: 'system', content: sys },
      ...history.slice(-6),
      {
        role: 'user',
        content: `Current slide elements:\n${JSON.stringify(elements)}\n\nInstruction: ${instruction}`,
      },
    ]

    const raw = await chat(messages, { jsonMode: true, model: advanced ? 'gpt-6.1-sol' : process.env.VERCEL_ENV !== 'production' && ['gpt-6-luna', 'gpt-5.5', 'gpt-6.1-sol', 'gpt-6.1-terra', 'claude-sonnet-5-5', 'claude-haiku-5-5'].includes(asked) ? asked : (process.env.SLIDE_EDIT_MODEL || 'gpt-6-luna'), effort: 'low', maxTokens: 32000 })
    let parsed: any
    try { parsed = JSON.parse(raw) }
    catch { res.status(200).json({ reply: raw || 'Sorry, I could not edit that.', elements: null }); return }

    if (patch) {
      const ok = Array.isArray(parsed.changes) || Array.isArray(parsed.added) || Array.isArray(parsed.removed)
      const arr = (v: any) => (Array.isArray(v) ? v : []) // a malformed field becomes empty, never a crash downstream
      res.status(200).json({ reply: parsed.reply || 'Done.', patch: ok ? { changes: arr(parsed.changes), added: arr(parsed.added), removed: arr(parsed.removed) } : null })
      return
    }
    res.status(200).json({
      reply: parsed.reply || 'Done.',
      elements: Array.isArray(parsed.elements) ? parsed.elements : null,
    })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
