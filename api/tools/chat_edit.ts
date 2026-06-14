import { chat, readBody } from '../_llm'

// Edit chatbot: given the current slide's elements + a natural-language instruction,
// return the updated elements array. The client merges the result onto the live slide.

const SYSTEM = `You are a slide-editing assistant for a PowerPoint-like editor.
You edit ONE slide. The canvas is {W}px wide and {H}px tall, origin at the top-left, units in px.

Each element is a JSON object with at least: id, type, left, top, width, height, rotate.
- type "text": has "content" (an HTML string, e.g. "<p>Hello</p>"), "defaultColor", "defaultFontName", optional "fill".
- type "image": has "src" (URL).
- type "shape": has "fill", "path", "viewBox".

Rules:
- Apply the user's instruction by modifying the elements.
- You MAY change text content, colors, sizes, positions; move/resize/rotate; add new elements; or delete elements.
- KEEP element "id" stable for elements you keep. Only omit an element if the user wants it deleted.
- For a NEW element, invent a short unique "id" and include all required fields for its type.
- Keep every element fully on-canvas (0..{W} horizontally, 0..{H} vertically).
- Preserve fields you are not changing.
- Text "content" must stay valid HTML.

Respond ONLY with JSON: {"reply": "<one short sentence to the user>", "elements": [ ...full updated array... ]}.`

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const { elements = [], instruction = '', history = [], slideSize } = await readBody(req)
    const W = Math.round(slideSize?.width || 1000)
    const H = Math.round(slideSize?.height || 562)

    const sys = SYSTEM.replaceAll('{W}', String(W)).replaceAll('{H}', String(H))
    const messages = [
      { role: 'system', content: sys },
      ...history.slice(-6),
      {
        role: 'user',
        content: `Current slide elements:\n${JSON.stringify(elements)}\n\nInstruction: ${instruction}`,
      },
    ]

    const raw = await chat(messages, { jsonMode: true, model: 'gpt-4o', temperature: 0.3, maxTokens: 8192 })
    let parsed: any
    try { parsed = JSON.parse(raw) }
    catch { res.status(200).json({ reply: raw || 'Sorry, I could not edit that.', elements: null }); return }

    res.status(200).json({
      reply: parsed.reply || 'Done.',
      elements: Array.isArray(parsed.elements) ? parsed.elements : null,
    })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
