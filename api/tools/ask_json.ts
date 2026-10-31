import { chat, readBody } from '../_llm.js'

// Small JSON helper for the AI panel (chart data, speaker notes, picking an element): one instruction in, one JSON
// object out.

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST only' }); return
  }
  try {
    const { instruction = '', strong = false } = await readBody(req)
    const raw = await chat([
      { role: 'system', content: 'Answer with one JSON object only, exactly in the shape the instruction asks for.' },
      { role: 'user', content: String(instruction).slice(0, 30000) },
    ], { model: strong ? 'gpt-5.5' : 'gpt-5.4-mini', effort: 'low', jsonMode: true, maxTokens: 8000 }) // numbers for charts/tables use the strong model
    res.status(200).json(JSON.parse(raw || '{}'))
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
