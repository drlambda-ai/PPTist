import { chat, metered, readBody } from '../_llm.js'

// Small JSON helper for the AI panel (chart data, speaker notes, picking an element): one instruction in, one JSON
// object out.

async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST only' }); return
  }
  try {
    const { instruction = '', strong = false, model: asked = '', image = '' } = await readBody(req)
    // a picture to read (the user's attached picture: "copy the text in my image"); only an https link or a picture
    const pic = /^(https:\/\/|data:image\/(png|jpeg|webp);base64,)/.test(String(image)) ? String(image) : ''
    const raw = await chat([
      { role: 'system', content: 'Answer with one JSON object only, exactly in the shape the instruction asks for.' },
      { role: 'user', content: pic ? [{ type: 'text', text: String(instruction).slice(0, 30000) }, { type: 'image_url', image_url: { url: pic } }] : String(instruction).slice(0, 30000) },
    ], { model: process.env.VERCEL_ENV !== 'production' && ['gpt-6-luna', 'gpt-5.5', 'gpt-6.1-sol', 'gpt-6.1-terra', 'claude-sonnet-5-5', 'claude-haiku-5-5'].includes(asked) ? asked : strong ? 'gpt-6.1-sol' : 'gpt-6-luna', effort: 'low', jsonMode: true, maxTokens: 8000 }) // numbers, notes, extensions: gpt-6.1-sol (exact with figures, 1/3 of gpt-5.5)
    res.status(200).json(JSON.parse(raw || '{}'))
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}

export default metered(handler)
