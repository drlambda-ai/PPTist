import { getKey, readBody } from '../_llm.js'

// AI image generation via OpenAI gpt-image-1. Returns a data URL the client can drop
// straight into an image element's src.

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const { prompt = '', size = '1536x1024' } = await readBody(req)
    if (!prompt) { res.status(400).json({ error: 'prompt required' }); return }

    const r = await fetch('https://api.openai.com/v1/images/generations', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${getKey()}`,
      },
      body: JSON.stringify({ model: 'gpt-image-1', prompt, size, n: 1 }),
    })
    if (!r.ok) { res.status(502).json({ error: `image api ${r.status}: ${await r.text()}` }); return }
    const data = await r.json()
    const b64 = data.data?.[0]?.b64_json
    if (!b64) { res.status(502).json({ error: 'no image returned' }); return }

    res.status(200).json({ src: `data:image/png;base64,${b64}` })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
