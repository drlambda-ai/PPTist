import { readBody } from '../_llm'

// Stock-image search proxy (Pexels). Maps Pexels results to the { data: [{id, src}], total }
// shape the ImageLibPanel expects.

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const key = process.env.PEXELS_API_KEY
    if (!key) { res.status(200).json({ data: [], total: 0 }); return }

    const { query = '', per_page = 20, page = 1, orientation = 'all' } = await readBody(req)
    const params = new URLSearchParams({ query, per_page: String(per_page), page: String(page) })
    if (orientation && orientation !== 'all') params.set('orientation', orientation)

    const r = await fetch(`https://api.pexels.com/v1/search?${params}`, {
      headers: { Authorization: key },
    })
    if (!r.ok) { res.status(200).json({ data: [], total: 0 }); return }
    const data = await r.json()
    res.status(200).json({
      data: (data.photos || []).map((p: any) => ({ id: String(p.id), src: p.src?.large || p.src?.original })),
      total: data.total_results || 0,
    })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
