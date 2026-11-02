import { readBody } from '../_llm.js'

// Stock-image search proxy (Pexels). Maps Pexels results to the { data: [{id, src}], total }
// shape the ImageLibPanel expects.

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const key = process.env.PEXELS_API_KEY
    const { query = '', per_page = 20, page = 1, orientation = 'all' } = await readBody(req)
    // Unsplash: the second source, used when Pexels is unset or finds nothing
    const unsplash = async () => {
      const id = process.env.UNSPLASH_CLIENT_ID
      if (!id) return null
      const q = new URLSearchParams({ query, per_page: String(per_page), page: String(page), client_id: id })
      if (orientation && orientation !== 'all') q.set('orientation', orientation === 'square' ? 'squarish' : orientation)
      const u = await fetch(`https://api.unsplash.com/search/photos?${q}`)
      if (!u.ok) return null
      const d = await u.json()
      return { data: (d.results || []).map((p: any) => ({ id: `u${p.id}`, src: p.urls?.regular, credit: p.user?.name ? `${p.user.name} / Unsplash` : 'Unsplash' })), total: d.total || 0 }
    }
    if (!key) { res.status(200).json((await unsplash()) || { data: [], total: 0 }); return }
    const params = new URLSearchParams({ query, per_page: String(per_page), page: String(page) })
    if (orientation && orientation !== 'all') params.set('orientation', orientation)

    const r = await fetch(`https://api.pexels.com/v1/search?${params}`, {
      headers: { Authorization: key },
    })
    const data = r.ok ? await r.json() : {}
    const photos = (data.photos || []).map((p: any) => ({ id: String(p.id), src: p.src?.large || p.src?.original }))
    if (!photos.length) { res.status(200).json((await unsplash()) || { data: [], total: 0 }); return }
    res.status(200).json({ data: photos, total: data.total_results || 0 })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
