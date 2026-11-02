import { getKey, readBody } from '../_llm.js'

// Real lookups for the chat agent, so nothing it calls a source or a photo of a real thing is made up.
//   {kind: "web", query}        -> {answer, sources: [{title, url}]}  (OpenAI web search; sources are the pages it read)
//   {kind: "wikimedia", query}  -> {src, title, credit, license, page} (Wikimedia Commons, freely licensed files)

const web = async (query: string) => {
  const r = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${getKey()}` },
    body: JSON.stringify({
      model: process.env.RESEARCH_MODEL || 'gpt-5.4-mini',
      tools: [{ type: 'web_search' }],
      tool_choice: 'required',
      input: `Search the web and answer for a presentation slide, in at most 8 short factual lines. Every fact must come from a page you read; give numbers with their year. Question: ${query}`,
    }),
  })
  if (!r.ok) throw new Error(`web search ${r.status}: ${(await r.text()).slice(0, 200)}`)
  const data = await r.json()
  const msg = (data.output || []).filter((o: any) => o.type === 'message').flatMap((o: any) => o.content || [])
  const answer = msg.map((c: any) => c.text || '').join('\n').trim()
  const seen = new Set<string>()
  const sources = msg.flatMap((c: any) => c.annotations || []).filter((a: any) => a.type === 'url_citation' && a.url && !seen.has(a.url) && seen.add(a.url))
    .map((a: any) => ({ title: a.title || a.url, url: String(a.url).replace(/[?&]utm_source=openai$/, '') }))
  return { answer, sources }
}

const wikimedia = async (query: string) => {
  const params = new URLSearchParams({
    action: 'query', format: 'json', generator: 'search', gsrsearch: `${query} filetype:bitmap`, gsrnamespace: '6', gsrlimit: '5',
    prop: 'imageinfo', iiprop: 'url|extmetadata|size|mime', iiurlwidth: '1600', origin: '*',
  })
  const r = await fetch(`https://commons.wikimedia.org/w/api.php?${params}`, { headers: { 'User-Agent': 'ChatSlide/1.0 (contact@chatslide.ai)' } })
  if (!r.ok) throw new Error(`wikimedia ${r.status}`)
  const pages = Object.values((await r.json()).query?.pages || {}) as any[]
  // photographs come as JPEG; PNG/SVG on Commons are mostly drawings, maps and diagrams: a "real photo" wants the former
  const photo = (p: any) => /jpe?g$/i.test(p.imageinfo?.[0]?.mime || '') ? 0 : 1
  const best = pages.sort((a, b) => photo(a) - photo(b) || (a.index || 0) - (b.index || 0)).find(p => p.imageinfo?.[0]?.thumburl)
  if (!best) return { src: '' }
  const i = best.imageinfo[0], m = i.extmetadata || {}
  const plain = (h = '') => String(h).replace(/<[^>]+>/g, '').trim()
  return {
    src: i.thumburl, title: String(best.title).replace(/^File:/, ''), page: i.descriptionurl,
    credit: plain(m.Artist?.value) || 'Wikimedia Commons', license: plain(m.LicenseShortName?.value),
  }
}

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const { kind = '', query = '' } = await readBody(req)
    if (!String(query).trim()) { res.status(400).json({ error: 'query required' }); return }
    if (kind === 'web') res.status(200).json(await web(String(query).slice(0, 500)))
    else if (kind === 'wikimedia') res.status(200).json(await wikimedia(String(query).slice(0, 200)))
    else res.status(400).json({ error: 'kind must be web or wikimedia' })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
