import { addCost, addUsage, chat, getKey, metered, readBody } from '../_llm.js'

// Real lookups for the chat agent, so nothing it calls a source or a photo of a real thing is made up.
//   {kind: "web", query}        -> {answer, sources: [{title, url}]}  (OpenAI web search; sources are the pages it read)
//   {kind: "wikimedia", query}  -> {src, title, credit, license, page} (Wikimedia Commons, freely licensed files)

const web = async (query: string) => {
  const r = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${getKey()}` },
    body: JSON.stringify({
      model: process.env.RESEARCH_MODEL || 'gpt-6-luna',
      tools: [{ type: 'web_search' }],
      tool_choice: 'required',
      input: `Search the web and answer for a presentation slide, in at most 8 short factual lines. Every fact must come from a page you read; give numbers with their year. Question: ${query}`,
    }),
  })
  if (!r.ok) throw new Error(`web search ${r.status}: ${(await r.text()).slice(0, 200)}`)
  const data = await r.json()
  addUsage(process.env.RESEARCH_MODEL || 'gpt-6-luna', data.usage)
  addCost(0.025 * Math.max(1, (data.output || []).filter((o: any) => o.type === 'web_search_call').length)) // per search, rounded up
  const msg = (data.output || []).filter((o: any) => o.type === 'message').flatMap((o: any) => o.content || [])
  const answer = msg.map((c: any) => c.text || '').join('\n').trim()
  const seen = new Set<string>()
  const sources = msg.flatMap((c: any) => c.annotations || []).filter((a: any) => a.type === 'url_citation' && a.url && !seen.has(a.url) && seen.add(a.url))
    .map((a: any) => ({ title: a.title || a.url, url: String(a.url).replace(/[?&]utm_source=openai$/, '') }))
  return { answer, sources }
}

const UA = { 'User-Agent': 'ChatSlide/1.0 (contact@chatslide.ai)' }
const commons = async (params: Record<string, string>) => {
  const r = await fetch(`https://commons.wikimedia.org/w/api.php?${new URLSearchParams({
    action: 'query', format: 'json', prop: 'imageinfo', iiprop: 'url|extmetadata|size|mime', iiurlwidth: '1600', origin: '*', ...params,
  })}`, { headers: UA })
  if (!r.ok) throw new Error(`wikimedia ${r.status}`)
  return (Object.values((await r.json()).query?.pages || {}) as any[]).filter(p => p.imageinfo?.[0]?.thumburl)
}

// Ground truth first: Wikidata lists each thing's own picture files (its photo, locator/detail maps, flag, logo).
// Commons full-text search returned 1800s engravings and a night-lights photo for "map of the United States".
const KIND_PROPS: Record<string, string[]> = { map: ['P1621', 'P1943', 'P242'], flag: ['P41'], logo: ['P154'], photo: ['P18'] }
const wikidataFiles = async (query: string) => {
  const plan = JSON.parse(await chat([{ role: 'user', content: `A slide needs a picture: "${query}". Return JSON {"subject": "<the real-world thing, its English name, e.g. United States>", "kind": "map|flag|logo|photo"}` }], { model: 'gpt-6-luna', jsonMode: true }))
  const props = KIND_PROPS[plan.kind] || KIND_PROPS.photo
  const found = await (await fetch(`https://www.wikidata.org/w/api.php?${new URLSearchParams({ action: 'wbsearchentities', search: String(plan.subject || query), language: 'en', format: 'json', limit: '1' })}`, { headers: UA })).json()
  const id = found.search?.[0]?.id
  if (!id) return []
  const claims = (await (await fetch(`https://www.wikidata.org/w/api.php?action=wbgetclaims&format=json&entity=${id}`, { headers: UA })).json()).claims || {}
  const rank = (c: any) => (c.rank === 'preferred' ? 0 : 1) // the current flag/map is marked preferred over historical ones
  return props.flatMap(p => (claims[p] || []).filter((c: any) => c.rank !== 'deprecated').sort((a: any, b: any) => rank(a) - rank(b)))
    .map((c: any) => c.mainsnak?.datavalue?.value).filter((v: unknown) => typeof v === 'string').slice(0, 8) as string[]
}

const wikimedia = async (query: string, nearest = false) => {
  const plain = (h = '') => String(h).replace(/<[^>]+>/g, '').trim()
  const files = await wikidataFiles(query).catch(() => [] as string[])
  const known = files.length ? await commons({ titles: files.map(f => `File:${f}`).join('|') }).catch(() => []) : []
  // photos AND drawings: maps, flags, diagrams and logos on Commons are mostly SVG (served as a PNG thumbnail)
  const search = async (q: string) => (await commons({ generator: 'search', gsrsearch: `${q} filetype:bitmap|drawing`, gsrnamespace: '6', gsrlimit: '8' }))
    .sort((a, b) => (a.index || 0) - (b.index || 0))
  // Search rank alone picked a 1755 colonial map and a night-lights photo for "map of the United States": a small
  // model reads the candidates' titles and descriptions and picks the one that IS the requested thing, or none.
  // close: the second round takes the nearest real image of the same subject (a plain barium swallow film for "barium
  // swallow in systemic sclerosis"), still better on a slide than a stock photo of a doctor
  const pick = async (pages: any[], close = false) => {
    if (!pages.length) return null
    try {
      const list = pages.map((p, k) => `${k}: ${String(p.title).replace(/^File:/, '')} | ${plain(p.imageinfo[0].extmetadata?.ImageDescription?.value).slice(0, 160)}`).join('\n')
      const out = JSON.parse(await chat([{ role: 'user', content: `A slide needs a picture of: ${query}\nWikimedia Commons files:\n${list}\n${close ? 'Which file is the closest real picture of the same subject (same disease, organ, scan type or thing; -1 only if none is)?' : 'Which file shows exactly that (current and clear; not a historical/old version, a detail, or something only related)?'} Return JSON {"pick": <number or -1 if none fits>}.` }], { model: 'gpt-6-luna', jsonMode: true }))
      return out.pick === -1 ? null : pages[out.pick] || pages[0]
    }
    catch { return pages[0] } // the search's own first result
  }
  const searched = await search(query).catch(() => [] as any[])
  let best = await pick([...known, ...searched.filter(p => !known.some(k => k.title === p.title))])
  // Commons search needs every word to match, in Commons' own words: "chest radiograph findings in pulmonary
  // hypertension" found nothing where "pulmonary hypertension x-ray" finds 7 real films (2026-11-11). Nothing found or
  // nothing fits: three other short searches, with the common terms and synonyms, then pick again.
  let close = false
  if (!best && nearest) { // medical pictures only: elsewhere a near match (a 1755 map for the US map) is worse than a stock photo
    const alts = await chat([{ role: 'user', content: `Wikimedia Commons file search found nothing for: ${query}\nGive 3 different short searches (2 to 4 English words each) that Commons file names would use: the common term, synonyms (X-ray for radiograph), the plain name of the thing. JSON {"queries": ["...", "...", "..."]}` }], { jsonMode: true, maxTokens: 2000 })
      .then(r => (JSON.parse(r).queries || []) as string[]).catch(() => [] as string[])
    const more = (await Promise.all(alts.slice(0, 3).map(q => search(String(q)).catch(() => [] as any[])))).flat()
    best = await pick([...searched, ...more].filter((p, k, all) => all.findIndex(x => x.title === p.title) === k).slice(0, 16), true)
    close = !!best
  }
  if (!best) return { src: '' }
  const i = best.imageinfo[0], m = i.extmetadata || {}
  return {
    src: i.thumburl, title: String(best.title).replace(/^File:/, ''), page: i.descriptionurl,
    credit: plain(m.Artist?.value) || 'Wikimedia Commons', license: plain(m.LicenseShortName?.value), close,
  }
}

// Academic papers (OpenAlex: free, no key, the same corpus as Semantic Scholar without its anonymous rate limit)
const scholar = async (query: string) => {
  const params = new URLSearchParams({ search: query, 'per-page': '6', select: 'display_name,publication_year,doi,authorships,primary_location,cited_by_count', mailto: 'contact@chatslide.ai' })
  const r = await fetch(`https://api.openalex.org/works?${params}`)
  if (!r.ok) throw new Error(`openalex ${r.status}`)
  const works = ((await r.json()).results || []) as any[]
  return {
    papers: works.map(w => ({
      title: w.display_name, year: w.publication_year, cited: w.cited_by_count,
      authors: (w.authorships || []).slice(0, 3).map((a: any) => a.author?.display_name).filter(Boolean).join(', ') + ((w.authorships || []).length > 3 ? ' et al.' : ''),
      venue: w.primary_location?.source?.display_name || '', url: w.doi || w.primary_location?.landing_page_url || '',
    })).filter(p => p.title),
  }
}

// A page or a YouTube video the user pasted: its text (Supadata), so the chat can build slides from it
const readUrl = async (url: string) => {
  const key = process.env.SUPADATA_API_KEY
  if (!key) throw new Error('reading links is not configured')
  if (!/^https?:\/\//i.test(url)) throw new Error('not a web link')
  const host = new URL(url).hostname.toLowerCase()
  // public sites only: no IP literals, localhost or internal names (the fetch runs on Supadata, but never point it inward)
  if (/^[\d.]+$|^\[|:|^localhost$|\.(local|internal|localhost|lan)$/.test(host) || !host.includes('.')) throw new Error('only public web links can be read')
  const yt = /(^|\.)(youtube\.com|youtu\.be)$/i.test(host)
  const endpoint = yt ? 'https://api.supadata.ai/v1/transcript' : 'https://api.supadata.ai/v1/web/scrape'
  const params = new URLSearchParams({ url, ...(yt ? { text: 'true' } : {}) })
  const r = await fetch(`${endpoint}?${params}`, { headers: { 'x-api-key': key } })
  if (!r.ok) throw new Error(`could not read that link (${r.status})`)
  const d = await r.json()
  return { title: d.name || d.title || url, text: String(d.content || '').slice(0, 60000), kind: yt ? 'video transcript' : 'web page' }
}

async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const { kind = '', query = '', nearest = false } = await readBody(req)
    if (!String(query).trim()) { res.status(400).json({ error: 'query required' }); return }
    if (kind === 'web') res.status(200).json(await web(String(query).slice(0, 500)))
    else if (kind === 'scholar') res.status(200).json(await scholar(String(query).slice(0, 300)))
    else if (kind === 'url') res.status(200).json(await readUrl(String(query).trim().slice(0, 2000)))
    else if (kind === 'wikimedia') res.status(200).json(await wikimedia(String(query).slice(0, 200), nearest === true))
    else res.status(400).json({ error: 'kind must be web, scholar, url or wikimedia' })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}

export default metered(handler)
