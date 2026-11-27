import { addCost, addUsage, chat, getKey, metered, readBody } from '../_llm.js'
import { stockSearch } from './img_search.js'

// Real lookups for the chat agent, so nothing it calls a source or a photo of a real thing is made up.
//   {kind: "web", query}        -> {answer, sources: [{title, url}]}  (OpenAI web search; sources are the pages it read)
//   {kind: "wikimedia", query}  -> {src, title, credit, license, page} (Wikimedia Commons, freely licensed files)
//   {kind: "stock", query, want} -> {data: [{src}], fits} (stock photos a model has looked at; the first `fits` match)

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

// the model that LOOKS at candidate pictures (stock photos, the nearest Wikimedia picture): see `stock` below
const LOOK_MODEL = process.env.LOOK_MODEL || 'gpt-6.1-terra'
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

const wikimedia = async (query: string, nearest = false, avoid: string[] = [], want = '') => {
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
    // never a picture already on the page: "small infant on oxygen", then "Indian infant on oxygen" both got the same
    // 1940s oxygen-meter photo the user had just rejected (2026-11-17)
    pages = pages.filter(p => !avoid.includes(p.imageinfo[0].thumburl))
    if (!pages.length) return null
    try {
      const list = pages.map((p, k) => `${k}: ${String(p.title).replace(/^File:/, '')} | ${plain(p.imageinfo[0].extmetadata?.ImageDescription?.value).slice(0, 160)}`).join('\n')
      const ask = `A slide needs a picture of: ${query}\nWikimedia Commons files:\n${list}\n${close ? 'Which file is the closest real picture of the same subject (same disease, organ, scan type or thing; current, never a historical/old photo: a 1940s oxygen meter came back for "infant on oxygen support", 2026-11-17; -1 if none is)?' : 'Which file shows exactly that (current and clear; not a historical/old version, a detail, or something only related)?'} Return JSON {"pick": <number or -1 if none fits>}.`
      // the nearest one is LOOKED at beside the user's words: by title alone a doctor's portrait and an air force
      // crew came back for "Indian infant on oxygen support" (replay 2026-11-19)
      const small = close ? await commons({ titles: pages.map(p => p.title).join('|'), iiurlwidth: '330' }).catch(() => [] as any[]) : []
      const viewable = (u: string) => /\.(jpe?g|png|webp)$/i.test(u.split('?')[0]) // an animated GIF cannot be looked at, and failed the whole look
      // the thumbnails go inline: Wikimedia answers OpenAI's own downloader with 429, the whole look failed and a stock
      // photo went on for "breast anatomy" twice (2026-11-23, reproduced 2026-11-24)
      const inline = async (u: string) => {
        const r = await fetch(u, { headers: UA }).catch(() => null)
        return r?.ok ? `data:${r.headers.get('content-type') || 'image/jpeg'};base64,${Buffer.from(await r.arrayBuffer()).toString('base64')}` : ''
      }
      const shown = (await Promise.all(pages.map(async (p, k) => {
        const x = small.find(x => x.title === p.title && viewable(x.imageinfo[0].thumburl))
        return { k, url: x ? await inline(x.imageinfo[0].thumburl) : '' }
      }))).filter(c => c.url)
      // a nearest picture nobody could look at is never placed: no thumbnails came, or only unviewable files
      if (close && !shown.length) return null
      const seen = shown.flatMap(c => [{ type: 'text', text: `${c.k}:` }, { type: 'image_url', image_url: { url: c.url, detail: 'low' } }])
      const out = JSON.parse(seen.length
        ? await chat([{ role: 'user', content: [{ type: 'text', text: `${ask}\nWhat the user wrote (any language; the latest message is last and wins):\n${want}\nThe files' pictures follow, numbered alike. Look at each: pick one only when its picture is shown and itself shows what the user asked for.` }, ...seen] }], { model: LOOK_MODEL, jsonMode: true, maxTokens: 4000 })
        : await chat([{ role: 'user', content: ask }], { model: 'gpt-6-luna', jsonMode: true }))
      if (close) return shown.some(c => c.k === out.pick) ? pages[out.pick] : null // only one of the pictures it was shown
      return out.pick === -1 ? null : pages[out.pick] || pages[0]
    }
    catch { return close ? null : pages[0] } // the search's own first result; never an unseen "nearest" one (a retinopathy animation for an infant on oxygen)
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

// Stock photos, looked at before one goes on a slide. The first search result went on unseen: seven backgrounds in a
// row for "a poverty-stricken area in a US city" were suburbs and Asian slums, "not a person" got people, "full of
// pedestrians" got two (2026-11-11..17). A vision model sees 15 candidates beside the user's own words and orders the
// ones that show it; none does -> one more search in a stock library's words, then the nearest, said as such.
// Model (2026-11-19, 5 real requests, 15 photos each, ~1,300 tokens in): gpt-6.1-terra picked the Los Angeles
// sidewalk tents for US city poverty and the crowded crossings for "full of pedestrians" (about $0.002 a look);
// gpt-6-luna ($0.0003) picked Ohio suburbs and two silhouettes; gpt-5.4-mini ($0.002) an Istanbul street.
const stock = async (query: string, want: string, orientation: string, avoid: string[]) => {
  const found = async (q: string) => ((await stockSearch({ query: q, per_page: 15, orientation })).data as any[]).filter(p => p.src)
  const fresh = (list: any[]) => list.filter(p => !avoid.includes(p.src)) // never a picture already on the page
  const look = async (list: any[]) => {
    const content = [{ type: 'text', text: `A presentation slide needs a picture.\nWhat the user wrote (any language; the latest message is last and wins):\n${want}\nThe editor's search words for this picture (they only say which picture of the message this is; the user's words decide what it must show): ${query}\n\nBelow are ${list.length} stock photos, numbered from 0, each with its caption. Look at each photo.\n"fits": the photos that show what the user asked for, with everything their words require (the subject, the place or country, how many people, what must NOT be in it), best first; [] when none really shows it.\n"nearest": the one closest photo when fits is [].\n"search": when fits is [], 2 to 4 other English words a stock photo library would file such a photo under; else "".\nReturn JSON {"fits": [numbers], "nearest": number, "search": "..."}` },
      ...list.flatMap((p, k) => [{ type: 'text', text: `${k}: ${String(p.alt || '').slice(0, 160)}` }, { type: 'image_url', image_url: { url: p.thumb || p.src, detail: 'low' } }])]
    const out = JSON.parse(await chat([{ role: 'user', content }], { model: LOOK_MODEL, jsonMode: true, maxTokens: 4000 }))
    const at = (xs: unknown) => [...new Set((Array.isArray(xs) ? xs : [xs]).map(Number))].filter(k => Number.isInteger(k) && !!list[k])
    return { fits: at(out.fits), nearest: at(out.nearest), search: String(out.search || '').trim() }
  }
  const first = await found(query)
  const ordered = (list: any[], lead: number[]) => [...lead.map(k => list[k]), ...list.filter((_, k) => !lead.includes(k)), ...first.filter(p => avoid.includes(p.src))]
  let list = fresh(first)
  if (!list.length) return { data: first } // nothing new to look at: the search's own order
  try {
    let seen = await look(list)
    if (!seen.fits.length && seen.search) {
      const more = fresh(await found(seen.search).catch(() => [] as any[])).filter(p => !list.some(x => x.src === p.src))
      const again = more.length ? await look(more) : seen
      if (again.fits.length) { list = more; seen = again }
    }
    return { data: ordered(list, seen.fits.length ? seen.fits : seen.nearest), fits: seen.fits.length }
  }
  catch (e) { // the look failed: the search's own order, as before, and no claim that anything was seen
    console.error('[research] picture look failed:', String((e as Error)?.message || e).slice(0, 160))
    return { data: ordered(list, []) }
  }
}

// Academic papers. OpenAlex first; since it meters requests without a key by the caller's IP (a shared daily budget),
// every paper search from production came back empty and the chat cited web pages instead (6 of 6, 2026-11-14/15).
// OPENALEX_API_KEY lifts that; Crossref (free, no key, the DOI registry itself) answers when OpenAlex does not.
const openalex = async (query: string) => {
  const params = new URLSearchParams({ search: query, 'per-page': '6', select: 'display_name,publication_year,doi,authorships,primary_location,cited_by_count', mailto: 'contact@chatslide.ai', ...(process.env.OPENALEX_API_KEY ? { api_key: process.env.OPENALEX_API_KEY } : {}) })
  const r = await fetch(`https://api.openalex.org/works?${params}`)
  if (!r.ok) throw new Error(`openalex ${r.status}`)
  const works = ((await r.json()).results || []) as any[]
  return works.map(w => ({
    title: w.display_name, year: w.publication_year, cited: w.cited_by_count,
    authors: (w.authorships || []).slice(0, 3).map((a: any) => a.author?.display_name).filter(Boolean).join(', ') + ((w.authorships || []).length > 3 ? ' et al.' : ''),
    venue: w.primary_location?.source?.display_name || '', url: w.doi || w.primary_location?.landing_page_url || '',
  })).filter(p => p.title)
}
export const crossref = async (query: string) => {
  const params = new URLSearchParams({ 'query.bibliographic': query, rows: '6', filter: 'type:journal-article', select: 'title,author,issued,container-title,DOI,is-referenced-by-count', mailto: 'contact@chatslide.ai' })
  const r = await fetch(`https://api.crossref.org/works?${params}`)
  if (!r.ok) throw new Error(`crossref ${r.status}`)
  const items = ((await r.json()).message?.items || []) as any[]
  return items.map(w => ({
    title: String(w.title?.[0] || '').replace(/<[^>]+>/g, ''), year: w.issued?.['date-parts']?.[0]?.[0] || '', cited: w['is-referenced-by-count'] || 0,
    authors: (w.author || []).slice(0, 3).map((a: any) => [a.given, a.family].filter(Boolean).join(' ') || a.name).filter(Boolean).join(', ') + ((w.author || []).length > 3 ? ' et al.' : ''),
    venue: String(w['container-title']?.[0] || ''), url: w.DOI ? `https://doi.org/${w.DOI}` : '',
  })).filter(p => p.title)
}
const scholarOnce = async (query: string) => {
  const found = await openalex(query).catch((e) => {
    console.error('[research] OpenAlex failed, using Crossref:', String(e?.message || e).slice(0, 160)) // a quota or outage must show in the logs
    return null
  })
  return { papers: found?.length ? found : await crossref(query).catch(() => found ?? []) } // an empty answer too: no match, or a quota
}

// OpenAlex matches every word: the router's 20-word queries ("systemic sclerosis contemporary guidelines epidemiology
// ... dazodalibep") and Russian ones found 0 papers where "systemic sclerosis epidemiology pathogenesis" finds 44,317
// (2026-11-12). Nothing found: up to three short English searches, the first that finds papers wins.
const scholar = async (query: string) => {
  const first = await scholarOnce(query)
  if (first.papers.length) return first
  const alts = await chat([{ role: 'user', content: `An academic paper search found nothing for: ${query}\nGive 3 different short English searches (3 to 5 words each, the core topic first) that would find the key papers. JSON {"queries": ["...", "...", "..."]}` }], { jsonMode: true, maxTokens: 2000 })
    .then(r => (JSON.parse(r).queries || []) as string[]).catch(() => [] as string[])
  for (const q of alts.slice(0, 3)) {
    const got = await scholarOnce(String(q)).catch(() => ({ papers: [] as any[] }))
    if (got.papers.length) return got
  }
  return first
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
  return { title: String(d.name || d.title || url).trim(), text: String(d.content || '').slice(0, 60000), kind: yt ? 'video transcript' : 'web page' }
}

// The pictures on a page. context.dev (brand.dev's web API, the key the chatslide.ai tools already use; 1 credit a
// page) renders the page, so script-loaded pictures are found: 91 / 85 / 91 on three brand sites where the Supadata
// markdown had 31 / 6 (2026-11-12). Supadata stays as the fallback.
const pagePictures = async (page: string): Promise<{ title: string, pictures: { src: string, alt: string }[] }> => {
  let host = ''
  try { const u = new URL(page); host = /^https?:$/.test(u.protocol) ? u.hostname.toLowerCase().replace(/\.$/, '') : '' }
  catch { /* not a URL: refused below */ }
  if (!host || /^[\d.]+$|^\[|:|^localhost$|\.(local|internal|localhost|lan)$/.test(host) || !host.includes('.')) throw new Error('only public web links can be read')
  const key = process.env.BRANDDEV_API_KEY
  if (key) {
    try {
      const r = await fetch(`https://api.context.dev/v1/web/scrape/images?${new URLSearchParams({ url: page })}`, { headers: { Authorization: `Bearer ${key}` }, signal: AbortSignal.timeout(45_000) }) // a slow render falls back
      const d = r.ok ? await r.json() : null
      const pictures = ((d?.images || []) as any[]).filter(x => x.type === 'url' && /^https:\/\//.test(String(x.src)))
        .map(x => ({ src: String(x.src), alt: String(x.alt || '').replace(/<[^>]+>/g, ' ').trim().slice(0, 120) }))
      if (pictures.length) return { title: host.replace(/^www\./, ''), pictures }
    }
    catch { /* the fallback below */ }
  }
  const read = await readUrl(page)
  return { title: read.title, pictures: [...read.text.matchAll(/!\[([^\]]*)\]\((https:\/\/[^)\s]+)\)/g)].map(m => ({ src: m[2], alt: m[1].slice(0, 120) })) }
}

// Real pictures from a brand's, company's or product's own website (a paid feature, gated in ChatSlide's proxy):
// the page is read by Supadata like any pasted link, and its pictures are picked by what the slide needs. Stock
// photos cannot show one brand's products ("more photos of the brand's products" got generic clothes, 2026-11-12).
// ponytail: the slide links to the site's own picture file; a site that blocks other origins will not export. Copy the
// files through our storage if that shows up.
const siteImages = async (what: string, url: string, brand: string, count: number, avoid: string[] = []) => {
  // one picture = its file, or what an image proxy points at (/_next/image?url=…), whatever size is asked for
  const fileOf = (src: string) => { try { const u = new URL(src); return u.searchParams.get('url') || u.origin + u.pathname } catch { return src } }
  const had = new Set(avoid.map(fileOf)) // already on the deck: "more" brings other pictures
  let page = /^https?:\/\//i.test(url) ? url : ''
  if (!page) {
    if (!brand.trim()) throw new Error('no website or brand named')
    const found = await web(`What is the official website (home page URL) of the brand or company "${brand}"? Give the URL.`)
    const official = await chat([{ role: 'user', content: `Brand: ${brand}\nSearch answer: ${found.answer}\nSources: ${found.sources.map(x => x.url).join(' ')}\nWhich URL is the brand's OWN official website (not a marketplace, wiki or news page)? JSON {"url": "<https://... or empty>"}` }], { model: 'gpt-6-luna', jsonMode: true, maxTokens: 2000 }).then(r => JSON.parse(r) || {}).catch(() => ({}))
    page = /^https?:\/\//i.test(String(official.url || '')) ? String(official.url) : ''
    if (!page) throw new Error(`no official website found for ${brand}`)
  }
  const read = await pagePictures(page)
  const seen = new Set<string>()
  const all: { src: string, alt: string }[] = []
  for (const x of read.pictures) {
    const file = fileOf(x.src)
    if (seen.has(file) || had.has(file) || /\.(svg|gif|ico)$/i.test(file.split('?')[0])) continue // one size of each picture; no vector icons
    seen.add(file)
    all.push(x)
  }
  if (!all.length) throw new Error(`no ${had.size ? 'more ' : ''}pictures found on ${page}`) // an error, so the call is not charged
  let picks = all.slice(0, count)
  try {
    const list = all.slice(0, 60).map((x, k) => `${k}: ${x.alt || '(no description)'} | ${x.src.split('?')[0].split('/').pop()}`).join('\n')
    const out = JSON.parse(await chat([{ role: 'user', content: `A slide needs ${count} picture(s) from the website "${read.title}": ${what}\nPictures on the page:\n${list}\nPick the ${count} that fit best (real photos of the products or the thing asked for; never logos, icons, payment badges or banners with text). JSON {"picks": [numbers]}` }], { model: 'gpt-6-luna', jsonMode: true, maxTokens: 2000 }))
    const chosen = (out.picks || []).map((k: number) => all[k]).filter(Boolean).slice(0, count)
    if (chosen.length) picks = chosen
  }
  catch { /* the page's first pictures */ }
  return { images: picks, page, title: String(read.title).trim() }
}

async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  try {
    const { kind = '', query = '', nearest = false, url = '', brand = '', count = 3, avoid = [], want = '', orientation = 'landscape' } = await readBody(req)
    if (kind === 'site_images') { res.status(200).json(await siteImages(String(query).slice(0, 300), String(url).trim().slice(0, 2000), String(brand).slice(0, 120), Math.min(Math.max(Number(count) || 3, 1), 4), (Array.isArray(avoid) ? avoid : []).slice(0, 200).map(String))); return }
    if (!String(query).trim()) { res.status(400).json({ error: 'query required' }); return }
    if (kind === 'web') res.status(200).json(await web(String(query).slice(0, 500)))
    else if (kind === 'scholar') res.status(200).json(await scholar(String(query).slice(0, 300)))
    else if (kind === 'url') res.status(200).json(await readUrl(String(query).trim().slice(0, 2000)))
    else if (kind === 'wikimedia') res.status(200).json(await wikimedia(String(query).slice(0, 200), nearest === true, Array.isArray(avoid) ? avoid.map(String) : [], String(want).slice(-1500)))
    else if (kind === 'stock') res.status(200).json(await stock(String(query).slice(0, 300), String(want).slice(-1500), orientation === 'portrait' ? 'portrait' : 'landscape', (Array.isArray(avoid) ? avoid : []).slice(0, 200).map(String)))
    else res.status(400).json({ error: 'kind must be web, scholar, url, wikimedia or stock' })
  }
  catch (e: any) {
    console.error('[research] failed:', String(e?.message || e).slice(0, 300)) // the 500 alone leaves the request log blank
    res.status(500).json({ error: String(e?.message || e) })
  }
}

export default metered(handler)
