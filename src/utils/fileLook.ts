import tinycolor from 'tinycolor2'

// The look of the user's own presentation file, as pptxtojson read it: its background, its own colours ranked by how
// much of the slides they cover, and its logo. A surgeon asked the chat to "use the colour and logo of my previous
// presentation" and was told to attach the file the project already had (2026-11-13).
export interface FileLook {
  background: string
  fontColor: string
  colors: string[] // the file's own colours, most used first, then tints (up to 10: a page with more colours than a theme's 6 gets no stray hue); its theme's when it has none of its own
  logo?: { src: string, left: number, top: number, width: number, height: number, ratio: number } // box in fractions of the slide; ratio = its width / height
}

const NEUTRAL = (c: tinycolor.Instance) => {
  const { s, l } = c.toHsl()
  return s < 0.15 || l > 0.93 || l < 0.1
}

export const fileLook = (json: any): FileLook => {
  const W = json.size.width, H = json.size.height
  // per colour: how much it covers, and on how many pages (a brand colour is on most pages; a red word or six orange
  // boxes are not the brand)
  const weight: Record<string, number> = {}, pagesOf: Record<string, Set<number>> = {}
  const add = (color: unknown, w: number, page: number) => {
    const c = tinycolor(String(color || ''))
    if (!color || !c.isValid() || c.getAlpha() < 0.5 || NEUTRAL(c)) return
    const k = c.toHexString()
    weight[k] = (weight[k] || 0) + w
    ;(pagesOf[k] = pagesOf[k] || new Set()).add(page)
  }
  const backgrounds: Record<string, number> = {}
  // a logo is a small picture inside a corner; the same one on many pages (a master's) beats one on the cover only
  const spots: Record<string, { pages: number, first: number, el: any }> = {}
  const inCorner = (e: any) => (e.left + e.width <= W * 0.35 || e.left >= W * 0.65) && (e.top + e.height <= H * 0.3 || e.top >= H * 0.7)

  const walk = (els: any[], page: number, seen: Set<string>) => {
    for (const e of els || []) {
      if (e.type === 'group') {
        walk(e.elements, page, seen)
        continue
      }
      const area = Math.min((e.width * e.height) / (W * H), 0.3)
      if (e.fill?.type === 'color') add(e.fill.value, area, page)
      if (e.borderWidth) add(e.borderColor, 0.01, page)
      for (const m of String(e.content || '').matchAll(/(?<![-\w])color:\s*([^;"]+)/g)) add(m[1], 0.01, page)
      const src = e.base64 || e.src
      if (e.type === 'image' && src && e.width * e.height <= W * H * 0.06 && inCorner(e)) {
        const k = `${Math.round(e.left / 10)},${Math.round(e.top / 10)},${Math.round(e.width / 10)}`
        if (seen.has(k)) continue
        seen.add(k)
        spots[k] = spots[k] || { pages: 0, first: page, el: e }
        spots[k].pages++
      }
    }
  }
  json.slides.forEach((sl: any, i: number) => {
    const bg = sl.fill?.type === 'color' ? tinycolor(sl.fill.value || '#fff').toHexString() : ''
    if (bg) backgrounds[bg] = (backgrounds[bg] || 0) + 1
    const seen = new Set<string>()
    walk(sl.layoutElements, i, seen)
    walk(sl.elements, i, seen)
  })

  const background = Object.keys(backgrounds).sort((a, b) => backgrounds[b] - backgrounds[a])[0] || '#ffffff'
  // its own colours: the main one and any other on a third of the pages, then lighter tints of the main one (text stays
  // readable on them); a file with none of its own gives its theme's
  const n = json.slides.length
  const ranked = Object.keys(weight).sort((a, b) => pagesOf[b].size - pagesOf[a].size || weight[b] - weight[a])
  const own = ranked.filter((c, i) => !i || pagesOf[c].size >= n / 3).slice(0, 3)
  const tints = own.length ? [15, 35, 55, 70, 82, 25, 45, 62, 76].map(k => tinycolor.mix(own[0], '#ffffff', k).toHexString()) : []
  const colors = own.length ? [...own, ...tints].slice(0, 10) : (json.themeColors || []).slice(0, 6)
  const logos = Object.values(spots).sort((a, b) => b.pages - a.pages || a.first - b.first)
  const logo = logos.find(l => l.pages > 1 || l.first === 0)?.el
  return {
    background,
    fontColor: tinycolor.mostReadable(background, ['#000000', '#ffffff']).toHexString(),
    colors,
    logo: logo && { src: logo.base64 || logo.src, left: logo.left / W, top: logo.top / H, width: logo.width / W, height: logo.height / H, ratio: logo.width / logo.height },
  }
}
