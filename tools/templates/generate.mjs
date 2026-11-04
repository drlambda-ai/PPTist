// Generates the template library: 10 palette families x 10 decoration motifs = 100 styles, each a full PPTist AI
// template (the page contract useAIPPT fills: cover, contents 2-6 items, transition, content 1-4 items, end).
// Deterministic, no model calls. Every template is validated (contrast, bounds, overlaps, contract) before it is
// written; a failing one is reported and skipped.
//
//   node tools/templates/generate.mjs            -> public/mocks/gen/<id>.json + public/mocks/gen/index.json
//   node tools/templates/generate.mjs --check    -> validate only
import fs from 'node:fs'
import path from 'node:path'

const W = 1000, H = 562.5
const OUT = path.join(path.dirname(new URL(import.meta.url).pathname), '../../public/mocks/gen')

// ---- styles -------------------------------------------------------------------------------------------------------
const PALETTES = [ // bg, surface (cards), primary, accent, text, muted
  { name: 'Violet', bg: '#ffffff', surface: '#f5f3ff', primary: '#7c3aed', accent: '#f59e0b', text: '#1f1b2e', muted: '#5b5670' },
  { name: 'Ocean', bg: '#f8fafc', surface: '#e0f2fe', primary: '#0369a1', accent: '#14b8a6', text: '#0f172a', muted: '#475569' },
  { name: 'Forest', bg: '#fbfaf5', surface: '#ecfccb', primary: '#3f6212', accent: '#ca8a04', text: '#1c1917', muted: '#57534e' },
  { name: 'Coral', bg: '#fffaf7', surface: '#ffe4dc', primary: '#c2410c', accent: '#0e7490', text: '#2b1d16', muted: '#6b5146' },
  { name: 'Slate', bg: '#ffffff', surface: '#f1f5f9', primary: '#1e293b', accent: '#e11d48', text: '#0f172a', muted: '#475569' },
  { name: 'Rose', bg: '#fff7f9', surface: '#ffe4ec', primary: '#be185d', accent: '#7c3aed', text: '#2a1520', muted: '#6b4a57' },
  { name: 'Sand', bg: '#faf7f0', surface: '#efe6d2', primary: '#92400e', accent: '#065f46', text: '#292114', muted: '#5f533c' },
  { name: 'Midnight', bg: '#0f172a', surface: '#1e293b', primary: '#38bdf8', accent: '#f472b6', text: '#f8fafc', muted: '#cbd5e1' },
  { name: 'Graphite', bg: '#18181b', surface: '#27272a', primary: '#a3e635', accent: '#fb923c', text: '#fafafa', muted: '#d4d4d8' },
  { name: 'Emerald Night', bg: '#052e2b', surface: '#0b3f3a', primary: '#34d399', accent: '#fbbf24', text: '#ecfdf5', muted: '#a7f3d0' },
]
const MOTIFS = ['sidebar', 'topband', 'corner', 'circles', 'frame', 'diagonal', 'underline', 'dots', 'split', 'minimal']
const FONTS = [ // heading, body: only fonts PPTist loads (src/configs/font.ts)
  ['Montserrat', 'OpenSans'], ['Merriweather', 'SourceSansPro'], ['Inter', 'Inter'], ['SourceSerif4', 'Inter'],
  ['Literata', 'Roboto'], ['DeYiHei', 'MiSans'], ['AlibabaPuHuiTi', 'AlibabaPuHuiTi'], ['SourceHanSerif', 'SourceHanSans'],
  ['YousheTitleBlack', 'MiSans'], ['Roboto', 'Roboto'],
]
const MOTIF_NAMES = { sidebar: 'Band', topband: 'Header', corner: 'Arc', circles: 'Orbit', frame: 'Frame', diagonal: 'Wedge', underline: 'Accent', dots: 'Grid', split: 'Split', minimal: 'Line' }

export const styles = () => PALETTES.flatMap((p, i) => MOTIFS.map((m, j) => ({
  id: `gen_${String(i * 10 + j + 1).padStart(3, '0')}`,
  name: `${p.name} ${MOTIF_NAMES[m]}`,
  palette: p, motif: m, fonts: FONTS[(i + j) % FONTS.length],
})))

// ---- elements -----------------------------------------------------------------------------------------------------
let seq = 0
const id = () => `g${(seq++).toString(36).padStart(6, '0')}`
const RECT = 'M 0 0 L 200 0 L 200 200 L 0 200 Z'
const CIRCLE = 'M 100 0 A 50 50 0 1 1 100 200 A 50 50 0 1 1 100 0 Z'
const TRI = 'M 0 200 L 200 0 L 200 200 Z'
const shape = (left, top, width, height, fill, pathD = RECT, opacity) =>
  ({ type: 'shape', id: id(), left, top, width, height, viewBox: [200, 200], path: pathD, fill, fixedRatio: false, rotate: 0, lock: true, ...(opacity ? { opacity } : {}) })
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;')
const text = (textType, left, top, width, height, value, { size, color, font, bold, align = 'left' }) => ({
  type: 'text', id: id(), left, top, width, height, rotate: 0, vertical: false, lineHeight: 1.2,
  defaultFontName: font, defaultColor: color, ...(textType ? { textType } : {}),
  content: `<p style="text-align: ${align};"><span style="font-size: ${size}px; color: ${color}; font-family: ${font};">${bold ? `<strong>${esc(value)}</strong>` : esc(value)}</span></p>`,
})

// decoration per motif: behind the text (first in the element list); `kind` lets covers be bolder than content pages
const decor = (st, kind) => {
  const { primary, accent, surface } = st.palette, big = kind === 'cover' || kind === 'end' || kind === 'transition'
  switch (st.motif) {
    case 'sidebar': return [shape(0, 0, big ? 300 : 24, H, primary)]
    case 'topband': return big ? [shape(0, 0, W, 190, primary)] : [shape(0, 0, W, 16, primary)]
    case 'corner': return [shape(W - (big ? 330 : 140), -(big ? 165 : 70), big ? 330 : 140, big ? 330 : 140, primary, CIRCLE), shape(-60, H - 60, 120, 120, accent, CIRCLE, 0.85)]
    case 'circles': return big
      ? [shape(640, 110, 300, 300, surface, CIRCLE), shape(760, 60, 140, 140, primary, CIRCLE, 0.9), shape(700, 360, 90, 90, accent, CIRCLE, 0.9)]
      : [shape(W - 90, 20, 50, 50, primary, CIRCLE, 0.9), shape(W - 55, 55, 26, 26, accent, CIRCLE)]
    case 'frame': return [shape(20, 20, W - 40, 4, primary), shape(20, H - 24, W - 40, 4, primary), shape(20, 20, 4, H - 40, primary), shape(W - 24, 20, 4, H - 40, primary)]
    case 'diagonal': return big ? [shape(W - 420, 0, 420, H, primary, TRI)] : [shape(W - 160, H - 90, 160, 90, primary, TRI)]
    case 'underline': return [] // drawn under the title by the page builders
    case 'dots': return Array.from({ length: big ? 25 : 9 }, (_, k) => shape(W - (big ? 200 : 110) + (k % (big ? 5 : 3)) * 30, 30 + Math.floor(k / (big ? 5 : 3)) * 30, 8, 8, k % 4 ? primary : accent, CIRCLE))
    case 'split': return big ? [shape(W / 2, 0, W / 2, H, surface), shape(W / 2, 0, 10, H, primary)] : [shape(0, 0, 10, H, primary)]
    case 'minimal': return [shape(60, H - 50, W - 120, 2, primary)]
  }
  return []
}

// where text may go on this style's page (decorations kept clear)
const area = (st, kind) => {
  const big = kind !== 'content' && kind !== 'contents'
  if (st.motif === 'sidebar') return big ? { x: 340, y: 60, w: W - 400, h: H - 120 } : { x: 70, y: 40, w: W - 120, h: H - 80 }
  if (st.motif === 'topband') return big ? { x: 60, y: 220, w: W - 120, h: H - 260 } : { x: 60, y: 46, w: W - 120, h: H - 86 }
  if (st.motif === 'split' && big) return { x: 50, y: 60, w: W / 2 - 70, h: H - 120 }
  if ((st.motif === 'diagonal' || st.motif === 'circles' || st.motif === 'corner') && big) return { x: 60, y: 60, w: 520, h: H - 120 }
  if (st.motif === 'dots' || st.motif === 'circles' || st.motif === 'corner') return { x: 60, y: 46, w: W - 260, h: H - 86 }
  return { x: 60, y: 50, w: W - 120, h: H - 100 }
}

// text colour on the cover/section area: dark themes and the top band need light text over primary-coloured blocks
const onBg = st => st.palette.text

const page = (st, type, elements) => ({ id: id(), type, background: { type: 'solid', color: st.palette.bg }, elements })

const titleBlock = (st, a, value, size = 30) => {
  const [hf] = st.fonts, { text: tc, primary } = st.palette
  const out = [text('title', a.x, a.y, a.w, 56, value, { size, color: tc, font: hf, bold: true })]
  if (st.motif === 'underline') out.push(shape(a.x, a.y + 60, 90, 6, primary))
  return out
}

// ---- page builders ------------------------------------------------------------------------------------------------
const cover = (st, v) => {
  const a = area(st, 'cover'), [hf, bf] = st.fonts, { muted, primary } = st.palette
  const top = v ? a.y + a.h / 2 - 80 : a.y + 40
  return page(st, 'cover', [...decor(st, 'cover'),
    text('title', a.x, top, a.w, 110, 'Presentation title', { size: a.w < 500 ? 38 : 44, color: onBg(st), font: hf, bold: true }),
    ...(st.motif === 'underline' ? [shape(a.x, top + 116, 120, 8, primary)] : []),
    text('content', a.x, top + 136, a.w, 60, 'A one-line subtitle for this presentation', { size: 20, color: muted, font: bf }),
  ])
}

const contents = (st, n) => {
  const a = area(st, 'contents'), [hf, bf] = st.fonts, { text: tc, primary, bg, surface } = st.palette
  const rowH = Math.min(64, (a.h - 90) / n), y0 = a.y + 90
  const els = [...decor(st, 'contents'), text('', a.x, a.y, a.w, 56, 'Contents', { size: 30, color: tc, font: hf, bold: true })]
  for (let k = 0; k < n; k++) {
    const y = y0 + k * rowH
    els.push(shape(a.x, y + 6, 44, 44, primary, CIRCLE),
      text('', a.x, y + 13, 44, 30, String(k + 1).padStart(2, '0'), { size: 16, color: bg === '#ffffff' || surface ? bg : '#ffffff', font: hf, bold: true, align: 'center' }),
      text('item', a.x + 64, y + 10, a.w - 64, 36, `Section ${k + 1}`, { size: 20, color: tc, font: bf }))
  }
  return page(st, 'contents', els)
}

const transition = st => {
  const a = area(st, 'transition'), [hf, bf] = st.fonts, { muted, primary } = st.palette
  return page(st, 'transition', [...decor(st, 'transition'),
    text('partNumber', a.x, a.y + 40, 200, 80, '01', { size: 64, color: primary, font: hf, bold: true }),
    text('title', a.x, a.y + 140, a.w, 70, 'Section title', { size: 38, color: onBg(st), font: hf, bold: true }),
    text('content', a.x, a.y + 220, a.w, 60, 'What this section covers', { size: 18, color: muted, font: bf }),
  ])
}

const content = (st, n) => {
  const a = area(st, 'content'), [hf, bf] = st.fonts, { text: tc, muted, surface, primary } = st.palette
  const els = [...decor(st, 'content'), ...titleBlock(st, a, 'Page title')]
  if (n === 1) {
    els.push(text('content', a.x, a.y + 100, a.w, a.h - 120, 'One paragraph of body text for this page.', { size: 20, color: tc, font: bf }))
    return page(st, 'content', els)
  }
  const gap = 24, cw = (a.w - gap * (n - 1)) / n, top = a.y + 100, ch = a.h - 110
  for (let k = 0; k < n; k++) {
    const x = a.x + k * (cw + gap)
    els.push(shape(x, top, cw, ch, surface), shape(x, top, cw, 6, primary),
      text('itemTitle', x + 16, top + 22, cw - 32, 40, `Point ${k + 1}`, { size: 20, color: tc, font: hf, bold: true }),
      text('item', x + 16, top + 70, cw - 32, ch - 90, 'A short explanation of this point.', { size: 15, color: muted, font: bf }))
  }
  return page(st, 'content', els)
}

const end = st => {
  const a = area(st, 'end'), [hf, bf] = st.fonts, { muted } = st.palette
  return page(st, 'end', [...decor(st, 'end'),
    text('', a.x, a.y + a.h / 2 - 70, a.w, 80, 'Thank you', { size: 48, color: onBg(st), font: hf, bold: true }),
    text('', a.x, a.y + a.h / 2 + 20, a.w, 40, 'Questions and discussion', { size: 20, color: muted, font: bf }),
  ])
}

export const build = st => ({
  title: st.name, width: W, height: H,
  theme: { themeColors: [st.palette.primary, st.palette.accent], fontColor: st.palette.text, fontName: st.fonts[1], backgroundColor: st.palette.bg,
    shadow: { h: 3, v: 3, blur: 2, color: '#808080' }, outline: { width: 2, color: '#525252', style: 'solid' } },
  slides: [cover(st, 0), cover(st, 1), ...[2, 3, 4, 5, 6].map(n => contents(st, n)), transition(st), ...[1, 2, 3, 4].map(n => content(st, n)), end(st)],
})

// ---- validation ---------------------------------------------------------------------------------------------------
const lum = hex => {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255).map(c => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4))
  return 0.2126 * r + 0.7152 * g + 0.0722 * b
}
export const contrast = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((m, n) => n - m); return (x + 0.05) / (y + 0.05) }
const inside = (e, s) => e.left + e.width / 2 >= s.left && e.left + e.width / 2 <= s.left + s.width && e.top + e.height / 2 >= s.top && e.top + e.height / 2 <= s.top + s.height
const colorOf = e => (e.content.match(/color: (#[0-9a-f]{6})/i) || [])[1] || e.defaultColor
const sizeOf = e => Number((e.content.match(/font-size: (\d+)px/) || [])[1] || 16)

export const validate = (st, tpl) => {
  const errs = []
  const items = (sl, t) => sl.elements.filter(e => e.textType === t).length
  const has = (type, test) => tpl.slides.some(sl => sl.type === type && test(sl))
  if (!has('cover', sl => items(sl, 'title') === 1 && items(sl, 'content') === 1)) errs.push('no cover with title + subtitle')
  for (const n of [2, 3, 4, 5, 6]) if (!has('contents', sl => items(sl, 'item') === n)) errs.push(`no contents page with ${n} items`)
  if (!has('transition', sl => items(sl, 'title') === 1 && items(sl, 'partNumber') === 1)) errs.push('no transition page')
  if (!has('content', sl => items(sl, 'title') === 1 && items(sl, 'content') === 1 && !items(sl, 'item'))) errs.push('no one-paragraph content page')
  for (const n of [2, 3, 4]) if (!has('content', sl => items(sl, 'item') === n && items(sl, 'itemTitle') === n)) errs.push(`no content page with ${n} items`)
  if (!tpl.slides.some(sl => sl.type === 'end')) errs.push('no end page')
  tpl.slides.forEach((sl, si) => {
    const texts = sl.elements.filter(e => e.type === 'text'), shapes = sl.elements.filter(e => e.type === 'shape')
    for (const e of sl.elements) {
      if (e.left < -1 || e.top < -1 || e.left + e.width > W + 1 || e.top + e.height > H + 1) {
        if (e.type === 'text') errs.push(`page ${si + 1}: text outside the slide`)
      }
    }
    for (const e of texts) {
      // the background behind a text: the top-most shape under its centre, else the slide
      const under = shapes.filter(s => !s.opacity && inside(e, s)).pop()
      const bg = under ? under.fill : sl.background.color
      const need = sizeOf(e) >= 24 ? 3 : 4.5
      const c = contrast(colorOf(e), bg)
      if (c < need) errs.push(`page ${si + 1}: "${e.content.replace(/<[^>]+>/g, '').slice(0, 30)}" contrast ${c.toFixed(1)} < ${need}`)
    }
    for (let i = 0; i < texts.length; i++) for (let j = i + 1; j < texts.length; j++) {
      const a = texts[i], b = texts[j]
      const ov = Math.max(0, Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)) * Math.max(0, Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top))
      if (ov > 0.15 * Math.min(a.width * a.height, b.width * b.height)) errs.push(`page ${si + 1}: text boxes overlap`)
    }
  })
  return [...new Set(errs)]
}

// ---- main ---------------------------------------------------------------------------------------------------------
if (import.meta.url === `file://${process.argv[1]}`) {
  const check = process.argv.includes('--check')
  const ok = [], bad = []
  for (const st of styles()) {
    seq = 0
    const tpl = build(st), errs = validate(st, tpl)
    if (errs.length) bad.push({ id: st.id, name: st.name, errs })
    else ok.push({ st, tpl })
  }
  if (!check) {
    fs.mkdirSync(OUT, { recursive: true })
    for (const { st, tpl } of ok) fs.writeFileSync(path.join(OUT, `${st.id}.json`), JSON.stringify(tpl))
    fs.writeFileSync(path.join(OUT, 'index.json'), JSON.stringify(ok.map(({ st }) => ({ id: st.id, name: st.name, motif: st.motif, palette: st.palette.name, fonts: st.fonts }))))
  }
  console.log(`${ok.length} templates pass, ${bad.length} fail`)
  for (const b of bad) console.log(`- ${b.id} ${b.name}: ${b.errs.slice(0, 4).join('; ')}${b.errs.length > 4 ? ` (+${b.errs.length - 4})` : ''}`)
  process.exit(bad.length ? 1 : 0)
}
