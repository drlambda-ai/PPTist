// the browser bundle, untyped: jszip's own .d.ts references Node's types, which retypes every setTimeout in the app
// @ts-expect-error no declaration for the bundle path
import JSZipBundle from 'jszip/dist/jszip.min.js'
type Zip = { files: Record<string, unknown>, file: { (name: string): { async(t: 'string'): Promise<string> } | null, (name: string, data: string): void }, generateAsync(o: { type: 'arraybuffer' }): Promise<ArrayBuffer> }
const JSZip = JSZipBundle as { loadAsync(buf: ArrayBuffer): Promise<Zip> }

// Valid PowerPoint that pptxtojson does not read, rewritten to the form it does (decks built by pptxgenjs, which is how
// ChatSlide exports, write both):
// 1. an absolute relationship target ("/ppt/charts/chart3.xml"): the part is not found and a chart silently vanishes;
// 2. chart categories as a one-level multiLvlStrRef: the labels are lost (bars read 0, 1, 2 instead of names).
const relative = (fromDir: string, abs: string) => {
  const from = fromDir.split('/').filter(Boolean), to = abs.split('/').filter(Boolean)
  let i = 0
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++
  return [...from.slice(i).map(() => '..'), ...to.slice(i)].join('/')
}

export const normalizePptx = async (buf: ArrayBuffer): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buf)
  let changed = false
  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name)
    if (!file) continue
    if (name.endsWith('.rels')) {
      const xml = await file.async('string')
      const dir = name.replace(/_rels\/[^/]+$/, '') // ppt/slides/_rels/slide9.xml.rels -> ppt/slides/
      const fixed = xml.replace(/Target="\/([^"]+)"/g, (_, abs) => `Target="${relative(dir, abs)}"`)
      if (fixed !== xml) {
        zip.file(name, fixed)
        changed = true
      }
    }
    else if (/^ppt\/charts\/chart[^/]*\.xml$/.test(name)) {
      const xml = await file.async('string')
      const fixed = xml.replace(
        /<c:multiLvlStrRef>\s*(<c:f>[^<]*<\/c:f>)?\s*<c:multiLvlStrCache>\s*(<c:ptCount[^>]*\/>)\s*<c:lvl>([\s\S]*?)<\/c:lvl>\s*<\/c:multiLvlStrCache>\s*<\/c:multiLvlStrRef>/g,
        (_, f = '', count, pts) => `<c:strRef>${f}<c:strCache>${count}${pts}</c:strCache></c:strRef>`,
      )
      if (fixed !== xml) {
        zip.file(name, fixed)
        changed = true
      }
    }
  }
  return changed ? zip.generateAsync({ type: 'arraybuffer' }) : buf
}
