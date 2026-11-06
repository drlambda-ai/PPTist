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
  // 3. a media name with & < > " ' (images named after a signed URL: "...&sig=..."): pptxtojson 2.2 HTML-escapes the
  // path before looking it up, finds nothing, and the whole import fails. Renamed, with the links to it.
  const renamed: [string, string][] = []
  for (const name of Object.keys(zip.files)) {
    if (!name.startsWith('ppt/media/') || !/[&<>"']/.test(name)) continue
    const safe = name.replace(/[&<>"']/g, '_').replace(/([^/]+)$/, `renamed-${renamed.length}-$1`) // never collides with a real name
    const z = zip as any
    z.file(safe, await z.file(name).async('arraybuffer'))
    z.remove(name)
    renamed.push([name.split('/').pop()!, safe.split('/').pop()!])
    changed = true
  }
  const xmlEscape = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name)
    if (!file) continue
    if (name.endsWith('.rels')) {
      const xml = await file.async('string')
      const dir = name.replace(/_rels\/[^/]+$/, '') // ppt/slides/_rels/slide9.xml.rels -> ppt/slides/
      let fixed = xml.replace(/Target="\/([^"]+)"/g, (_, abs) => `Target="${relative(dir, abs)}"`)
      // only internal links to the renamed file, whether written escaped or, by pptxgenjs, raw; never an external link
      fixed = fixed.replace(/<Relationship\b[^>]*>/g, (rel) => {
        if (/TargetMode="External"/.test(rel)) return rel
        for (const [from, to] of renamed) {
          for (const form of [xmlEscape(from), from]) {
            if (rel.includes(`/${form}"`)) return rel.replace(`/${form}"`, `/${to}"`)
          }
        }
        return rel
      })
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
