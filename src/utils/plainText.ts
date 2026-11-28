// Text and slide-HTML helpers written with string methods. Product code uses no regular expressions (standing rule);
// each helper gives the same result as the pattern it replaced, checked on the text boxes of 300 real saved decks.

const isWordChar = (ch: string) => !!ch && ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9') || ch === '_')
const isDigit = (ch: string) => !!ch && ch >= '0' && ch <= '9'
const isHexDigit = (ch: string) => !!ch && '0123456789abcdefABCDEF'.includes(ch)
const isSpace = (ch: string) => !!ch && !ch.trim()

/** A CSS property's value in inline styles: the first `name:` (not part of a longer name when `whole`), skipping spaces, up to ; or " */
const propAt = (html: string, name: string, whole: boolean, from = 0) => {
  for (let at = html.indexOf(`${name}:`, from); at >= 0; at = html.indexOf(`${name}:`, at + 1)) {
    if (whole && at > 0 && (html[at - 1] === '-' || isWordChar(html[at - 1]))) continue
    let k = at + name.length + 1
    while (k < html.length && isSpace(html[k])) k++
    return { at, valueAt: k }
  }
  return null
}
const valueEnd = (html: string, k: number, quot = false) => {
  while (k < html.length && html[k] !== ';' && html[k] !== '"') {
    if (quot && html.startsWith('&quot;', k)) k += 6
    else k++
  }
  return k
}

/** the first inline `color:` value (not background-color), as the old /(?<![-\w])color:\s*([^;"]+)/ */
export const cssColorOf = (html: string) => {
  const p = propAt(html, 'color', true)
  if (!p) return ''
  const end = valueEnd(html, p.valueAt)
  return end > p.valueAt ? html.slice(p.valueAt, end) : ''
}
/** every inline `color:` set to `to`, as the old .replace(/(?<![-\w])color:\s*[^;"]+/g, `color: ${to}`) */
export const setCssColor = (html: string, to: string) => {
  let out = '', last = 0
  for (let p = propAt(html, 'color', true); p; p = propAt(html, 'color', true, p.at + 1)) {
    const end = valueEnd(html, p.valueAt)
    if (end === p.valueAt) continue
    out += `${html.slice(last, p.at)}color: ${to}`
    last = end
  }
  return out + html.slice(last)
}
/** the first colour after `color:` that is #rrggbb or rgb()/rgba(), as the old /color:\s*(#[0-9a-fA-F]{6}|rgba?\([^)]*\))/ */
export const titleColorOf = (html: string) => {
  for (let at = html.indexOf('color:'); at >= 0; at = html.indexOf('color:', at + 1)) {
    let k = at + 6
    while (k < html.length && isSpace(html[k])) k++
    if (html[k] === '#' && [...html.slice(k + 1, k + 7)].length === 6 && [...html.slice(k + 1, k + 7)].every(isHexDigit)) return html.slice(k, k + 7)
    for (const f of ['rgba(', 'rgb(']) {
      if (html.startsWith(f, k)) {
        const close = html.indexOf(')', k)
        if (close > 0) return html.slice(k, close + 1)
      }
    }
  }
  return ''
}
/** the first inline font-family value with &quot; kept, as the old /font-family:\s*((?:&quot;|[^;"])+)/ */
export const fontFamilyOf = (html: string) => {
  const p = propAt(html, 'font-family', false)
  if (!p) return ''
  const end = valueEnd(html, p.valueAt, true)
  return end > p.valueAt ? html.slice(p.valueAt, end) : ''
}
/** every inline font-family set to `font`, as the old .replace(/font-family:\s*(?:&quot;|[^;"])*\/g, ...) */
export const setFontFamily = (html: string, font: string) => {
  let out = '', last = 0
  for (let p = propAt(html, 'font-family', false); p; p = propAt(html, 'font-family', false, p.at + 1)) {
    out += `${html.slice(last, p.at)}font-family: ${font}`
    last = valueEnd(html, p.valueAt, true)
  }
  return out + html.slice(last)
}
/** every number run in a string ("rgb(1, 2, 3)" -> [1, 2, 3]) */
export const numbersIn = (t: string) => {
  const out: number[] = []
  let cur = ''
  for (const ch of `${t} `) {
    if (isDigit(ch)) cur += ch
    else if (cur) { out.push(Number(cur)); cur = '' }
  }
  return out
}
/** the first opening tag whose name starts with `name` and is followed by a space or > ("<p ...>", "<span ...>"), or '' */
export const firstTag = (html: string, name: string, exact = false) => {
  for (let at = html.indexOf(`<${name}`); at >= 0; at = html.indexOf(`<${name}`, at + 1)) {
    const next = html[at + name.length + 1]
    if (exact && next !== '>' && !isSpace(next)) continue
    const close = html.indexOf('>', at)
    if (close > 0) return html.slice(at, close + 1)
  }
  return ''
}
/** each <p> or <p ...> gets font-size first in its style (the old <p(\s[^>]*)?> rewrite) */
export const sizeParagraphs = (html: string, px: number) => {
  let out = '', last = 0
  for (let at = html.indexOf('<p'); at >= 0; at = html.indexOf('<p', at + 1)) {
    const next = html[at + 2]
    if (next !== '>' && !isSpace(next)) continue
    const close = html.indexOf('>', at)
    if (close < 0) break
    const attrs = html.slice(at + 2, close)
    const tag = attrs.includes('style="') ? `<p${attrs.replace('style="', `style="font-size: ${px}px; `)}>` : `<p${attrs} style="font-size: ${px}px;">`
    out += html.slice(last, at) + tag
    last = close + 1
  }
  return out + html.slice(last)
}
/** the text after the first opening tag replaced (the old /(<[^/][^>]*>)([^<]+)(<)/ rewrite): the title's own styling, new words */
export const replaceFirstText = (html: string, text: string) => {
  for (let at = html.indexOf('<'); at >= 0; at = html.indexOf('<', at + 1)) {
    if (html[at + 1] === '/' || at + 1 >= html.length) continue
    const close = html.indexOf('>', at + 1)
    if (close < 0) return html
    const next = html.indexOf('<', close + 1)
    if (next > close + 1) return html.slice(0, close + 1) + text + html.slice(next)
  }
  return html
}
/** the first ">  12 <" (1-3 digits between tags) set to the number given (the old />\s*\d{1,3}\s*</ rewrite) */
export const replaceFirstNumber = (html: string, n: number) => {
  for (let at = html.indexOf('>'); at >= 0; at = html.indexOf('>', at + 1)) {
    const next = html.indexOf('<', at + 1)
    if (next < 0) return html
    const inner = html.slice(at + 1, next).trim()
    if (inner.length >= 1 && inner.length <= 3 && [...inner].every(isDigit)) return `${html.slice(0, at)}>${n}<${html.slice(next + 1)}`
  }
  return html
}
/** "&quot;Font Name&quot;" inside HTML written as 'Font Name' (the old /&quot;([^&]*?)&quot;/g rewrite) */
export const quotesToSingle = (html: string) => {
  let out = '', k = 0
  for (let at = html.indexOf('&quot;'); at >= 0; at = html.indexOf('&quot;', k)) {
    const end = html.indexOf('&quot;', at + 6)
    if (end < 0) break
    const inner = html.slice(at + 6, end)
    if (inner.includes('&')) { out += html.slice(k, at + 6); k = at + 6; continue }
    out += `${html.slice(k, at)}'${inner.split('\'').join('\\\'')}'`
    k = end + 6
  }
  return out + html.slice(k)
}
/** HTML comments removed */
export const dropComments = (t: string) => {
  let out = '', k = 0
  for (let at = t.indexOf('<!--'); at >= 0; at = t.indexOf('<!--', k)) {
    const end = t.indexOf('-->', at + 4)
    if (end < 0) break
    out += t.slice(k, at)
    k = end + 3
  }
  return out + t.slice(k)
}
/** the page or slide number a reader's marker line gives: "[page 3]", "[slide 3] ...", "Slide 3:", else 0 */
export const pageMark = (line: string) => {
  const num = (t: string) => (t.length && [...t].every(isDigit) ? Number(t) : 0)
  if (line.startsWith('[page ') && line.endsWith(']')) return num(line.slice(6, -1))
  if (line.startsWith('[slide ')) {
    const close = line.indexOf(']')
    if (close > 7 && (close === line.length - 1 || line[close + 1] === ' ')) return num(line.slice(7, close))
  }
  if (line.startsWith('Slide ') && line.endsWith(':')) return num(line.slice(6, -1))
  return 0
}
/** a line's reader marker removed ("[page 3]", "[slide 3]", "Slide 3:") */
export const dropPageMark = (line: string) => {
  for (const [open, close] of [['[page ', ']'], ['[slide ', ']'], ['Slide ', ':']]) {
    if (!line.startsWith(open)) continue
    const end = line.indexOf(close, open.length)
    if (end > open.length && [...line.slice(open.length, end)].every(isDigit)) return line.slice(end + 1)
  }
  return line
}
/** the words of a text, split at spaces, slashes and dashes, each with where it starts */
export const wordsAt = (t: string) => {
  const out: { word: string, at: number }[] = []
  let at = -1
  for (let k = 0; k <= t.length; k++) {
    const ch = t[k]
    const breaks = k === t.length || isSpace(ch) || '/-–—'.includes(ch)
    if (!breaks && at < 0) at = k
    if (breaks && at >= 0) { out.push({ word: t.slice(at, k), at }); at = -1 }
  }
  return out
}
/** scripts written without spaces between words (Han, kana, Thai, Lao, Khmer, Myanmar) */
export const hasUnspacedScript = (t: string) => [...t].some((ch) => {
  const c = ch.codePointAt(0)!
  return (c >= 0x3400 && c <= 0x9FFF) || (c >= 0xF900 && c <= 0xFAFF) || (c >= 0x20000 && c <= 0x2FA1F) || (c >= 0x3040 && c <= 0x30FF) || (c >= 0x31F0 && c <= 0x31FF)
    || (c >= 0x0E00 && c <= 0x0EFF) || (c >= 0x1780 && c <= 0x17FF) || (c >= 0x19E0 && c <= 0x19FF) || (c >= 0x1000 && c <= 0x109F) || (c >= 0xA9E0 && c <= 0xA9FF) || (c >= 0xAA60 && c <= 0xAA7F)
})
/** a text split after . ! ? … 。！？ followed by spaces, the old .split(/(?<=[.!?…。！？])\s+/) */
export const sentencesOf = (t: string) => {
  const out: string[] = []
  let from = 0
  for (let k = 1; k < t.length; k++) {
    if (isSpace(t[k]) && '.!?…。！？'.includes(t[k - 1])) {
      out.push(t.slice(from, k))
      while (k < t.length && isSpace(t[k])) k++
      from = k
    }
  }
  out.push(t.slice(from))
  return out
}
/** trailing . ! ? … 。！？ removed */
export const dropEndPunct = (t: string) => {
  let k = t.length
  while (k > 0 && '.!?…。！？'.includes(t[k - 1])) k--
  return t.slice(0, k)
}
/** "<file-content" / "</web-results" and the like (any case) written as "&lt;file-content" so a fence cannot be closed */
export const escapeFenceTags = (t: string) => {
  let out = '', k = 0
  // compared piece by piece: lower-casing the whole text first shifts every position after a letter like "İ"
  for (let at = t.indexOf('<'); at >= 0; at = t.indexOf('<', at + 1)) {
    const name = t[at + 1] === '/' ? at + 2 : at + 1
    const hit = ['file-content', 'web-results'].find(n => t.slice(name, name + n.length).toLowerCase() === n && !isWordChar(t[name + n.length]))
    if (!hit) continue
    out += `${t.slice(k, at)}&lt;${t.slice(name, name + hit.length)}`
    k = name + hit.length
  }
  return out + t.slice(k)
}
