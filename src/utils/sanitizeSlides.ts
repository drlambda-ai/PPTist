import DOMPurify from 'dompurify'
import type { PPTElement, Slide } from '@/types/slides'

export const cleanHtml = (html: string) => DOMPurify.sanitize(html, { ADD_ATTR: ['target'] })

const isWebLink = (target: string) => {
  try {
    return ['http:', 'https:'].includes(new URL(target, location.href).protocol)
  }
  catch {
    return false
  }
}

// Decks enter the store through here. Their HTML is rendered with v-html and their links opened with
// window.open on the app's own origin, and a deck can come from a link anyone can send.
// Table cells are not touched: they hold plain text ("p<0.05") and formatText rewrites them before
// display, which can undo a cleaning done here. Each table sink calls cleanHtml on the final string.
export const sanitizeElements = (elements: PPTElement[]) => {
  for (const el of elements) {
    if (el.type === 'text') el.content = cleanHtml(el.content)
    if (el.type === 'shape' && el.text) el.text.content = cleanHtml(el.text.content)
    if (el.link?.type === 'web' && !isWebLink(el.link.target)) delete el.link
  }
}

export const sanitizeSlides = (slides: Slide[]): Slide[] => {
  for (const slide of slides) {
    if (slide.remark) slide.remark = cleanHtml(slide.remark)
    sanitizeElements(slide.elements || [])
  }
  return slides
}
