// Inside ChatSlide's editor page: PPTist is served same-origin at /api/pptist/ and the page hands it this object.
// API calls go through ChatSlide's proxy with the user's token; the deck is loaded and saved per project.
export interface ChatSlideHost {
  projectId: string
  locale?: string
  title: string
  token: () => Promise<string>
  canExport: () => boolean // false: ChatSlide shows its upgrade prompt
  exported?: (format: string) => void // a file was saved: ChatSlide records it (an export counts against eligibility)
  load: () => Promise<{ session: any, slides: any[], theme: any, viewportRatio?: number } | null>
  save: (state: { session: any, slides: any[], theme: any, viewportRatio?: number }) => Promise<void>
  // the deck back into the project; returns the picture links of the pages it uploaded, by key
  sync: (pages: { key: string, url?: string, image?: Blob, text: string, notes: string }[]) => Promise<Record<string, string>>
  flush?: () => Promise<void> // set by the editor: sync now
  versions?: () => Promise<{ url: string, at: string }[]> // the deck's saved copies, newest first (version history)
  // a thumbs up/down on a chat reply (with an optional comment), stored where ChatSlide keeps its chat feedback and
  // tied to the turn (the x-turn-id of that message's calls, as in pptist_turns and the editor_turn log)
  feedback?: (turn: string, good: boolean, comment: string, reply: string) => Promise<boolean>
  pptx: () => Promise<string> // the project's deck as a .pptx (blob: URL), for the first open
  // chat actions done by ChatSlide itself (absent when the editor runs alone)
  share?: () => Promise<{ url: string, copied: boolean }>
  goto?: (step: string) => Promise<void>
  support?: (message: string) => Promise<string>
  canUseAdvanced?: () => boolean // the advanced model is on paid plans
  offerAdvanced?: () => void // ChatSlide opens its paywall for it
  offerPaid?: () => void // ChatSlide opens its paywall for a paid feature (pictures from a website)
  featureRequest?: (what: string, message: string) => Promise<boolean> // true once it is recorded
  outOfCredits?: () => void // ChatSlide opens its plans/paywall
  uploadImage?: (image: Blob) => Promise<string> // a picture as its own file (the deck stores the link, not the bytes)
  classic?: () => void // back to ChatSlide's classic editor
  // the figures inside a PDF/PPTX/DOCX (ChatSlide backend extract-images), with the page each came from
  extractImages?: (file: File) => Promise<{ images?: { url: string, page?: number, name?: string }[], error?: string }>
  // the files the user gave this project in ChatSlide (the text read from them), so the chat uses them without a re-attach
  files?: () => Promise<{ name: string, text: string }[]>
  // what the user asked for when the project was made (their topic or pasted text, and the slide count they chose)
  brief?: () => Promise<{ text: string, pages?: number } | null>
  // one of those files as the user uploaded it (its colours, logo and pictures are not in its text)
  file?: (name: string) => Promise<Blob | null>
  // a question answered from the project's files in full (a long file is more than one read holds)
  askFiles?: (question: string) => Promise<{ answer: string, source: string } | null>
  // AI picture through ChatSlide's image route (its credits, gates and moderation); { src } or { error }
  generateImage?: (prompt: string, aspectRatio: '16:9' | '9:16') => Promise<{ src?: string, error?: string }>
  // ChatSlide's own templates, for rebuilding a page with one: what can be chosen (own = the palette this project was
  // made with, on its template) ...
  restyleOptions?: () => { templates: { name: string, label: string, art: boolean, palettes: string[], own?: string }[], layouts: { name: string, fits: string }[] }
  // ... and page `page` (1-based) as a one-page .pptx (blob: URL): the template's design and layout ('' = the one that
  // fits) holding this text, these pictures (links, each with its width / height on the page) and the page's own
  // full-page photo ('' = none). null when it cannot be built
  restyle?: (page: number, look: { template: string, palette: string, layout: string, texts: RestyleTexts, images: { src: string, ratio: number }[], photo: string }) => Promise<string | null>
}
export interface RestyleTexts { title: string, items: { heading: string, text: string }[] }

// The Share button (desktop and phone bars): the editor had none, so shares fell to zero at launch (2026-11-10).
export async function shareDeck(t: (k: string, v?: Record<string, string | number>) => string, say: (text: string) => void) {
  if (!host?.share || !host.canExport()) return // canExport() false opens ChatSlide's upgrade prompt
  const { url, copied } = await host.share()
  say(t('Sharing is on: {p0}', { p0: url }) + (copied ? t(' (link copied)') : ''))
}

// A proxied call that ChatSlide refused for lack of credits (HTTP 402).
export class OutOfCredits extends Error {
  credits = 0
  required = 0
}

// Credits the proxy charged since the last takeSpend() (read by the chat panel for its footnote).
let spent = 0
let left = -1
// ChatSlide charges a chat message what its AI calls cost, at least a credit: every call of the message carries its
// name (newTurn starts the next message)
let turnId = crypto.randomUUID()
export const newTurn = () => { turnId = crypto.randomUUID() }
export const currentTurn = () => turnId
export const takeSpend = () => {
  const out = { charged: spent, left }
  spent = 0
  return out
}

export const host: ChatSlideHost | null = (() => {
  try {
    const p = window.parent as any
    return p !== window && p.__chatslide ? p.__chatslide : null
  }
  catch {
    return null // cross-origin parent: not ChatSlide
  }
})()

// PPTist's own endpoints (fixed paths in this app), moved under the proxy
const OWN = /^\/api\/(deck|deck_config|deck_upload|deck_put|record_turn|tools\/)/
export const route = (url: string) => {
  const u = new URL(url, location.origin)
  return u.origin === location.origin && OWN.test(u.pathname) ? `/api/pptist${u.pathname}${u.search}` : null
}

if (host) {
  const fetch0 = window.fetch.bind(window)
  window.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url
    const to = route(url)
    if (!to) return fetch0(input, init)
    const req = new Request(to, input instanceof Request ? input : undefined)
    const headers = new Headers(init?.headers || req.headers)
    headers.set('Authorization', `Bearer ${await host.token()}`)
    headers.set('x-turn-id', turnId)
    headers.set('x-project-id', host.projectId) // ChatSlide records each chat turn per project, for review
    const r = await fetch0(new Request(req, { ...init, headers }))
    if (r.status === 402) {
      const data = await r.json().catch(() => ({}))
      if (!headers.get('X-Quiet-402')) host.outOfCredits?.() // quiet: the caller has a free way to do it
      const e = new OutOfCredits('out of credits')
      e.credits = Number(data.credits) || 0
      e.required = Number(data.required) || 0
      throw e
    }
    const charged = Number(r.headers.get('x-credits-charged') || 0)
    if (charged > 0) {
      spent += charged
      left = Number(r.headers.get('x-credits-left') || -1)
    }
    return r
  }
}

export default host
