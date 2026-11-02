// Inside ChatSlide's editor page: PPTist is served same-origin at /api/pptist/ and the page hands it this object.
// API calls go through ChatSlide's proxy with the user's token; the deck is loaded and saved per project.
export interface ChatSlideHost {
  projectId: string
  locale?: string
  title: string
  token: () => Promise<string>
  canExport: () => boolean // false: ChatSlide shows its upgrade prompt
  load: () => Promise<{ session: any, slides: any[], theme: any } | null>
  save: (state: { session: any, slides: any[], theme: any }) => Promise<void>
  // the deck back into the project; returns the picture links of the pages it uploaded, by key
  sync: (pages: { key: string, url?: string, image?: Blob, text: string, notes: string }[]) => Promise<Record<string, string>>
  flush?: () => Promise<void> // set by the editor: sync now
  pptx: () => Promise<string> // the project's deck as a .pptx (blob: URL), for the first open
  // chat actions done by ChatSlide itself (absent when the editor runs alone)
  share?: () => Promise<{ url: string, copied: boolean }>
  goto?: (step: 'scripts' | 'video') => Promise<void>
  support?: (message: string) => Promise<string>
  outOfCredits?: () => void // ChatSlide opens its plans/paywall
  // the figures inside a PDF/PPTX/DOCX (ChatSlide backend extract-images), with the page each came from
  extractImages?: (file: File) => Promise<{ images?: { url: string, page?: number, name?: string }[], error?: string }>
  // AI picture through ChatSlide's image route (its credits, gates and moderation); { src } or { error }
  generateImage?: (prompt: string, aspectRatio: '16:9' | '9:16') => Promise<{ src?: string, error?: string }>
}

// A proxied call that ChatSlide refused for lack of credits (HTTP 402).
export class OutOfCredits extends Error {
  credits = 0
  required = 0
}

// Credits the proxy charged since the last takeSpend() (read by the chat panel for its footnote).
let spent = 0
let left = -1
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
const OWN = /^\/api\/(deck|deck_config|deck_upload|deck_put|tools\/)/
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
    const r = await fetch0(new Request(req, { ...init, headers }))
    if (r.status === 402) {
      const data = await r.json().catch(() => ({}))
      host.outOfCredits?.()
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
