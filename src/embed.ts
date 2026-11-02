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
    return fetch0(new Request(req, { ...init, headers }))
  }
}

export default host
