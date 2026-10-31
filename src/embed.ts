import axios from '@/services/axios'

// Inside ChatSlide's editor page: PPTist is served same-origin at /api/pptist/ and the page hands it this object.
// API calls go through ChatSlide's proxy with the user's token; the deck is loaded and saved per project.
export interface ChatSlideHost {
  projectId: string
  token: () => Promise<string>
  canExport: () => boolean // false: ChatSlide shows its upgrade prompt
  load: () => Promise<{ session: any, slides: any[], theme: any } | null>
  save: (state: { session: any, slides: any[], theme: any }) => Promise<void>
  pptx: () => Promise<string> // the project's deck as a .pptx (blob: URL), for the first open
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
const route = (url: string) => {
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
  axios.interceptors.request.use(async config => {
    const to = config.url && route(config.url)
    if (to) {
      config.url = to
      config.headers.set('Authorization', `Bearer ${await host.token()}`)
    }
    return config
  })
}
