// Eval only (LLM_CACHE_DIR=<dir>, imported by scripts/e2e/server.mts, never by api/ or Vercel): every model request the
// editor's handlers send (api.openai.com, api.anthropic.com; answered by the API or by Codex in LLM_VIA_CODEX=1 mode) is
// keyed by a hash of its URL and full body, and an identical request gets the saved response back. A build that does
// not touch a turn's path replays it exactly at $0; a change re-asks only the calls downstream of what it altered.
// Only 2xx answers are kept, each with the request it answered (a miss that should have hit: diff the two requests).
// A fresh or no dir = independent runs. Log: <dir>/cache.log, one "hit"/"miss" line per call.
import { createHash } from 'node:crypto'
import { appendFileSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

export function install(dir: string) {
  mkdirSync(dir, { recursive: true })
  const inner = globalThis.fetch
  const log = (s: string) => appendFileSync(join(dir, 'cache.log'), `${new Date().toISOString().slice(11, 19)} ${s}\n`)
  globalThis.fetch = (async (u: any, init?: any) => {
    const url = String(u?.url ?? u)
    if (!/^https:\/\/api\.(openai|anthropic)\.com\//.test(url)) return inner(u, init)
    const body = typeof init?.body === 'string' ? init.body : init?.body == null ? '' : Buffer.from(init.body).toString()
    const key = createHash('sha256').update(`${url}\n${body}`).digest('hex')
    const f = join(dir, `${key}.json`)
    if (existsSync(f)) {
      const r = JSON.parse(readFileSync(f, 'utf8'))
      log(`hit  ${key.slice(0, 12)} ${url.replace(/^https:\/\/[^/]+/, '')}`)
      return new Response(r.body, { status: r.status, headers: { 'Content-Type': r.type } })
    }
    const res = await inner(u, init)
    const text = await res.text() // a stream is read whole: the caller then reads the same bytes in one go
    const type = res.headers.get('content-type') || 'application/json'
    if (res.ok) { writeFileSync(`${f}.tmp`, JSON.stringify({ url, status: res.status, type, body: text, request: body })); renameSync(`${f}.tmp`, f) }
    log(`miss ${key.slice(0, 12)} ${url.replace(/^https:\/\/[^/]+/, '')} ${res.status}`)
    return new Response(text, { status: res.status, headers: { 'Content-Type': type } })
  }) as typeof fetch
  console.log(`LLM_CACHE_DIR: model responses cached in ${dir}`)
}
