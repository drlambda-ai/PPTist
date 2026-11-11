// Click-through tests: the real editor build + the real PPTist AI handlers on [::1]:3001, with a stand-in for ChatSlide
// (scripts/e2e/harness.html: ?deck=<file in decks/>&paid=1). Logs every AI call to $E2E_LOG. Run: scripts/e2e/serve.sh
import http from 'http'
import { readFileSync, existsSync, appendFileSync } from 'fs'
import { extname, join } from 'path'
const DIST = process.env.E2E_DIST || '/tmp/pptist-e2e/dist'
const LOG = process.env.E2E_LOG || '/tmp/pptist-e2e/calls.log'
const types: Record<string, string> = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon', '.woff2': 'font/woff2' }
const handlers: Record<string, any> = {}
http.createServer(async (req, res0) => {
  const url = new URL(req.url || '/', 'http://x')
  const path = url.pathname.replace(/^\/api\/pptist/, '')
  const m = path.match(/^\/api\/tools\/(\w+)$/)
  if (m && m[1] === 'research' && process.env.BREAK_RESEARCH === '1') { res0.writeHead(200, { 'Content-Type': 'application/json' }); res0.end('true{broken'); return }
  if (m) {
    handlers[m[1]] ||= (await import(new URL(`../../api/tools/${m[1]}.ts`, import.meta.url).href)).default
    let code = 200
    const res: any = { setHeader: (k: string, v: string) => res0.setHeader(k, v), status(n: number) { code = n; return res }, json(o: any) { res0.writeHead(code, { 'Content-Type': 'application/json' }); res0.end(JSON.stringify(o)); appendFileSync(LOG, `${new Date().toISOString().slice(11, 19)} ${m[1]} ${code} ${JSON.stringify(o).slice(0, 3000)}\n`) }, end(b?: any) { res0.writeHead(code); res0.end(b) } }
    if (m[1] === 'deck_chat') { const chunks: Buffer[] = []; for await (const c of req) chunks.push(c as Buffer); const body = JSON.parse(Buffer.concat(chunks).toString() || '{}'); (req as any).body = body; if (!body.localize) appendFileSync(LOG, `>> current=${body.current} pending=${JSON.stringify(body.pending || null).slice(0, 160)} msg=${String(body.instruction).slice(0, 60)}\n`) }
    try { await handlers[m[1]](req, res) } catch (e: any) { res0.writeHead(500); res0.end(String(e?.message || e)) }
    return
  }
  if (path.startsWith('/api/')) { res0.writeHead(200, { 'Content-Type': 'application/json' }); res0.end('{}'); return } // record_turn, deck saves
  const f = join(DIST, path === '/' ? 'index.html' : path)
  if (!existsSync(f)) { res0.writeHead(404); res0.end(); return }
  res0.writeHead(200, { 'Content-Type': types[extname(f)] || 'application/octet-stream' }); res0.end(readFileSync(f))
}).listen(3001, '::1', () => console.log('ready'))
