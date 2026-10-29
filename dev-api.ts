// Dev only: serve this repo's Vercel functions from the vite dev server, so the whole app runs on one port.
// api/*.ts and api/tools/*.ts run in-process; api/deck.py runs as `python3 -m deck.service <action>`;
// uploads land in .deck-data and are served under /blob (Vercel Blob in production).
import fs from 'node:fs'
import path from 'node:path'
import { randomUUID } from 'node:crypto'
import { spawn } from 'node:child_process'
import type { Plugin, ViteDevServer } from 'vite'

const ROOT = path.dirname(new URL(import.meta.url).pathname)
const DATA = path.join(ROOT, '.deck-data')

const readRaw = (req: any): Promise<Buffer> => new Promise((resolve) => {
  const chunks: Buffer[] = []
  req.on('data', (c: Buffer) => chunks.push(c))
  req.on('end', () => resolve(Buffer.concat(chunks)))
})

const runPython = (action: string, payload: Buffer): Promise<string> => new Promise((resolve) => {
  const p = spawn('python3', ['-m', 'deck.service', action], { cwd: ROOT, env: process.env })
  let out = ''
  let err = ''
  p.stdout.on('data', (d) => {
    out += d 
  })
  p.stderr.on('data', (d) => {
    err += d 
  })
  p.on('close', () => resolve(out.trim() || JSON.stringify({ error: err.trim().split('\n').pop() || 'python failed' })))
  p.stdin.end(payload)
})

export default function devApi(): Plugin {
  return {
    name: 'dev-api',
    apply: 'serve',
    configureServer(server: ViteDevServer) {
      server.middlewares.use(async (req: any, res: any, next) => {
        const url: string = (req.url || '').split('?')[0]
        res.status = (code: number) => {
          res.statusCode = code; return res 
        }
        res.json = (obj: any) => {
          res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify(obj)) 
        }
        try {
          if (url.startsWith('/blob/')) {
            const file = path.normalize(path.join(DATA, decodeURIComponent(url.slice(6))))
            if (!file.startsWith(DATA) || !fs.existsSync(file)) return res.status(404).end()
            return res.end(fs.readFileSync(file))
          }
          if (url === '/api/deck_put') {
            const name = decodeURIComponent(req.headers['x-filename'] || 'file').replace(/[/\\]/g, '_')
            const rel = `decks/${randomUUID().replace(/-/g, '')}/${name}`
            fs.mkdirSync(path.join(DATA, path.dirname(rel)), { recursive: true })
            fs.writeFileSync(path.join(DATA, rel), await readRaw(req))
            return res.json({ url: `/blob/${rel}` })
          }
          if (url === '/api/deck') {
            const body = JSON.parse((await readRaw(req)).toString() || '{}')
            const action = body.action
            delete body.action
            res.setHeader('Content-Type', 'application/json')
            return res.end(await runPython(action, Buffer.from(JSON.stringify(body))))
          }
          if (url.startsWith('/api/')) {
            const mod = await server.ssrLoadModule(path.join(ROOT, `${url}.ts`))
            return await mod.default(req, res)
          }
        }
        catch (e: any) {
          return res.status(500).json({ error: String(e?.message || e) })
        }
        next()
      })
    },
  }
}
