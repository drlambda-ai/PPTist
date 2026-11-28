// Eval only: the router (api/tools/deck_chat.ts) on recorded inputs, with no browser and no server. Reads JSON lines
// {"id", "body"} (body = the editor_turn record's "router" field, the exact request the editor sent) and writes
// {"id", "calls", "usd"} per line. LLM_VIA_CODEX=1 answers every model request through Codex on the ChatGPT plan (free);
// without it the real API key is spent (~$0.002 a call). Used by scripts/deck-eval/route_eval.py (app.chatslide.ai).
//   npx tsx scripts/e2e/route.mts < in.jsonl > out.jsonl
import { createInterface } from 'node:readline'

if (process.env.LLM_VIA_CODEX === '1') (await import('./codex-llm.mts')).install()
const { default: handler } = await import('../../api/tools/deck_chat.ts')

const one = (body: unknown) => new Promise<{ calls: unknown, usd: number, error?: string }>(resolve => {
  let usd = 0
  const res: any = {
    headersSent: false,
    setHeader: (k: string, v: string) => { if (k === 'x-ai-cost-usd') usd = Number(v) },
    status: () => res,
    json: (o: any) => resolve({ calls: o.calls, usd, ...(o.error ? { error: String(o.error) } : {}) }),
    send: () => resolve({ calls: null, usd, error: 'no json' }), end: () => resolve({ calls: null, usd, error: 'no json' }),
  }
  Promise.resolve(handler({ method: 'POST', body, headers: {} }, res)).catch(e => resolve({ calls: null, usd, error: String(e) }))
})

const lines: string[] = []
for await (const l of createInterface({ input: process.stdin })) if (l.trim()) lines.push(l)
const jobs = Number(process.env.ROUTE_JOBS || 6)
let next = 0
await Promise.all(Array.from({ length: jobs }, async () => {
  while (next < lines.length) {
    const { id, body } = JSON.parse(lines[next++])
    process.stdout.write(`${JSON.stringify({ id, ...(await one(body)) })}\n`)
  }
}))
