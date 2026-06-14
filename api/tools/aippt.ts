import { chatStream, readBody } from '../_llm'

// Streams slides as JSONL (one AIPPTSlide object per line). The frontend parses each
// line and maps it onto a template. We buffer deltas and only flush complete lines so a
// JSON object is never split across network chunks.

const SYSTEM = `You convert a presentation outline into slide data, emitted as JSONL: ONE JSON object per line, no code fences, no extra prose.
Allowed line shapes:
{"type":"cover","data":{"title":"...","text":"..."}}
{"type":"contents","data":{"items":["...","..."]}}
{"type":"transition","data":{"title":"...","text":"..."}}
{"type":"content","data":{"title":"...","items":[{"title":"...","text":"..."}]}}
{"type":"end"}
Order: exactly one "cover" first, then one "contents", then several "content" slides (use "transition" between major sections), then exactly one "end".
Each "content" slide should have 2-4 items. Write in {LANG}. Output ONLY JSONL.`

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  const { content = '', language = 'English' } = await readBody(req)

  res.setHeader('Content-Type', 'application/octet-stream')
  res.setHeader('Cache-Control', 'no-cache')

  const messages = [
    { role: 'system', content: SYSTEM.replace('{LANG}', language) },
    { role: 'user', content: `Outline:\n${content}` },
  ]

  let buffer = ''
  try {
    await chatStream(messages, (delta) => {
      buffer += delta
      const lines = buffer.split('\n')
      buffer = lines.pop() || ''
      for (const line of lines) {
        if (line.trim()) res.write(line + '\n')
      }
    }, { temperature: 0.6, maxTokens: 8192 })
    if (buffer.trim()) res.write(buffer + '\n')
    res.end()
  }
  catch (e: any) {
    res.write(`\n[error] ${String(e?.message || e)}`)
    res.end()
  }
}
