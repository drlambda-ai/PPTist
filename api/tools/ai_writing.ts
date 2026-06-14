import { chatStream, readBody } from '../_llm'

// Streams an AI text rewrite (raw text) for the in-editor "AI writing" action.

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  const { content = '', command = '' } = await readBody(req)

  res.setHeader('Content-Type', 'application/octet-stream')
  res.setHeader('Cache-Control', 'no-cache')

  const messages = [
    {
      role: 'system',
      content: 'You rewrite or generate slide text. Return only the resulting text, no preamble, no quotes.',
    },
    { role: 'user', content: `${command}\n\nText:\n${content}` },
  ]

  try {
    await chatStream(messages, (delta) => res.write(delta), { temperature: 0.7 })
    res.end()
  }
  catch (e: any) {
    res.write(`\n[error] ${String(e?.message || e)}`)
    res.end()
  }
}
