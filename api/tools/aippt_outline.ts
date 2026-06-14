import { chatStream, readBody } from '../_llm'

// Streams a markdown outline for a presentation topic (consumed raw by the frontend).

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  const { content = '', language = 'English' } = await readBody(req)

  res.setHeader('Content-Type', 'application/octet-stream')
  res.setHeader('Cache-Control', 'no-cache')

  const messages = [
    {
      role: 'system',
      content: `You are a presentation outline writer. Output a clean Markdown outline only (no preface, no code fences). Use "#" for the deck title, "##" for each slide/section title, and "-" bullets for key points. Write in ${language}. Aim for 5-8 sections.`,
    },
    { role: 'user', content: `Topic: ${content}` },
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
