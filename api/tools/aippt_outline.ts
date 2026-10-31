import { chatStream, readBody } from '../_llm.js'

// Streams a markdown outline for a presentation topic (consumed raw by the frontend).

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  const { content = '', language = 'English' } = await readBody(req)

  res.setHeader('Content-Type', 'application/octet-stream')
  res.setHeader('Cache-Control', 'no-cache')

  const messages = [
    {
      role: 'system',
      content: `You write the content of a presentation as a Markdown outline: "#" deck title, "##" one per slide, "-" bullets.
Every bullet says something, it never describes what will be said. Write "- **Germany**: EV share of new cars fell to 13.5% in 2024 after the subsidy ended", never "- **Germany**: analyse the German market". Banned bullet shapes: "介绍/分析/探讨/概述/研究 X", "overview of X", "discuss X".
Each content slide has 3-4 bullets: a bold short heading, a colon, then one sentence with the concrete fact, number (with its year), named example or implication.
Numbers: when the request includes material, take every number from it and invent none. Otherwise use only well-established public figures you are sure of, with the year; if unsure, make the point without a number.
The second line after the deck title is a one-sentence subtitle stating the deck's main conclusion. No placeholders anywhere (XXX, 姓名, 日期, presenter, your company).
Write in ${language}. Output only the outline, no preface, no code fences. If the user states a number of slides, plan exactly that many in total, counting a cover and a closing slide; otherwise aim for 5-8 slides.`,
    },
    { role: 'user', content: `Topic: ${content}` },
  ]

  try {
    await chatStream(messages, (delta) => res.write(delta), { model: 'gpt-5.5', effort: 'low' })
    res.end()
  }
  catch (e: any) {
    res.write(`\n[error] ${String(e?.message || e)}`)
    res.end()
  }
}
