import { chatStream, readBody } from '../_llm.js'

// Streams a markdown outline for a presentation topic (consumed raw by the frontend).

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST only' }); return }
  const { content = '', language = 'English', as_written = false } = await readBody(req)

  res.setHeader('Content-Type', 'application/octet-stream')
  res.setHeader('Cache-Control', 'no-cache')

  // the user's own pasted text (an essay, a report) laid out as it is: told only "take every number from it", a pasted
  // essay came back regrouped and retold, with claims it never made (2026-11-19)
  const asWritten = `You lay out the user's own text as a presentation, as a Markdown outline: "#" deck title (the text's own title, or its subject in its own words), "##" one per section of the text, in its order, with the section's own heading, "-" one bullet per paragraph or list item of that section, in order, every one.
Copy every word as written: never reword, summarise, merge, shorten or add. Each bullet is a bold heading of the paragraph's first few words, a colon, then the whole paragraph. A section with many paragraphs simply has many bullets. Leave out only source lines that are a bare link or site name, and the user's request to the editor ("make a presentation from this").
The second line after the deck title is one sentence copied from the text. Output only the outline, no preface, no code fences.`
  const messages = [
    {
      role: 'system',
      content: as_written ? asWritten : `You write the content of a presentation as a Markdown outline: "#" deck title, "##" one per slide, "-" bullets.
Every bullet says something, it never describes what will be said. Write "- **Germany**: EV share of new cars fell to 13.5% in 2024 after the subsidy ended", never "- **Germany**: analyse the German market". Banned bullet shapes: "介绍/分析/探讨/概述/研究 X", "overview of X", "discuss X".
Each content slide has 3-4 bullets: a bold short heading, a colon, then one sentence with the concrete fact, number (with its year), named example or implication.
Numbers: when the request includes material, take every number from it and invent none. Otherwise use only well-established public figures you are sure of, with the year; if unsure, make the point without a number.
The second line after the deck title is a one-sentence subtitle stating the deck's main conclusion. No placeholders anywhere (XXX, 姓名, 日期, presenter, your company).
Write in ${language}. Output only the outline, no preface, no code fences. If the user states a number of slides, plan exactly that many in total, counting a cover and a closing slide; otherwise aim for 5-8 slides.`,
    },
    { role: 'user', content: `${as_written ? 'Text' : 'Topic'}: ${content}` },
  ]

  try {
    await chatStream(messages, (delta) => res.write(delta), { model: 'gpt-6.1-sol', effort: 'low' })
    res.end()
  }
  catch (e: any) {
    res.write(`\n[error] ${String(e?.message || e)}`)
    res.end()
  }
}
