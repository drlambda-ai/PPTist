import { getKey, readBody } from '../_llm.js'

// Deck-level chat router (same design as ChatSlide's assistant): the model picks ONE tool per turn from a fixed menu;
// the editor applies it (page operations directly, slide rewrites through chat_edit). Parts of a multi-part request
// that one tool cannot do go into followup_note so they are acknowledged, never silently dropped.

const MODEL = process.env.DECK_ROUTER_MODEL || 'gpt-5.4-mini'

const page = { type: 'integer', description: '1-based page number in the current deck' }
const note = { type: 'string', description: 'ONLY parts of the user\'s own request this tool leaves undone. Empty when everything asked is done. Never offers or suggestions.' }
const fn = (name: string, description: string, properties: Record<string, any>, required: string[]) => ({
  type: 'function',
  function: {
    name, description,
    parameters: { type: 'object', properties: { ...properties, followup_note: note }, required, additionalProperties: false },
  },
})

const TOOLS = [
  fn('add_page', 'Add a new page with new content after a given page, built on the layout of a similar page.', {
    instruction: { type: 'string', description: 'What the new page must say, specific and complete' },
    after: page,
    layout_from: { ...page, description: 'Existing page whose layout fits the new content best' },
  }, ['instruction', 'after', 'layout_from']),
  fn('delete_page', 'Delete one or more pages.', { pages: { type: 'array', items: page } }, ['pages']),
  fn('duplicate_page', 'Duplicate a page (the copy goes right after it).', { index: page }, ['index']),
  fn('move_page', 'Move a page to a new position.', { from: page, to: page }, ['from', 'to']),
  fn('update_slide', 'Change the content, wording, styling or arrangement of one page.', {
    index: page, instruction: { type: 'string', description: 'The edit, specific and complete' },
  }, ['index', 'instruction']),
  fn('update_all_pages', 'Apply the same kind of change to every page (tone, language, font colour, shorten...).', {
    instruction: { type: 'string' },
  }, ['instruction']),
  fn('answer_question', 'Answer a question about the deck or the editor without changing anything.', {
    answer: { type: 'string' },
  }, ['answer']),
]

const SYSTEM = `You route a user's request about their slide deck to exactly one tool.
The deck is listed below, one line per page: number, then its text. The user is looking at page {CURRENT}.

Decision rules:
1. "this page/slide", "here" means page {CURRENT}. Page numbers the user gives are 1-based.
2. Adding a page with new content -> add_page; pick layout_from = the existing page whose structure fits best.
3. Removing pages -> delete_page (several pages in one call is fine). Never delete the only page.
4. Edits to one page (text, numbers, colours, size, position, add a text box) -> update_slide.
5. The same change on every page -> update_all_pages.
6. Questions ("how do I...", "what does page 3 say") -> answer_question; do not change the deck.
7. If the request has several parts, do the most important part and put the rest in followup_note. Otherwise leave
   followup_note empty: no offers, no suggestions.
8. Answer in the user's language.`

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST only' }); return 
  }
  try {
    const { instruction = '', pages = [], current = 1, history = [] } = await readBody(req)
    const deck = pages.map((p: string, i: number) => `${i + 1}. ${String(p).slice(0, 300)}`).join('\n')
    const r = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${getKey()}` },
      body: JSON.stringify({
        model: MODEL,
        tools: TOOLS,
        tool_choice: 'required',
        parallel_tool_calls: false,
        messages: [
          { role: 'system', content: SYSTEM.replaceAll('{CURRENT}', String(current)) + `\n\nDeck:\n${deck}` },
          ...history.slice(-6),
          { role: 'user', content: instruction },
        ],
      }),
    })
    if (!r.ok) throw new Error(`OpenAI ${r.status}: ${await r.text()}`)
    const call = (await r.json()).choices?.[0]?.message?.tool_calls?.[0]
    let args: any = {}
    try {
      args = JSON.parse(call?.function?.arguments || '{}') 
    }
    catch {
      args = {} 
    } // malformed arguments become an empty call, never a 500
    res.status(200).json({ action: call?.function?.name || 'answer_question', args })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
