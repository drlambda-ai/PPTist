import { getKey, readBody } from '../_llm.js'

// Deck-level chat router, modelled on ChatSlide's assistant menu (chatslide-backend util/assistant/slide_tools.json)
// but able to do several steps: the editor calls it in a loop, applying each round's calls and sending back what was
// done, until the model calls finish. Page numbers always refer to the deck as listed in that round.

const MODEL = process.env.DECK_ROUTER_MODEL || 'gpt-5.5'

const page = { type: 'integer', description: '1-based page number in the deck as listed' }
const pages = { type: 'array', items: page }
const fn = (name: string, description: string, properties: Record<string, any>, required: string[]) => ({
  type: 'function',
  function: { name, description, parameters: { type: 'object', properties, required, additionalProperties: false } },
})

const TOOLS = [
  fn('add_page', 'Add one or more new pages with new content after a given page ("add 3 slides on X", "expand into more slides").', {
    instruction: { type: 'string', description: 'What the new page(s) must say, specific and complete' },
    count: { type: 'integer', description: 'How many new pages, usually 1' },
    after: page,
    layout_from: { ...page, description: 'Existing page whose layout fits the new content best' },
  }, ['instruction', 'count', 'after', 'layout_from']),
  fn('expand_page', 'Extend / expand / go deeper on an existing page ("扩展/展开/扩充第3页", "extend page 3", "more detail on this slide"): new analysis that goes beyond what the page already says, on a continuation page right after it. The page itself stays as it is.', {
    index: page,
    instruction: { type: 'string', description: 'What the user wants more of (or "" for general depth)' },
  }, ['index', 'instruction']),
  fn('add_chart_page', 'Add a new page whose main content is a chart (bar, column, line, pie, ring, area, radar, scatter).', {
    instruction: { type: 'string', description: 'What the chart shows, with the numbers if the user or deck gives them' },
    after: page,
    chart_type: { type: 'string', enum: ['bar', 'column', 'line', 'pie', 'ring', 'area', 'radar', 'scatter'] },
  }, ['instruction', 'after', 'chart_type']),
  fn('add_table_page', 'Add a new page whose main content is a table.', {
    instruction: { type: 'string', description: 'What the table shows (rows, columns, values) as the user or deck gives it' },
    after: page,
  }, ['instruction', 'after']),
  fn('delete_pages', 'Delete one or more pages.', { pages }, ['pages']),
  fn('duplicate_pages', 'Duplicate pages; each copy goes right after its original.', { pages }, ['pages']),
  fn('move_pages', 'Move one or more pages (kept in their order) so the first of them becomes page `to`. Use it for "move forward/back", "put X at the end/front" (a swap is swap_pages).', {
    pages, to: { type: 'integer', description: 'The page number the (first) moved page has AFTER the move: 1 = front, the deck length = end. "Move page 5 forward one" -> to 4; "move the last page to page 2" -> to 2.' },
  }, ['pages', 'to']),
  fn('reorder_pages', 'Put the whole deck in a new order (when the user describes the full order or a sort).', {
    order: { ...pages, description: 'Every current page number exactly once, in the new order' },
  }, ['order']),
  fn('swap_pages', 'Swap the positions of two pages.', { a: page, b: page }, ['a', 'b']),
  fn('edit_page', 'Change what is ON one page: wording, numbers, font family/size/colour/bold/italic/alignment, moving or resizing text, images or shapes, deleting elements, adding a text box or shape.', {
    index: page, instruction: { type: 'string', description: 'The edit, specific and complete' },
  }, ['index', 'instruction']),
  fn('set_font', 'Change the font family of text, on some or all pages (done directly, keeps every layout).', {
    pages: { ...pages, description: 'Pages; empty = every page' },
    font: { type: 'string', description: 'Font family name, e.g. "Microsoft YaHei", "Arial", "SimHei"' },
    scope: { type: 'string', enum: ['all', 'titles', 'body'], description: 'all text, only titles, or only non-title text' },
  }, ['pages', 'font', 'scope']),
  fn('edit_all_pages', 'Apply the same element-level change to every page (font, text colour, tone, translate, shorten...).', {
    instruction: { type: 'string' },
  }, ['instruction']),
  fn('change_image', 'Replace a picture on a page with a new one: a stock photo search, or an AI-generated image.', {
    index: page,
    image: { type: 'integer', description: '1-based number of the image on that page, in the listed order; 1 if there is only one' },
    query: { type: 'string', description: 'Stock search words (English) or the image-generation prompt' },
    source: { type: 'string', enum: ['photo', 'ai'], description: 'photo unless the user asks for an illustration, AI image or something no photo shows' },
  }, ['index', 'image', 'query', 'source']),
  fn('remove_images', 'Remove pictures from pages (done directly). Never logos unless the user names the logo.', {
    pages: { ...pages, description: 'Pages; empty = every page' },
    images: { type: 'array', items: { type: 'integer' }, description: 'Image numbers as listed (applied on each given page); empty = every non-logo image' },
  }, ['pages', 'images']),
  fn('move_element_to_page', 'Move or copy an image or text box from one page to another page (same position).', {
    from: page, to: page,
    element: { type: 'string', description: 'Which element, as the user described it (e.g. "the photo", "image 2", "the title")' },
    copy: { type: 'boolean', description: 'true = keep the original too (copy), false = move' },
  }, ['from', 'to', 'element', 'copy']),
  fn('add_image', 'Insert a NEW picture on a page (stock photo or AI-generated), at a given place.', {
    index: page,
    query: { type: 'string', description: 'Stock search words (English) or the image-generation prompt' },
    source: { type: 'string', enum: ['photo', 'ai'] },
    box: {
      type: 'object', additionalProperties: false, required: ['x', 'y', 'w', 'h'],
      description: 'Where, in percent of the slide (0-100): left, top, width, height. Pick free space on that page.',
      properties: { x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' } },
    },
  }, ['index', 'query', 'source', 'box']),
  fn('set_background', 'Set the background of pages: a colour, or a photo ("background photo", "更好看的背景", "fondo").', {
    pages: { ...pages, description: 'Pages to change; empty = every page' },
    color: { type: 'string', description: '#RRGGBB, or "" when a photo is wanted' },
    photo_query: { type: 'string', description: 'Stock photo search words (English) for a background photo, or ""' },
  }, ['pages', 'color', 'photo_query']),
  fn('write_notes', 'Write, expand, shorten or restyle the SPEAKER NOTES / script (what the presenter says), not the slide text.', {
    pages: { ...pages, description: 'Pages; empty = every page' },
    instruction: { type: 'string', description: 'What to do with the notes ("expand to about 1 minute", "funnier", "write them"), or ""' },
  }, ['pages', 'instruction']),
  fn('set_transition', 'Set the page transition effect.', {
    pages: { ...pages, description: 'Pages; empty = every page' },
    effect: { type: 'string', enum: ['no', 'fade', 'slideX', 'slideY', 'random', 'slideX3D', 'slideY3D', 'rotate', 'scaleY', 'scaleX', 'scale', 'scaleReverse'] },
  }, ['pages', 'effect']),
  fn('animate_pages', 'Give the elements of pages an entrance animation, appearing one after another in reading order (or remove animations).', {
    pages: { ...pages, description: 'Pages; empty = every page' },
    effect: { type: 'string', enum: ['none', 'fadeIn', 'fadeInUp', 'fadeInLeft', 'zoomIn', 'bounceIn', 'slideInLeft', 'slideInUp', 'flipInX', 'backInUp'] },
    trigger: { type: 'string', enum: ['click', 'auto'], description: 'click = one per click; auto = one after another by itself' },
  }, ['pages', 'effect', 'trigger']),
  fn('rebuild_from_template', 'Regenerate the whole deck from the company template with new or changed material/data (only when a company template is in use).', {
    instruction: { type: 'string' },
  }, ['instruction']),
  fn('undo', 'Undo the last change.', {}, []),
  fn('redo', 'Redo the last undone change.', {}, []),
  fn('finish', 'Everything the user asked for is done (only after at least one step was applied).', {}, []),
  fn('answer_question', 'Answer a question about the deck or the editor without changing anything, or ask what the user means when the request is truly ambiguous.', {
    answer: { type: 'string' },
  }, ['answer']),
]

const SYSTEM = `You turn a user's request about their slide deck into tool calls.
The deck is listed below, one line per page: number, then its elements. The user is looking at page {CURRENT}.
{TEMPLATE}
Rules:
1. "this page/slide", "here" = page {CURRENT}. "next/previous page" are relative to it. Page numbers are 1-based and
   ALWAYS refer to the deck exactly as listed below, even when you make several calls.
2. A request with several parts is done one step at a time: call the tool for the NEXT part only; you will be asked
   again with the updated deck. Do every part; never drop one.
   Steps already applied for this request are listed under "Done so far"; the deck below already reflects them.
   Call the tool(s) for what remains, or finish when nothing remains.
   "Extend/expand/展开/扩展/扩充 page N" -> expand_page (never add_page, never edit_page) unless the user asks to rewrite
   page N itself or to add a page about a NEW topic.
3. Reordering ("move page 5 forward", "put 2 and 3 at the end", "swap 2 and 4") -> move_pages. Forward = earlier.
   Exchanging two pages -> swap_pages. A full new order or a sort -> reorder_pages.
4. Anything about what is ON a page (text, fonts, colours, sizes, positions of text/images/shapes) -> edit_page for that page.
   Only a font FAMILY change (字体/font/Schrift/шрифт) -> set_font, never edit_page/edit_all_pages.
   ALL changes to one page go into ONE edit_page call with one combined instruction.
   The same change on every page -> edit_all_pages.
   Removing pictures -> remove_images (not edit_page).
5. A new or different picture -> change_image, ONE call for the one image meant: "the picture/photo" is the image marked "the largest
   image here" on that page; never an icon or a [logo] unless the user names it. If that page has no image except
   icons/logos, use add_image instead. Replace several images only when asked. Moving a picture/text to another page -> move_element_to_page.
   Merging pages A and B: edit_page on A (instruction: combine B's content into A, giving its full text), then delete_pages B.
   "Translate this slide" means the visible slide text (edit_page / edit_all_pages); "script/notes/讲稿/备注" means write_notes.
   Overflowing or overlapping text -> edit_page ("shorten or resize so nothing overflows or overlaps").
   "Only N slides / reduce to N / shorter deck / 只要N页" -> delete_pages: the least essential pages until EXACTLY N remain
   (the count wins: keep the cover and the closing page only while N leaves room for them after the content; prefer
   deleting pages that repeat others). Never write_notes for it. When the cut removes half the deck or
   more (e.g. "only one page"), FIRST edit_page the content page that will stay so it summarises the key points of the
   pages being removed, THEN delete_pages; never leave only an agenda or cover.
   Notes/script tools only when the message is about the script, speaker notes, 讲稿, Sprechertext, guion, narration text.
   "Add/include another slide" with no topic -> add_page on the next logical topic of the deck, after the current page; don't ask.
   A message that is just pasted content (a paragraph of facts, no instruction) -> edit_page on the current page: put that
   content on it (or add_page if the current page is a cover/agenda/closing page).
   Video, voice-over, narration audio, avatars, export to video: this editor cannot make them -> answer_question saying so
   (speaker notes can be written for a script).
   Vague dislike with no target ("no me gusta", "bad slides", "不好看"): answer_question with ONE short question offering
   2-3 concrete options (e.g. shorter text / different layout / new picture). Never redesign the whole deck on a guess.
   The user's OWN logo or image ("my logo", "mi logotipo", "我的图片") is not in the chat: answer_question telling them to
   insert it once with the toolbar's image button on one page, then ask to copy it to the pages they want (only the pages
   they name, or all if they said all/the slides).
   "Review the deck AND change/fix what is wrong" -> edit_all_pages (fix overflow, overlap, alignment, typos; keep content).
   A review with no request to change anything -> answer_question listing the issues.
   "Fix/revise/修改/sửa lại page N" with no detail -> edit_page on N: fix overflow, overlap, alignment and typos; change
   wording only where it is broken; keep the content and design.
   Several wishes in one message ("prettier, remove the table of contents, add more") -> do each CONCRETE wish as its own
   step; for a vague one ("prettier", "more stuff"), end with answer_question asking what exactly, with 2-3 options.
6. Questions ("how do I...", "what does page 3 say") -> answer_question only. Facts about this editor: Download (AI panel)
   saves the deck as .pptx; the ☰ menu -> 导出 also exports PDF, images or JSON; ▶ at the top right presents;
   "+ Template" uploads a company .pptx template, "+ Content" uploads material (xlsx, docx, pdf, md...); every change
   can be undone with Undo. Say you don't know rather than inventing prices, limits or features.
7. Never delete every page.`

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST only' }); return
  }
  try {
    const { instruction = '', pages: deckPages = [], current = 1, history = [], template = false, done = [] } = await readBody(req)
    const deck = deckPages.map((p: string, i: number) => `${i + 1}. ${String(p).slice(0, 400)}`).join('\n')
    const system = SYSTEM.replaceAll('{CURRENT}', String(current)).replace('{TEMPLATE}', template
      ? 'The deck was built from the user\'s company template.'
      : 'No company template is in use: never call rebuild_from_template.')
    const r = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${getKey()}` },
      body: JSON.stringify({
        model: MODEL,
        tools: TOOLS,
        tool_choice: 'required',
        parallel_tool_calls: false,
        messages: [
          { role: 'system', content: system + `\n\nDeck:\n${deck}` },
          ...history.slice(-6),
          { role: 'user', content: instruction + (done.length ? `\n\nDone so far:\n${done.map((d: string, i: number) => `${i + 1}. ${d}`).join('\n')}` : '') },
        ],
      }),
    })
    if (!r.ok) throw new Error(`OpenAI ${r.status}: ${await r.text()}`)
    const calls = ((await r.json()).choices?.[0]?.message?.tool_calls || []).map((c: any) => {
      let args: any = {}
      try {
        args = JSON.parse(c.function?.arguments || '{}')
      }
      catch {
        args = {}
      } // malformed arguments become an empty call, never a 500
      return { action: c.function?.name || 'answer_question', args }
    })
    res.status(200).json({ calls: calls.length ? calls : [{ action: 'answer_question', args: {} }] })
  }
  catch (e: any) {
    res.status(500).json({ error: String(e?.message || e) })
  }
}
