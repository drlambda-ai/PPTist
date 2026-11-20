// The record the check of a finished chat message reads, fitted under ask_json's cap (it cuts an instruction at
// 30,000 characters, silently). Cut bare at the end, a 21-page build after a 6,000-character request lost the changes
// of pages 17-21 while the reply ("21 pages") stayed: a finished turn then reads as a claimed change the list lacks,
// and is undone and built again (review of 926c1d0b). So the context is shortened first (the earlier conversation,
// the answered request, the deck as it was, the steps), then each changed page's text; every changed page keeps its
// line, and whatever was shortened or left out is said in the record, so absence there is never a ground.
// tools/turn-check-prompt.test.mjs builds that case.
export const CHECK_LIMIT = 29500

export interface CheckParts {
  head: string // the rules, the slide count asked, the user's files, the message: never cut
  reply: string
  request: (max: number) => string // the question the message answers, with its request cut to max ('' when none)
  earlier: (userMax: number) => string // the earlier conversation, a user message cut to userMax
  deck: (perPage: number) => string[] // the deck as it was, one line a page
  steps: (max: number) => string
  changes: (room: number) => string[] // one entry per changed page, its text cut to room characters a side
}

const LEVELS = [
  { request: 6000, earlier: 6000, deck: 500, steps: 5000 },
  { request: 1500, earlier: 300, deck: 500, steps: 5000 },
  { request: 1500, earlier: 300, deck: 100, steps: 2000 },
  { request: 600, earlier: 0, deck: 0, steps: 1000 },
]
const SHORTENED = 'For length, the earlier conversation, the answered request, the deck as it was and the steps are shortened or left out above and below: what is missing there for that reason is no ground.\n'

export const fitCheck = (p: CheckParts, room0: number, limit = CHECK_LIMIT) => {
  const put = (k: number, list: string[], more = '') => {
    const l = LEVELS[k]
    const deck = l.deck ? p.deck(l.deck).join('\n') : '(left out for length)'
    const earlier = l.earlier ? p.earlier(l.earlier) || '(none)' : '(left out for length)'
    return `${p.head}${p.request(l.request)}${k ? SHORTENED : ''}`
      + `The deck when the message came, as the editor saw it (text that repeats on most pages is left out):\n${deck}\n`
      + `The editor's steps, each with its result:\n${p.steps(l.steps) || '(none)'}\nThe editor's reply: ${JSON.stringify(p.reply.slice(0, 1500))}\n`
      + `What changed in the deck:\n${list.join('\n') || '(nothing)'}\n${more}`
      + `Earlier conversation:\n${earlier}`
  }
  const last = LEVELS.length - 1
  for (let k = 0; k <= last; k++) {
    // the context goes first; only with all of it shortened does a changed page get less text than it would
    for (const room of k < last ? [room0] : [room0, ...[120, 60, 30, 0].filter(r => r < room0)]) {
      const text = put(k, p.changes(room))
      if (text.length <= limit) return text
    }
  }
  // more changed pages than fit even as bare lines: the list stops early and says how many more changed
  const all = p.changes(0)
  const more = (n: number) => `(${n} more changed page(s) are not listed here for length: they changed too, and a change missing from this list is no ground.)\n`
  let n = all.length
  while (n > 0 && put(last, all.slice(0, n), more(all.length - n)).length > limit) n--
  return put(last, all.slice(0, n), more(all.length - n))
}
