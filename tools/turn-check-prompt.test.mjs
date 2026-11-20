// The check of a finished chat message must see every changed page under ask_json's 30,000-character cap.
// Run: node --test tools/turn-check-prompt.test.mjs
import { test } from 'node:test'
import { Buffer } from 'node:buffer'
import assert from 'node:assert/strict'
import { build } from 'esbuild'
const built = await build({ entryPoints: ['src/views/Editor/turnCheckPrompt.ts'], bundle: true, write: false, format: 'esm', platform: 'node' })
const { fitCheck, CHECK_LIMIT } = await import(`data:text/javascript;base64,${Buffer.from(built.outputFiles[0].text).toString('base64')}`)

const cut = (t, n) => (t.length > n ? `${t.slice(0, n)}… ${n < 120 ? '(cut)' : '(cut here for length: the rest is there, not shown)'}` : t)
// the parts as the panel builds them: `pages` changed pages of long text, a pending request of `request` characters
const parts = (pages, request, stepChars = 60) => ({
  head: `${'R'.repeat(3300)}\nThe user's message (they were looking at page 1 of ${pages}): "Add after the current presentation"\n`,
  reply: `Number of pages updated: ${pages}.`,
  request: max => `The message answers the editor's question "Replace the deck or add after it?" about the user's request ${JSON.stringify(cut('q'.repeat(request), max))}: that request is what was asked.\n`,
  earlier: userMax => `user: ${'q'.repeat(request).slice(0, userMax)}\nassistant: Replace the deck or add after it?`,
  deck: per => Array.from({ length: pages }, (_, k) => `page ${k + 1}: ${cut('d'.repeat(900), Math.max(60, Math.min(per, Math.floor(6000 / pages))))}`),
  steps: max => cut(Array.from({ length: pages }, (_, k) => `edit_page {"index":${k + 1}} -> ${'s'.repeat(stepChars)}`).join('\n'), max),
  changes: room => Array.from({ length: pages }, (_, k) => `page ${k + 1} CHANGED\n  before: ${cut('b'.repeat(900), room)} [1 picture(s)]\n  after: ${cut('a'.repeat(900), room)} [1 picture(s)]`),
})
const room0 = pages => Math.max(120, Math.min(700, Math.floor(6000 / pages)))

test('21 changed pages after a 6,000-character request: every page is in the record, under the cap, and the cut is said', () => {
  const p = parts(21, 6000)
  // the record as 926c1d0b sent it: past the cap, so ask_json dropped the last changed pages
  const bare = `${p.head}${p.request(6000)}${p.deck(500).join('\n')}\n${p.steps(5000)}\n${p.reply}\n${p.changes(room0(21)).join('\n')}\n${p.earlier(6000)}`
  assert.ok(bare.length > 30000, `the case is past the cap unfitted (${bare.length})`)
  const text = fitCheck(p, room0(21))
  assert.ok(text.length <= CHECK_LIMIT && CHECK_LIMIT <= 30000, `fits (${text.length})`)
  for (let k = 1; k <= 21; k++) assert.ok(text.includes(`page ${k} CHANGED\n  before: b`) && text.includes('  after: a'), `page ${k} is listed with its text`)
  assert.equal(text.split('\n  after: a').length - 1, 21)
  assert.ok(text.includes('shortened or left out') && text.includes('is no ground'), 'the cut is said')
  assert.ok(text.includes('Number of pages updated: 21.') && text.startsWith(p.head), 'the rules, the message and the reply are whole')
  assert.ok(text.includes(`b… (cut here for length`), 'a changed page keeps the text it would have had: the context was cut first')
})

test('a short record is sent as it is, with no note', () => {
  const text = fitCheck(parts(2, 200), room0(2))
  assert.ok(!text.includes('shortened or left out') && !text.includes('more changed page'))
  assert.equal(text.split('\n  after: a').length - 1, 2)
})

test('120 changed pages: all listed with less text each', () => {
  const text = fitCheck(parts(120, 6000), room0(120))
  assert.ok(text.length <= CHECK_LIMIT)
  assert.equal(text.split(' CHANGED\n').length - 1, 120)
  assert.ok(text.includes('shortened or left out'))
})

test('more changed pages than fit as bare lines: the list says how many more changed', () => {
  const text = fitCheck(parts(600, 6000), room0(600))
  assert.ok(text.length <= CHECK_LIMIT)
  const listed = text.split(' CHANGED\n').length - 1
  assert.ok(listed > 0 && listed < 600)
  assert.ok(text.includes(`(${600 - listed} more changed page(s) are not listed here for length: they changed too, and a change missing from this list is no ground.)`))
})
