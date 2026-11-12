# PPTist (the ChatSlide advanced editor)

Pushing `master` deploys production (Vercel project `pptist`, Enception team), served at
`app.chatslide.ai/api/pptist/`. Use `/usr/bin/git`; commit directly to `master`.

## Every problem a user hits in the editor chat gets fixed

Quanlai, 2026-11-11: "make sure that they are always fixed unless something is genuinely out of
our capacity." An unhappy turn found in the sessions is a work order, not a finding: fix it in
code, test it in the real editor, review, push. Report as not fixable only what our tools
genuinely cannot do (a picture no source holds), with the reason in one line.

1. Read the turns: `scripts/editor-turns.sql` (prod, read-only). A turn is unhappy when the user
   repeats or rephrases, complains, or the reply says nothing changed. Run the query alone in its
   call, with one `-f` and no `<` anywhere in the command: a second input makes the DB guard open
   the approval window (unattended, that is a denial).
2. Reproduce and verify in the browser: `scripts/e2e/serve.sh`, then Playwright on
   `http://[::1]:3001/harness.html?deck=<file>` (`&paid=1` for a paid user). A user's deck is
   rebuilt only in `/tmp`, never committed.
3. Every fixed failure becomes a regression case, and the cases run before the push: add a row
   to `~/proj/app.chatslide.ai/scripts/deck-eval/editor-cases.jsonl` (a synthetic fixture, a
   measured `expect_*` where one exists), then with `serve.sh` up:
   `PPTIST_EVAL_URL='http://[::1]:3001/' python3 scripts/deck-eval/run.py --file editor-cases.jsonl --no-grade -j 4`
   (22 cases, about 3 minutes, cents). A fix that fails another case is not done. Fifty patches
   in two days went out checked by hand only, and two of them broke earlier fixes (2026-11-12).
4. A change to the router loop, its prompt or a model is measured on the 100 replayed real
   conversations (`--file data/luna100-cases.jsonl`, then `regrade.py`): pass rate, cases with a
   repeated action, seconds. State the table in the commit.
5. New UI strings go into all 12 `src/locales/*.json` ("deck" is a presentation, never a deck of cards).
6. `npx vue-tsc --noEmit -p tsconfig.app.json`, one commit per fix, a note on
   `refs/notes/review`, then push.

The daily job `pptist-editor-sessions` (cron-skills) runs this; `pptist-push` pushes reviewed commits.

## How the chat turn works (keep it this way)

- The router (`api/tools/deck_chat.ts`) runs a real tool loop: the steps already run go back to the
  model as its own tool calls with their results, and after a step it may answer in plain text,
  which ends the message. Never send them as a text list with a forced tool call: measured on 100
  replayed conversations, that proposed a step it had just run in 34 of them (now 1-2), with the
  same pass rate.
- Destructive steps (delete several pages, restore, a new deck) change nothing until the user
  confirms, and every save is a version: ChatSlide keeps the newest 20 copies and one per hour for
  7 days, and the chat offers them (`version_history`).
