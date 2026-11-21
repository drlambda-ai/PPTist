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
3. Test only with real users' input (Quanlai, 2026-11-20: "We need to only test with real users'
   input and figure out why"). Before a push, replay on the old and the new build, each on the
   user's own deck as saved at send time (`make_backtest.py --accept <accept.py output> --ok <n>`):
   the real turns users rejected (did the fix help, and why does each one still fail) and real
   turns users accepted (a fix that breaks one is not done). Read why each failing turn fails;
   a grade alone is not a reason.
   Replays spend the OpenAI key; run them with `LLM_VIA_CODEX=1` and the keys unset (`env -u OPENAI_API_KEY
   -u ANTHROPIC_API_KEY`, see `scripts/e2e/codex-llm.mts`) whenever the question is which tools a fix runs: $0, about
   4x slower, first tool as faithful as the API, multi-step sequences less so. `LLM_CACHE_DIR=<dir>` on the server answers a request
   it has seen before from disk (`scripts/e2e/llm-cache.mts`): an unchanged path replays exactly at $0; use a fresh dir
   per run when runs must be independent. API replays, for pass rates only,
   share one $5 a day for this work (Quanlai, 2026-11-21).
   Which tool the router picks is a backend question: post the real turn's request straight to
   `/api/tools/deck_chat` (Codex mode), no browser. The browser harness only for the full turn or a screenshot, one
   run at a time, `-j 2` at most: parallel runs once held 124 headless browsers, 7.4 GB (2026-11-21, "please kill
   all of them").
4. Success is what real users do, on their real decks (Quanlai, 2026-11-19: "Why don't you just
   use whatever the user actually uses?", then "focus on all of the new prompts recently, real
   prompts. Don't use fabricated content"). Only recent real prompts on the user's saved deck: no
   generated decks, invented messages or synthetic cases in any measurement or regression check. The success rate is the share of production turns the
   user accepted, judged from their next message in the same deck (`accept.py` in app.chatslide.ai
   `scripts/deck-eval`). A change to the router loop, its prompt or a model is backtested on
   users' own decks as they were at send time: `make_backtest.py` (its docstring has the
   commands), the same file on the old and the new build, graded alike, at least 3 runs each
   (one run swings about 10 points). State the table in the commit.
5. New UI strings go into all 12 `src/locales/*.json` ("deck" is a presentation, never a deck of cards).
6. `npx vue-tsc --noEmit -p tsconfig.app.json`, one commit per fix, a note on
   `refs/notes/review`, then push.

The daily job `pptist-editor-sessions` (cron-skills) runs this; `pptist-push` pushes reviewed commits.

### Name the cause of each unhappy turn before fixing it

Quanlai, 2026-11-18. Every unhappy turn gets exactly one cause, and the cause picks the fix:

1. **No credits.** The user cannot pay for the step. Fix: the paywall opens, never an error or a silent no-op.
2. **No capability.** Our tools cannot do it. Fix: build the tool, or `say_unable` in one line.
3. **Capability exists, the AI does not know it.** The router never picked the tool. Fix: the tool's description or the router prompt.
4. **Capability exists, the AI picked it and did it wrong.** Fix: the tool's own code or its edit prompt.
5. **The AI did not have what the user submitted.** An uploaded file, an attached image, an earlier message, the deck's current state or the selected page never reached the model. Fix: pass that context into the turn.
6. **The search could not get what was asked.** A picture of a named person or thing, a paper, a page, a figure in a file came back wrong or empty. Fix: the source or the query for that kind of thing.
7. **Guessed where it should have asked, or asked where it should have acted.** Fix: the router rule for that request.
8. **A step errored.** Fix: the code that threw.
9. **Done, but the user did not see it.** Another page, too small a change, a reply that does not say where. Fix: show the changed page and say what changed.

`python3 ~/proj/app.chatslide.ai/scripts/deck-eval/sessions.py --hours 720` labels every unhappy turn with its cause
(Codex, cached) and prints the count per cause, sorted by cause. Turns it marks `not_a_failure` (the next message was
just the next request) are left out of the count. Report the count per cause alongside the fixes.

## How the chat turn works (keep it this way)

- The router (`api/tools/deck_chat.ts`) runs a real tool loop: the steps already run go back to the
  model as its own tool calls with their results, and after a step it may answer in plain text,
  which ends the message. Never send them as a text list with a forced tool call: measured on 100
  replayed conversations, that proposed a step it had just run in 34 of them (now 1-2), with the
  same pass rate.
- Destructive steps (delete several pages, restore, a new deck) change nothing until the user
  confirms, and every save is a version: ChatSlide keeps the newest 20 copies and one per hour for
  7 days, and the chat offers them (`version_history`).
