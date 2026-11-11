# PPTist (the ChatSlide advanced editor)

Pushing `master` deploys production (Vercel project `pptist`, Enception team), served at
`app.chatslide.ai/api/pptist/`. Use `/usr/bin/git`; commit directly to `master`.

## Every problem a user hits in the editor chat gets fixed

Quanlai, 2026-11-11: "make sure that they are always fixed unless something is genuinely out of
our capacity." An unhappy turn found in the sessions is a work order, not a finding: fix it in
code, test it in the real editor, review, push. Report as not fixable only what our tools
genuinely cannot do (a picture no source holds), with the reason in one line.

1. Read the turns: `scripts/editor-turns.sql` (prod, read-only). A turn is unhappy when the user
   repeats or rephrases, complains, or the reply says nothing changed.
2. Reproduce and verify in the browser: `scripts/e2e/serve.sh`, then Playwright on
   `http://[::1]:3001/harness.html?deck=<file>` (`&paid=1` for a paid user). A user's deck is
   rebuilt only in `/tmp`, never committed.
3. New UI strings go into all 12 `src/locales/*.json`.
4. `npx vue-tsc --noEmit -p tsconfig.app.json`, one commit per fix, a note on
   `refs/notes/review`, then push.

The daily job `pptist-editor-sessions` (cron-skills) runs this; `pptist-push` pushes reviewed commits.
