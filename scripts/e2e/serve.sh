#!/bin/sh
# Click-through tests of the chat in the real editor: builds the editor, serves it with the real AI handlers on
# http://[::1]:3001/harness.html?deck=coffee.json (&paid=1 for a paid user; BREAK_RESEARCH=1 makes the picture
# service return broken JSON, to test the crash path). Drive it with Playwright; the AI calls are logged to
# $E2E_LOG. Needs OPENAI_API_KEY (source ~/.agents/.env.global). Stop: pkill -f scripts/e2e/server.mts
set -e
cd "$(dirname "$0")/../.."
OUT=${E2E_OUT:-/tmp/pptist-e2e}
mkdir -p "$OUT"
npx vite build --outDir "$OUT/dist" > "$OUT/build.log" 2>&1
cp scripts/e2e/harness.html "$OUT/dist/harness.html"
mkdir -p "$OUT/dist/decks" && cp scripts/e2e/decks/*.json "$OUT/dist/decks/"
: > "$OUT/calls.log"
E2E_DIST="$OUT/dist" E2E_LOG="$OUT/calls.log" nohup npx tsx scripts/e2e/server.mts > "$OUT/server.log" 2>&1 &
~/.agents/tools/wait-until.sh 60 2 grep -q ready "$OUT/server.log"
echo "ready: http://[::1]:3001/harness.html?deck=coffee.json  (log: $OUT/calls.log)"
