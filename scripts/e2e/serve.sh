#!/bin/sh
# Click-through tests of the chat in the real editor: builds the editor, serves it with the real AI handlers on
# http://[::1]:3001/harness.html?deck=coffee.json (&paid=1 for a paid user; BREAK_RESEARCH=1 makes the picture
# service return broken JSON, to test the crash path). Drive it with Playwright; the AI calls are logged to
# $E2E_LOG. Needs OPENAI_API_KEY (source ~/.agents/.env.global). Stop: pkill -f scripts/e2e/server.mts
# E2E_APP=<the app.chatslide.ai checkout>: the stand-in also restyles pages with ChatSlide's templates (&restyle=1), by
# the app's own code (components/slides/editorRestyle.ts) bundled for the page.
set -e
cd "$(dirname "$0")/../.."
OUT=${E2E_OUT:-/tmp/pptist-e2e}
mkdir -p "$OUT"
npx vite build --outDir "$OUT/dist" > "$OUT/build.log" 2>&1
cp scripts/e2e/harness.html "$OUT/dist/harness.html"
mkdir -p "$OUT/dist/decks" && cp scripts/e2e/decks/*.json "$OUT/dist/decks/"
if [ -n "$E2E_APP" ]; then
  (cd "$E2E_APP" && node_modules/.bin/esbuild scripts/deck-eval/restyle-host.ts --bundle --format=iife --platform=browser \
    --outfile="$OUT/dist/app-restyle.js" --tsconfig=tsconfig.json --jsx=automatic --log-level=error --define:global=globalThis \
    '--define:process.env.NODE_ENV="production"' '--banner:js=var process = { env: {} };' --loader:.css=empty --loader:.scss=empty \
    --loader:.svg=dataurl --loader:.png=dataurl --loader:.webp=dataurl --loader:.jpg=dataurl --loader:.woff2=empty)
fi
: > "$OUT/calls.log"
E2E_DIST="$OUT/dist" E2E_LOG="$OUT/calls.log" nohup npx tsx scripts/e2e/server.mts > "$OUT/server.log" 2>&1 &
~/.agents/tools/wait-until.sh 60 2 grep -q ready "$OUT/server.log"
echo "ready: http://[::1]:3001/harness.html?deck=coffee.json  (log: $OUT/calls.log)"
