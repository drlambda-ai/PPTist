// Eval only (LLM_VIA_CODEX=1, imported by scripts/e2e/server.mts, never by api/ or Vercel): every OpenAI request the
// editor's handlers make is answered by Codex on the ChatGPT plan instead, so a replay spends no API credit
// (Quanlai, 2026-11-21, after ~$49 of API spend on replays in one day). The handlers are unchanged: this replaces the
// global fetch for api.openai.com and maps each request to one ~/.agents/tools/codex-exec-once.sh run and back to
// the response shape the caller reads. A request it cannot emulate throws; nothing reaches api.openai.com or
// api.anthropic.com from this process. Log: $CODEX_LOG (one JSON line per call).
import { execFile } from 'node:child_process'
import { createHash } from 'node:crypto'
import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { jsonrepair } from 'jsonrepair'

const ROOT = process.env.CODEX_ROOT || '/tmp/backtest/cx'
const RUN = process.env.LLM_VIA_CODEX_RUN || new Date().toISOString().replace(/\D/g, '').slice(0, 14) // a new server = fresh runs
const WORK = join(ROOT, 'codex-work', RUN) // codex-exec-once keys a run on prompt + args + cwd: same prompt in one run = cached
const OUT = join(ROOT, 'codex-out'), IMG = join(ROOT, 'img')
const LOG = process.env.CODEX_LOG || join(ROOT, 'codex.log')
const ONCE = join(homedir(), '.agents/tools/codex-exec-once.sh')
// Codex is an agent: its own tools stay off, or it answers by calling them (it tried a ChatGPT app's "edit_page" for
// the router's edit_page and came back empty, 2026-11-21). Read-only sandbox besides; the prompt says answer directly.
const BARE = ['apps', 'plugins', 'remote_plugin', 'browser_use', 'browser_use_external', 'computer_use', 'image_generation', 'multi_agent',
  'shell_tool', 'unified_exec', 'goals', 'skill_search', 'tool_suggest', 'hooks', 'in_app_browser', 'code_mode_host', 'sleep_tool'].flatMap(f => ['--disable', f])
// what the ChatGPT plan serves to Codex (probed 2026-11-21: gpt-6.1-sol, gpt-6.1-terra and gpt-6 are refused)
const SERVED = ['gpt-6-luna', 'gpt-6-astra', 'gpt-5.6-luna', 'gpt-5.6-terra']
const NEAREST: Record<string, string> = { 'gpt-6.1-sol': 'gpt-5.6-terra', 'gpt-6.1-terra': 'gpt-5.6-terra' }
const served = (m: string) => (SERVED.includes(m) ? m : NEAREST[m] || 'gpt-6-luna')

const log = (o: Record<string, unknown>) => appendFileSync(LOG, `${JSON.stringify({ t: new Date().toISOString().slice(11, 19), ...o })}\n`)
const sha = (s: string) => createHash('sha256').update(s).digest('hex').slice(0, 20)
const run = (args: string[], timeout: number, force = false) => new Promise<number>(resolve =>
  execFile(ONCE, args, { cwd: WORK, timeout, maxBuffer: 1 << 26, env: { ...process.env, ...(force ? { CODEX_ONCE_FORCE: '1' } : {}) } }, e => resolve(e ? (typeof (e as any).code === 'number' ? (e as any).code : 1) : 0)))

const text = (c: any): string => typeof c === 'string' ? c : Array.isArray(c) ? c.map(text).join('\n') : c?.text ?? (c == null ? '' : JSON.stringify(c))

// image parts become files passed with -i, in order; the prompt says where each one sits
async function images(content: any, files: string[]): Promise<string> {
  if (!Array.isArray(content)) return text(content)
  const parts: string[] = []
  for (const p of content) {
    const url = p?.type === 'image_url' ? p.image_url?.url : p?.type === 'input_image' ? p.image_url : null
    if (!url) { parts.push(text(p)); continue }
    const m = /^data:image\/(\w+);base64,(.*)$/s.exec(url)
    const ext = m ? m[1].replace('jpeg', 'jpg') : (/\.(png|jpe?g|webp|gif)(\?|$)/i.exec(url)?.[1] || 'jpg').toLowerCase()
    const f = join(IMG, `${sha(url)}.${ext}`)
    if (!existsSync(f)) {
      if (m) writeFileSync(f, Buffer.from(m[2], 'base64'))
      else {
        const r = await realFetch(url, { headers: { 'User-Agent': 'ChatSlide/1.0 (contact@chatslide.ai)' } })
        if (!r.ok || !/^image\//.test(r.headers.get('content-type') || '')) { parts.push(`[an image here could not be downloaded: ${r.status}]`); continue }
        writeFileSync(f, Buffer.from(await r.arrayBuffer()))
      }
    }
    files.push(f)
    parts.push(`[attached image #${files.length}]`)
  }
  return parts.join('\n')
}

// chat-completions messages and Responses input items, as one transcript
async function transcript(items: any[], files: string[]): Promise<string> {
  const out: string[] = []
  for (const m of items) {
    if (m?.type === 'function_call') out.push(`=== ASSISTANT (you) called tool ${m.name} with arguments ${m.arguments}`)
    else if (m?.type === 'function_call_output') out.push(`=== TOOL RESULT of that call\n${text(m.output)}`)
    else if (m?.role === 'tool') out.push(`=== TOOL RESULT\n${text(m.content)}`)
    else {
      const calls = (m?.tool_calls || []).map((c: any) => `\n[called tool ${c.function?.name} with arguments ${c.function?.arguments}]`).join('')
      out.push(`=== ${String(m?.role || 'user').toUpperCase()}\n${await images(m?.content, files)}${calls}`)
    }
  }
  return out.join('\n\n')
}

const TOOL_FMT = (tools: any[], required: boolean) => `\n\n=== TOOLS YOU CAN CALL (JSON Schema of each tool's arguments)\n${JSON.stringify(tools)}\n\n`
  + `=== HOW TO ANSWER\nYou are the assistant in the conversation above, answering its last turn. ${required
    ? 'You MUST call exactly one tool. Reply with ONLY {"tool_calls":[{"name":"<tool name>","arguments":{...}}]}'
    : 'Either call exactly one tool, replying with ONLY {"tool_calls":[{"name":"<tool name>","arguments":{...}}]}, or answer in plain text without a tool, replying with ONLY {"content":"<your text>"}'}. `
  + 'The arguments must match that tool\'s schema. One JSON object, no prose, no code fences.'

async function codex(model: string, effort: string, prompt: string, files: string[], kind: string, web = false, needJson = true) {
  const use = served(model)
  const full = `You stand in for the OpenAI API model "${model}" inside an automated test. Do not run any command or read any file${web ? '' : ' or browse'}: answer directly from the text${files.length ? ' and the attached images' : ''} below${web ? ' and the web pages you search' : ''}.\n\n${prompt}`
  if (Buffer.byteLength(full) > 800_000) throw new Error(`codex-eval: prompt of ${Buffer.byteLength(full)} bytes is too long for one argv`)
  const out = join(OUT, `${sha(`${RUN}\n${use}\n${effort}\n${files.join(',')}\n${full}`)}.out`)
  const args = [out, full, ...files.flatMap(f => ['-i', f]), '-m', use, '-c', `model_reasoning_effort="${effort}"`, ...(web ? ['-c', 'web_search="live"'] : []), '-c', 'notify=[]', ...BARE, '--sandbox', 'read-only', '--skip-git-repo-check', '--ephemeral']
  const t0 = Date.now()
  let rc = await run(args, 300_000), retried = false
  while (rc === 4 && Date.now() - t0 < 360_000) { // the same request is in flight from a parallel case: its answer is cached when done
    await new Promise(r => setTimeout(r, 2000))
    rc = await run(args, 300_000)
  }
  // an answer with no JSON where JSON was asked for: run once more, not again (codex-exec-once refuses a bare retry)
  if (needJson && [0, 3].includes(rc) && existsSync(out) && !readFileSync(out, 'utf8').includes('{')) {
    retried = true
    rc = await run(args, 300_000, true)
  }
  const err = existsSync(`${out}.err`) ? readFileSync(`${out}.err`, 'utf8') : ''
  const r429 = /\b429\b|rate.?limit|usage limit|too many requests/i.test(err)
  log({ kind, model, served: use, rc, secs: Math.round((Date.now() - t0) / 1000), images: files.length, chars: full.length, ...(retried ? { retried } : {}), ...(r429 ? { r429 } : {}), ...(use !== model ? { mapped: true } : {}) })
  if (![0, 3].includes(rc) || !existsSync(out)) throw new Error(`codex-eval: codex failed (exit ${rc}${r429 ? ', 429' : ''}), see ${out}.err`)
  return readFileSync(out, 'utf8').trim()
}

const obj = (s: string) => {
  const a = s.indexOf('{'), b = s.lastIndexOf('}')
  if (a < 0 || b <= a) throw new Error(`codex-eval: no JSON in answer: ${s.slice(0, 200)}`)
  const j = s.slice(a, b + 1)
  try { return JSON.parse(j) }
  catch { return JSON.parse(jsonrepair(j)) }
}

// {"tool_calls":[...]} or {"content":...} from Codex; a forced call with no valid tool is an error, never an answer
function toolAnswer(raw: string, tools: { name: string }[], required: boolean) {
  const j = obj(raw)
  const calls = (Array.isArray(j.tool_calls) ? j.tool_calls : []).filter((c: any) => tools.some(t => t.name === c?.name)).slice(0, 1)
  if (!calls.length && (required || typeof j.content !== 'string')) throw new Error(`codex-eval: no valid tool call in answer: ${raw.slice(0, 300)}`)
  return { calls: calls.map((c: any, i: number) => ({ id: `call_cx_${Date.now()}_${i}`, name: c.name, arguments: JSON.stringify(c.arguments ?? {}) })), content: calls.length ? '' : j.content }
}

const json = (o: unknown) => new Response(JSON.stringify(o), { status: 200, headers: { 'Content-Type': 'application/json' } })
const usage = { prompt_tokens: 0, completion_tokens: 0, input_tokens: 0, output_tokens: 0 } // $0: nothing billed per token

async function chatCompletions(b: any) {
  const files: string[] = []
  const effort = b.reasoning_effort || 'low'
  let prompt = await transcript(b.messages || [], files)
  if (b.tools?.length) {
    const tools = b.tools.map((t: any) => t.function)
    const required = b.tool_choice === 'required'
    const a = toolAnswer(await codex(b.model, effort, prompt + TOOL_FMT(tools, required), files, 'chat+tools'), tools, required)
    return json({ choices: [{ message: { role: 'assistant', content: a.content || null, tool_calls: a.calls.length ? a.calls.map(c => ({ id: c.id, type: 'function', function: { name: c.name, arguments: c.arguments } })) : undefined } }], usage })
  }
  const jsonMode = b.response_format?.type === 'json_object'
  prompt += `\n\n=== HOW TO ANSWER\nYou are the assistant in the conversation above: write its reply to the last turn${jsonMode ? ', as ONLY one JSON object, no prose, no code fences' : ', and nothing else'}.`
  let content = await codex(b.model, effort, prompt, files, b.stream ? 'chat-stream' : jsonMode ? 'chat-json' : 'chat', false, jsonMode)
  if (jsonMode) content = JSON.stringify(obj(content))
  if (b.stream) { // one delta, then [DONE]: the shape chatStream reads
    return new Response(`data: ${JSON.stringify({ choices: [{ delta: { content } }] })}\n\ndata: [DONE]\n\n`, { status: 200, headers: { 'Content-Type': 'text/event-stream' } })
  }
  return json({ choices: [{ message: { role: 'assistant', content } }], usage })
}

async function responses(b: any) {
  const files: string[] = []
  const effort = b.reasoning?.effort || 'low'
  const input = typeof b.input === 'string' ? [{ role: 'user', content: b.input }] : b.input || []
  const prompt = await transcript(input, files)
  const fns = (b.tools || []).filter((t: any) => t.type === 'function')
  if (fns.length) {
    const required = b.tool_choice === 'required'
    const a = toolAnswer(await codex(b.model, effort, prompt + TOOL_FMT(fns.map((t: any) => ({ name: t.name, description: t.description, parameters: t.parameters })), required), files, 'responses+tools'), fns, required)
    return json({ output: a.calls.length ? a.calls.map(c => ({ type: 'function_call', call_id: c.id, name: c.name, arguments: c.arguments })) : [{ type: 'message', content: [{ type: 'output_text', text: a.content }] }], usage })
  }
  // research.ts "web": Codex's own live web search; the sources come back as the url_citation annotations it reads
  if ((b.tools || []).some((t: any) => t.type === 'web_search')) {
    const j = obj(await codex(b.model, effort, `${prompt}\n\n=== HOW TO ANSWER\nSearch the web first. Reply with ONLY {"answer":"<your answer>","sources":[{"title":"<page title>","url":"<url of a page you read>"}]}, no prose, no code fences.`, files, 'web_search', true))
    const annotations = (Array.isArray(j.sources) ? j.sources : []).filter((x: any) => x?.url).map((x: any) => ({ type: 'url_citation', url: String(x.url), title: String(x.title || x.url) }))
    return json({ output: [{ type: 'message', content: [{ type: 'output_text', text: String(j.answer || ''), annotations }] }], usage })
  }
  const content = await codex(b.model, effort, `${prompt}\n\n=== HOW TO ANSWER\nWrite the assistant's reply to the last turn, and nothing else.`, files, 'responses', false, false)
  return json({ output: [{ type: 'message', content: [{ type: 'output_text', text: content }] }], usage })
}

let realFetch: typeof fetch
export function install() {
  for (const d of [WORK, OUT, IMG]) mkdirSync(d, { recursive: true })
  // getKey() only checks that a key is set; this one is not a key, and no request carrying it leaves this process
  process.env.OPENAI_API_KEY = 'codex-eval-mode-not-a-key'
  delete process.env.ANTHROPIC_API_KEY
  realFetch = globalThis.fetch
  globalThis.fetch = (async (u: any, init?: any) => {
    const url = String(u?.url ?? u)
    if (/^https:\/\/api\.anthropic\.com\//.test(url)) { log({ kind: 'anthropic', refused: true }); throw new Error('codex-eval: Anthropic is not called in Codex eval mode') }
    if (!/^https:\/\/api\.openai\.com\//.test(url)) return realFetch(u, init)
    const b = JSON.parse(String(init?.body || '{}'))
    if (url.endsWith('/v1/chat/completions')) return chatCompletions(b)
    if (url.endsWith('/v1/responses')) return responses(b)
    log({ kind: url.replace('https://api.openai.com', ''), refused: true })
    throw new Error(`codex-eval: ${url} is not available in Codex eval mode (no image generation)`)
  }) as typeof fetch
  console.log(`LLM_VIA_CODEX: OpenAI requests go to Codex (run ${RUN}, log ${LOG})`)
}
