// Shared OpenAI helpers (no SDK dependency — uses global fetch on Vercel Node runtime).
import { AsyncLocalStorage } from 'node:async_hooks'

// What each request cost us, for ChatSlide to charge credits by (Quanlai, 2026-11-10: "charge based on the real cost
// on our side"). $ per 1M tokens [input, cached input, output], read off our OpenAI bill for 2026-11-03..10 and
// rounded up; Claude from its price list. A model not listed is priced as the dearest, never as free.
const PRICES: [string, [number, number, number]][] = [
  ['gpt-6-luna', [0.15, 0.03, 0.6]], ['gpt-5.6-luna', [0.15, 0.03, 1.2]],
  ['gpt-6.1-sol', [2.5, 0.5, 10]], ['gpt-6.1-terra', [2.5, 0.5, 10]], ['gpt-5.6-terra', [2.5, 0.5, 10]],
  ['gpt-6-astra', [2.5, 0.5, 15]], ['gpt-5.5', [5, 0.5, 30]], ['gpt-5.4-mini', [0.4, 0.04, 4.5]], ['gpt-5.4', [2.5, 0.25, 15]],
  ['gpt-4o-mini', [0.15, 0.075, 0.6]], ['gpt-4o', [2.5, 1.25, 10]],
  ['claude-haiku', [1, 0.1, 5]], ['claude-sonnet', [3, 0.3, 15]], ['claude-opus', [5, 0.5, 25]],
]
const meter = new AsyncLocalStorage<{ usd: number }>()
export const addCost = (usd: number) => {
  const m = meter.getStore()
  if (m) m.usd += usd
}
// usage as OpenAI chat (prompt_tokens includes cached), OpenAI responses (input_tokens includes cached) or Anthropic
// (input_tokens excludes cache reads and writes) returns it
export const addUsage = (model: string, u: any) => {
  if (!u) return
  const [pin, pcached, pout] = PRICES.find(([k]) => model.startsWith(k))?.[1] ?? [10, 1, 40]
  const anthropic = 'cache_read_input_tokens' in u || 'cache_creation_input_tokens' in u
  const cached = anthropic ? u.cache_read_input_tokens || 0 : u.prompt_tokens_details?.cached_tokens ?? u.input_tokens_details?.cached_tokens ?? 0
  const fresh = anthropic ? (u.input_tokens || 0) + (u.cache_creation_input_tokens || 0) : (u.prompt_tokens ?? u.input_tokens ?? 0) - cached
  const out = u.completion_tokens ?? u.output_tokens ?? 0
  addCost((fresh * pin + cached * pcached + out * pout) / 1e6)
}
// A handler whose response carries what it cost (header x-ai-cost-usd), read by ChatSlide's credit proxy.
export const metered = (handler: (req: any, res: any) => Promise<unknown>) => (req: any, res: any) => meter.run({ usd: 0 }, () => {
  const store = meter.getStore()!
  for (const k of ['json', 'send', 'end'] as const) {
    const send = res[k]?.bind(res)
    if (send) res[k] = (...a: any[]) => {
      if (!res.headersSent) res.setHeader('x-ai-cost-usd', store.usd.toFixed(6))
      return send(...a)
    }
  }
  return handler(req, res)
})

const OPENAI_URL = 'https://api.openai.com/v1/chat/completions'

export function getKey(): string {
  const key = process.env.OPENAI_API_KEY
  if (!key) throw new Error('OPENAI_API_KEY not set')
  return key
}

interface ChatOpts {
  model?: string
  temperature?: number
  jsonMode?: boolean
  maxTokens?: number
  effort?: 'low' | 'medium' | 'high'
}

// gpt-5/6 models take max_completion_tokens + reasoning_effort and reject a custom temperature.
const params = (opts: ChatOpts, model: string, fallbackTemp: number) => model.startsWith('gpt-5') || model.startsWith('gpt-6')
  ? { max_completion_tokens: opts.maxTokens ?? 16000, reasoning_effort: opts.effort ?? 'low' }
  : { temperature: opts.temperature ?? fallbackTemp, max_tokens: opts.maxTokens ?? 4096 }

// Claude via the Anthropic Messages API: system messages become the (cached) system prompt; JSON mode is asked for in
// the prompt and the reply is cut to its outermost {...}.
async function claude(messages: any[], opts: ChatOpts): Promise<string> {
  const key = process.env.ANTHROPIC_API_KEY
  if (!key) throw new Error('ANTHROPIC_API_KEY not set')
  const system = messages.filter(m => m.role === 'system').map(m => m.content).join('\n\n')
    + (opts.jsonMode ? '\n\nReply with one JSON object only, no prose, no code fences.' : '')
  const turns = messages.filter(m => m.role !== 'system').map(m => ({ role: m.role === 'assistant' ? 'assistant' : 'user', content: String(m.content) }))
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    body: JSON.stringify({
      model: opts.model, max_tokens: Math.min(opts.maxTokens ?? 8192, 32000),
      system: [{ type: 'text', text: system, cache_control: { type: 'ephemeral' } }], messages: turns,
    }),
  })
  if (!res.ok) throw new Error(`Anthropic ${res.status}: ${await res.text()}`)
  const data = await res.json()
  addUsage(String(opts.model), data.usage)
  const text = (data.content || []).filter((c: any) => c.type === 'text').map((c: any) => c.text).join('')
  if (!opts.jsonMode) return text
  const a = text.indexOf('{'), b = text.lastIndexOf('}')
  return a >= 0 && b > a ? text.slice(a, b + 1) : '{}'
}

// Non-streaming chat completion → returns the assistant text.
// Anthropic down or out of credit (2026-11-02: "credit balance is too low" took the whole chat down): the same call on
// OpenAI's cheap model, so the editor degrades instead of failing.
export const FALLBACK_MODEL = process.env.FALLBACK_MODEL || 'gpt-6-luna'
export const vendorDown = (e: unknown) => /Anthropic (400|401|402|403|429|5\d\d)/.test(String((e as Error)?.message || e))

export async function chat(messages: any[], opts: ChatOpts = {}): Promise<string> {
  if (opts.model?.startsWith('claude')) {
    try {
      return await claude(messages, opts)
    }
    catch (e) {
      if (!vendorDown(e)) throw e
      console.error('[llm] Anthropic unavailable, falling back:', String((e as Error).message).slice(0, 160))
      opts = { ...opts, model: FALLBACK_MODEL }
    }
  }
  const res = await fetch(OPENAI_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${getKey()}`,
    },
    body: JSON.stringify({
      model: opts.model || 'gpt-4o',
      ...params(opts, opts.model || 'gpt-4o', 0.4),
      messages,
      ...(opts.jsonMode ? { response_format: { type: 'json_object' } } : {}),
    }),
  })
  if (!res.ok) throw new Error(`OpenAI ${res.status}: ${await res.text()}`)
  const data = await res.json()
  addUsage(opts.model || 'gpt-4o', data.usage)
  return data.choices?.[0]?.message?.content ?? ''
}

// Streaming chat completion → invokes onDelta(textChunk) for each content delta.
export async function chatStream(
  messages: any[],
  onDelta: (text: string) => void,
  opts: ChatOpts = {},
): Promise<void> {
  const res = await fetch(OPENAI_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${getKey()}`,
    },
    body: JSON.stringify({
      model: opts.model || 'gpt-4o',
      ...params(opts, opts.model || 'gpt-4o', 0.6),
      stream: true,
      messages,
    }),
  })
  if (!res.ok || !res.body) throw new Error(`OpenAI ${res.status}: ${await res.text()}`)

  const reader = res.body.getReader()
  const decoder = new TextDecoder('utf-8')
  let buffer = ''
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop() || ''
    for (const line of lines) {
      const trimmed = line.trim()
      if (!trimmed.startsWith('data:')) continue
      const payload = trimmed.slice(5).trim()
      if (payload === '[DONE]') return
      try {
        const json = JSON.parse(payload)
        const delta = json.choices?.[0]?.delta?.content
        if (delta) onDelta(delta)
      }
      catch { /* ignore keep-alive / partial */ }
    }
  }
}

export function readBody(req: any): Promise<any> {
  if (req.body && typeof req.body === 'object') return Promise.resolve(req.body)
  return new Promise((resolve) => {
    let data = ''
    req.on('data', (c: any) => { data += c })
    req.on('end', () => {
      try { resolve(JSON.parse(data || '{}')) }
      catch { resolve({}) }
    })
  })
}
