// Shared OpenAI helpers (no SDK dependency — uses global fetch on Vercel Node runtime).

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
  const text = ((await res.json()).content || []).filter((c: any) => c.type === 'text').map((c: any) => c.text).join('')
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
