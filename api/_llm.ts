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
}

// Non-streaming chat completion → returns the assistant text.
export async function chat(messages: any[], opts: ChatOpts = {}): Promise<string> {
  const res = await fetch(OPENAI_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${getKey()}`,
    },
    body: JSON.stringify({
      model: opts.model || 'gpt-4o',
      temperature: opts.temperature ?? 0.4,
      max_tokens: opts.maxTokens ?? 4096,
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
      temperature: opts.temperature ?? 0.6,
      max_tokens: opts.maxTokens ?? 4096,
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
