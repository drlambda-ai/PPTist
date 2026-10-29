<template>
  <MoveablePanel
    class="template-deck-panel"
    :width="340"
    :height="520"
    title="Template deck"
    :left="-360"
    :top="60"
    @close="close()"
  >
    <div class="body">
      <div class="messages" ref="messagesRef">
        <div class="hint" v-if="!sid">
          Upload the company's .pptx template. The deck is built on it, in its own fonts, colours and layouts.
        </div>
        <div class="hint" v-else-if="!messages.length">
          Attach the data, then say what to make, e.g. "用附件里的数据做一份年度复盘，用模板的风格。"
        </div>
        <div class="msg" :class="m.role" v-for="(m, i) in messages" :key="i">{{ m.text }}</div>
        <div class="msg assistant loading" v-if="busy">{{ busy }}</div>
      </div>

      <div class="composer">
        <div class="files">
          <FileInput accept=".pptx" @change="files => uploadTemplate(files[0])">
            <Button :disabled="!!busy">{{ sid ? 'New template' : 'Template .pptx' }}</Button>
          </FileInput>
          <FileInput v-if="sid" accept=".xlsx,.csv,.txt,.md,.docx,.pdf,.pptx,.json" @change="files => uploadData(files[0])">
            <Button :disabled="!!busy">+ Data</Button>
          </FileInput>
        </div>
        <div class="docs" v-if="docs.length">{{ docs.join(', ') }}</div>
        <TextArea v-model:value="input" :rows="3" :disabled="!sid" placeholder="What to make, or what to change" />
        <div class="actions">
          <Button :disabled="!!busy || !version || version.n < 2" @click="undo()">Undo</Button>
          <a v-if="version" :href="version.pptx" :download="`deck-v${version.n}.pptx`"><Button>Download</Button></a>
          <Button type="primary" :disabled="!!busy || !sid || !input.trim()" @click="send()">Send</Button>
        </div>
      </div>
    </div>
  </MoveablePanel>
</template>

<script lang="ts" setup>
import { nextTick, onMounted, ref } from 'vue'
import { useMainStore } from '@/store'
import useImport from '@/hooks/useImport'
import message from '@/utils/message'
import Button from '@/components/Button.vue'
import TextArea from '@/components/TextArea.vue'
import FileInput from '@/components/FileInput.vue'
import MoveablePanel from '@/components/MoveablePanel.vue'

// Template-deck server (chatslide-template-deck): fills the template natively, so the downloaded .pptx is the
// client's own file; the editor shows each version through PPTist's importer.
const API = import.meta.env.VITE_DECK_API || '/deck'
const SID_KEY = 'template-deck-sid'

interface Version { n: number; pptx: string }

const mainStore = useMainStore()
const { importPPTXFile } = useImport()

const sid = ref<string>('')
const messages = ref<{ role: string; text: string }[]>([])
const docs = ref<string[]>([])
const version = ref<Version | null>(null)
const input = ref('')
const busy = ref('')
const messagesRef = ref<HTMLElement>()

const close = () => mainStore.setTemplateDeckPanelState(false)

const scrollBottom = () => nextTick(() => {
  if (messagesRef.value) messagesRef.value.scrollTop = messagesRef.value.scrollHeight
})

const call = async (path: string, body?: BodyInit, headers: Record<string, string> = {}) => {
  const res = await fetch(`${API}${path}`, body === undefined ? undefined : {
    method: 'POST', body, headers: { 'X-Session': sid.value, ...headers },
  })
  const data = await res.json()
  if (data.error) throw new Error(data.error)
  return data
}

const show = async (v: Version | null) => {
  version.value = v
  const blob = await (await fetch(v ? v.pptx : `/data/${sid.value}/template.pptx`)).blob()
  importPPTXFile([new File([blob], 'deck.pptx')], { cover: true })
}

const load = async (id: string) => {
  const st = await call(`/state/${id}`)
  sid.value = id
  messages.value = st.history
  docs.value = st.docs
  await show(st.version)
  scrollBottom()
}

const run = async (label: string, fn: () => Promise<void>) => {
  busy.value = label
  scrollBottom()
  try {
    await fn()
  }
  catch (e) {
    message.error((e as Error).message)
  }
  busy.value = ''
  scrollBottom()
}

const uploadTemplate = (file: File) => run('Learning the template (about a minute)…', async () => {
  const st = await call('/session', file, { 'X-Filename': encodeURIComponent(file.name) })
  localStorage.setItem(SID_KEY, st.sid)
  await load(st.sid)
})

const uploadData = (file: File) => run(`Reading ${file.name}…`, async () => {
  const r = await call('/upload', file, { 'X-Filename': encodeURIComponent(file.name) })
  docs.value = r.docs
})

const send = () => {
  const text = input.value.trim()
  messages.value.push({ role: 'user', text })
  input.value = ''
  return run('Building the deck (about a minute)…', async () => {
    const r = await call('/chat', JSON.stringify({ message: text }), { 'Content-Type': 'application/json' })
    messages.value.push({ role: 'assistant', text: r.reply })
    if (r.version) await show(r.version)
  })
}

const undo = () => run('Reverting…', async () => {
  const st = await call('/undo', '')
  messages.value = st.history
  await show(st.version)
})

onMounted(() => {
  const saved = localStorage.getItem(SID_KEY)
  if (saved) load(saved).catch(() => localStorage.removeItem(SID_KEY))
})
</script>

<style lang="scss" scoped>
.body {
  height: 100%;
  display: flex;
  flex-direction: column;
}
.messages {
  flex: 1;
  overflow-y: auto;
  padding: 4px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.hint {
  font-size: 12px;
  color: #888;
  line-height: 1.6;
}
.msg {
  font-size: 13px;
  line-height: 1.5;
  padding: 6px 10px;
  border-radius: 8px;
  max-width: 85%;
  word-break: break-word;
  white-space: pre-wrap;

  &.user {
    align-self: flex-end;
    background: #d8e6ff;
  }
  &.assistant {
    align-self: flex-start;
    background: #f1f1f1;
  }
  &.loading {
    opacity: 0.6;
  }
}
.composer {
  padding-top: 8px;
  border-top: 1px solid #eee;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.files, .actions {
  display: flex;
  gap: 8px;
}
.actions {
  justify-content: flex-end;
}
.docs {
  font-size: 12px;
  color: #888;
}
</style>
