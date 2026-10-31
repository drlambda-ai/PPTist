<template>
  <MoveablePanel
    class="template-deck-panel"
    :width="360"
    :height="560"
    title="AI"
    :left="-380"
    :top="60"
    @close="close()"
  >
    <div class="body">
      <div class="messages" ref="messagesRef">
        <div class="hint" v-if="!s.history.length">
          Describe the deck you want. Optional: upload your company's .pptx template (the deck keeps its exact look)
          and the content to use (xlsx, docx, pdf, md...). Once it exists, ask for any change: add, delete, move or
          rewrite pages.
        </div>
        <div class="msg" :class="m.role" v-for="(m, i) in s.history" :key="i">{{ m.text }}</div>
        <div class="msg assistant loading" v-if="busy">{{ busy }}</div>
      </div>

      <div class="composer">
        <div class="chips" v-if="s.template || s.docs.length">
          <span class="chip" v-if="s.template">Template: {{ s.template.name }}
            <i @click="removeTemplate()" v-if="!busy && !s.mode">×</i></span>
          <span class="chip" v-for="(d, i) in s.docs" :key="d.name + i">{{ d.name }}
            <i @click="s.docs.splice(i, 1); save()" v-if="!busy">×</i></span>
        </div>
        <TextArea v-model:value="input" :rows="3" placeholder="e.g. 用附件里的数据做一份年度复盘 / delete page 3 / add a page on Q3 risks" />
        <div class="actions">
          <FileInput accept=".pptx" @change="files => uploadTemplate(files[0])" v-if="!s.mode">
            <Button :disabled="!!busy">+ Template</Button>
          </FileInput>
          <FileInput accept=".xlsx,.csv,.txt,.md,.docx,.pdf,.pptx,.json" @change="files => uploadContent(files[0])">
            <Button :disabled="!!busy">+ Content</Button>
          </FileInput>
          <span class="grow"></span>
          <Button v-if="canUndo" :disabled="!!busy" @click="undo()">Undo</Button>
          <Button v-if="s.mode" :disabled="!!busy" @click="download()">Download</Button>
          <Button type="primary" :disabled="!!busy || !input.trim()" @click="send()">Send</Button>
        </div>
        <div class="new" v-if="s.mode && !busy"><span @click="restart()">Start a new deck</span></div>
      </div>
    </div>
  </MoveablePanel>
</template>

<script lang="ts" setup>
import { computed, nextTick, onMounted, reactive, ref } from 'vue'
import { storeToRefs } from 'pinia'
import { nanoid } from 'nanoid'
import { useMainStore, useSlidesStore } from '@/store'
import useImport from '@/hooks/useImport'
import useAIPPT from '@/hooks/useAIPPT'
import useSlideHandler from '@/hooks/useSlideHandler'
import useHistorySnapshot from '@/hooks/useHistorySnapshot'
import api from '@/services'
import type { Slide } from '@/types/slides'
import type { AIPPTSlide } from '@/types/AIPPT'
import { jsonrepair } from 'jsonrepair'
import { upload } from '@/utils/blobUpload'
import Button from '@/components/Button.vue'
import TextArea from '@/components/TextArea.vue'
import FileInput from '@/components/FileInput.vue'
import MoveablePanel from '@/components/MoveablePanel.vue'

// One AI panel for both paths:
// - with a company template: the deck server (api/deck.py) fills the template natively, every version is a real
//   .pptx in the template; the editor shows it through the .pptx importer;
// - without one: PPTist's own AI deck generation. Edits after that go through a ChatSlide-style tool router
//   (api/tools/deck_chat.ts) applied to the editor's slides; slide rewrites use chat_edit.

interface Session {
  mode: '' | 'template' | 'native'
  template: { name: string, templateUrl: string, analysis: any, profile: any, chrome: any[] } | null
  docs: { name: string, text: string }[]
  history: { role: 'user' | 'assistant', text: string }[]
  versions: { spec: any, deckUrl: string, h: number }[]
}

const KEY = 'ai-deck-session'
const blank = (): Session => ({ mode: '', template: null, docs: [], history: [], versions: [] })
const loadSession = (): Session => {
  try {
    return { ...blank(), ...JSON.parse(localStorage.getItem(KEY) || '{}') } 
  }
  catch {
    return blank() 
  }
}

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const { slides, slideIndex, viewportSize, viewportRatio } = storeToRefs(slidesStore)
const { importPPTXFile } = useImport()
const { AIPPT, getMdContent } = useAIPPT()
const { resetSlides } = useSlideHandler()
const { addHistorySnapshot, undo: undoEditor } = useHistorySnapshot()

const s = reactive<Session>(loadSession())
const input = ref('')
const busy = ref('')
const messagesRef = ref<HTMLElement>()
let blobMode = ''

const SLIDES_KEY = 'ai-deck-slides'
const save = () => {
  try {
    localStorage.setItem(KEY, JSON.stringify(s))
    // PPTist keeps no slides across reloads; a generated deck is kept here (a template deck reloads from its .pptx)
    if (s.mode === 'native') localStorage.setItem(SLIDES_KEY, JSON.stringify(slides.value))
  }
  catch { /* storage full: the session still works until reload */ }
}
const say = (role: 'user' | 'assistant', text: string) => {
  s.history.push({ role, text })
  save()
  nextTick(() => {
    if (messagesRef.value) messagesRef.value.scrollTop = messagesRef.value.scrollHeight 
  })
}
const close = () => mainStore.setTemplateDeckPanelState(false)
const canUndo = computed(() => s.mode === 'template' ? s.versions.length > 1 : s.mode === 'native')

const run = async (label: string, fn: () => Promise<void>) => {
  busy.value = label
  try {
    await fn()
  }
  catch (e) {
    say('assistant', 'That failed: ' + ((e as Error).message || e))
  }
  busy.value = ''
}

const deckApi = async (action: string, body: any) => {
  const r = await fetch('/api/deck', { method: 'POST', body: JSON.stringify({ action, ...body }) })
  const data = await r.json()
  if (data.error) throw new Error(data.error)
  return data
}

// Big files go straight to blob storage (function bodies are capped at 4.5MB on Vercel).
const uploadFile = async (file: File): Promise<string> => {
  if (!blobMode) blobMode = (await (await fetch('/api/deck_config')).json()).blob
  if (blobMode === 'vercel') {
    return (await upload(`deck/${file.name}`, file, { access: 'public', handleUploadUrl: '/api/deck_upload' })).url
  }
  const r = await fetch('/api/deck_put', { method: 'POST', body: file, headers: { 'X-Filename': encodeURIComponent(file.name) } })
  return (await r.json()).url
}

const showPptx = async (url: string) => {
  const blob = await (await fetch(url)).blob()
  importPPTXFile([new File([blob], 'deck.pptx')], { cover: true })
}

const uploadTemplate = (file: File) => run('Learning the template (about a minute)…', async () => {
  const templateUrl = await uploadFile(file)
  const learned = await deckApi('learn', { templateUrl })
  s.template = { name: file.name, ...learned }
  save()
  await showPptx(templateUrl)
  say('assistant', `Template ready: ${learned.analysis.length} layouts. The deck will be built on it.`)
})

const removeTemplate = () => {
  s.template = null
  save()
}

const uploadContent = (file: File) => run(`Reading ${file.name}…`, async () => {
  const r = await deckApi('read', { url: await uploadFile(file), name: file.name })
  s.docs.push({ name: r.name, text: r.text })
  save()
})

// ---- template path: the deck server builds every version ----
const templateTurn = async (text: string) => {
  const t = s.template!
  const last = s.versions[s.versions.length - 1]
  const r = await deckApi('turn', {
    message: text,
    session: { ...t, docs: s.docs, history: s.history.slice(0, -1), spec: last?.spec || null },
  })
  say('assistant', r.reply)
  if (r.deckUrl) {
    s.versions.push({ spec: r.spec, deckUrl: r.deckUrl, h: s.history.length - 2 })
    s.mode = 'template'
    save()
    await showPptx(r.deckUrl)
  }
}

// ---- native path: PPTist's own generation, then routed edits ----
const readStream = async (res: any, onText: (t: string) => void) => {
  const reader: ReadableStreamDefaultReader = res.body.getReader()
  const decoder = new TextDecoder('utf-8')
  for (;;) {
    const { done, value } = await reader.read()
    if (done) return
    onText(decoder.decode(value, { stream: true }))
  }
}

const material = () => s.docs.map(d => `--- ${d.name} ---\n${d.text}`).join('\n\n')
const languageOf = (text: string) => (/[\u4e00-\u9fff]/.test(text) ? '中文' : 'English')

const generateNative = async (text: string) => {
  const brief = s.docs.length ? `${text}\n\nUse only this material (never invent numbers):\n${material()}` : text
  let outline = ''
  await readStream(await api.AIPPT_Outline({ content: brief, language: languageOf(text), model: '' }), t => {
    outline += t 
  })
  outline = getMdContent(outline).replace(/<!--[\s\S]*?-->/g, '')
  const template = await api.getMockData('template_1')
  resetSlides()
  let pending = ''
  const flush = (line: string) => {
    const t = line.replace('```jsonl', '').replace('```json', '').replace('```', '').trim()
    if (!t) return
    try {
      AIPPT(template.slides, [JSON.parse(jsonrepair(t)) as AIPPTSlide]) 
    }
    catch { /* a partial line; the stream completes it */ }
  }
  await readStream(await api.AIPPT({ content: `User request: ${text}\n\n${outline}`, language: languageOf(text), style: '通用', model: '' }), (chunk) => {
    const lines = (pending + chunk).split(/\n+/)
    pending = lines.pop() || ''
    lines.forEach(flush)
  })
  flush(pending)
  slidesStore.setTheme(template.theme)
  s.mode = 'native'
  say('assistant', languageOf(text) === '中文' ? `已生成 ${slides.value.length} 页。可以继续要求增加、删除、移动或改写页面。` : `Generated ${slides.value.length} pages. Ask for any change: add, delete, move or rewrite pages.`)
}

const pageText = (sl: Slide) => sl.elements.map((e: any) => (e.content || e.text?.content || '').replace(/<[^>]+>/g, ' '))
  .join(' ').replace(/\s+/g, ' ').trim()
const slideSize = () => ({ width: viewportSize.value, height: Math.round(viewportSize.value * viewportRatio.value) })

// Rewrite one slide through chat_edit, merging kept elements by id (same contract as the slide chat panel).
const editSlide = async (index: number, instruction: string) => {
  const slide = slides.value[index]
  const ret = await api.AI_Edit({ elements: slide.elements, instruction, slideSize: slideSize() })
  if (!ret.elements) return ret.reply
  const byId: Record<string, any> = Object.fromEntries(slide.elements.map(e => [e.id, e]))
  slidesStore.updateSlide({ elements: ret.elements.map((e: any) => (byId[e.id] ? { ...byId[e.id], ...e } : e)) }, slide.id)
  return ret.reply
}

const cloneSlide = (sl: Slide): Slide => ({
  ...JSON.parse(JSON.stringify(sl)), id: nanoid(10),
  elements: sl.elements.map(e => ({ ...JSON.parse(JSON.stringify(e)), id: nanoid(10) })),
})

const nativeTurn = async (text: string) => {
  const { action, args } = await (await fetch('/api/tools/deck_chat', {
    method: 'POST',
    body: JSON.stringify({
      instruction: text, pages: slides.value.map(pageText), current: slideIndex.value + 1,
      history: s.history.slice(-7, -1).map(h => ({ role: h.role, content: h.text })),
    }),
    headers: { 'Content-Type': 'application/json' },
  })).json()
  const n = slides.value.length
  const at = (p: number) => Math.min(Math.max((p || 1) - 1, 0), n - 1)
  let reply = ''
  if (action === 'add_page' && s.mode === 'native') {
    // same generator as the deck, so the layout matches the item count and the content follows the same rules
    const ask = `One single content slide for an existing deck (no cover, no closing slide, no contents slide). Deck so far: ${slides.value.map(pageText).join(' | ').slice(0, 1500)}. The new slide: ${args.instruction}`
    const brief = s.docs.length ? `${ask}\n\nUse only this material (never invent numbers):\n${material()}` : ask
    let outline = ''
    await readStream(await api.AIPPT_Outline({ content: brief, language: languageOf(text), model: '' }), t => {
      outline += t
    })
    let jsonl = ''
    await readStream(await api.AIPPT({ content: `User request: ${ask}\n\n${getMdContent(outline)}`, language: languageOf(text), style: '通用', model: '' }), t => {
      jsonl += t
    })
    const item = jsonl.split('\n').map(l => {
      try {
        return JSON.parse(jsonrepair(l.trim())) as AIPPTSlide
      }
      catch {
        return null
      }
    }).find(x => x?.type === 'content')
    if (!item) reply = languageOf(text) === '中文' ? '没能生成这一页，请换个说法再试。' : 'Could not write that page, please rephrase.'
    else {
      const template = await api.getMockData('template_1')
      AIPPT(template.slides, [item]) // appends at the end (2 slides if it splits 5+ items); move them into place
      const list = [...slides.value]
      list.splice(at(args.after) + 1, 0, ...list.splice(n))
      slidesStore.setSlides(list)
      slidesStore.updateSlideIndex(at(args.after) + 1)
      reply = `Added page ${at(args.after) + 2}.`
    }
  }
  else if (action === 'add_page') {
    const copy = cloneSlide(slides.value[at(args.layout_from || args.after)])
    const list = [...slides.value]
    list.splice(at(args.after) + 1, 0, copy)
    slidesStore.setSlides(list)
    await editSlide(at(args.after) + 1, `This is a NEW page. Replace all its text with: ${args.instruction}. Keep the layout and styling. Write in the language of that text.`)
    slidesStore.updateSlideIndex(at(args.after) + 1)
    reply = `Added page ${at(args.after) + 2}.`
  }
  else if (action === 'delete_page') {
    const ids = (args.pages || []).map((p: number) => slides.value[at(p)]?.id).filter(Boolean)
    if (ids.length >= n) reply = 'A deck needs at least one page, so I kept it.'
    else {
      slidesStore.deleteSlide(ids)
      reply = `Deleted ${ids.length} page(s).`
    }
  }
  else if (action === 'duplicate_page') {
    const list = [...slides.value]
    list.splice(at(args.index) + 1, 0, cloneSlide(slides.value[at(args.index)]))
    slidesStore.setSlides(list)
    reply = `Duplicated page ${at(args.index) + 1}.`
  }
  else if (action === 'move_page') {
    const list = [...slides.value]
    const [moved] = list.splice(at(args.from), 1)
    list.splice(at(args.to), 0, moved)
    slidesStore.setSlides(list)
    reply = `Moved page ${at(args.from) + 1} to position ${at(args.to) + 1}.`
  }
  else if (action === 'update_slide') {
    reply = await editSlide(at(args.index), args.instruction)
    slidesStore.updateSlideIndex(at(args.index))
  }
  else if (action === 'update_all_pages') {
    for (let i = 0; i < n; i++) {
      busy.value = `Updating page ${i + 1} of ${n}…`
      await editSlide(i, args.instruction)
    }
    reply = `Updated all ${n} pages.`
  }
  else reply = args.answer || 'Could you say that more specifically?'
  if (action !== 'answer_question') addHistorySnapshot()
  say('assistant', reply + (args.followup_note ? `\n(Not done yet: ${args.followup_note})` : ''))
}

const send = () => {
  const text = input.value.trim()
  input.value = ''
  say('user', text)
  if (s.template && s.mode !== 'native') return run('Building the deck (about a minute)…', () => templateTurn(text))
  if (!s.mode) return run('Writing the deck…', () => generateNative(text))
  const before = JSON.parse(JSON.stringify(slides.value))
  return run('Working…', async () => {
    try {
      await nativeTurn(text)
    }
    catch (e) { // never leave a half-applied operation (e.g. an inserted page that was never rewritten)
      slidesStore.setSlides(before)
      throw e
    }
  })
}

const undo = () => run('Reverting…', async () => {
  if (s.mode === 'native') {
    undoEditor()
    return nextTick(save)
  }
  const undone = s.versions.pop()!
  s.history = s.history.slice(0, undone.h)
  say('assistant', `Reverted to version ${s.versions.length}.`)
  await showPptx(s.versions[s.versions.length - 1].deckUrl)
})

const download = () => {
  if (s.mode === 'template') {
    const a = document.createElement('a')
    a.href = s.versions[s.versions.length - 1].deckUrl
    a.download = 'deck.pptx'
    a.click()
  }
  else mainStore.setDialogForExport('pptx')
}

const restart = () => {
  Object.assign(s, blank())
  localStorage.removeItem(SLIDES_KEY)
  save()
}

// Bring the deck back after a page reload.
onMounted(() => {
  if (s.mode === 'template' && s.versions.length) {
    run('Loading your deck…', () => showPptx(s.versions[s.versions.length - 1].deckUrl))
  }
  else if (s.mode === 'native') {
    try {
      const saved = JSON.parse(localStorage.getItem(SLIDES_KEY) || 'null')
      if (saved?.length) slidesStore.setSlides(saved)
    }
    catch { /* nothing saved */ }
  }
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
.actions {
  display: flex;
  gap: 6px;
  align-items: center;
  white-space: nowrap;
}
.grow {
  flex: 1;
}
.chips {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
}
.chip {
  font-size: 12px;
  background: #f1f1f1;
  border-radius: 10px;
  padding: 2px 8px;

  i {
    cursor: pointer;
    margin-left: 4px;
    font-style: normal;
  }
}
.new {
  font-size: 12px;
  color: #888;
  text-align: right;

  span {
    cursor: pointer;
    text-decoration: underline;
  }
}
</style>
