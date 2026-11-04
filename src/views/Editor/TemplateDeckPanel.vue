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
          Ask for any change to this deck, or describe a new one. 📎 attaches your company's .pptx template or your
          material (xlsx, docx, pdf, md…).
        </div>
        <div class="msg" :class="m.role" v-for="(m, i) in s.history" :key="i">{{ m.text }}</div>
        <div class="msg assistant loading" v-if="busy">{{ busy }}</div>
        <div class="choices" v-if="!busy && s.pending?.options?.length">
          <span class="choice" v-for="o in s.pending.options" :key="o" @click="input = o; send()">{{ o }}</span>
        </div>
      </div>

      <div class="composer">
        <div class="chips" v-if="s.template || s.docs.length || s.images?.length">
          <span class="chip" v-if="s.template">Template: {{ s.template.name }}
            <i @click="removeTemplate()" v-if="!busy && !s.mode">×</i></span>
          <span class="chip" v-for="(m, i) in s.images || []" :key="'img' + m.name + i">🖼 {{ m.name }}
            <i @click="s.images!.splice(i, 1); save()" v-if="!busy">×</i></span>
          <span class="chip" v-for="(d, i) in s.docs" :key="d.name + i">{{ d.name }}
            <i @click="s.docs.splice(i, 1); save()" v-if="!busy">×</i></span>
        </div>
        <div class="input-box" @dragover.prevent @drop.prevent="e => attach(e.dataTransfer?.files)">
          <TextArea v-model:value="input" :rows="3" placeholder="e.g. 用附件里的数据做一份年度复盘 / delete page 3 / add a page on Q3 risks" @enter="onEnter" />
          <div class="input-row">
            <FileInput accept=".pptx,.xlsx,.csv,.txt,.md,.docx,.pdf,.json,.png,.jpg,.jpeg,.webp,.gif" @change="files => attach(files)">
              <span class="attach" :class="{ disabled: !!busy }" aria-label="Attach a company template (.pptx) or material"><i-icon-park-outline:link /></span>
            </FileInput>
            <span class="grow"></span>
            <Button type="primary" size="small" :disabled="!!busy || !input.trim()" @click="send()">Send</Button>
          </div>
        </div>
        <div class="new" v-if="s.mode && !busy"><span @click="restart()">Start a new deck</span></div>
      </div>
    </div>
  </MoveablePanel>
  <!-- ChatSlide: pages being turned into pictures (off-screen, outside the panel so a closed panel still renders) -->
  <Teleport to="body">
    <div class="chatslide-render" ref="renderRef" v-if="toRender.length">
      <ThumbnailSlide v-for="sl in toRender" :key="sl.id" :slide="sl" :size="1600" />
    </div>
  </Teleport>
</template>

<script lang="ts" setup>
import { nextTick, onMounted, reactive, ref, watch } from 'vue'
import { host as chatslide, OutOfCredits, takeSpend } from '@/embed'
import { storeToRefs } from 'pinia'
import { nanoid } from 'nanoid'
import { useMainStore, useSlidesStore, useSnapshotStore } from '@/store'
import useExport from '@/hooks/useExport'
import useImport from '@/hooks/useImport'
import useAIPPT from '@/hooks/useAIPPT'
import useSlideHandler from '@/hooks/useSlideHandler'
import useHistorySnapshot from '@/hooks/useHistorySnapshot'
import api from '@/services'
import type { Slide, PPTImageElement } from '@/types/slides'
import type { AIPPTSlide } from '@/types/AIPPT'
import { jsonrepair } from 'jsonrepair'
import { upload } from '@/utils/blobUpload'
import Button from '@/components/Button.vue'
import TextArea from '@/components/TextArea.vue'
import FileInput from '@/components/FileInput.vue'
import ThumbnailSlide from '@/views/components/ThumbnailSlide/index.vue'
import { toJpeg } from 'html-to-image'
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
  images?: { name: string, url: string }[] // pictures the user attached, placed on slides by add_image source "attached"
  history: { role: 'user' | 'assistant', text: string }[]
  versions: { spec: any, deckUrl: string, h: number }[]
  edited?: boolean // template deck changed in the editor since the last build
  synced?: { keys: string[], urls: Record<string, string> } // ChatSlide: pages last sent to the project, and their pictures
  pending?: { request: string, question: string, options?: string[] } | null // an ask_user waiting for the user's answer
}

const KEY = 'ai-deck-session'
const blank = (): Session => ({ mode: '', template: null, docs: [], history: [], versions: [] })
const loadSession = (): Session => {
  if (chatslide) return blank() // inside ChatSlide the session comes with the project (onMounted)
  try {
    return { ...blank(), ...JSON.parse(localStorage.getItem(KEY) || '{}') } 
  }
  catch {
    return blank() 
  }
}

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const snapshotStore = useSnapshotStore() // undo/redo are async (IndexedDB): await them before saving
const { slides, slideIndex, viewportSize, viewportRatio, theme } = storeToRefs(slidesStore)
const { importPPTXFile } = useImport()
const { exportPPTX } = useExport()
const { AIPPT, getMdContent } = useAIPPT()
const { resetSlides } = useSlideHandler()
const { addHistorySnapshot } = useHistorySnapshot()

const s = reactive<Session>(loadSession())
const input = ref('')
const busy = ref('')
const messagesRef = ref<HTMLElement>()
let blobMode = ''

const SLIDES_KEY = 'ai-deck-slides'
let saveTimer = 0
let syncTimer = 0
let saveFailed = false
const save = () => {
  if (chatslide) { // per project, debounced; also after manual edits (watch below)
    clearTimeout(saveTimer)
    clearTimeout(syncTimer)
    syncTimer = window.setTimeout(() => syncToChatSlide().catch(() => undefined), 15000)
    saveTimer = window.setTimeout(async () => {
      await externalizeImages().catch(() => undefined) // pictures saved as files: the deck file stays small and loads fast
      chatslide!.save({ session: JSON.parse(JSON.stringify(s)), slides: slides.value, theme: theme.value })
        .then(() => saveFailed = false)
        .catch(() => { // said once (saying saves again): the next edit retries
          if (!saveFailed) say('assistant', 'Saving failed; your last edits may not be kept. Check your connection.')
          saveFailed = true
        })
    }, 2000)
    return
  }
  try {
    localStorage.setItem(KEY, JSON.stringify(s))
    // PPTist keeps no slides across reloads; a generated deck is kept here (a template deck reloads from its .pptx)
    if (s.mode === 'native' || s.edited) {
      localStorage.setItem(`${SLIDES_KEY}-theme`, JSON.stringify(theme.value)) // tiny, first: chart/table colours follow it after a reload
      localStorage.setItem(SLIDES_KEY, JSON.stringify(slides.value))
    }
  }
  catch { // storage full: the session still works until reload; never leave an older saved deck behind
    localStorage.removeItem(SLIDES_KEY)
  }
}
// Inline pictures (data: URLs from AI images, attachments, pasted images) become their own files, uploaded once; the
// slide keeps the link. A 10-page deck with a few AI pictures was 6 MB and took ~4 s to open; links keep it small.
const uploaded = new Map<string, string>() // data URL -> file link, this session
const externalizeImages = async () => {
  if (!chatslide?.uploadImage) return
  const swap = async (src: string) => {
    if (!src?.startsWith('data:image/') || src.length < 20000) return src
    if (!uploaded.has(src)) uploaded.set(src, await chatslide!.uploadImage!(await (await fetch(src)).blob()))
    return uploaded.get(src)!
  }
  for (const sl of slides.value) {
    for (const e of sl.elements as any[]) {
      if (e.type === 'image' && e.src?.startsWith('data:image/')) {
        const url = await swap(e.src)
        if (url !== e.src) slidesStore.updateElement({ id: e.id, props: { src: url }, slideId: sl.id })
      }
    }
    const bg = sl.background as any
    if (bg?.type === 'image' && bg.image?.src?.startsWith('data:image/')) {
      const url = await swap(bg.image.src)
      if (url !== bg.image.src) slidesStore.updateSlide({ background: { ...bg, image: { ...bg.image, src: url } } }, sl.id)
    }
  }
}

const say = (role: 'user' | 'assistant', text: string) => {
  s.history.push({ role, text })
  save()
  nextTick(() => {
    if (messagesRef.value) messagesRef.value.scrollTop = messagesRef.value.scrollHeight 
  })
}
const close = () => mainStore.setTemplateDeckPanelState(false)

const run = async (label: string, fn: () => Promise<void>) => {
  busy.value = label
  try {
    await fn()
  }
  catch (e) {
    if (e instanceof OutOfCredits) say('assistant', zh(userText) ? `积分不够了（还需 ${e.required}，剩 ${e.credits}）。升级方案后可以继续。` : `You're out of credits (this needs ${e.required}, you have ${e.credits}). Upgrade your plan to keep going.`)
    else if (/\b(Anthropic|OpenAI) \d{3}\b/.test(String((e as Error).message))) say('assistant', zh(userText) ? 'AI 服务暂时不可用，这条没有做任何修改。请稍后再试。' : 'The AI service is unavailable right now; nothing was changed. Please try again in a moment.')
    else say('assistant', 'That failed: ' + ((e as Error).message || e))
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

// The user's own picture: uploaded once, then placed by the chat ("put my photo on slide 3")
const uploadPicture = (file: File) => run(`Uploading ${file.name}…`, async () => {
  // re-encoded through a canvas: only raster pixels are published (never SVG/HTML or a spoofed type), at most 2000 px
  const bmp = await createImageBitmap(file).catch(() => null)
  if (!bmp) fail(zh(userText) ? `「${file.name}」不是可用的图片。` : `${file.name} is not a picture I can use.`)
  const k = Math.min(1, 2000 / Math.max(bmp.width, bmp.height))
  const c = document.createElement('canvas')
  c.width = Math.round(bmp.width * k)
  c.height = Math.round(bmp.height * k)
  c.getContext('2d')!.drawImage(bmp, 0, 0, c.width, c.height)
  const blob: Blob = await new Promise((ok, no) => c.toBlob(b => (b ? ok(b) : no(new Error('encode failed'))), 'image/png'))
  // a unique name: two attachments with the same file name stay separately addressable
  const taken = new Set((s.images || []).map(m => m.name))
  let name = file.name.replace(/\.[^.]+$/, '') + '.png'
  for (let n = 2; taken.has(name); n++) name = `${file.name.replace(/\.[^.]+$/, '')} (${n}).png`
  s.images = [...(s.images || []), { name, url: await uploadFile(new File([blob], name, { type: 'image/png' })) }]
  save()
})

const docFiles = new Map<string, File>() // this session's attached documents, for pulling their figures out
const figuresOf = new Map<string, { url: string, page?: number }[]>() // extracted once per file
const uploadContent = (file: File) => run(`Reading ${file.name}…`, async () => {
  const r = await deckApi('read', { url: await uploadFile(file), name: file.name })
  s.docs.push({ name: r.name, text: r.text })
  docFiles.set(r.name, file)
  figuresOf.delete(r.name) // a new file under the same name: its own figures, not the old file's
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
    s.edited = false
    localStorage.removeItem(SLIDES_KEY)
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

const material = () => s.docs.map(d => `--- ${d.name} ---\n${d.text.slice(0, Math.floor(80000 / s.docs.length))}`).join('\n\n') // read_file reaches the rest
const languageOf = (text: string) => (/[\u4e00-\u9fff]/.test(text) ? '中文' : 'English') // for the panel's own short replies
// What the generators write in: a new deck follows the request's language; pages added later follow the DECK's language
// (a Chinese deck stays Chinese when someone types an English or Russian request), unless the request names a language.
const requestLanguage = (text: string) => `the language of this request: "${text.slice(0, 120)}"`
const deckLanguage = (text: string) => `the language of the existing deck (sample: "${deckLines().join(' ').slice(0, 160)}"), unless this request names another language: "${text.slice(0, 120)}"`

const generateNative = async (text: string) => {
  const brief = s.docs.length ? `${text}\n\nUse only this material (never invent numbers):\n${material()}` : text
  let outline = ''
  await readStream(await api.AIPPT_Outline({ content: brief, language: requestLanguage(text), model: '' }), t => {
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
  await readStream(await api.AIPPT({ content: `User request: ${text}\n\n${outline}`, language: requestLanguage(text), style: '通用', model: '' }), (chunk) => {
    const lines = (pending + chunk).split(/\n+/)
    pending = lines.pop() || ''
    lines.forEach(flush)
  })
  flush(pending)
  slidesStore.setTheme(template.theme)
  s.mode = 'native'
  say('assistant', languageOf(text) === '中文' ? `已生成 ${slides.value.length} 页。可以继续要求增加、删除、移动或改写页面。` : `Generated ${slides.value.length} pages. Ask for any change: add, delete, move or rewrite pages.`)
}

// One line per page for the router: text in reading order (top to bottom), numbered images/charts so "image 2" and
// "the chart" resolve. Text repeated on most pages (footer, deck name) is left out so pages are told apart by content.
const elText = (e: any) => (e.content || e.text?.content || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()
const readingOrder = (sl: Slide) => [...sl.elements].sort((x, y) => (x.top - y.top) || (x.left - y.left))
const imgSpot = (e: any) => `${Math.round(e.left / 10)},${Math.round(e.top / 10)},${Math.round(e.width / 10)}`
const logoSpots = () => { // an image at the same spot on most pages is a logo, never "the picture"
  const n: Record<string, number> = {}
  const { width: W, height: H } = slideSize()
  const small = (e: any) => e.width * e.height < W * H * 0.04 // a repeated big photo is content, not a logo
  for (const sl of slides.value) for (const k of new Set(sl.elements.filter(e => e.type === 'image' && small(e)).map(imgSpot))) n[k] = (n[k] || 0) + 1
  return new Set(Object.keys(n).filter(k => slides.value.length > 2 && n[k] > slides.value.length / 2))
}
const pageText = (sl: Slide, common = new Set<string>(), logos = new Set<string>()) => {
  let img = 0
  const biggest = Math.max(0, ...sl.elements.filter(e => e.type === 'image' && !logos.has(imgSpot(e))).map((e: any) => e.width * e.height))
  return readingOrder(sl).map((e: any) => {
    if (e.type === 'image') { // size and place let the router tell a photo from an icon or a logo
      const { width: W, height: H } = slideSize()
      const area = (e.width * e.height) / (W * H)
      if (logos.has(imgSpot(e))) return `[image ${++img}: logo]`
      const kind = `${area < 0.005 ? 'icon' : area < 0.04 ? 'small image' : 'large image'}${e.width * e.height === biggest ? ', the largest image here' : ''}`
      const where = `${e.top + e.height / 2 < H / 3 ? 'top' : e.top + e.height / 2 > (2 * H) / 3 ? 'bottom' : 'middle'}-${e.left + e.width / 2 < W / 3 ? 'left' : e.left + e.width / 2 > (2 * W) / 3 ? 'right' : 'centre'}`
      return `[image ${++img}: ${kind}, ${Math.round(area * 100)}% of slide, ${where}]`
    }
    if (e.type === 'chart') return '[chart]'
    const t = elText(e)
    return common.has(t) ? '' : t
  }).join(' ').replace(/\s+/g, ' ').trim()
}
const deckLines = () => {
  const count: Record<string, number> = {}
  for (const sl of slides.value) for (const t of new Set(sl.elements.map(elText))) if (t) count[t] = (count[t] || 0) + 1
  const common = new Set(Object.keys(count).filter(t => slides.value.length > 2 && count[t] > slides.value.length / 2))
  const logos = logoSpots()
  return slides.value.map(sl => pageText(sl, common, logos))
}
const slideSize = () => ({ width: viewportSize.value, height: Math.round(viewportSize.value * viewportRatio.value) })

// Rewrite one slide through chat_edit, merging kept elements by id. Image sources (often MBs of base64) never go to
// the model: it sees a short token and the real src is restored on merge.
let userText = '' // the message being handled: replies are written in its language
let lastEditTerse = '' // what code knows about the last slide edit; shown instead of the model's sentence in multi-step turns
const editedThisTurn = new Set<string>() // slide ids changed by an earlier step of the current turn
const touched = (ids: string[]) => ids.forEach(id => editedThisTurn.add(id))
const picturedThisTurn = new Set<string>() // pages that got a NEW picture in the current message
const editSlide = async (index: number, instruction: string) => {
  const slide = slides.value[index]
  // reading order so "the first point" means the top-left one; the merge below keeps the slide's own stacking order
  // a shape's words (every text box of an imported ChatSlide/PowerPoint deck) are shown like a text box's: the model
  // edits "content" the same way on both, and the merge puts it back inside the shape's text
  const flat = (e: any) => (e.type === 'shape' && e.text ? (({ text, ...r }) => ({ ...r, content: text.content, defaultColor: text.defaultColor, defaultFontName: text.defaultFontName }))(e) : e)
  // the model copies what it sees, so a new box often comes back as a "shape" with "content": put the words where
  // the editor draws them (a shape's text, or a plain text box when it has no outline)
  const unflat = (e: any) => {
    if (e.type !== 'shape' || e.content === undefined) return e
    const { content, defaultColor, defaultFontName, ...rest } = e
    if (!rest.path) return { ...rest, type: 'text', content, defaultColor: defaultColor || '#333', defaultFontName: defaultFontName || '', rotate: rest.rotate || 0 }
    return { ...rest, text: { content, defaultColor: defaultColor || '#333', defaultFontName: defaultFontName || '', align: 'middle' } }
  }
  const elements = readingOrder(slide).map((e: any) => (e.type === 'image' ? { ...e, src: `image:${e.id}` }
    : e.type === 'shape' && e.path?.length > 80 ? { ...flat(e), path: `path:${e.id}` } : flat(e))) // long SVG paths: token, restored below
  const r = await fetch('/api/tools/chat_edit', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ patch: true, elements, slideSize: slideSize(), instruction: `${instruction}\n(Write the "reply" sentence in the language of this message: "${userText.slice(0, 120)}")` }),
  })
  const ret = await r.json()
  if (!ret.patch) fail(ret.reply || ret.error || 'That edit could not be applied.')
  const unchanged = () => new NoChange(editedThisTurn.has(slide.id)
    ? (zh(userText) ? `第 ${index + 1} 页没有进一步的改动。` : `No further change on page ${index + 1}.`)
    : (zh(userText) ? `第 ${index + 1} 页没有改动：这个修改我没能做成。` : `Page ${index + 1} was not changed: I could not make that edit.`))
  if (!ret.patch.changes?.length && !ret.patch.added?.length && !ret.patch.removed?.length) throw unchanged() // say so, don't echo the request
  const live = slides.value.find(x => x.id === slide.id) || slide // merged onto the slide as it is now
  const changes: Record<string, any> = Object.fromEntries((ret.patch.changes || []).filter((c: any) => c?.id).map((c: any) => [c.id, c]))
  const removed = new Set<string>(ret.patch.removed || [])
  const seen = new Set<string>()
  const merged = [
    ...live.elements.filter(e => !removed.has(e.id)).map((e: any) => {
      const c = changes[e.id]
      if (!c) return e
      const { src: _src, path, ...rest } = c // image sources and tokenised paths are never taken from the model
      void _src
      if (e.type === 'shape' && e.text) {
        const { content, defaultColor, defaultFontName, text, ...shape } = rest
        const t = { ...e.text, ...(text || {}), ...(content !== undefined ? { content } : {}), ...(defaultColor ? { defaultColor } : {}), ...(defaultFontName ? { defaultFontName } : {}) }
        return { ...e, ...shape, text: t, ...(path && !String(path).startsWith('path:') ? { path } : {}) }
      }
      return { ...e, ...rest, ...(path && !String(path).startsWith('path:') ? { path } : {}) }
    }),
    ...(ret.patch.added || []).filter((e: any) => e?.id && e.type !== 'image' && !live.elements.some(x => x.id === e.id) && !seen.has(e.id) && seen.add(e.id)).map(unflat),
  ]
  // a reply that describes changes the merge did not make is never shown as done
  if (JSON.stringify(merged) === JSON.stringify(live.elements)) throw unchanged()
  const before: Record<string, number> = Object.fromEntries(live.elements.filter((e: any) => e.type === 'text').map((e: any) => [e.id, e.height]))
  slidesStore.updateSlide({ elements: merged }, slide.id)
  await fitText(slide.id, Object.keys(changes).filter(id => before[id]), before)
  // what actually landed: elements that differ, plus those removed (not the raw patch, which may name ids the merge ignored)
  const liveById = new Map(live.elements.map((e: any) => [e.id, JSON.stringify(e)]))
  const n = merged.filter((e: any) => liveById.get(e.id) !== JSON.stringify(e)).length + live.elements.filter((e: any) => !merged.some((m: any) => m.id === e.id)).length
  void n
  editedThisTurn.add(slide.id)
  lastEditTerse = zh(userText) ? `已修改第 ${index + 1} 页。` : `Updated page ${index + 1}.`
  return lastEditTerse // the model's own sentence misstated the edit too often (audited 2026-11-02): code reports
}

// A text box grows to fit its content when its page is shown. After an edit, show the page and shrink the font sizes
// of any changed box that grew past its old height (translations run longer), so it never spills onto other elements.
let fitQueue: Promise<void> = Promise.resolve() // one page shown at a time, even when pages are edited in parallel
const fitText = (slideId: string, ids: string[], before: Record<string, number>) => {
  fitQueue = fitQueue.then(() => fitOne(slideId, ids, before)).catch(() => undefined)
  return fitQueue
}
const fitOne = async (slideId: string, ids: string[], before: Record<string, number>) => {
  if (!ids.length) return
  const at = slides.value.findIndex(x => x.id === slideId)
  if (at < 0) return
  slidesStore.updateSlideIndex(at)
  for (let round = 0; round < 3; round++) {
    await nextTick()
    await new Promise(r => setTimeout(r, 350)) // the editor measures and stores the new heights
    const sl = slides.value.find(x => x.id === slideId)
    let changed = false
    for (const e of (sl?.elements || []) as any[]) {
      if (!ids.includes(e.id) || e.type !== 'text' || e.height <= before[e.id] * 1.06 + 4) continue
      const k = Math.max(0.6, before[e.id] / e.height)
      const scale = (v: number) => Math.max(9, Math.round(v * k * 10) / 10)
      let content = String(e.content).replace(/font-size:\s*([\d.]+)px/g, (_m: string, v: string) => `font-size: ${scale(Number(v))}px`)
      if (content === e.content) { // no inline sizes: scale the size the editor actually renders, set on each paragraph
        const node = document.querySelector(`#editable-element-${e.id} .ProseMirror, #editable-element-${e.id} .text`)
        const px = node ? parseFloat(getComputedStyle(node).fontSize) : 0
        if (!px) continue
        content = content.replace(/<p(\s[^>]*)?>/g, (_m: string, attrs = '') => (/style="/.test(attrs)
          ? `<p${attrs.replace(/style="/, `style="font-size: ${scale(px)}px; `)}>` : `<p${attrs} style="font-size: ${scale(px)}px;">`))
        if (content === e.content) continue
      }
      slidesStore.updateElement({ id: e.id, props: { content }, slideId })
      changed = true
    }
    if (!changed) return
  }
}

const cloneSlide = (sl: Slide): Slide => {
  const newId: Record<string, string> = Object.fromEntries(sl.elements.map(e => [e.id, nanoid(10)]))
  return {
    ...JSON.parse(JSON.stringify(sl)), id: nanoid(10),
    elements: sl.elements.map(e => ({ ...JSON.parse(JSON.stringify(e)), id: newId[e.id] })),
    animations: (sl.animations || []).filter(a => newId[a.elId]).map(a => ({ ...a, id: nanoid(10), elId: newId[a.elId] })),
  }
}

const zh = (text: string) => languageOf(text) === '中文'
// A step that cannot be done stops the turn (the router would otherwise retry it with new arguments).
class StepFailed extends Error {}
class NoChange extends StepFailed {} // the page already is what was asked (or the edit changed nothing visible)
function fail(msg: string): never {
  throw new StepFailed(msg)
}

// A new content page from the same outline + JSONL generator as a prompt-built deck (layout matches item count).
const generatePage = async (text: string, instruction: string, at: number, count = 1) => {
  const ask = `Exactly ${count} content slide(s) for an existing deck (no cover, no closing slide, no contents slide). Deck so far: ${deckLines().join(' | ').slice(0, 1500)}. The new slide(s): ${instruction}`
  const brief = s.docs.length ? `${ask}\n\nUse only this material (never invent numbers):\n${material()}` : ask
  let outline = ''
  await readStream(await api.AIPPT_Outline({ content: brief, language: deckLanguage(text), model: '' }), t => {
    outline += t
  })
  let jsonl = ''
  await readStream(await api.AIPPT({ content: `User request: ${ask}\n\n${getMdContent(outline)}`, language: deckLanguage(text), style: '通用', model: '' }), t => {
    jsonl += t
  })
  const items = jsonl.split('\n').map(l => {
    try {
      return JSON.parse(jsonrepair(l.trim())) as AIPPTSlide
    }
    catch {
      return null
    }
  }).filter((x): x is AIPPTSlide => x?.type === 'content').slice(0, count)
  if (!items.length) return false
  const n = slides.value.length
  AIPPT((await api.getMockData('template_1')).slides, items) // appends at the end (an item with 5+ points may split in 2)
  const list = [...slides.value]
  list.splice(at, 0, ...list.splice(n))
  slidesStore.setSlides(list)
  return true
}

const CHART = `Title and labels in the deck's language. Return JSON {"title": "...", "labels": ["..."], "legends": ["series name"], "series": [[numbers per label] per legend], "note": "one line on the source, or empty"}.
Copy every number exactly, each under its own year/label. Use only numbers the user, the deck or the material give; if there are none, return {"error": "<one sentence asking for the numbers>"}.`

const addChartPage = async (text: string, instruction: string, chartType: string, at: number) => {
  const res = await fetch('/api/tools/ask_json', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ strong: true, instruction: `${CHART}\n\nChart: ${instruction}\n\nDeck: ${deckLines().join(' | ').slice(0, 3000)}\n\nMaterial: ${material().slice(0, 12000)}` }),
  })
  const d = await res.json()
  if (d.error || !d.labels?.length) fail(d.error || (zh(text) ? '缺少图表数据。' : 'No numbers to chart.'))
  const { width: W, height: H } = slideSize()
  insertPage(at, d.title, d.note, {
    type: 'chart', id: nanoid(10), chartType, left: W * 0.08, top: H * 0.2, width: W * 0.84, height: H * (d.note ? 0.66 : 0.72),
    rotate: 0, themeColors: theme.value.themeColors, textColor: theme.value.fontColor,
    data: { labels: d.labels, legends: d.legends || [''], series: d.series },
  })
  return ''
}

// A new page in the deck's look: a copy of a content page with its title, footer and logo kept and its body replaced by
// one main element (chart, table). Footer = elements in the bottom 12% or repeated on most pages.
const insertPage = (at: number, heading: string, note: string, main: any) => {
  const { width: W, height: H } = slideSize()
  const base = slides.value[Math.min(Math.max(1, at - 1), slides.value.length - 1)] // a content page, not the cover
  const count: Record<string, number> = {}
  for (const sl of slides.value) for (const t of new Set(sl.elements.map(elText))) if (t) count[t] = (count[t] || 0) + 1
  // footer = what repeats across pages (same text, or the same spot on most pages: logo, page number), never a page's own data
  const spot = (e: any) => `${e.type},${Math.round(e.left / 10)},${Math.round(e.top / 10)},${Math.round((e.width || 0) / 10)}`
  const spots: Record<string, number> = {}
  for (const sl of slides.value) for (const k of new Set(sl.elements.map(spot))) spots[k] = (spots[k] || 0) + 1
  const isFooter = (e: any) => (elText(e) && count[elText(e)] > slides.value.length / 2)
    || (e.top > H * 0.85 && (spots[spot(e)] > slides.value.length / 2 || /^\d{1,3}$/.test(elText(e))))
  const texts = base.elements.filter((e: any) => (e.type === 'text' || e.type === 'shape') && elText(e) && !isFooter(e))
  const title: any = texts.find((e: any) => e.textType === 'title') || [...texts].sort((x, y) => x.top - y.top)[0]
  // header band: anything that starts above the title's bottom edge (+4%), e.g. an icon box or the rule under the title,
  // except big pictures that belong to the old body
  const band = title ? title.top + title.height + H * 0.04 : 0
  const keep = base.elements.filter((e: any) => isFooter(e) || e === title
    || (e.top <= band && !(e.type === 'text' && e !== title) && (e.width || 0) * (e.height || 0) < W * H * 0.15))
    .map(e => ({ ...JSON.parse(JSON.stringify(e)), id: nanoid(10) }))
  const t = keep.find((e: any) => title && e.left === title.left && e.top === title.top && elText(e) === elText(title)) as any
  const esc = (x: string) => String(x).replace(/</g, '&lt;')
  if (t) { // the deck's own title box and styling, new words
    const html = t.type === 'shape' ? t.text.content : t.content
    const styled = html.replace(/(<[^/][^>]*>)([^<]+)(<)/, (_m: string, open: string, _w: string, close: string) => `${open}${esc(heading)}${close}`)
    if (t.type === 'shape') t.text.content = styled
    else t.content = styled
  }
  const top = t ? t.top + t.height + H * 0.03 : H * 0.18
  const bottom = Math.min(H * 0.86, ...keep.filter(isFooter).map((e: any) => e.top - H * 0.02))
  const body = { ...main, left: W * 0.06, top, width: W * 0.88, height: Math.max(bottom - top - (note ? H * 0.06 : 0), H * 0.3) }
  if (body.type === 'table') body.height = Math.min(body.height, body.data.length * 40)
  // the colour the title actually shows: its inline CSS colour (a fixed grammar: #hex or rgb()), else its default colour
  const titleHtml = t ? (t.type === 'shape' ? t.text.content : t.content) : ''
  const accent = titleHtml.match(/color:\s*(#[0-9a-fA-F]{6}|rgba?\([^)]*\))/)?.[1] || (t?.defaultColor || t?.text?.defaultColor || '')
  // a grey or black title is no brand colour: then the deck theme's first colour leads
  const isGrey = (c: string) => { // #rrggbb or rgb(r, g, b): channels within 24 of each other
    const v = c.startsWith('#') ? (c.length === 4 ? [1, 2, 3].map(k => parseInt(c[k] + c[k], 16)) : [1, 3, 5].map(k => parseInt(c.slice(k, k + 2), 16)))
      : (c.match(/\d+/g) || []).slice(0, 3).map(Number)
    return v.length === 3 && Math.max(...v) - Math.min(...v) < 24
  }
  const lead = accent && !isGrey(accent) ? accent : (theme.value.themeColors.find(c => !isGrey(c)) || theme.value.themeColors[0])
  if (body.type === 'chart') body.themeColors = [lead, ...theme.value.themeColors.filter(c => c !== lead)]
  if (body.type === 'table') body.theme = { ...body.theme, color: lead }
  const slide: Slide = {
    ...JSON.parse(JSON.stringify(base)), id: nanoid(10), animations: [],
    elements: [
      ...(t ? keep : [{
        type: 'text', id: nanoid(10), left: W * 0.06, top: H * 0.06, width: W * 0.88, height: H * 0.1, rotate: 0, textType: 'title',
        content: `<p style="font-size: 28px;"><strong>${esc(heading)}</strong></p>`, defaultFontName: theme.value.fontName, defaultColor: theme.value.fontColor,
      } as any, ...keep]),
      body,
      ...(note ? [{
        type: 'text', id: nanoid(10), left: W * 0.06, top: body.top + body.height + H * 0.01, width: W * 0.88, height: H * 0.05, rotate: 0,
        content: `<p style="font-size: 11px;">${esc(note)}</p>`, defaultFontName: theme.value.fontName, defaultColor: theme.value.fontColor,
      } as any] : []),
    ],
  }
  const list = [...slides.value]
  list.splice(at, 0, slide)
  slidesStore.setSlides(list)
  return lead
}

const addTablePage = async (text: string, instruction: string, at: number) => {
  const res = await fetch('/api/tools/ask_json', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ strong: true, instruction: `Title and cells in the deck's language. Return JSON {"title": "...", "rows": [["header 1", "header 2"], ["cell", "cell"]], "note": "one line on the source, or empty"}. At most 8 rows and 6 columns, short cells. Copy every number exactly from the deck or material and keep each value under its own year/column (check direction: which year is higher). Use only facts the user, the deck or the material give; if there are none, return {"error": "<one sentence asking for them>"}.\n\nTable: ${instruction}\n\nDeck: ${deckLines().join(' | ').slice(0, 3000)}\n\nMaterial: ${material().slice(0, 12000)}` }),
  })
  const d = await res.json()
  if (d.error || !d.rows?.length) fail(d.error || (zh(text) ? '缺少表格内容。' : 'Nothing to put in a table.'))
  const { width: W, height: H } = slideSize()
  const cols = Math.max(...d.rows.map((r: any[]) => r.length))
  const style = { fontname: theme.value.fontName, color: theme.value.fontColor, fontsize: '14px' }
  insertPage(at, d.title, d.note, {
    type: 'table', id: nanoid(10), left: W * 0.08, top: H * 0.22, width: W * 0.84, height: d.rows.length * 40, rotate: 0,
    colWidths: new Array(cols).fill(1 / cols), cellMinHeight: 36,
    outline: { width: 1, style: 'solid', color: '#cccccc' },
    theme: { color: theme.value.themeColors[0], rowHeader: true, rowFooter: false, colHeader: false, colFooter: false },
    data: d.rows.map((r: any[], i: number) => Array.from({ length: cols }, (_, j) => ({
      id: nanoid(10), colspan: 1, rowspan: 1, text: String(r[j] ?? ''),
      style: i === 0 ? { ...style, color: '#ffffff', bold: true } : style, // header row sits on the accent colour
    }))),
  })
}

// "Extend page N": new points that go deeper than page N (strong model, grounded in page N, the deck and the material),
// laid out as cards on a continuation page with the deck's header, footer and colours. Page N is left as it is.
const expandPage = async (text: string, index: number, instruction: string) => {
  const sl = slides.value[index]
  const res = await fetch('/api/tools/ask_json', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ strong: true, instruction: `The user wants page ${index + 1} extended: ${instruction || 'more depth'}.
Page ${index + 1} says: ${pageText(sl)}
Write a continuation page in the page's language: a title (page ${index + 1}'s title + a short qualifier) and 3-4 points, each a short bold heading and one or two sentences of NEW content: what the numbers mean, causes, comparisons, implications, next steps. Never repeat a statement or number layout already on the page; use a number only when it is on the page or in the material, exactly as given; never invent figures.
Return JSON {"title": "...", "points": [{"heading": "...", "text": "..."}]}.
Deck: ${deckLines().join(' | ').slice(0, 2500)}
Material: ${material().slice(0, 12000)}` }),
  })
  const d = await res.json()
  if (!d.points?.length) fail(zh(text) ? '没能写出扩展内容。' : 'Could not write the extension.')
  const { width: W, height: H } = slideSize()
  const k = Math.min(d.points.length, 4)
  const cols = k === 4 ? 2 : k, rows = k === 4 ? 2 : 1
  const esc = (x: string) => String(x).replace(/\*\*/g, '').replace(/</g, '&lt;') // markdown bold markers never show
  const color = insertPage(index + 1, d.title, '', { type: '__cards' }) // the deck's accent colour
  const page = slides.value[index + 1]
  const body: any = page.elements.find((e: any) => e.type === '__cards')
  const gap = W * 0.02, cw = (body.width - gap * (cols - 1)) / cols, ch = (body.height - gap * (rows - 1)) / rows
  const cards = d.points.slice(0, 4).flatMap((p: any, i: number) => {
    const left = body.left + (i % cols) * (cw + gap), top = body.top + Math.floor(i / cols) * (ch + gap)
    return [
      { type: 'shape', id: nanoid(10), left, top, width: cw, height: ch, rotate: 0, viewBox: [200, 200], path: 'M 0 0 L 200 0 L 200 200 L 0 200 Z', fixedRatio: false, fill: '#f4f5f8', outline: { width: 0 } },
      { type: 'shape', id: nanoid(10), left, top, width: W * 0.006, height: ch, rotate: 0, viewBox: [200, 200], path: 'M 0 0 L 200 0 L 200 200 L 0 200 Z', fixedRatio: false, fill: color },
      { type: 'text', id: nanoid(10), left: left + W * 0.02, top: top + H * 0.02, width: cw - W * 0.035, height: ch - H * 0.04, rotate: 0,
        defaultFontName: theme.value.fontName, defaultColor: theme.value.fontColor,
        content: `<p><span style="font-size: ${k > 2 ? 18 : 20}px; color: ${color};"><strong>${esc(p.heading)}</strong></span></p><p><span style="font-size: ${k > 2 ? 14 : 16}px;">${esc(p.text)}</span></p>` },
    ]
  })
  slidesStore.updateSlide({ elements: [...page.elements.filter((e: any) => e.type !== '__cards'), ...cards] }, page.id)
}

// The free rectangle nearest to the requested box (same aspect, shrinking down to 40%) on a 24x14 occupancy grid;
// full-slide backgrounds don't count as occupied. Null when nothing is free.
const freeBox = (sl: Slide, want: { left: number, top: number, width: number, height: number }) => {
  const { width: W, height: H } = slideSize()
  const C = 24, R = 14, cw = W / C, rh = H / R
  const busyCell = Array.from({ length: R }, () => new Array(C).fill(false))
  for (let r = Math.floor(R * 0.86); r < R; r++) busyCell[r].fill(true) // the footer band
  for (const e of sl.elements as any[]) {
    if (!e.width || !e.height || e.width * e.height > W * H * 0.6) continue
    for (let r = Math.max(0, Math.floor(e.top / rh)); r < Math.min(R, Math.ceil((e.top + e.height) / rh)); r++) {
      for (let c = Math.max(0, Math.floor(e.left / cw)); c < Math.min(C, Math.ceil((e.left + e.width) / cw)); c++) busyCell[r][c] = true
    }
  }
  const cx = want.left + want.width / 2, cy = want.top + want.height / 2
  for (let k = 1; k >= 0.4; k -= 0.1) {
    const w = Math.max(1, Math.round(want.width * k / cw)), h = Math.max(1, Math.round(want.height * k / rh))
    let best: any = null
    for (let r = 0; r + h <= R; r++) {
      for (let c = 0; c + w <= C; c++) {
        let ok = true
        for (let i = r; i < r + h && ok; i++) for (let j = c; j < c + w && ok; j++) ok = !busyCell[i][j]
        const d = Math.hypot((c + w / 2) * cw - cx, (r + h / 2) * rh - cy)
        if (ok && (!best || d < best.d)) best = { d, left: c * cw, top: r * rh, width: w * cw, height: h * rh }
      }
    }
    if (best) return best
  }
  return null // nowhere free: never cover the page's content
}

// Fetch a picture: a stock photo first (unless an AI image is asked for), else gpt-image.
// Fetch a picture from the source asked for; never quietly another kind. `credit` is said in the reply.
let pictureNote = ''
let pictureSource: 'photo' | 'wikimedia' | 'ai' | 'attached' = 'photo' // what the last picture really came from (a stock search can fall back to AI)
const findPicture = async (query: string, source: string, landscape: boolean) => {
  pictureNote = ''
  pictureSource = source === 'wikimedia' ? 'wikimedia' : source === 'ai' ? 'ai' : 'photo'
  const Z = zh(userText)
  if (source === 'attached') { // the query names one of the user's attached pictures
    const pics = s.images || []
    const pic = pics.find(m => m.name === query) || (pics.length === 1 ? pics[0] : undefined)
    if (!pic) fail(pics.length ? (Z ? `没有找到附件图片「${query}」。现有：${pics.map(m => m.name).join('、')}。` : `There is no attached picture "${query}". Attached: ${pics.map(m => m.name).join(', ')}.`)
      : (Z ? '还没有附加图片：请用 📎 图标上传你的图片。' : 'No picture is attached yet: attach it with the 📎 icon.'))
    pictureSource = 'attached'
    return pic.url
  }
  if (source === 'wikimedia') {
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'wikimedia', query }) })).json()
    if (!r.src) fail(Z ? `维基共享资源上没有找到「${query}」的图片。要我改用图库照片或 AI 生成吗？` : `Wikimedia Commons has no picture of "${query}". Should I use a stock photo or generate one instead?`)
    pictureNote = Z ? `图片来自维基共享资源：${r.title}（${r.credit}${r.license ? `，${r.license}` : ''}）${r.page}` : `Picture from Wikimedia Commons: ${r.title} (${r.credit}${r.license ? `, ${r.license}` : ''}) ${r.page}`
    return r.src as string
  }
  if (source === 'photo') {
    const r = await api.searchImage({ query, per_page: 1, orientation: landscape ? 'landscape' : 'portrait' } as any)
    const src = r.data?.[0]?.src // the services axios instance already unwraps response.data
    if (src) return src as string
    pictureNote = Z ? `没有匹配的图库照片，所以用 AI 生成了一张。` : 'No stock photo matched, so I generated one with AI.'
    pictureSource = 'ai'
  }
  if (chatslide?.generateImage) { // inside ChatSlide: its image route bills, gates and moderates
    const g = await chatslide.generateImage(query, landscape ? '16:9' : '9:16')
    if (g.error === 'credits') {
      const e = new OutOfCredits('out of credits')
      e.required = 10
      throw e
    }
    if (g.error) fail(Z ? `图片没能生成（${g.error}）。` : `The picture could not be generated (${g.error}).`)
    return g.src || ''
  }
  const r = await fetch('/api/tools/ai_image', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ prompt: query, size: landscape ? '1536x1024' : '1024x1536' }),
  })
  return ((await r.json()).src || '') as string
}

const changeImage = async (text: string, slideId: string, image: number, query: string, source: string) => {
  const sl = slides.value.find(x => x.id === slideId)
  const el = sl && readingOrder(sl).filter((e): e is PPTImageElement => e.type === 'image')[Math.max(0, (image || 1) - 1)]
  if (!sl || !el) fail(zh(text) ? '这一页没有图片。' : 'That page has no picture.')
  const src = await findPicture(query, source, el.width >= el.height)
  if (!src) fail(zh(text) ? '没有找到合适的图片。' : 'Could not get a picture.')
  slidesStore.updateElement({ id: el.id, props: { src } as any, slideId })
  touched([slideId])
  return pictureNote
}

const writeNotes = async (ids: string[], instruction: string) => {
  let written = 0
  for (const id of ids) {
    const sl = slides.value.find(x => x.id === id)!
    busy.value = `Writing notes ${ids.indexOf(id) + 1} of ${ids.length}…`
    const r = await fetch('/api/tools/ask_json', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ strong: true, instruction: `Speaker notes for THIS one slide only (never mention page numbers or other slides): what the presenter says aloud, adding context the slide only implies, in the slide's language. Every number must be exactly as on the slide. Default length 3-5 short sentences. Request: ${instruction || 'write them'}\nSlide: ${pageText(sl)}\nCurrent notes: ${(sl.remark || '').replace(/<[^>]+>/g, ' ').trim() || '(none)'}\nReturn JSON {"notes": "..."} (paragraphs separated by \\n).` }),
    })
    const notes = String((await r.json()).notes || '').trim()
    if (notes) {
      slidesStore.updateSlide({ remark: String(notes).split('\n').filter(Boolean).map(l => `<p>${l.replace(/</g, '&lt;')}</p>`).join('') }, id)
      written++
    }
  }
  return written
}

// New pages never land after the closing page ("Thank you"): an insert at the very end goes before it.
// a new page's title, for replies: page numbers shift when later steps insert pages, titles do not
const titleAt = (i: number) => (slides.value[i] ? (readingOrder(slides.value[i]).map(elText).find(Boolean) || '').slice(0, 50) : '') || (zh(userText) ? `第 ${i + 1} 页` : `page ${i + 1}`)
const beforeClosing = (at: number) => {
  const last = slides.value[slides.value.length - 1]
  // a generated deck marks it (type 'end'); in an imported deck the closing page is a last page with almost no text
  const words = last ? pageText(last).replace(/\[[^\]]*\]/g, '').trim() : ''
  const closing = last && (last.type === 'end' || (slides.value.length > 2 && words.length > 0 && words.length < 30
    && !last.elements.some(e => e.type === 'chart' || e.type === 'table')))
  return closing && at >= slides.value.length ? slides.value.length - 1 : at
}

const runCall = async (text: string, action: string, args: any, ids: string[]) => {
  const idOf = (p: number) => {
    if (!(p >= 1 && p <= ids.length)) fail(zh(text) ? `没有第 ${p} 页（共 ${ids.length} 页）。` : `There is no page ${p} (the deck has ${ids.length}).`)
    return ids[p - 1]
  }
  const pos = (id: string) => slides.value.findIndex(x => x.id === id)
  const nums = (ps: number[]) => [...new Set((ps || []).map(idOf))].filter(id => pos(id) >= 0)
  const n = slides.value.length
  const Z = zh(text)
  if (action === 'add_page') {
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    if (s.mode === 'native') {
      const n0 = slides.value.length
      if (!await generatePage(text, args.instruction, after, Math.min(Math.max(args.count || 1, 1), 10))) fail(Z ? '没能生成这一页。' : 'Could not write that page.')
      const added = slides.value.length - n0
      const capped = (args.count || 1) > 10 ? (Z ? '（一次最多写 10 个主题，再说一次可继续加）' : ' (at most 10 topics per request; ask again for more)') : ''
      return (Z ? `已在第 ${after} 页后加了 ${added} 页。` : `Added ${added} page(s) after page ${after}.`) + capped
    }
    const count = Math.min(Math.max(args.count || 1, 1), 10)
    for (let k = 0; k < count; k++) {
      busy.value = `Writing page ${k + 1} of ${count}…`
      const list = [...slides.value]
      list.splice(after + k, 0, cloneSlide(slides.value[pos(idOf(args.layout_from || args.after))]))
      slidesStore.setSlides(list)
      await editSlide(after + k, `This is NEW page ${k + 1} of ${count} on: ${args.instruction}. Replace all its text with that page's own content (no repeats across pages). Keep the layout and styling. Write in the language of that text.`)
    }
    const capped = (args.count || 1) > 10 ? (Z ? '（一次最多加 10 页，再说一次可继续加）' : ' (10 is the most per request; ask again for more)') : ''
    return (Z ? `已在第 ${after} 页后加了 ${count} 页。` : `Added ${count} page(s) after page ${after}.`) + capped
  }
  if (action === 'expand_page') {
    const i = pos(idOf(args.index))
    busy.value = Z ? '正在扩展…' : 'Extending…'
    await expandPage(text, i, args.instruction || '')
    slidesStore.updateSlideIndex(i + 1)
    return Z ? `已在第 ${i + 1} 页后加了扩展页（第 ${i + 2} 页），原页保持不变。` : `Added an extension of page ${i + 1} as page ${i + 2}; page ${i + 1} is unchanged.`
  }
  if (action === 'add_chart_page') {
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    return (await addChartPage(text, args.instruction, args.chart_type || 'column', after)) || (Z ? `已加入图表页「${titleAt(after)}」。` : `Added a chart page "${titleAt(after)}".`)
  }
  if (action === 'add_table_page') {
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    await addTablePage(text, args.instruction, after)
    slidesStore.updateSlideIndex(after)
    return Z ? `已加入表格页「${titleAt(after)}」。` : `Added a table page "${titleAt(after)}".`
  }
  if (args.page_named === false && typeof args.index === 'number' && ['use_file_images', 'add_image', 'change_image'].includes(action)) args.index = slideIndex.value + 1 // no page named: the one the user is looking at
  if (action === 'use_file_images') { // figures from the user's own PDF/PPTX/DOCX, extracted by ChatSlide's backend
    const file = docFiles.get(args.file) || (docFiles.size === 1 ? [...docFiles.values()][0] : undefined)
    if (!file) fail(s.docs.some(d => d.name === args.file)
      ? (Z ? `「${args.file}」需要重新上传一次才能提取其中的图片。` : `Attach ${args.file} again so I can pull its pictures out (it was attached in an earlier session).`)
      : (Z ? '还没有附件：请用 📎 图标上传文件。' : 'No file is attached yet: attach it with the 📎 icon.'))
    if (!chatslide?.extractImages) fail(Z ? '在 ChatSlide 里打开这个项目才能提取文件中的图片。' : 'Pulling pictures out of a file works inside ChatSlide: open this project there.')
    busy.value = Z ? '正在提取文件中的图片…' : 'Pulling the pictures out of the file…'
    let figs = figuresOf.get(args.file) || figuresOf.get(file.name)
    if (!figs) {
      const r = await chatslide.extractImages(file)
      if (r.error === 'plan') fail(Z ? '从文件中提取图片需要付费方案，已为你打开方案页面。' : 'Pulling pictures out of files needs a paid plan; I opened the plans for you.')
      if (r.error) fail(Z ? `没能提取图片（${r.error}）。` : `I could not pull the pictures out (${r.error}).`)
      figs = r.images || []
      figuresOf.set(file.name, figs)
    }
    const want: number[] = args.file_pages || []
    const chosen = (want.length ? figs.filter(m => m.page && want.includes(m.page)) : figs).slice(0, Math.min(Math.max(args.count || 1, 1), 6))
    if (!chosen.length) fail(figs.length
      ? (Z ? `「${file.name}」第 ${want.join('、')} 页没有图片（共提取到 ${figs.length} 张，在第 ${[...new Set(figs.map(m => m.page))].filter(Boolean).join('、')} 页）。` : `There are no pictures on page(s) ${want.join(', ')} of ${file.name} (it has ${figs.length}, on pages ${[...new Set(figs.map(m => m.page))].filter(Boolean).join(', ')}).`)
      : (Z ? `「${file.name}」里没有可用的图片。` : `${file.name} has no pictures I can use.`))
    const id = idOf(args.index)
    const { width: W, height: H } = slideSize()
    let placed = 0
    for (const m of chosen) {
      const box = freeBox(slides.value[pos(id)], { left: W * 0.55, top: H * 0.2, width: W * 0.4, height: H * 0.6 })
      if (!box) break
      slidesStore.updateSlide({ elements: [...slides.value[pos(id)].elements, { type: 'image', id: nanoid(10), src: m.url, fixedRatio: true, rotate: 0, ...box } as any] }, id)
      placed++
    }
    touched([id])
    slidesStore.updateSlideIndex(pos(id))
    if (!placed) fail(Z ? `第 ${pos(id) + 1} 页没有空位放图片。` : `Page ${pos(id) + 1} has no free space for the pictures.`)
    return (Z ? `已把「${file.name}」中的 ${placed} 张图片放到第 ${pos(id) + 1} 页。` : `Placed ${placed} picture(s) from ${file.name} on page ${pos(id) + 1}.`) + (placed < chosen.length ? (Z ? `（其余 ${chosen.length - placed} 张没有空位）` : ` (no room for the other ${chosen.length - placed})`) : '')
  }
  if (action === 'add_image') {
    const id = idOf(args.index)
    const { width: W, height: H } = slideSize()
    const b = args.box || { x: 55, y: 20, w: 40, h: 60 }
    const clamp = (v: number) => Math.min(Math.max(v || 0, 0), 100) / 100
    const sl = slides.value[pos(id)]
    const box = freeBox(sl, { left: clamp(b.x) * W, top: clamp(b.y) * H, width: Math.max(clamp(b.w), 0.05) * W, height: Math.max(clamp(b.h), 0.05) * H })
    if (!box) fail(Z ? `第 ${pos(id) + 1} 页已经排满，没有放图片的空位；可以先删掉或缩小一些内容，或者把图片放到新的一页。` : `Page ${pos(id) + 1} is full; there is no free space for a picture. Remove or shrink something first, or put the picture on a new page.`)
    busy.value = args.source === 'ai' ? 'Drawing the picture…' : 'Finding a picture…'
    const src = await findPicture(args.query, args.source, box.width >= box.height)
    if (!src) fail(Z ? '没有找到合适的图片。' : 'Could not get a picture.')
    const now = freeBox(slides.value[pos(id)], box) // re-checked: the page may have changed while the picture loaded
    if (!now) fail(Z ? `第 ${pos(id) + 1} 页已经没有空位了。` : `Page ${pos(id) + 1} has no free space any more.`)
    slidesStore.updateSlide({ elements: [...slides.value[pos(id)].elements, { type: 'image', id: nanoid(10), src, fixedRatio: true, rotate: 0, ...now } as any] }, id)
    touched([id])
    picturedThisTurn.add(id)
    slidesStore.updateSlideIndex(pos(id))
    const how = pictureSource === 'attached' ? (Z ? '你上传的' : 'your attached') : pictureSource === 'ai' ? (Z ? 'AI 生成' : 'AI-generated') : pictureSource === 'wikimedia' ? (Z ? '维基共享资源' : 'Wikimedia') : (Z ? '图库照片' : 'stock photo')
    const cx = (now.left + now.width / 2) / W, cy = (now.top + now.height / 2) / H
    const askedX = clamp(b.x) + clamp(b.w) / 2 // where the router wanted it; the free-space search may have moved it
    const moved = (askedX > 2 / 3 && cx <= 2 / 3) || (askedX < 1 / 3 && cx >= 1 / 3)
    const movedNote = moved ? (Z ? `（${askedX > 2 / 3 ? '右' : '左'}侧没有空位，所以放在了这里）` : ` (no free space on the ${askedX > 2 / 3 ? 'right' : 'left'}, so it went here)`) : ''
    const zhWhere = (cy < 1 / 3 || cy > 2 / 3 ? (cx < 1 / 3 ? '左' : cx > 2 / 3 ? '右' : '') + (cy < 1 / 3 ? '上' : '下') + (cx >= 1 / 3 && cx <= 2 / 3 ? '方' : '') : cx < 1 / 3 ? '左侧' : cx > 2 / 3 ? '右侧' : '正中')
    const where = Z ? zhWhere : `${cy < 1 / 3 ? 'top' : cy > 2 / 3 ? 'bottom' : 'middle'} ${cx < 1 / 3 ? 'left' : cx > 2 / 3 ? 'right' : 'centre'}`
    return (Z ? `已在页面${where}插入“${String(args.query || '').slice(0, 60)}”的${how}图片${movedNote}。` : `Added a picture (${how}) for "${String(args.query || '').slice(0, 60)}" at the ${where}${movedNote}.`) + (pictureNote ? ` ${pictureNote}` : '')
  }
  if (action === 'remove_images') {
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    const logos = logoSpots()
    let n = 0
    for (const id of targets) {
      const imgs = readingOrder(slides.value[pos(id)]).filter(e => e.type === 'image') // the numbering the router saw
      const drop = new Set((args.images?.length ? args.images.map((k: number) => imgs[k - 1]) : imgs.filter(e => !logos.has(imgSpot(e))))
        .filter(Boolean).map((e: any) => e.id))
      if (!drop.size) continue
      n += drop.size
      slidesStore.updateSlide({ elements: slides.value[pos(id)].elements.filter(e => !drop.has(e.id)) }, id)
      touched([id])
    }
    if (!n) fail(Z ? '没有可删除的图片。' : 'There is no picture to remove.')
    return Z ? `已删除 ${n} 张图片。` : `Removed ${n} picture(s).`
  }
  if (action === 'set_transition') {
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    for (const id of targets) slidesStore.updateSlide({ turningMode: args.effect }, id)
    return Z ? `已设置 ${targets.length} 页的切换效果。` : `Transition set on ${targets.length} page(s).`
  }
  if (action === 'animate_pages') {
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    for (const id of targets) {
      const sl = slides.value[pos(id)]
      const animations = args.effect === 'none' ? [] : readingOrder(sl).filter((e: any) => e.type !== 'line').map((e, k) => ({
        id: nanoid(10), elId: e.id, effect: args.effect, type: 'in' as const, duration: 600,
        trigger: (k === 0 || args.trigger === 'click' ? 'click' : 'auto') as 'click' | 'auto',
      }))
      slidesStore.updateSlide({ animations }, id)
    }
    return Z ? (args.effect === 'none' ? '已移除动画。' : `已给 ${targets.length} 页加上动画。`) : (args.effect === 'none' ? 'Animations removed.' : `Animated ${targets.length} page(s).`)
  }
  if (action === 'delete_pages') {
    const del = nums(args.pages)
    if (del.length >= n) fail(Z ? '至少要保留一页。' : 'A deck needs at least one page, so I kept it.')
    slidesStore.deleteSlide(del)
    return Z ? `已删除 ${del.length} 页，现在共 ${slides.value.length} 页。` : `Deleted ${del.length} page(s); ${slides.value.length} remain.`
  }
  if (action === 'read_file') {
    const doc = s.docs.find(d => d.name === args.file) || (s.docs.length === 1 ? s.docs[0] : undefined)
    if (!doc) fail(s.docs.length ? (Z ? `没有找到附件「${args.file}」。现有附件：${s.docs.map(d => d.name).join('、')}。` : `There is no attached file "${args.file}". Attached: ${s.docs.map(d => d.name).join(', ')}.`)
      : (Z ? '还没有附件：请用输入框的 📎 图标上传文件。' : 'No file is attached yet: attach it with the 📎 icon in the chat box.'))
    const parts = doc.text.split(/^(?=\[page \d+\]$)/m)
    const want: number[] = args.pages || []
    const total = parts.filter(p => /^\[page \d+\]/.test(p)).length
    const missing = want.filter(n => n < 1 || n > total)
    if (want.length && total && missing.length === want.length) fail(Z ? `「${doc.name}」只有 ${total} 页。` : `${doc.name} has only ${total} pages.`)
    const got = want.filter(n => !missing.includes(n))
    const picked = want.length && total ? parts.filter(p => got.includes(+(p.match(/^\[page (\d+)\]/)?.[1] || 0))).join('\n') : doc.text
    if (!picked.replace(/\[page \d+\]/g, '').trim()) fail(Z ? `「${doc.name}」这几页读不出文字（可能是扫描图片）。` : `No text could be read from ${doc.name}${want.length ? ` page(s) ${want.join(', ')}` : ''} (it may be a scanned image).`)
    memo = `<file-content name="${doc.name.replace(/["<>]/g, '')}"${want.length ? ` pages="${want.join(',')}"` : ''}>\n${fence(picked.slice(0, 12000))}${picked.length > 12000 ? '\n[…cut; read fewer pages for the rest]' : ''}\n</file-content>`
    const skipped = missing.length && total ? (Z ? `（没有第 ${missing.join('、')} 页，共 ${total} 页）` : ` (no page ${missing.join(', ')}; it has ${total})`) : ''
    return (Z ? `已读取「${doc.name}」${want.length ? `第 ${got.join('、')} 页` : ''}。` : `Read ${doc.name}${want.length ? `, page(s) ${got.join(', ')}` : ''}.`) + skipped
  }
  if (action === 'find_papers') {
    busy.value = Z ? '正在查找学术文献…' : 'Finding papers…'
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'scholar', query: args.query }) })).json()
    if (r.error || !r.papers?.length) fail(Z ? `没有找到关于“${args.query}”的学术文献，所以没有添加引用。` : `I found no papers on "${args.query}", so I added no citations.`)
    memo = `<web-results>\n${fence(r.papers.map((p: any, i: number) => `[${i + 1}] ${p.authors} (${p.year}). ${p.title}. ${p.venue}. ${p.url} (cited ${p.cited}x)`).join('\n'))}\n</web-results>`
    return Z ? `已找到 ${r.papers.length} 篇相关文献。` : `Found ${r.papers.length} relevant paper(s).`
  }
  if (action === 'read_link') {
    busy.value = Z ? '正在读取链接…' : 'Reading the link…'
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'url', query: args.url }) })).json()
    if (r.error || !String(r.text || '').trim()) fail(Z ? `没能读取这个链接（${r.error || '没有内容'}）。` : `I could not read that link (${r.error || 'no content'}).`)
    memo = `<file-content name="${String(r.title).replace(/["<>]/g, '').slice(0, 80)}">\n${fence(String(r.text).slice(0, 12000))}\n</file-content>`
    return Z ? `已读取${r.kind === 'video transcript' ? '视频字幕' : '网页'}：${String(r.title).slice(0, 60)}。` : `Read the ${r.kind}: ${String(r.title).slice(0, 60)}.`
  }
  if (action === 'web_search') {
    busy.value = Z ? '正在网上搜索…' : 'Searching the web…'
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'web', query: args.query }) })).json()
    if (r.error || !r.sources?.length) fail(Z ? `网上没有找到可靠的来源（${args.query}），所以没有添加引用。` : `I found no reliable sources online for "${args.query}", so I added no citations.`)
    memo = `<web-results>\n${fence(`${r.answer}\nSources:\n${r.sources.slice(0, 6).map((x: any, i: number) => `[${i + 1}] ${x.title} ${x.url}`).join('\n')}`)}\n</web-results>`
    return Z ? '已在网上查到资料。' : 'Searched online.'
  }
  if (action === 'merge_pages') {
    const ids = nums(args.pages)
    if (ids.length < 2) fail(Z ? '请说明要合并哪几页。' : 'Which pages should be combined?')
    const [keep, ...rest] = ids
    // the router sees each page in one short line; the merge gets every word of the pages being folded in
    const extra = rest.map(id => readingOrder(slides.value[pos(id)]).map(elText).filter(Boolean).join('\n')).join('\n\n')
    await editSlide(pos(keep), `Combine the content below into this page: keep this page's title, add the points of the other page(s) condensed so everything fits without overflow, and rearrange the layout if needed. ${args.instruction || ''}\n\nContent of the other page(s):\n${extra}`)
    slidesStore.deleteSlide(rest)
    return Z ? `已把 ${ids.length} 页合并为第 ${pos(keep) + 1} 页，现在共 ${slides.value.length} 页。` : `Combined ${ids.length} pages into page ${pos(keep) + 1}; ${slides.value.length} remain.`
  }
  if (action === 'duplicate_pages') {
    const list = [...slides.value]
    for (const id of nums(args.pages)) {
      const i = list.findIndex(x => x.id === id)
      list.splice(i + 1, 0, cloneSlide(list[i]))
    }
    slidesStore.setSlides(list)
    return Z ? '已复制。' : 'Duplicated.'
  }
  if (action === 'move_pages') {
    const move = nums(args.pages)
    const moved = slides.value.filter(x => move.includes(x.id)) // keep deck order
    const rest = slides.value.filter(x => !move.includes(x.id))
    const at = Math.min(Math.max((args.to || 1) - 1, 0), rest.length)
    rest.splice(at, 0, ...moved)
    if (rest.map(x => x.id).join() === slides.value.map(x => x.id).join()) fail(Z ? '这些页已经在那个位置了，顺序没有变化。' : 'Those pages are already there; the order is unchanged.')
    slidesStore.setSlides(rest)
    return Z ? `已移动到第 ${at + 1} 页。` : `Moved to position ${at + 1}.`
  }
  if (action === 'reorder_pages') {
    const order = nums(args.order)
    const list = [...order, ...slides.value.map(x => x.id).filter(id => !order.includes(id))] // pages left out keep their relative order at the end
      .map(id => slides.value[pos(id)])
    if (list.map(x => x.id).join() === slides.value.map(x => x.id).join()) fail(Z ? '这已经是当前的顺序，没有变化。' : 'That is already the current order; nothing changed.')
    slidesStore.setSlides(list)
    return Z ? '已重新排序。' : 'Reordered.'
  }
  if (action === 'swap_pages') {
    const list = [...slides.value]
    const i = pos(idOf(args.a)), j = pos(idOf(args.b))
    ;[list[i], list[j]] = [list[j], list[i]]
    slidesStore.setSlides(list)
    return Z ? `已对调第 ${i + 1} 页和第 ${j + 1} 页。` : `Swapped pages ${i + 1} and ${j + 1}.`
  }
  if (action === 'edit_page') {
    // no page named: the page the user is looking at, like the classic assistant
    const target = args.page_named === false && args.index !== slideIndex.value + 1 ? slideIndex.value + 1 : args.index
    const i = pos(idOf(target))
    const reply = await editSlide(i, args.instruction)
    slidesStore.updateSlideIndex(i)
    return reply
  }
  if (action === 'set_font') { // direct: element default + every inline font-family, no model, no layout change
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    touched(targets)
    const font = String(args.font || '').replace(/[^\p{L}\p{N} ,_-]/gu, '').trim().slice(0, 60) // a font name, nothing CSS can parse further
    if (!font) fail(Z ? '请说明要换成哪种字体。' : 'Which font?')
    const setIn = (html: string) => html.replace(/font-family:\s*[^;"]*/g, `font-family: ${font}`)
    let n = 0
    for (const id of targets) {
      const sl = slides.value[pos(id)]
      const title = (e: any) => e.textType === 'title' || e.text?.type === 'title' || e === readingOrder(sl).find((x: any) => elText(x))
      for (const e of sl.elements as any[]) {
        if (!(e.type === 'text' || (e.type === 'shape' && e.text)) || !elText(e)) continue
        if (args.scope === 'titles' ? !title(e) : args.scope === 'body' ? title(e) : false) continue
        if (e.type === 'text') slidesStore.updateElement({ id: e.id, props: { defaultFontName: font, content: setIn(e.content) }, slideId: id })
        else slidesStore.updateElement({ id: e.id, props: { text: { ...e.text, defaultFontName: font, content: setIn(e.text.content) } }, slideId: id })
        n++
      }
    }
    return Z ? `已把 ${targets.length} 页上 ${n} 处文字换成 ${font}。` : `Set ${font} on ${n} text boxes across ${targets.length} page(s).`
  }
  if (action === 'edit_all_pages') { // 5 pages in flight, the next starts as soon as one ends; merged by slide id
    const all = slides.value.map(x => x.id)
    const viewing = slides.value[slideIndex.value]?.id // fitting shows each page; come back to this one
    let done = 0, next = 0, stop = false
    const same: number[] = [], failed: number[] = []
    const worker = async () => {
      while (next < all.length && !stop) {
        const id = all[next++]
        try {
          await editSlide(pos(id), args.instruction)
        }
        catch (e) { // one page never stops the others; an unexpected error stops them all
          if (e instanceof NoChange) same.push(pos(id) + 1)
          else if (e instanceof StepFailed) failed.push(pos(id) + 1)
          else {
            stop = true
            throw e
          }
        }
        busy.value = `Working on the pages: ${++done} of ${all.length}…`
      }
    }
    // wait for every worker (an edit in flight still lands), then report the first unexpected error
    const settled = await Promise.allSettled(Array.from({ length: Math.min(5, all.length) }, worker))
    const err = settled.find((x): x is PromiseRejectedResult => x.status === 'rejected')
    if (err) throw err.reason
    await fitQueue
    if (viewing && pos(viewing) >= 0) slidesStore.updateSlideIndex(pos(viewing))
    const earlier = same.filter(p => editedThisTurn.has(all[p - 1])) // done by an earlier step of this turn
    const untouched = same.filter(p => !editedThisTurn.has(all[p - 1]))
    const changedNow = all.length - same.length - failed.length
    if (!changedNow && earlier.length && !untouched.length && !failed.length) return Z ? '这些页在前一步已经改好，没有进一步的改动。' : 'No further change: the earlier step already covered these pages.'
    const ok = changedNow + earlier.length
    const list = (xs: number[]) => xs.sort((a, b) => a - b).join(Z ? '、' : ', ')
    const what = String(args.instruction || '').slice(0, 60)
    const why = (failed.length ? (Z ? `第 ${list(failed)} 页修改失败。` : ` Page(s) ${list(failed)} failed.`) : '')
      + (untouched.length && ok ? (Z ? `第 ${list(untouched)} 页没有可改的内容。` : ` Page(s) ${list(untouched)} had nothing for that.`) : '')
    void earlier
    if (!ok) fail((Z ? `这个修改在任何一页都没做成：“${what}”。` : `I could not make that change on any page: "${what}".`) + why)
    return (Z ? `已更新 ${ok} 页。` : `Updated ${ok} page${ok > 1 ? 's' : ''}.`) + why
  }
  if (action === 'change_image') {
    busy.value = args.source === 'ai' ? 'Drawing the picture…' : 'Finding a picture…'
    const id = idOf(args.index)
    slidesStore.updateSlideIndex(pos(id))
    const note = await changeImage(text, id, args.image, args.query, args.source)
    const how = pictureSource === 'attached' ? (Z ? '你上传的' : 'your attached') : pictureSource === 'ai' ? (Z ? 'AI 生成' : 'AI-generated') : pictureSource === 'wikimedia' ? (Z ? '维基共享资源' : 'Wikimedia') : (Z ? '图库照片' : 'stock photo')
    return (Z ? `图片已更换为“${String(args.query || '').slice(0, 60)}”的${how}图片，请看看是否合适。` : `Picture replaced (${how}) for "${String(args.query || '').slice(0, 60)}"; check it is what you meant.`) + (note ? ` ${note}` : '')
  }
  if (action === 'move_element_to_page') {
    const from = slides.value[pos(idOf(args.from))], to = idOf(args.to)
    const list = readingOrder(from) // same order and image numbering the router saw
    let img = 0
    const r = await fetch('/api/tools/ask_json', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ instruction: `Which element is "${args.element}"? Elements: ${JSON.stringify(list.map((e: any, k) => ({ k, type: e.type === 'image' ? `image ${++img}` : e.type, text: elText(e).slice(0, 60), left: Math.round(e.left), top: Math.round(e.top), width: Math.round(e.width), height: Math.round(e.height || 0) })))}. "The picture" is the largest image. Return JSON {"k": <index>} or {"k": -1}.` }),
    })
    const el = list[(await r.json()).k]
    if (!el) fail(Z ? '没找到这个元素。' : 'Could not find that element.')
    if (!args.copy) slidesStore.updateSlide({ elements: from.elements.filter(e => e.id !== el.id) }, from.id)
    const target = slides.value[pos(to)]
    slidesStore.updateSlide({ elements: [...target.elements, { ...JSON.parse(JSON.stringify(el)), id: nanoid(10) }] }, to)
    slidesStore.updateSlideIndex(pos(to))
    return args.copy ? (Z ? '已复制过去。' : 'Copied.') : (Z ? '已移动。' : 'Moved.')
  }
  if (action === 'set_background') {
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    let background: any = { type: 'solid', color: args.color }
    if (args.photo_query) {
      busy.value = 'Finding a picture…'
      const src = await findPicture(args.photo_query, 'photo', true)
      if (!src) fail(Z ? '没有找到合适的背景图。' : 'Could not get a background picture.')
      background = { type: 'image', image: { src, size: 'cover' } }
    }
    else if (!/^#[0-9a-fA-F]{6}$/.test(args.color || '')) fail(Z ? '请给出背景颜色。' : 'Which background colour?')
    for (const id of targets) slidesStore.updateSlide({ background }, id)
    touched(targets)
    const missed: number[] = []
    // a page whose background is mostly painted over by big shapes would look unchanged: recolour those shapes too
    if (!args.photo_query) {
      const { width: W, height: H } = slideSize()
      const covered = targets.filter(id => slides.value[pos(id)].elements
        .filter((e: any) => e.type === 'shape' && !elText(e) && e.width >= W * 0.9 && e.width * (e.height || 0) >= W * H * 0.25) // full-width bands only, never cards or diagrams
        .reduce((a: number, e: any) => a + Math.min(e.width, W) * Math.min(e.height, H), 0) >= W * H * 0.45)
      for (let i = 0; i < covered.length; i += 3) {
        await Promise.all(covered.slice(i, i + 3).map(id => editSlide(pos(id), `Change the page's BACKGROUND to ${args.color}: recolour only the large decorative background shapes (full-width bands, full-page panels) to ${args.color} or a lighter/darker tint of it; keep cards, text boxes, pictures and their colours unchanged and keep text readable.`)
          .catch(() => {
            missed.push(pos(id) + 1)
          })))
      }
    }
    const note = missed.length ? (Z ? `（第 ${missed.join('、')} 页的大色块没能改色，可能仍会挡住背景）` : ` (the large shapes on page(s) ${missed.join(', ')} could not be recoloured and may still cover it)`) : ''
    const bgHow = args.photo_query
      ? (pictureSource === 'ai' ? (Z ? `“${String(args.photo_query).slice(0, 60)}”的 AI 生成图片` : `an AI-generated picture for "${String(args.photo_query).slice(0, 60)}"`) : (Z ? `“${String(args.photo_query).slice(0, 60)}”的图库照片` : `a stock photo for "${String(args.photo_query).slice(0, 60)}"`))
      : (Z ? `颜色 ${args.color}` : `colour ${args.color}`)
    return (Z ? `已把 ${targets.length} 页背景设为${bgHow}。` : `Background of ${targets.length} page(s) set to ${bgHow}.`) + note
  }
  if (action === 'write_notes') {
    const hasNotes = (id: string) => !!(slides.value[pos(id)]?.remark || '').replace(/<[^>]+>/g, '').trim()
    const asked = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    const targets = args.existing_only ? asked.filter(hasNotes) : asked // revising touches only pages that have notes
    if (args.existing_only && !targets.length) fail(Z ? '这些页还没有演讲备注，所以没有可以修改的；说“写备注”我就先写出来。' : 'These pages have no speaker notes yet, so there is nothing to revise; say "write notes" and I will write them first.')
    const written = await writeNotes(targets, args.instruction || '')
    if (!written) fail(Z ? '演讲备注没能写出来。' : 'The speaker notes could not be written.')
    return (Z ? `已写好 ${written} 页的演讲备注。` : `Wrote speaker notes for ${written} page(s).`) + (written < targets.length ? (Z ? `（${targets.length - written} 页没写成）` : ` (${targets.length - written} page(s) failed)`) : '')
  }
  if (action === 'download_deck') {
    if (chatslide && !chatslide.canExport()) return Z ? '下载需要付费方案，已为你打开方案页面。' : 'Downloading needs a paid plan; I opened the plans for you.'
    if (s.mode === 'template' && !s.edited) { // the engine's own .pptx is exact; after editor changes, export the editor
      const a = document.createElement('a')
      a.href = s.versions[s.versions.length - 1].deckUrl
      a.download = 'deck.pptx'
      a.click()
      return Z ? '正在下载 .pptx。' : 'Downloading the .pptx now.'
    }
    if (args.format === 'pdf' || args.format === 'images') { // these need the print/render dialogs
      mainStore.setDialogForExport(args.format === 'pdf' ? 'pdf' : 'image')
      return Z ? '已打开导出窗口，点“导出”即可下载。' : 'I opened the Export window; click Export to download.'
    }
    busy.value = Z ? '正在导出 .pptx…' : 'Exporting the .pptx…'
    exportPPTX(slides.value, true, false) // the browser saves the file; no dialog
    const asked = String(args.asked || '').trim()
    const other = asked && !/pptx|powerpoint|ppt/i.test(asked) // a format we cannot export to directly (Google Slides, Keynote…)
    return other
      ? (Z ? `这里不能直接导出为 ${asked}；正在下载 .pptx，${asked} 可以直接打开它。` : `There is no direct export to ${asked} here; downloading a .pptx, which ${asked} opens directly.`)
      : (Z ? '正在下载 .pptx。' : 'Downloading the .pptx now.')
  }
  if (action === 'share_deck') {
    if (!chatslide?.share) fail(Z ? '这里不能分享：请在 ChatSlide 里打开这个项目再分享。' : 'Sharing is not available here: open this project in ChatSlide to share it.')
    if (!chatslide.canExport()) return Z ? '分享需要付费方案，已为你打开方案页面。' : 'Sharing needs a paid plan; I opened the plans for you.'
    const { url, copied } = await chatslide.share()
    return (Z ? `已开启分享：${url}` : `Sharing is on: ${url}`) + (copied ? (Z ? '（链接已复制）' : ' (link copied)') : '')
  }
  if (action === 'open_step') {
    if (!chatslide?.goto && args.step !== 'video') { // no Scripts step outside ChatSlide: the notes here ARE the script
      const targets = slides.value.map(x => x.id)
      const written = await writeNotes(targets, text)
      if (!written) fail(Z ? '演讲备注没能写出来。' : 'The speaker notes could not be written.')
      return Z ? `已写好 ${written} 页的演讲备注（脚本）。` : `Wrote the speaker notes (the script) for ${written} page(s).`
    }
    if (!chatslide?.goto) fail(Z ? '这里没有该步骤：请在 ChatSlide 里打开这个项目。' : 'That step is not available here: open this project in ChatSlide.')
    await chatslide.goto(args.step === 'video' ? 'video' : 'scripts')
    return args.step === 'video' ? (Z ? '正在打开视频步骤。' : 'Opening the Video step.') : (Z ? '正在打开脚本步骤。' : 'Opening the Scripts step.')
  }
  if (action === 'talk_to_human') {
    if (!chatslide?.support) fail(Z ? '请发邮件到 contact@chatslide.ai 联系支持团队。' : 'Please email contact@chatslide.ai to reach the support team.')
    busy.value = Z ? '正在联系支持…' : 'Contacting support…'
    return (await chatslide.support(args.message || userText)) || (Z ? '已转给支持团队。' : 'Sent to the support team.')
  }
  if (action === 'go_to_page') {
    const id = idOf(args.page)
    slidesStore.updateSlideIndex(pos(id))
    return Z ? `已切换到第 ${pos(id) + 1} 页。` : `Showing page ${pos(id) + 1}.`
  }
  if (action === 'ask_user') {
    const opts = (args.options || []).filter(Boolean).slice(0, 4)
    s.pending = { request: userText, question: args.question || '', options: opts }
    save()
    return args.question || (Z ? '能再具体一点吗？' : 'Could you say that more specifically?')
  }
  if (action === 'undo' || action === 'redo') {
    if (action === 'undo') {
      if (restoreDeckBeforeCreate()) { /* a replaced deck is back */ }
      else if (s.mode === 'template' && !s.edited && s.versions.length > 1) { // undo a template rebuild: previous version
        s.versions.pop()
        save()
        await showPptx(s.versions[s.versions.length - 1].deckUrl)
      }
      else if (snapshotStore.canUndo) await snapshotStore.unDo()
      else return Z ? '没有可以撤销的修改。' : 'There is nothing to undo yet.'
    }
    else await snapshotStore.reDo()
    return Z ? (action === 'undo' ? '已撤销。' : '已重做。') : (action === 'undo' ? 'Undone.' : 'Redone.')
  }
  return args.answer || (Z ? '能再具体一点吗？' : 'Could you say that more specifically?')
}

// Footer page numbers are static text in an imported deck: after pages move, set each to its page's position.
// A bare 1-3 digit number near the bottom, at a spot shared by most pages, is that page's number.
const renumber = () => {
  const H = slideSize().height
  const isNum = (e: any) => e.top + (e.height || 0) / 2 >= H * 0.85 && /^\d{1,3}$/.test(elText(e))
  const at = (e: any) => `${Math.round(e.left / 10)},${Math.round(e.top / 10)}`
  const spots: Record<string, number> = {}
  for (const sl of slides.value) for (const k of new Set(sl.elements.filter(isNum).map(at))) spots[k] = (spots[k] || 0) + 1
  slides.value.forEach((sl, i) => {
    for (const e of sl.elements as any[]) { // the page-number box sits at the same spot on most pages
      if (!isNum(e) || spots[at(e)] <= slides.value.length / 2 || elText(e) === String(i + 1)) continue
      const swap = (html: string) => html.replace(/>\s*\d{1,3}\s*</, `>${i + 1}<`)
      if (e.type === 'text') slidesStore.updateElement({ id: e.id, props: { content: swap(e.content) }, slideId: sl.id })
      else if (e.type === 'shape' && e.text) slidesStore.updateElement({ id: e.id, props: { text: { ...e.text, content: swap(e.text.content) } }, slideId: sl.id })
    }
  })
}

let deckBeforeCreate: Slide[] | null = null
// A new deck replaced the one in the editor: Undo (button or chat) puts that one back in a single step and keeps it
// across reloads (marked edited, so the saved slides win over a built template version).
const restoreDeckBeforeCreate = () => {
  if (!deckBeforeCreate) return false
  slidesStore.setSlides(deckBeforeCreate)
  slidesStore.updateSlideIndex(0)
  deckBeforeCreate = null
  s.edited = true
  addHistorySnapshot()
  save()
  return true
}

// Earlier turns for context: only user messages that got a reply (a request left unanswered by a closed tab must
// never be redone on the next message), the last 3 exchanges, excluding the message being handled now
const answeredHistory = () => {
  const h = s.history.slice(0, -1)
  const out: { role: 'user' | 'assistant', text: string }[] = []
  for (let i = 0; i < h.length; i++) {
    if (h[i].role === 'user' && h[i + 1]?.role !== 'assistant') continue
    out.push(h[i])
  }
  return out.slice(-6)
}

// What a read_file / web_search step found: the router reads it in "Done so far", the user never sees it raw.
let memo = ''
const fence = (text: string) => text.replace(/<\/?(file-content|web-results)\b/gi, '&lt;$1') // no closing tag inside the fence

// Agent loop: route, apply one step, report back what was done, until the router calls finish (or 8 rounds).
const deckTurn = async (text: string) => {
  userText = text
  editedThisTurn.clear()
  picturedThisTurn.clear()
  const pagesAtStart = slides.value.length
  takeSpend() // spend from an earlier, failed turn never lands in this footnote
  const pendingAtStart = s.pending || null
  const done: string[] = []
  const replies: string[] = []
  const terse: string[] = [] // per step: the edit model's own sentence can contradict the rest of a multi-step turn
  let changed = false
  for (let round = 0; round < 10; round++) {
    const { calls, error } = await (await fetch('/api/tools/deck_chat', {
      method: 'POST',
      body: JSON.stringify({
        instruction: text, pages: deckLines(), current: slideIndex.value + 1, template: !!s.template, done,
        pending: s.pending || undefined,
        files: s.docs.map(d => ({ name: d.name, pages: (d.text.match(/^\[page \d+\]$/gm) || []).length || undefined, chars: d.text.length })),
        pictures: (s.images || []).map(m => m.name),
        history: answeredHistory().map(h => ({ role: h.role, content: h.text })),
      }),
      headers: { 'Content-Type': 'application/json' },
    })).json()
    if (error) throw new Error(error)
    if (calls.some((c: any) => c.action === 'create_deck') && !done.length) {
      deckBeforeCreate = JSON.parse(JSON.stringify(slides.value)) // one Undo click brings this deck back
      busy.value = s.template ? 'Building the deck (about a minute)…' : 'Writing the deck…'
      await (s.template ? templateTurn(text) : generateNative(text))
      addHistorySnapshot() // the new deck is a history step of its own
      return
    }
    if (calls.some((c: any) => c.action === 'rebuild_from_template') && s.template) return templateTurn(text)
    // one step per round: when the model batches calls it numbers later ones as if earlier ones had run, so the rest
    // is re-planned against the updated deck in the next round
    const work = calls.filter((c: any) => !['finish', 'rebuild_from_template', 'create_deck'].includes(c.action)).slice(0, 1)
    const left = calls.find((c: any) => c.action === 'finish')?.args?.not_done
    if (!work.length && left && String(left).trim()) { // never pretend the rest is done
      replies.push(String(left).trim())
      terse.push(String(left).trim())
    }
    if (!work.length) break
    const ids = slides.value.map(x => x.id) // page numbers in this round's calls refer to this deck
    if (work.every((c: any) => done.some(d => d.startsWith(`${c.action} ${JSON.stringify(c.args || {})} ->`)))) break // repeating itself
    // the same page edited twice already this turn: a third pass is the router looping (seen: 7 edits of page 2)
    if (work.some((c: any) => c.action === 'edit_page' && done.filter(d => d.startsWith('edit_page ') && new RegExp(`"index":${Number(c.args?.index)}[,}]`).test(d)).length >= 2)) break
    // one new picture per page per message: a second one on the same page was an unasked extra (seen: a turkey).
    // Compared on the page it would really land on; skipped (and told to the router), never ending the turn.
    const landsOn = (c: any) => ids[(c.args?.page_named === false ? slideIndex.value + 1 : Number(c.args?.index)) - 1]
    const extra = work.find((c: any) => c.action === 'add_image' && picturedThisTurn.has(landsOn(c)))
    if (extra) {
      done.push(`add_image ${JSON.stringify(extra.args || {})} -> skipped: that page already got a new picture in this message`)
      continue
    }
    // one deck-wide pass per turn: after pages are deleted the numbering shifts and a re-plan deleted MORE pages
    // ("delete 25-28" took 6); a second edit_all_pages re-edits every page for nothing (30 pages, twice)
    if (work.some((c: any) => ['delete_pages', 'edit_all_pages', 'reorder_pages', 'set_background'].includes(c.action) && done.some(d => d.startsWith(`${c.action} `)))) {
      replies.push(zh(text) ? '这一轮只做一次整体删除/重排/背景/全页修改；如果还要别的改动，请再说一次。' : 'One round of deleting, reordering, background or deck-wide editing per message; if more should change, say which.')
      terse.push(replies[replies.length - 1])
      break
    }
    for (const c of work) {
      busy.value = `Working (step ${done.length + 1})…`
      let reply = ''
      try {
        reply = await runCall(text, c.action, c.args || {}, ids)
      }
      catch (e) {
        if (!(e instanceof StepFailed)) throw e
        replies.push(e.message)
        terse.push(e.message)
        round = 99 // stop: report what could not be done instead of retrying it
        break
      }
      replies.push(reply)
      terse.push(c.action === 'edit_page' ? lastEditTerse : reply)
      done.push(`${c.action} ${JSON.stringify(c.args || {})} -> ${reply}${memo ? `\n${memo}` : ''}`)
      memo = ''
      if (!['answer_question', 'ask_user', 'undo', 'redo', 'download_deck', 'read_file', 'web_search', 'find_papers', 'read_link', 'go_to_page'].includes(c.action)) changed = true
    }
    if (work.every((c: any) => ['answer_question', 'ask_user', 'go_to_page'].includes(c.action))) break // a question is asked/answered once; showing a page ends the turn
  }
  if (pendingAtStart && s.pending === pendingAtStart) { // answered; but never wipe a NEW question asked this turn
    s.pending = null
    save()
  }
  if (changed) {
    deckBeforeCreate = null // after another edit, Undo steps back normally
    renumber()
    addHistorySnapshot()
    if (s.mode !== 'native') s.edited = true // a template deck or a deck opened in the editor: keep the edited copy
  }
  const lines = terse.length > 1 ? replies.map((r, i) => terse[i] ?? r) : replies // one step: the model's sentence is fine
  // counts quoted by earlier steps go stale when a later step adds or removes pages: the final size closes the turn
  if (slides.value.length !== pagesAtStart) lines.push(zh(text) ? `现在共 ${slides.value.length} 页。` : `The deck now has ${slides.value.length} pages.`)
  const reply = lines.filter(Boolean).join('\n') || (zh(text) ? '已完成。' : 'Done.')
  const sp = takeSpend()
  const fee = sp.charged > 0 ? (zh(text) ? `\n（用了 ${sp.charged} 积分${sp.left >= 0 && sp.left <= 9000000 ? `，剩 ${sp.left}` : ''}）` : `\n(Used ${sp.charged} credit${sp.charged > 1 ? 's' : ''}${sp.left >= 0 && sp.left <= 9000000 ? ` · ${sp.left} left` : ''})`) : ''
  say('assistant', await finalize(text, reply, done) + fee)
}

// One small call on every turn: the reply in the user's language, plus any part of the request that none of the steps
// covered, said out loud (a Haiku router sometimes calls finish without naming what it skipped).
const finalize = async (instruction: string, reply: string, steps: string[]) => {
  try {
    // the steps (tool, arguments, result) are the evidence of what was done; the reply alone is too terse to judge
    const evidence = steps.map(d => d.split('\n')[0].slice(0, 300))
    const r = await (await fetch('/api/tools/deck_chat', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ localize: reply, instruction, steps: evidence }) })).json()
    void steps // the generated "not done" sentence was wrong more often than right (audited 2026-11-02); the router's not_done and the code's own replies carry that
    return r.reply || reply
  }
  catch {
    return reply
  }
}

const send = () => {
  const text = input.value.trim()
  input.value = ''
  say('user', text)
  // every message goes through the router and edits the deck in the editor; only an explicit request for a new
  // presentation builds one (create_deck), and that stays undoable
  const before = JSON.parse(JSON.stringify(slides.value))
  return run('Working…', async () => {
    try {
      await deckTurn(text)
      save()
    }
    catch (e) { // never leave a half-applied operation (e.g. an inserted page that was never rewritten)
      slidesStore.setSlides(before)
      throw e
    }
  })
}

// One attach control (or drop files on the box): a .pptx becomes the company template while none is set and no deck
// was built yet; everything else is material
const attach = (files?: FileList | File[] | null) => {
  if (busy.value || !files) return
  for (const f of Array.from(files)) {
    if (/^image\//.test(f.type)) uploadPicture(f)
    else if (/\.pptx$/i.test(f.name) && !s.template && !s.mode) uploadTemplate(f)
    else uploadContent(f)
  }
}

// Enter sends, Shift+Enter is a new line; Enter while an IME is composing (picking Chinese characters) never sends
const onEnter = (e: KeyboardEvent) => {
  if (e.shiftKey || e.isComposing || e.keyCode === 229) return
  e.preventDefault()
  if (!busy.value && input.value.trim()) send()
}

const download = (): 'started' | 'dialog' | 'plans' => {
  if (chatslide && !chatslide.canExport()) return 'plans' // ChatSlide has opened its plans
  if (s.mode === 'template' && !s.edited) { // the engine's own .pptx is exact; after editor changes, export the editor
    const a = document.createElement('a')
    a.href = s.versions[s.versions.length - 1].deckUrl
    a.download = 'deck.pptx'
    a.click()
    return 'started'
  }
  mainStore.setDialogForExport('pptx')
  return 'dialog'
}

const restart = () => {
  Object.assign(s, blank())
  localStorage.removeItem(SLIDES_KEY)
  save()
}

// ---- ChatSlide: the deck goes back into the project (each page as a picture, notes as its script) so Scripts,
// Video and sharing use it. Only pages that changed since the last sync are rendered and uploaded.
const renderRef = ref<HTMLElement>()
const toRender = ref<Slide[]>([])
const hashOf = (x: string) => {
  let h = 5381
  for (let i = 0; i < x.length; i++) h = ((h * 33) ^ x.charCodeAt(i)) >>> 0
  return h.toString(36)
}
const plain = (html = '') => new DOMParser().parseFromString(html, 'text/html').body.textContent?.trim() || ''
const plainText = (sl: Slide) => sl.elements
  .map((el: any) => el.type === 'text' ? plain(el.content) : el.type === 'shape' ? plain(el.text?.content) : '')
  .filter(Boolean).join('\n')
const pageKeys = (deck: Slide[]) => deck.map(sl => `${sl.id}-${hashOf(JSON.stringify(sl))}`)
let syncing: Promise<void> | null = null
const syncOnce = async () => {
  const deck: Slide[] = JSON.parse(JSON.stringify(slides.value)) // what is sent is what was rendered
  const keys = pageKeys(deck)
  if (s.synced && s.synced.keys.join() === keys.join()) return
  const urls = s.synced?.urls || {}
  const need = deck.filter((_, i) => !urls[keys[i]])
  const images: Record<string, Blob> = {}
  if (need.length) {
    toRender.value = need
    await nextTick()
    await new Promise(r => setTimeout(r, 400)) // images inside the pages load
    const nodes = [...(renderRef.value?.children || [])] as HTMLElement[]
    for (const [j, sl] of need.entries()) {
      nodes[j].querySelectorAll('foreignObject [xmlns]').forEach(e => e.removeAttribute('xmlns'))
      const dataUrl = await toJpeg(nodes[j], { quality: 0.9, width: 1600, fontEmbedCSS: '' })
      images[keys[deck.indexOf(sl)]] = await (await fetch(dataUrl)).blob()
    }
    toRender.value = []
  }
  const got = await chatslide!.sync(deck.map((sl, i) => ({
    key: keys[i], url: urls[keys[i]], image: images[keys[i]], text: plainText(sl), notes: plain(sl.remark),
  })))
  s.synced = { keys, urls: Object.fromEntries(keys.map(k => [k, urls[k] || got[k]])) }
  save()
}
// Repeats until the project has the deck as it is now: edits made during a sync are sent by the next round
// (a caller that joins a running sync, e.g. flush before Scripts, waits for that too).
const syncToChatSlide = () => syncing ||= (async () => {
  try {
    for (let round = 0; round < 5; round++) {
      await syncOnce()
      if (s.synced?.keys.join() === pageKeys(slides.value).join()) break
    }
  }
  finally {
    syncing = null
    toRender.value = []
  }
})()

// Bring the deck back after a page reload.
onMounted(() => {
  if (chatslide) { // ChatSlide project: its saved editor state, or on first open its deck converted from ChatSlide
    if (chatslide.title) slidesStore.setTitle(chatslide.title)
    chatslide.flush = () => syncToChatSlide() // before Scripts: the latest pages first
    run('Opening your deck…', async () => {
      const saved = await chatslide!.load()
      if (saved?.slides?.length) {
        Object.assign(s, blank(), saved.session || {})
        if (saved.theme) slidesStore.setTheme(saved.theme)
        slidesStore.setSlides(saved.slides)
        slidesStore.updateSlideIndex(0)
      }
      else {
        const link = await chatslide!.pptx()
        const loaded = new Promise<void>(res => {
          const stop = watch(slides, () => {
            stop()
            setTimeout(res, 1500) // let the import settle
          })
          setTimeout(res, 20000)
        })
        await showPptx(link)
        await loaded
        s.edited = true // the editor copy is the deck from now on
        s.synced = { keys: pageKeys(slides.value), urls: {} } // unchanged: the project keeps its own slides until an edit
        URL.revokeObjectURL(link)
      }
      watch([slides, theme], save, { deep: true }) // manual edits are saved too
    })
    return
  }
  // opened from ChatSlide ("Edit in advanced editor"): ?pptx=<blob url> loads that deck as a fresh session
  const link = new URLSearchParams(location.search).get('pptx')
  let host = ''
  try {
    host = link ? new URL(link).hostname : ''
  }
  catch { /* not a URL */ }
  if (link && host.endsWith('.public.blob.vercel-storage.com')) {
    history.replaceState(null, '', location.pathname)
    run('Opening your deck…', async () => {
      await showPptx(link)
      Object.assign(s, blank()) // only once the deck is in: a failed import keeps the previous session
      localStorage.removeItem(SLIDES_KEY)
      save()
      deckApi('drop', { url: link }).catch(() => undefined) // the one-time copy is not kept in public storage
    })
    return
  }
  let saved: Slide[] | null = null
  try {
    saved = (s.mode === 'native' || s.edited) ? JSON.parse(localStorage.getItem(SLIDES_KEY) || 'null') : null
  }
  catch { /* nothing saved */ }
  if (saved?.length) {
    slidesStore.setSlides(saved)
    try {
      const t = JSON.parse(localStorage.getItem(`${SLIDES_KEY}-theme`) || 'null')
      if (t) slidesStore.setTheme(t)
    }
    catch { /* no saved theme */ }
  }
  else if (s.mode === 'template' && s.versions.length) {
    s.edited = false // the editor copy was lost; the last built version is what is shown
    run('Loading your deck…', () => showPptx(s.versions[s.versions.length - 1].deckUrl))
  }
})
</script>

<style lang="scss" scoped>
.chatslide-render {
  position: fixed;
  left: -100000px;
  top: 0;
}
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
.input-box {
  border: 1px solid $borderColor;
  border-radius: $borderRadius;
  padding: 4px;
  :deep(.textarea) {
    border: 0;
    box-shadow: none;
  }
}
.input-row {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 0 4px;
}
.attach {
  cursor: pointer;
  font-size: 18px;
  color: #666;
  display: flex;
  &.disabled {
    opacity: .4;
    pointer-events: none;
  }
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
.choices {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin: 2px 0 4px;

  .choice {
    border: 1px solid $themeColor;
    color: $themeColor;
    border-radius: $borderRadius;
    padding: 3px 10px;
    font-size: 12px;
    cursor: pointer;

    &:hover {
      background-color: rgba($color: $themeColor, $alpha: .08);
    }
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
