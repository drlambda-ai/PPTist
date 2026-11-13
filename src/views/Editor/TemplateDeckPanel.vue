<template>
  <component
    :is="inline ? 'div' : MoveablePanel"
    :dir="direction"
    class="template-deck-panel"
    :class="{ inline }"
    :key="chatLeft"
    v-bind="inline ? {} : { width: 320, height: 560, title: chatslide ? 'ChatSlide' : 'AI', left: chatLeft, top: 84 }"
    @close="close()"
  >
    <div class="body">
      <div class="messages" ref="messagesRef">
        <div class="hint" v-if="!s.history.length">
          {{ t("Hi! Ask me to change anything in this deck: text, pictures, charts, pages or language. Attach your company's PowerPoint template or a PDF, Word, Excel or Markdown file with 📎 and I'll build from it.") }}
        </div>
        <div class="msg" :class="m.role" v-for="(m, i) in s.history" :key="i">{{ m.text }}<div class="rate" v-if="m.turn && chatslide?.feedback && (m.rating !== undefined || i === s.history.length - 1)"><span :class="{ on: m.rating === 1 }" :aria-label="t('Good reply')" @click="rate(m, true)"><i-icon-park-outline:thumbs-up /></span><span :class="{ on: m.rating === 0 }" :aria-label="t('Bad reply')" @click="rate(m, false)"><i-icon-park-outline:thumbs-down /></span></div><input class="why" v-if="m.rating === 0 && m.comment === undefined" v-model="why" :placeholder="t('What went wrong? (optional)')" @keydown.enter="comment(m)" @blur="comment(m)" /></div>
        <div class="msg assistant loading" v-if="busy">{{ busy }}</div>
        <div class="choices" v-if="!busy && s.pending?.options?.length">
          <span class="choice" v-for="o in s.pending.options" :key="o" @click="input = o; send()">{{ o }}</span>
        </div>
      </div>

      <div class="composer">
        <div class="chips" v-if="s.template || s.docs.length || s.images?.length">
          <span class="chip" v-if="s.template">{{ t('Template:') }} {{ s.template.name }}
            <i @click="removeTemplate()" v-if="!busy && !s.mode">×</i></span>
          <span class="chip" v-for="(m, i) in s.images || []" :key="'img' + m.name + i">🖼 {{ m.name }}
            <i @click="s.images!.splice(i, 1); save()" v-if="!busy">×</i></span>
          <span class="chip" v-for="(d, i) in s.docs" :key="d.name + i">{{ d.name }}
            <i @click="s.docs.splice(i, 1); save()" v-if="!busy">×</i></span>
        </div>
        <div class="input-box" @dragover.prevent @drop.prevent="e => attach(e.dataTransfer?.files)">
          <TextArea v-model:value="input" :rows="inline ? 2 : 3" :placeholder="t('e.g. Make page 3 shorter / Add a page on Q3 risks / Use the attached data')" @enter="onEnter" />
          <div class="input-row">
            <FileInput accept=".pptx,.xlsx,.csv,.txt,.md,.docx,.pdf,.json,.png,.jpg,.jpeg,.webp,.gif" @change="files => attach(files)">
              <span class="attach" :class="{ disabled: !!busy }" :aria-label="t('Attach a company template (.pptx) or material')"><i-icon-park-outline:link /></span>
            </FileInput>
            <span class="attach" v-if="chatslide?.versions" :class="{ disabled: !!busy }" :aria-label="t('Version history')" :title="t('Version history')" @click="showVersions()"><i-icon-park-outline:history /></span>
            <span class="grow"></span>
            <Button type="primary" size="small" :disabled="!!busy || !input.trim()" @click="send()">{{ t('Send') }}</Button>
          </div>
        </div>
        <!-- free users see what a paid plan adds; a tap opens ChatSlide's plans -->
        <div class="paid-strip" v-if="chatslide?.offerPaid && !chatslide.canUseAdvanced?.()" @click="chatslide.offerPaid()">
          <span>{{ t('Paid plans: advanced AI model · real pictures from your brand’s website · download as PowerPoint or PDF') }}</span>
          <b>{{ t('Upgrade') }}</b>
        </div>
        <div class="new" v-if="s.mode && !busy"><span @click="restart()">{{ t('Start a new deck') }}</span></div>
      </div>
    </div>
  </component>
  <!-- ChatSlide: pages being turned into pictures (off-screen, outside the panel so a closed panel still renders) -->
  <Teleport to="body">
    <div class="chatslide-render" ref="renderRef" v-if="toRender.length">
      <ThumbnailSlide v-for="sl in toRender" :key="sl.id" :slide="sl" :size="1600" />
    </div>
  </Teleport>
</template>

<script lang="ts" setup>
import { PRESET_THEMES } from '@/configs/theme'
import { t, direction, lang } from '@/i18n'

defineProps<{ inline?: boolean }>() // on a phone: part of the page, not a floating window
// The chat opens at the right edge, over the Design sidebar: in the middle it covered half the slide on a laptop and
// users could not see what it changed (2026-11-10). Asking for the Design panel moves it to the left edge.
const chatLeft = ref(-8)

import { nextTick, onMounted, reactive, ref, watch } from 'vue'
import { host as chatslide, currentTurn, newTurn, OutOfCredits, takeSpend } from '@/embed'
import { storeToRefs } from 'pinia'
import { nanoid } from 'nanoid'
import { isPC } from '@/utils/common'
import { useMainStore, useSlidesStore, useSnapshotStore } from '@/store'
import { db } from '@/utils/database'
import { ToolbarStates } from '@/types/toolbar'
import { saveAs } from 'file-saver'
import { jpegsToPdf, type JpegPage } from '@/utils/imagesToPdf'
// @ts-expect-error no declaration for the bundle path (see utils/normalizePptx.ts)
import JSZip from 'jszip/dist/jszip.min.js'
import useExport from '@/hooks/useExport'
import useImport from '@/hooks/useImport'
import useAIPPT from '@/hooks/useAIPPT'
import useSlideHandler from '@/hooks/useSlideHandler'
import useHistorySnapshot from '@/hooks/useHistorySnapshot'
import useSlideTheme from '@/hooks/useSlideTheme'
import api from '@/services'
import type { Slide, PPTImageElement, PPTElement } from '@/types/slides'
import type { AIPPTSlide } from '@/types/AIPPT'
import { jsonrepair } from 'jsonrepair'
import { upload } from '@/utils/blobUpload'
import Button from '@/components/Button.vue'
import TextArea from '@/components/TextArea.vue'
import FileInput from '@/components/FileInput.vue'
import ThumbnailSlide from '@/views/components/ThumbnailSlide/index.vue'
import { getFontEmbedCSS, toJpeg } from 'html-to-image'
import MoveablePanel from '@/components/MoveablePanel.vue'

// One AI panel for both paths:
// - with a company template: the deck server (api/deck.py) fills the template natively, every version is a real
//   .pptx in the template; the editor shows it through the .pptx importer;
// - without one: PPTist's own AI deck generation. Edits after that go through a ChatSlide-style tool router
//   (api/tools/deck_chat.ts) applied to the editor's slides; slide rewrites use chat_edit.

interface Session {
  mode: '' | 'template' | 'native'
  template: { name: string, templateUrl: string, analysis: any, profile: any, chrome: any[] } | null
  docs: { name: string, text: string, project?: boolean }[] // project: a file the user gave the project in ChatSlide
  projectFiles?: boolean // the project's files were added once (a removed one stays removed)
  images?: { name: string, url: string }[] // pictures the user attached, placed on slides by add_image source "attached"
  // turn: the chat message this reply answers (its x-turn-id); rating 1 up, 0 down; comment: sent with a thumbs-down
  history: { role: 'user' | 'assistant', text: string, turn?: string, rating?: 0 | 1, comment?: string }[]
  versions: { spec: any, deckUrl: string, h: number }[]
  edited?: boolean // template deck changed in the editor since the last build
  synced?: { keys: string[], urls: Record<string, string> } // ChatSlide: pages last sent to the project, and their pictures
  // offer: a button for an optional next step; it reaches the router only when the user taps it (an untouched offer
  // pulled the next, unrelated message onto its page, 2026-11-11). A question (ask_user) holds for any reply.
  pending?: { request: string, question: string, options?: string[], done?: string[], offer?: boolean, localize?: boolean, repeat?: { action: string, args: any }[], pages?: number[], from?: number } | null // an ask_user waiting for the user's answer; done: steps already run for it
  removeAsked?: string // pages (ids, joined) the user was asked to confirm removing pictures from
  deleteAsked?: string[] // pages the user was asked to confirm deleting; only the next message can delete them
  savedVersions?: Record<string, string> // version history offered in the chat: button label -> that saved copy
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
// eval hook (scripts/deck-eval): a replayed real conversation's earlier messages go straight into the chat history
;(window as any).__deckPanel = { s, resetUndo: async () => {
  await db.snapshots.clear(); await snapshotStore.initSnapshotDatabase()
} }
const input = ref('')
const busy = ref('')
const messagesRef = ref<HTMLElement>()
let blobMode = ''

const SLIDES_KEY = 'ai-deck-slides'
let saveTimer = 0
let savePending = false // a debounced save has not reached ChatSlide yet
let syncTimer = 0
let saveFailed = false
let loadingPage = '' // while a project's deck is being converted: never saved, or the blank page replaces the deck
const save = () => {
  if (chatslide && loadingPage && slides.value.map(x => x.id).join() === loadingPage) return
  if (chatslide) { // per project, debounced; also after manual edits (watch below)
    clearTimeout(saveTimer)
    clearTimeout(syncTimer)
    syncTimer = window.setTimeout(() => syncToChatSlide().catch(() => undefined), 15000)
    savePending = true
    saveTimer = window.setTimeout(async () => {
      savePending = false
      await externalizeImages().catch(() => undefined) // pictures saved as files: the deck file stays small and loads fast
      chatslide!.save({ session: JSON.parse(JSON.stringify(s)), slides: slides.value, theme: theme.value, viewportRatio: viewportRatio.value })
        .then(() => saveFailed = false)
        .catch(() => { // said once (saying saves again): the next edit retries
          if (!saveFailed) say('assistant', t('Saving failed; your last edits may not be kept. Check your connection.'))
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
// The deck on screen saved now (a save still waiting on its timer goes at once): restoring a saved version first keeps
// the current deck as the newest copy, so a restore never loses work (a user's 21 pages went to 17 with no way back,
// 2026-11-12). Throws when the save fails.
const saveNow = async () => {
  if (!chatslide || !savePending) return
  clearTimeout(saveTimer)
  savePending = false
  await externalizeImages().catch(() => undefined)
  await chatslide.save({ session: JSON.parse(JSON.stringify(s)), slides: slides.value, theme: theme.value, viewportRatio: viewportRatio.value })
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
        const src = e.src
        const url = await swap(src)
        // only the picture that was uploaded: a newer one placed while it uploaded stays
        const live: any = slides.value.find(x => x.id === sl.id)?.elements.find(x => x.id === e.id)
        if (url !== src && live?.src === src) slidesStore.updateElement({ id: e.id, props: { src: url }, slideId: sl.id })
      }
    }
    const bg = sl.background as any
    if (bg?.type === 'image' && bg.image?.src?.startsWith('data:image/')) {
      const src = bg.image.src
      const url = await swap(src)
      const live = slides.value.find(x => x.id === sl.id)?.background as any
      if (url !== src && live?.image?.src === src) slidesStore.updateSlide({ background: { ...live, image: { ...live.image, src: url } } }, sl.id)
    }
  }
}

// The full-slide background picture keeps its box: "make the picture bigger" also stretched it to the canvas.
// Backgrounds change through set_background / change_image(image 0), never as a side effect of a page edit; the one
// exception is a page whose only picture it is, where "the picture" can mean nothing else.
const keepUnaskedSizes = (i: number, was: PPTElement[]) => {
  const sl = slides.value[i]
  if ((was as any[]).filter(e => e.type === 'image').length < 2) return
  for (const old of was as any[]) {
    const now = sl.elements.find(e => e.id === old.id) as any
    if (now && old.type === 'image' && isBackdrop(old) && ['left', 'top', 'width', 'height'].some(k => now[k] !== old[k])) {
      slidesStore.updateElement({ id: old.id, props: { left: old.left, top: old.top, width: old.width, height: old.height }, slideId: sl.id })
    }
  }
}

let replyTurn = '' // the chat message being answered: its replies get the thumbs, tied to its turn id
const say = (role: 'user' | 'assistant', text: string) => {
  s.history.push(role === 'assistant' && replyTurn ? { role, text, turn: replyTurn } : { role, text })
  save()
  nextTick(() => {
    if (messagesRef.value) messagesRef.value.scrollTop = messagesRef.value.scrollHeight 
  })
}
const { applyPresetTheme } = useSlideTheme()
// a new background colour shows through the page's full-slide picture; the picture stays, faded, and undo brings it back
const fadeBackdrop = (id: string) => {
  const sl = slides.value.find(x => x.id === id)
  for (const e of (sl?.elements || []) as any[]) {
    if (e.type === 'image' && isBackdrop(e)) slidesStore.updateElement({ id: e.id, props: { filters: { ...e.filters, opacity: '12%' } }, slideId: id })
  }
}
const hostSteps = () => chatslide || (window as any).__evalHost as typeof chatslide // the eval stands in for ChatSlide's own steps
// a dropped connection on an AI call is retried once: it showed the user "Failed to fetch" for nothing else
if (!(window as any).__deckFetchRetry) {
  (window as any).__deckFetchRetry = true
  const fetch0 = window.fetch.bind(window)
  window.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
    try {
      return await fetch0(input, init)
    }
    catch (e) {
      if (!(e instanceof TypeError) || !String((input as any).url || input).includes('/api/tools/')) throw e
      await new Promise(r => setTimeout(r, 1500))
      return fetch0(input, init)
    }
  }
}
const close = () => mainStore.setTemplateDeckPanelState(false)

const run = async (label: string, fn: () => Promise<void>) => {
  busy.value = label
  try {
    await fn()
  }
  catch (e) {
    // the steps finished before the failure are kept and saved: every message says so, or a retry redoes paid work
    const kept = !!(e as any)?.keptSteps
    const keptLine = kept ? ` ${t('The changes made before it are kept.')}` : ''
    if (e instanceof OutOfCredits) say('assistant', t("You're out of credits (this needs {p0}, you have {p1}). Upgrade your plan to keep going.", { p0: e.required, p1: e.credits }) + keptLine)
    else if (/\b(Anthropic|OpenAI) \d{3}\b/.test(String((e as Error).message))) {
      say('assistant', kept ? t('The AI service is unavailable right now. Please try again in a moment.') + keptLine : t('The AI service is unavailable right now; nothing was changed. Please try again in a moment.'))
    }
    // a bug of ours, never shown raw to the user (it is in the turn record with its stack)
    else if (e instanceof SyntaxError || e instanceof TypeError || e instanceof ReferenceError || e instanceof RangeError) {
      say('assistant', (e as any).keptSteps
        ? t('Something went wrong on our side partway through; the changes made before it are kept. Please ask for the rest again.')
        : t('Something went wrong on our side; nothing was changed. Please try again.'))
    }
    else say('assistant', t('That failed: ') + ((e as Error).message || e) + keptLine)
  }
  lastGoodSlides = null // the copy is only for the turn that just ended
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

const uploadTemplate = (file: File) => run(t('Learning the template (about a minute)…'), async () => {
  const templateUrl = await uploadFile(file)
  const learned = await deckApi('learn', { templateUrl })
  s.template = { name: file.name, ...learned }
  save()
  await showPptx(templateUrl)
  say('assistant', t('Template ready: {p0} layouts. The deck will be built on it.', { p0: learned.analysis.length }))
})

const removeTemplate = () => {
  s.template = null
  save()
}

// The user's own picture: uploaded once, then placed by the chat ("put my photo on slide 3")
const uploadPicture = (file: File) => run(t('Uploading {p0}…', { p0: file.name }), async () => {
  // re-encoded through a canvas: only raster pixels are published (never SVG/HTML or a spoofed type), at most 2000 px
  const bmp = await createImageBitmap(file).catch(() => null)
  if (!bmp) fail(t('{p0} is not a picture I can use.', { p0: file.name }))
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
const uploadContent = (file: File) => run(t('Reading {p0}…', { p0: file.name }), async () => {
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
// A file attached in the chat is the source; the project's own files (the deck was made from them) where they fit
const materialBrief = () => `${s.docs.every(d => d.project) ? 'Material from the user\'s files, to use where it fits this request' : 'Use only this material'} (never invent numbers):\n${material()}`
// What the generators write in: a new deck follows the request's language; pages added later follow the DECK's language
// (a Chinese deck stays Chinese when someone types an English or Russian request), unless the request names a language.
const requestLanguage = (text: string) => `the language of this request: "${text.slice(0, 120)}"`
const deckLanguage = (text: string) => `the language of the existing deck (sample: "${deckLines().join(' ').slice(0, 160)}"), unless this request names another language: "${text.slice(0, 120)}"`

const generateNative = async (text: string) => {
  const brief = s.docs.length ? `${text}\n\n${materialBrief()}` : text
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
  say('assistant', t('Generated {p0} pages. Ask for any change: add, delete, move or rewrite pages.', { p0: slides.value.length }))
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
// A picture covering (nearly) the whole slide is its backdrop (ChatSlide templates draw backgrounds as pictures).
// "The picture" means the content pictures; the backdrop only when it is the only picture on the page.
const isBackdrop = (e: any) => {
  const { width: W, height: H } = slideSize()
  return e.width * e.height >= 0.8 * W * H
}
const firstBackdrop = (sl: Slide) => sl.elements.find((e: any) => e.type === 'image' && isBackdrop(e)) as PPTImageElement | undefined
// No content pictures: the page's picture is the lowest full-slide one. A full-slide picture above it can only be a
// see-through design layer (an opaque one would hide it): "replace all pictures on slide 1" put a photo over the cover
// photo, hiding it (2026-11-10).
const pictures = (sl: Slide) => {
  const content = readingOrder(sl).filter((e): e is PPTImageElement => e.type === 'image' && !isBackdrop(e))
  const backdrop = firstBackdrop(sl)
  return content.length ? content : backdrop ? [backdrop] : []
}
// One numbering everywhere (page list, selection, change_image, remove_images, "already replaced"): content pictures are
// 1.., the full-slide picture 0; on a page with no content picture the lowest full-slide picture is 1, and 0 means it too
// (counted as two pictures, the cover photo was replaced twice in one message).
const pictureNo = (sl: Slide, image: number | undefined) => (image === 0 && !pictures(sl).some(e => !isBackdrop(e)) && firstBackdrop(sl) ? 1 : (image ?? 1))
// A page that is one flat picture: ChatSlide's pages rendered to one image (Blank_layout background_url, which the
// editor's first open imports as the slide's background picture) with nothing on it. Its text and pictures are inside
// that image, so none can be changed on its own, and replacing or repainting the background wipes the page. A photo
// slide made in the editor is a picture element, and stays a normal picture.
const isFlat = (sl: Slide) => sl.background?.type === 'image' && !sl.elements.some((e: any) => e.type !== 'shape' || elText(e))

const pageText = (sl: Slide, common = new Set<string>(), logos = new Set<string>()) => {
  if (isFlat(sl)) return '[image 1: one picture covering the whole page; the page has no separate text]'
  const numbered = pictures(sl) // same numbering change_image uses
  const backdrop = firstBackdrop(sl)
  const biggest = Math.max(0, ...numbered.filter(e => !logos.has(imgSpot(e))).map((e: any) => e.width * e.height))
  return readingOrder(sl).map((e: any) => {
    if (e.type === 'image') { // size and place let the router tell a photo from an icon or a logo
      const { width: W, height: H } = slideSize()
      const area = (e.width * e.height) / (W * H)
      const n = numbered.indexOf(e) + 1
      if (!n) return e === backdrop ? '[image 0: background picture covering the whole slide]' : '[a see-through design layer covering the whole slide, not a picture]'
      const img = n - 1
      if (logos.has(imgSpot(e))) return `[image ${img + 1}: logo]`
      const kind = `${area < 0.005 ? 'icon' : area < 0.04 ? 'small image' : 'large image'}${e.width * e.height === biggest ? ', the largest image here' : ''}`
      const where = `${e.top + e.height / 2 < H / 3 ? 'top' : e.top + e.height / 2 > (2 * H) / 3 ? 'bottom' : 'middle'}-${e.left + e.width / 2 < W / 3 ? 'left' : e.left + e.width / 2 > (2 * W) / 3 ? 'right' : 'centre'}`
      return `[image ${img + 1}: ${kind}, ${Math.round(area * 100)}% of slide, ${where}]`
    }
    if (e.type === 'chart') return `[${e.chartType} chart: ${(e.data?.labels || []).slice(0, 8).join(', ')}]` // an existing chart, editable in place
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
// page|query of each NEW picture in the current message: the same picture is not added twice, but "examples of A, B, C"
// adds one per item (a one-per-page rule cut a doctor's four clinical examples to one, 2026-11-11)
const picturedThisTurn = new Set<string>()
const roomMadeThisTurn = new Set<string>() // pages whose text was moved aside for a picture in the current message
const picKey = (page: string, query: unknown) => `${page}|${String(query || '').trim().toLowerCase()}`
const replacedThisTurn = new Set<string>() // "<page id>:<picture number>" replaced in the current message
// The words the router quoted as the user asking for the background, kept only when the user wrote them in THIS
// message, or in the previous one when that message changed the background ("again"). The router answered "the
// background: yes" for "replace these photos" and swapped the background (2026-11-10, again in the 2026-11-13 replay);
// an older "make the background blue" must not license it. The router still decides what the words mean.
const saidBackground = (text: string, words: unknown) => {
  const norm = (x: string) => x.toLowerCase().replace(/\s+/g, ' ').trim()
  const w = typeof words === 'string' ? norm(words) : ''
  if (!w) return false
  const users = s.history.filter(h => h.role === 'user').map(h => String(h.text || ''))
  const prev = users[users.length - 1] === text ? users[users.length - 2] : users[users.length - 1]
  const prevChangedBackground = lastTurnCalls.some(k => k.startsWith('set_background ') || (k.startsWith('change_image ') && /"image":0[,}]/.test(k)))
  return [text, ...(prevChangedBackground && prev ? [prev] : [])].some(x => norm(x).includes(w))
}
let currentAtSend = '' // id of the page the user was on when they sent the message
const createdThisMessage = new Set<string>() // pages added since the message was sent
let lastTurnCalls: string[] = [] // the previous message's steps, as stepKey: what an Undo just reverted
// a step's TARGET (action + where), not its wording: a redo of the undone change rarely repeats the query word for word
const stepKey = (action: string, args: any) => `${action} ${JSON.stringify(Object.fromEntries(Object.entries(args || {})
  .filter(([k]) => !['answer', 'source'].includes(k)).sort()))}` // what to change stays in: "no, red instead" is a new step
// text boxes that run into a picture or into another plain text box (a paid redesign left a heading on top of the
// card text, 2026-11-11), as "a|b" keys. Text on a card shape is a design, not a collision.
const textOnPictures = (els: any[]) => {
  const pics = els.filter((e: any) => (e.type === 'image' && !isBackdrop(e)))
  const texts = els.filter((e: any) => e.type === 'text' && elText(e))
  const tt = texts.flatMap((a: any, i: number) => texts.slice(i + 1).filter((b: any) => {
    const w = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
    const h = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top)
    return w > 0 && h > 0 && w * h > 0.15 * Math.min(a.width * a.height, b.width * b.height)
  }).map((b: any) => [a.id, b.id].sort().join('|')))
  return new Set([...tt, ...els.filter((e: any) => e.type !== 'image' && elText(e)).flatMap((a: any) => pics.filter((b: any) => {
    const w = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
    const h = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top)
    return w > 0 && h > 0 && w * h > 0.05 * Math.min(a.width * a.height, b.width * b.height)
  }).map((b: any) => `${a.id}|${b.id}`))])
}
// what overlaps on the page, measured: the model reads boxes badly, and a user complaining about it is seeing it
const overlapNote = (sl: Slide) => {
  const els = readingOrder(sl).filter((e: any) => (e.type === 'image' && !isBackdrop(e)) || elText(e))
  const pairs: string[] = []
  els.forEach((a: any, i) => els.slice(i + 1).forEach((b: any) => {
    const w = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
    const h = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top)
    const small = Math.min(a.width * a.height, b.width * b.height)
    if (w > 0 && h > 0 && small > 0 && w * h > 0.1 * small) pairs.push(`${a.id} and ${b.id}`)
  }))
  return pairs.length ? `\n(Measured on this page, these elements overlap: ${pairs.slice(0, 8).join('; ')}. Fix them when the request is about layout, overlap, or text that is covered or cut off.)` : ''
}

// Out of credits for the advanced model: the standard one (no credit) does the edit, without the paywall.
// fallback: out of credits for the advanced model, use the standard one. Off for a redesign: the standard model's
// "redesign" is a few pixels (backtest 2026-11-11), so the user is told about the credits instead.
let lastEditAdvanced = false // the model the last editSlide actually ran: callers decide retries on this, not re-derive it
const editSlide = async (index: number, instruction: string, advanced = false, fallback = true) => {
  lastEditAdvanced = false
  if (advanced) {
    try {
      const r = await editSlideWith(index, instruction, true)
      lastEditAdvanced = true
      return r
    }
    catch (e) {
      if (!(e instanceof OutOfCredits)) { lastEditAdvanced = true; throw e }
      if (!fallback) throw e
    }
  }
  return editSlideWith(index, instruction, false)
}
const editSlideWith = async (index: number, instruction: string, advanced: boolean) => {
  const slide = slides.value[index]
  // reading order so "the first point" means the top-left one; the merge below keeps the slide's own stacking order
  // a shape's words (every text box of an imported ChatSlide/PowerPoint deck) are shown like a text box's: the model
  // edits "content" the same way on both, and the merge puts it back inside the shape's text
  const flat = (e: any) => (e.type === 'shape' && e.text ? (({ text, ...r }) => ({ ...r, content: text.content, defaultColor: text.defaultColor, defaultFontName: text.defaultFontName }))(e) : e)
  // the model copies what it sees, so a new box often comes back as a "shape" with "content": put the words where
  // the editor draws them (a shape's text, or a plain text box when it has no outline)
  const unflat = (e0: any) => {
    const e = { ...e0, ...Object.fromEntries(GEOMETRY.filter(k => k in e0).map(k => [k, Number.isFinite(toNum(e0[k])) ? toNum(e0[k]) : 0])) }
    delete e.layer // a stacking instruction, applied by the merge
    if (e.type !== 'shape' || e.content === undefined) return e
    const { content, defaultColor, defaultFontName, ...rest } = e
    if (!rest.path) return { ...rest, type: 'text', content, defaultColor: defaultColor || '#333', defaultFontName: defaultFontName || '', rotate: rest.rotate || 0 }
    return { ...rest, text: { content, defaultColor: defaultColor || '#333', defaultFontName: defaultFontName || '', align: 'middle' } }
  }
  const elements = readingOrder(slide).map((e: any) => (e.type === 'image' ? { ...e, src: `image:${e.id}` }
    : e.type === 'shape' && e.path?.length > 80 ? { ...flat(e), path: `path:${e.id}` } : flat(e))) // long SVG paths: token, restored below
  const r = await fetch('/api/tools/chat_edit', {
    method: 'POST', headers: { 'Content-Type': 'application/json', ...(advanced ? { 'X-Quiet-402': '1' } : {}) },
    body: JSON.stringify({ patch: true, advanced, elements, slideSize: slideSize(), instruction: `${instruction}${overlapNote(slide)}\n(If you change nothing, the "reply" sentence says why, e.g. what you found on the page; never claim a change you did not make. Add "too_hard": true only when the request applies but you could not do it well; nothing needed is not too hard. Write the "reply" sentence in the language of this message: "${userText.slice(0, 120)}")` }),
  })
  const ret = await r.json()
  if (!ret.patch) fail(ret.reply || ret.error || t('That edit could not be applied.'))
  const unchanged = () => Object.assign(new NoChange(editedThisTurn.has(slide.id)
    ? (t('No further change on page {p0}.', { p0: index + 1 }))
    : (t('Page {p0} was not changed.', { p0: index + 1 })) + (ret.reply ? ` ${ret.reply}` : '')), { tooHard: ret.too_hard === true })
  if (!ret.patch.changes?.length && !ret.patch.added?.length && !ret.patch.removed?.length) throw unchanged() // say so, don't echo the request
  const live = slides.value.find(x => x.id === slide.id) || slide // merged onto the slide as it is now
  const changes: Record<string, any> = Object.fromEntries((ret.patch.changes || []).filter((c: any) => c?.id).map((c: any) => [c.id, c]))
  const removed = new Set<string>(ret.patch.removed || [])
  const seen = new Set<string>()
  const merged = [
    ...live.elements.filter(e => !removed.has(e.id)).map((e: any) => {
      const c = changes[e.id]
      if (!c) return e
      const { src: _src, path, layer: _layer, ...rest } = c // image sources and tokenised paths are never taken from the model
      void _src; void _layer // layer is a stacking instruction, applied below, not an element field
      // a position or size that is not a number is dropped: top ":" from the model threw a heading into the corner
      // of the slide (2026-11-11, a user's "improve the design")
      for (const k of GEOMETRY) {
        if (!(k in rest)) continue
        const v = toNum(rest[k])
        if (Number.isFinite(v)) rest[k] = v
        else delete rest[k]
      }
      if (e.type === 'shape' && e.text) {
        const { content, defaultColor, defaultFontName, text, ...shape } = rest
        const t = { ...e.text, ...(text || {}), ...(content !== undefined ? { content } : {}), ...(defaultColor ? { defaultColor } : {}), ...(defaultFontName ? { defaultFontName } : {}) }
        return { ...e, ...shape, text: t, ...(path && !String(path).startsWith('path:') ? { path } : {}) }
      }
      return { ...e, ...rest, ...(path && !String(path).startsWith('path:') ? { path } : {}) }
    }),
    // an added element without a real position and size is dropped rather than drawn at 0,0
    ...(ret.patch.added || []).filter((e: any) => e?.id && GEOMETRY.slice(0, 4).every(k => Number.isFinite(toNum(e[k]))) && e.type !== 'image' && !live.elements.some(x => x.id === e.id) && !seen.has(e.id) && seen.add(e.id)).map(unflat),
  ]
  // stacking order: "layer" front/back moves that element (a user asked six times to put the title on top of a white
  // box; the merge kept the old order every time, 2026-11-11). Only the layered elements move; everything else keeps
  // its order. "back" lands just above the slide's background layers (backdrop pictures, full-slide shapes).
  const { width: SW, height: SH } = slideSize()
  const fullSlide = (e: any) => (e.type === 'image' && isBackdrop(e)) || (e.type === 'shape' && !elText(e) && e.width >= SW * 0.95 && e.height >= SH * 0.95)
  const isBox = (e: any) => e.type === 'shape' && !String(e.content ?? e.text?.content ?? '').replace(/<[^>]+>/g, '').trim() && e.width >= 20 && e.height >= 20
  // a NEW box without words goes behind the text by default ("a white rectangle behind the title" landed on top of it);
  // lines, arrows and icons (thinner than 20px) and anything already on the slide keep their place
  const liveIds = new Set(live.elements.map(e => e.id))
  const addedLayer: Record<string, string> = Object.fromEntries((ret.patch.added || []).filter((e: any) => e?.id && !liveIds.has(e.id))
    .map((e: any) => [e.id, e.layer || (isBox(e) ? 'back' : '')]))
  const layer = (id: string) => changes[id]?.layer || addedLayer[id]
  if (merged.some((e: any) => layer(e.id) === 'front' || layer(e.id) === 'back')) {
    const moved = (e: any) => layer(e.id) === 'front' || layer(e.id) === 'back'
    const texts = merged.filter((e: any) => elText(e) && e.type !== 'image')
    // a new box behind text covers the text it mostly sits under (30%+ of that text's area): the model drew a 90px
    // square behind a 340px title. Kept inside the slide.
    const fit = (b: any) => {
      if (!addedLayer[b.id] || layer(b.id) !== 'back' || !isBox(b)) return b
      const area = (x: any) => x.width * x.height
      const ov = (t: any) => Math.max(0, Math.min(b.left + b.width, t.left + t.width) - Math.max(b.left, t.left)) * Math.max(0, Math.min(b.top + b.height, t.top + t.height) - Math.max(b.top, t.top))
      const under = texts.filter((t: any) => ov(t) >= 0.3 * area(t))
      if (!under.length) return b
      const pad = 12
      const l = Math.max(0, Math.min(b.left, ...under.map((t: any) => t.left - pad))), tp = Math.max(0, Math.min(b.top, ...under.map((t: any) => t.top - pad)))
      const r = Math.min(SW, Math.max(b.left + b.width, ...under.map((t: any) => t.left + t.width + pad))), bt = Math.min(SH, Math.max(b.top + b.height, ...under.map((t: any) => t.top + t.height + pad)))
      return { ...b, left: l, top: tp, width: r - l, height: bt - tp }
    }
    const stay = merged.filter((e: any) => !moved(e))
    const back = merged.filter((e: any) => layer(e.id) === 'back').map(fit), front = merged.filter((e: any) => layer(e.id) === 'front')
    let at = 0 // just above the background layers at the bottom of the stack
    while (at < stay.length && fullSlide(stay[at])) at++
    merged.splice(0, merged.length, ...stay.slice(0, at), ...back, ...stay.slice(at), ...front)
  }
  // a reply that describes changes the merge did not make is never shown as done
  if (JSON.stringify(merged) === JSON.stringify(live.elements)) throw unchanged()
  const before: Record<string, number> = Object.fromEntries(live.elements.filter((e: any) => e.type === 'text').map((e: any) => [e.id, e.height]))
  if (JSON.stringify(merged) === JSON.stringify(live.elements)) throw unchanged()
  slidesStore.updateSlide({ elements: merged }, slide.id)
  await fitText(slide.id, Object.keys(changes).filter(id => before[id]), before)
  // what actually landed: elements that differ, plus those removed (not the raw patch, which may name ids the merge ignored)
  const liveById = new Map(live.elements.map((e: any) => [e.id, JSON.stringify(e)]))
  const n = merged.filter((e: any) => liveById.get(e.id) !== JSON.stringify(e)).length + live.elements.filter((e: any) => !merged.some((m: any) => m.id === e.id)).length
  void n
  editedThisTurn.add(slide.id)
  lastEditTerse = t('Updated page {p0}.', { p0: index + 1 })
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

// Version history: the deck's saved copies (ChatSlide keeps the newest 20 and one per hour for 7 days), offered as
// buttons, from the chat ("yesterday's version", "верните версию от 11.11": day = that date) or the history icon. A doctor
// asked for "yesterday's version" and "the version history" after a day's work was replaced, and the editor had neither
// (2026-11-12).
// page numbers as runs: 2-28, 30
const spans = (ps: number[]) => {
  const out: string[] = []
  for (const p of [...ps].sort((a, b) => a - b)) {
    const last = out[out.length - 1]?.split('–')
    if (last && Number(last[last.length - 1]) === p - 1) out[out.length - 1] = `${last[0]}–${p}`
    else out.push(String(p))
  }
  return out.join(', ')
}
// Thumbs up/down on a chat reply: recorded at once, tied to its turn; a thumbs-down then takes an optional comment,
// sent as its own record of the same turn (the classic chat records a 👎 and then its reason the same way)
const why = ref('')
const rate = (m: Session['history'][number], good: boolean) => {
  if (!m.turn || m.rating !== undefined) return
  m.rating = good ? 1 : 0
  save()
  chatslide?.feedback?.(m.turn, good, '', m.text).catch(() => undefined)
}
const comment = (m: Session['history'][number]) => {
  if (m.comment !== undefined) return
  m.comment = why.value.trim().slice(0, 300)
  why.value = ''
  save()
  if (m.comment && m.turn) chatslide?.feedback?.(m.turn, false, m.comment, m.text).catch(() => undefined)
}
const localDay = (at: string | number) => {
  const d = new Date(at)
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}
const listVersions = async (text: string, day = ''): Promise<string> => {
  if (!chatslide?.versions) fail(t('Version history is not available here. Undo goes back through the changes of this session.'))
  busy.value = t('Looking up the saved versions…')
  // one candidate per hour, newest first: the copies of the last minutes are steps Undo already covers, and listed
  // first they would hide yesterday's version
  const hours = new Set<string>()
  const perHour = (await chatslide.versions()).filter(v => { const h = String(v.at).slice(0, 13); if (hours.has(h)) return false; hours.add(h); return true })
  const all = (day ? perHour.filter(v => localDay(v.at) === day) : perHour).slice(0, 24)
  if (day && !all.length) return t('There is no saved version from {p0}. Saved versions go back 7 days.', { p0: new Date(`${day}T12:00:00`).toLocaleDateString(lang, { month: 'short', day: 'numeric' }) })
  const sig = (d: any[]) => { // content, places, colours, backgrounds, notes
    // what a person would see as a difference; not the fields the editor fills in on load
    const j = JSON.stringify(d.map(x => [x.background?.color || x.background?.image?.src || x.background?.type || '', x.remark || '',
      (x.elements || []).map((e: any) => [e.type, e.content ?? e.text?.content ?? e.src ?? '', Math.round(e.left || 0), Math.round(e.top || 0), Math.round(e.width || 0), e.type === 'text' ? 0 : Math.round(e.height || 0), e.fill || '', e.defaultColor || ''])])) // (a text box's height is refitted when it is shown)
    let h = 0
    for (let i = 0; i < j.length; i++) h = (h * 31 + j.charCodeAt(i)) | 0
    return `${d.length}|${j.length}|${h}`
  }
  const seen = new Set([sig(slides.value)]) // versions that differ from the deck on screen, one per distinct state
  const choices: Record<string, string> = {}
  for (let k = 0; k < all.length && Object.keys(choices).length < 6; k += 8) {
    const batch = await Promise.all(all.slice(k, k + 8).map(v => fetch(v.url, { cache: 'no-store' }).then(r => r.json()).then(d => ({ v, d })).catch(() => null)))
    for (const x of batch) {
      if (!x?.d?.slides?.length || seen.has(sig(x.d.slides)) || Object.keys(choices).length >= 6) continue
      seen.add(sig(x.d.slides))
      const when = new Date(x.v.at).toLocaleString(lang, { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })
      choices[`${when} · ${t('{p0} pages', { p0: x.d.slides.length })}`] = x.v.url
    }
  }
  if (!Object.keys(choices).length) return t('There is no earlier saved version that differs from the deck you have now.')
  s.savedVersions = choices
  s.pending = { request: text, question: '', options: Object.keys(choices), offer: true }
  return t('These are the earlier saved versions of this deck (you have {p0} pages now). Tap one to go back to it; Undo brings the current deck back.', { p0: slides.value.length })
}
// the history icon: the same list, no AI call
const showVersions = () => { if (!busy.value) run(t('Looking up the saved versions…'), async () => { say('assistant', await listVersions('')); save() }) }

const cloneSlide = (sl: Slide): Slide => {
  const newId: Record<string, string> = Object.fromEntries(sl.elements.map(e => [e.id, nanoid(10)]))
  return {
    ...JSON.parse(JSON.stringify(sl)), id: nanoid(10),
    elements: sl.elements.map(e => ({ ...JSON.parse(JSON.stringify(e)), id: newId[e.id] })),
    animations: (sl.animations || []).filter(a => newId[a.elId]).map(a => ({ ...a, id: nanoid(10), elId: newId[a.elId] })),
  }
}

// A step that cannot be done stops the turn (the router would otherwise retry it with new arguments).
class StepFailed extends Error {}
class NoChange extends StepFailed {} // the page already is what was asked (or the edit changed nothing visible)
function fail(msg: string): never {
  throw new StepFailed(msg)
}

// A new content page from the same outline + JSONL generator as a prompt-built deck (layout matches item count).
const generatePage = async (text: string, instruction: string, at: number, count = 1) => {
  const ask = `Exactly ${count} content slide(s) for an existing deck (no cover, no closing slide, no contents slide). Deck so far: ${deckLines().join(' | ').slice(0, 1500)}. The new slide(s): ${instruction}`
  const brief = s.docs.length ? `${ask}\n\n${materialBrief()}` : ask
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
  if (d.error || !d.labels?.length) fail(d.error || (t('No numbers to chart.')))
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
  if (d.error || !d.rows?.length) fail(d.error || (t('Nothing to put in a table.')))
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
  if (!d.points?.length) fail(t('Could not write the extension.'))
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
const findPicture = async (query: string, source: string, landscape: boolean, avoid: string[] = [], reuse = false) => {
  pictureNote = ''
  if (source === 'direct') { // a picture file already chosen from a website: only the paid add_site_pictures step may place one
    if (!directAllowed) fail(t('Real pictures from a brand or company website are a paid feature. Upgrade your plan and I will take them from the site.'))
    pictureSource = 'photo'
    return query
  }
  const nearest = source === 'medical' // a medical picture may be the nearest real one of the same subject
  if (nearest) source = 'wikimedia'
  pictureSource = source === 'wikimedia' ? 'wikimedia' : source === 'ai' ? 'ai' : 'photo'
  if (source === 'attached') { // the query names one of the user's attached pictures
    const pics = s.images || []
    const pic = pics.find(m => m.name === query) || (pics.length === 1 ? pics[0] : undefined)
    if (!pic) {
      fail(pics.length ? (t('There is no attached picture "{p0}". Attached: {p1}.', { p0: query, p1: pics.map(m => m.name).join(', ') }))
        : (t('No picture is attached yet: attach it with the 📎 icon.')))
    }
    pictureSource = 'attached'
    return pic.url
  }
  if (source === 'wikimedia') {
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'wikimedia', query, nearest }) })).json()
    if (!r.src) { // no match: a stock photo, said plainly; a question here stopped a turn with three more pictures to add
      const sp = await api.searchImage({ query, per_page: 5, orientation: landscape ? 'landscape' : 'portrait' } as any).catch(() => ({ data: [] }))
      const alt = (sp.data || []).map((x: any) => x.src).find((u: string) => u && !avoid.includes(u))
      if (!alt) fail(t('Wikimedia Commons has no picture of "{p0}". Should I use a stock photo or generate one instead?', { p0: query }))
      pictureSource = 'photo'
      pictureNote = t('Wikimedia Commons had no picture of "{p0}", so this is a stock photo.', { p0: query.slice(0, 60) })
      return alt as string
    }
    pictureNote = t('Picture from Wikimedia Commons: {p0} ({p1}{p2}) {p3}', { p0: r.title, p1: r.credit, p2: r.license ? t(', {p0}', { p0: r.license }) : '', p3: r.page })
      + (r.close ? ` ${t('No exact picture exists there, so this is the nearest one of the same subject: check it is right for your slide.')}` : '')
    return r.src as string
  }
  if (source === 'photo') {
    // a few results, so "all the pictures" with similar searches do not get the same photo twice
    const r = await api.searchImage({ query, per_page: 5, orientation: landscape ? 'landscape' : 'portrait' } as any)
    const src = (r.data || []).map((x: any) => x.src).find((u: string) => u && !avoid.includes(u)) // the services axios instance already unwraps response.data
    if (src) return src as string
    // every result already used: the first again, rather than an AI picture nobody asked for (10 credits)
    if (reuse && r.data?.[0]?.src) return r.data[0].src as string
    pictureNote = t('No stock photo matched, so I generated one with AI.')
    pictureSource = 'ai'
  }
  if (chatslide?.generateImage) { // inside ChatSlide: its image route bills, gates and moderates
    const g = await chatslide.generateImage(query, landscape ? '16:9' : '9:16')
    if (g.error === 'credits') {
      const e = new OutOfCredits('out of credits')
      e.required = 10
      throw e
    }
    if (g.error) fail(t('The picture could not be generated ({p0}).', { p0: g.error }))
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
  // picture 0 = the full-slide backdrop ("change the background picture"); 1.. = the content pictures
  const el = sl && (image === 0 ? firstBackdrop(sl) : pictures(sl)[Math.max(0, (image || 1) - 1)])
  if (!sl || !el) fail(t('That page has no picture.'))
  const src = await findPicture(query, source, el.width >= el.height, sl.elements.filter((e: any) => e.type === 'image').map((e: any) => e.src))
  if (!src) fail(t('Could not get a picture.'))
  slidesStore.updateElement({ id: el.id, props: { src } as any, slideId })
  if (isBackdrop(el)) showBackdrop(slideId, el.id)
  touched([slideId])
  return { note: pictureNote, backdrop: isBackdrop(el) }
}
// A background picture the user just chose is seen: a colour change had faded it to 12% and left a full-slide colour
// layer over it ("I don't see changes", twice). The fade goes and the picture moves above such layers, under the text.
const showBackdrop = (slideId: string, picId: string) => {
  const sl = slides.value.find(x => x.id === slideId)
  if (!sl) return
  const { width, height } = slideSize()
  const els = sl.elements.map((e: any) => (e.id !== picId ? e : { ...e, filters: (({ opacity, ...rest }) => rest)(e.filters || {}) }))
  const at = els.findIndex((e: any) => e.id === picId)
  // the colour layers above it go just under it; everything else keeps its order (text never ends up beneath it)
  const layer = (e: any, k: number) => k > at && e.type === 'shape' && !elText(e) && e.width >= width * 0.95 && e.height >= height * 0.95
  const layers = els.filter(layer)
  const rest = els.filter((e: any, k) => !layer(e, k))
  rest.splice(rest.findIndex((e: any) => e.id === picId), 0, ...layers)
  slidesStore.updateSlide({ elements: rest }, slideId)
}

const writeNotes = async (ids: string[], instruction: string) => {
  let written = 0
  for (const id of ids) {
    const sl = slides.value.find(x => x.id === id)!
    busy.value = t('Writing notes {p0} of {p1}…', { p0: ids.indexOf(id) + 1, p1: ids.length })
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
const titleAt = (i: number) => (slides.value[i] ? (readingOrder(slides.value[i]).map(elText).find(Boolean) || '').slice(0, 50) : '') || (t('page {p0}', { p0: i + 1 }))
// the new pages by number and title: the router saw only "Added 1 page(s)" and added the same TikTok page again
const newTitles = (at: number, count: number) => Array.from({ length: count }, (_, k) => `${at + k + 1} "${titleAt(at + k)}"`).join(', ')
const beforeClosing = (at: number) => {
  const last = slides.value[slides.value.length - 1]
  // a generated deck marks it (type 'end'); in an imported deck the closing page is a last page with almost no text
  const words = last ? pageText(last).replace(/\[[^\]]*\]/g, '').trim() : ''
  const closing = last && (last.type === 'end' || (slides.value.length > 2 && words.length > 0 && words.length < 30
    && !last.elements.some(e => e.type === 'chart' || e.type === 'table')))
  return closing && at >= slides.value.length ? slides.value.length - 1 : at
}

const runCall = async (text: string, action: string, args: any, ids: string[]): Promise<string> => {
  const idOf = (p: number) => {
    if (!(p >= 1 && p <= ids.length)) fail(t('There is no page {p0} (the deck has {p1}).', { p0: p, p1: ids.length }))
    return ids[p - 1]
  }
  const pos = (id: string) => slides.value.findIndex(x => x.id === id)
  const nums = (ps: number[]) => [...new Set((ps || []).map(idOf))].filter(id => pos(id) >= 0)
  const n = slides.value.length
  if (['add_page', 'add_chart_page', 'add_table_page'].includes(action) && Number(args.after) > slides.value.length) args.after = slides.value.length
  if (['add_page', 'add_chart_page', 'add_table_page'].includes(action) && Number(args.layout_from) > slides.value.length) args.layout_from = slides.value.length
  // pages from the user's own file, and none is attached: ask for it (the request waits for the file), never invent
  // the content ("buatkan ppt terkait makalah ini" generated a deck about nothing, 2026-11-12)
  if ((action === 'add_page' || action === 'create_deck') && args.from_users_file && !s.docs.length && !(s.images || []).length) {
    return runCall(text, 'ask_user', { question: t('No file is attached yet: attach it with the 📎 icon in the chat box.'), options: [] }, ids)
  }
  if (action === 'add_page') {
    // "after this one" / no place named: right after the page the user was on (the router once said "after 30")
    if (args.after_named === false && currentAtSend) {
      const made = slides.value.map((x, i) => (createdThisMessage.has(x.id) ? i : -1)).filter(i => i >= 0)
      args.after = (made.length ? Math.max(...made) : pos(currentAtSend)) + 1 // a 2nd new page follows the 1st
    }
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    if (s.mode === 'native') {
      const n0 = slides.value.length
      const topics = Array.isArray(args.topics) && args.topics.length > 1 ? ` One page per topic, in this order: ${args.topics.join('; ')}.` : ''
      if (!await generatePage(text, args.instruction + topics, after, Math.min(Math.max(args.count || 1, 1), 10))) fail(t('Could not write that page.'))
      const added = slides.value.length - n0
      const capped = (args.count || 1) > 10 ? (t(' (at most 10 topics per request; ask again for more)')) : ''
      return (t('Added {p0} page(s) after page {p1}: {p2}.', { p0: added, p1: after, p2: newTitles(after, added) })) + capped
    }
    const count = Math.min(Math.max(args.count || 1, 1), 10)
    for (let k = 0; k < count; k++) {
      busy.value = t('Writing page {p0} of {p1}…', { p0: k + 1, p1: count })
      const list = [...slides.value]
      const source = slides.value[pos(idOf(args.layout_from || args.after))]
      list.splice(after + k, 0, cloneSlide(source))
      slidesStore.setSlides(list)
      // each new page from its own topic: with one shared instruction "TikTok and Reddit" became two TikTok pages
      const topic = Array.isArray(args.topics) && args.topics.length === count ? `${args.topics[k]} (part of: ${args.instruction})` : args.instruction
      await editSlide(after + k, `This is NEW page ${k + 1} of ${count} on: ${topic}. Replace all its text with that page's own content (no repeats across pages). Keep the layout and styling. Write in the language of that text.`, !!args.hard && !!(chatslide?.canUseAdvanced?.() ?? (window as any).__evalAdvanced))
      // a line left exactly as on the copied page (the small header "Key Takeaway: Match Platform to Your Goal" stayed
      // on a new TikTok page): one more pass names those lines
      const copied = new Set(source.elements.map(elText).filter(Boolean))
      const stale = slides.value[after + k].elements.map(elText).filter(t => t && copied.has(t))
      if (stale.length) await editSlide(after + k, `These lines are still the text of the page this one was copied from; rewrite each for THIS page's topic (${args.instruction}), same length and style: ${stale.map(t => `"${t.slice(0, 120)}"`).join('; ')}`).catch(() => undefined)
      // pictures still the copied page's (a new "coffee prices" page kept slide 2's barista photos, 2026-11-11): a stock
      // photo for this page's topic in the same frame, so the layout stays and nothing repeats
      const copiedPics = new Set(source.elements.filter((e: any) => e.type === 'image' && !isBackdrop(e)).map((e: any) => e.src))
      const page = slides.value[after + k]
      const { width: SW, height: SH } = slideSize()
      for (const [n, e] of pictures(page).entries()) {
        if (isBackdrop(e) || !copiedPics.has(e.src) || e.width * e.height < 0.02 * SW * SH) continue // keep backdrops and logos
        await changeImage(text, page.id, n + 1, String(Array.isArray(args.topics) ? args.topics[k] : args.instruction).slice(0, 80), 'photo').catch(() => undefined)
      }
    }
    const capped = (args.count || 1) > 10 ? (t(' (10 is the most per request; ask again for more)')) : ''
    return (t('Added {p0} page(s) after page {p1}: {p2}.', { p0: count, p1: after, p2: newTitles(after, count) })) + capped
  }
  if (action === 'expand_page') {
    const i = pos(idOf(args.index))
    busy.value = t('Extending…')
    await expandPage(text, i, args.instruction || '')
    slidesStore.updateSlideIndex(i + 1)
    return t('Added an extension of page {p0} as page {p1}; page {p2} is unchanged.', { p0: i + 1, p1: i + 2, p2: i + 1 })
  }
  if (action === 'add_chart_page') {
    // changing a chart that is already there ("this chart -> pie"): edit it in place, never a new chart page
    const there = Number(args.existing_chart_page) || 0
    if (there > 0 && slides.value[pos(idOf(there))]?.elements.some(e => e.type === 'chart')) {
      const i = pos(idOf(there))
      const reply = await editSlide(i, `Change the chart on this page: ${args.instruction}${args.chart_type ? ` (chart type: ${args.chart_type})` : ''}. Keep its data unless asked.`)
      slidesStore.updateSlideIndex(i)
      return reply
    }
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    return (await addChartPage(text, args.instruction, args.chart_type || 'column', after)) || (t('Added a chart page "{p0}".', { p0: titleAt(after) }))
  }
  if (action === 'add_table_page') {
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    await addTablePage(text, args.instruction, after)
    slidesStore.updateSlideIndex(after)
    return t('Added a table page "{p0}".', { p0: titleAt(after) })
  }
  if (action === 'use_file_images') { // figures from the user's own PDF/PPTX/DOCX, extracted by ChatSlide's backend
    const file = docFiles.get(args.file) || (docFiles.size === 1 ? [...docFiles.values()][0] : undefined)
    if (!file) {
      fail(s.docs.some(d => d.name === args.file)
        ? (t('Attach {p0} again so I can pull its pictures out (it was attached in an earlier session).', { p0: args.file }))
        : (t('No file is attached yet: attach it with the 📎 icon.')))
    }
    if (!chatslide?.extractImages) fail(t('Pulling pictures out of a file works inside ChatSlide: open this project there.'))
    busy.value = t('Pulling the pictures out of the file…')
    let figs = figuresOf.get(args.file) || figuresOf.get(file.name)
    if (!figs) {
      const r = await chatslide.extractImages(file)
      if (r.error === 'plan') fail(t('Pulling pictures out of files needs a paid plan; I opened the plans for you.'))
      if (r.error) fail(t('I could not pull the pictures out ({p0}).', { p0: r.error }))
      figs = r.images || []
      figuresOf.set(file.name, figs)
    }
    const want: number[] = args.file_pages || []
    const chosen = (want.length ? figs.filter(m => m.page && want.includes(m.page)) : figs).slice(0, Math.min(Math.max(args.count || 1, 1), 6))
    if (!chosen.length) {
      fail(figs.length
        ? (t('There are no pictures on page(s) {p0} of {p1} (it has {p2}, on pages {p3}).', { p0: want.join(', '), p1: file.name, p2: figs.length, p3: [...new Set(figs.map(m => m.page))].filter(Boolean).join(', ') }))
        : (t('{p0} has no pictures I can use.', { p0: file.name })))
    }
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
    if (!placed) fail(t('Page {p0} has no free space for the pictures.', { p0: pos(id) + 1 }))
    return (t('Placed {p0} picture(s) from {p1} on page {p2}.', { p0: placed, p1: file.name, p2: pos(id) + 1 })) + (placed < chosen.length ? (t(' (no room for the other {p0})', { p0: chosen.length - placed })) : '')
  }
  if (action === 'add_image' && args.replaces_existing) { // "change the image" on a page that has one replaces it
    const sl = slides.value[pos(idOf(args.index))]
    if (sl?.elements.some((e: any) => e.type === 'image' && !isBackdrop(e))) return runCall(text, 'change_image', { index: pos(sl.id) + 1, image: 1, query: args.query, source: args.source, page_named: true, said_background: '' }, ids)
  }
  // several listed things: one picture each, in a row along the free bottom or right of the page (asked as separate
  // calls, the router added one picture and stopped, 2026-11-11)
  // the language of new slide text, named by the router and stated in the instruction itself: instructions written in
  // Chinese gave an English user's new pages Chinese text (2026-11-12, "generate more slides upto 30")
  // (a translation is written in its target language, never the deck's: the two would contradict)
  const slideLang = String(args.translate_to || args.language || '').slice(0, 40)
  if (['add_page', 'edit_page', 'edit_all_pages'].includes(action) && slideLang && !String(args.instruction || '').includes('Write all slide text in')) {
    args = { ...args, instruction: `${args.instruction}\n(Write all slide text in ${slideLang}.)` }
  }
  // a medical picture is a real one (Wikimedia, which falls back to a labelled stock photo): asked in prose, the router
  // still chose stock photos 6 of 6 times for "a picture for each variant of pulmonary hypertension" (2026-11-11)
  // an AI picture only when the user's words asked for one: it costs them 10 credits each ("сделать картинку на каждый
  // вариант" drew 4 AI pictures, 40 credits, unasked, 2026-11-11)
  if ((action === 'add_image' || action === 'change_image') && args.source === 'ai' && !args.asked_for_ai) args = { ...args, source: 'photo' }
  if ((action === 'add_image' || action === 'change_image') && args.medical && ['photo', 'wikimedia'].includes(args.source)) args = { ...args, source: 'medical' }
  // a picture of one brand's own things asked for as an ordinary picture: it comes from that brand's website (the router
  // split "more photos of the brand's products" into generic stock searches when the message had a second wish)
  if ((action === 'add_image' || action === 'change_image') && String(args.own_brand || '').trim() && args.source !== 'direct' && args.source !== 'attached' && !args.asked_for_ai) {
    const items = (args.items || []).filter(Boolean)
    return runCall(text, 'add_site_pictures', { index: args.index, page_named: args.page_named, site: '', brand: args.own_brand, what: [args.query, ...items].filter(Boolean).join('; '), count: items.length || 1 }, ids)
  }
  // Real pictures from a brand's or company's own website: a paid feature. Free users are told so and get the plans
  // (Quanlai, 2026-11-12: "say that this is a paid feature and ask them to pay", "an agent reading a certain website").
  if (action === 'add_site_pictures') {
    if (!(chatslide?.canUseAdvanced?.() ?? (window as any).__evalAdvanced)) {
      (chatslide?.offerPaid || chatslide?.offerAdvanced)?.()
      // said as a step that did not happen: the other parts of the message still run
      fail(t('Real pictures from a brand or company website are a paid feature. Upgrade your plan and I will take them from the site.'))
    }
    busy.value = t('Reading the website for pictures…')
    const onDeck = slides.value.flatMap(sl => sl.elements.filter((e: any) => e.type === 'image').map((e: any) => String(e.src)))
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ kind: 'site_images', query: args.what, url: args.site, brand: args.brand, count: args.count, avoid: onDeck.filter(u => u.startsWith('https://')) }) })).json()
    if (r.error || !r.images?.length) fail(t('I could not get pictures from that website ({p0}).', { p0: r.error || r.page || '' }))
    // only pictures that really load go on the slide (a site may refuse other sites, or serve a format browsers skip)
    const loads = (src: string) => new Promise<boolean>((ok) => { const im = new Image(); const tm = setTimeout(() => ok(false), 8000); im.onload = () => { clearTimeout(tm); ok(im.naturalWidth > 80) }; im.onerror = () => { clearTimeout(tm); ok(false) }; im.src = src })
    const srcs: string[] = []
    for (const x of r.images as any[]) if (await loads(x.src)) srcs.push(x.src)
    if (!srcs.length) fail(t('I could not get pictures from that website ({p0}).', { p0: r.page || '' }))
    const pg = Number(args.index)
    const pictures = () => slides.value.reduce((n, sl) => n + sl.elements.filter(e => e.type === 'image').length, 0)
    const before = pictures()
    directAllowed = true
    try { await runCall(text, 'add_image', { index: pg, page_named: true, query: srcs[0], items: srcs.length > 1 ? srcs : [], source: 'direct', box: { x: 55, y: 20, w: 40, h: 60 }, replaces_existing: false }, ids) }
    finally { directAllowed = false }
    return t('Added {p0} picture(s) from {p1}.', { p0: pictures() - before, p1: String(r.page).replace(/^https?:\/\/(www\.)?/, '').replace(/\/$/, '') })
  }
  if (action === 'add_image' && Array.isArray(args.items) && args.items.filter(Boolean).length > 1) {
    const items = args.items.filter(Boolean).slice(0, 6) // "6 pictures" got 4 (prod 2026-11-12, b1b)
    const pg = Number(args.index)
    const at = pos(idOf(pg))
    const setKey = picKey(ids[at], items.join('|'))
    if (picturedThisTurn.has(setKey)) fail(t('Those pictures were already added in this message.'))
    const { width: W, height: H } = slideSize()
    const row = { left: 0.05 * W, top: 0.6 * H, width: 0.9 * W, height: 0.3 * H }
    const fb = freeBox(slides.value[at], row)
    const roomy = items.length <= 4 && !!fb && fb.width >= row.width * 0.8 && fb.height >= row.height * 0.8 // five or six: a grid page
    // a full page: the pictures go on a new page right after it, in a grid, under that page's title (on a full slide
    // four pictures landed in the corners and off the edge, 2026-11-11)
    let pageId = ''
    if (!roomy) {
      const src = slides.value[at]
      const title: any = readingOrder(src).find((e: any) => elText(e))
      const page: any = { id: nanoid(10), background: JSON.parse(JSON.stringify(src.background || {})), elements: title ? [{ ...JSON.parse(JSON.stringify(title)), id: nanoid(10) }] : [] }
      const list = [...slides.value]
      list.splice(at + 1, 0, page)
      slidesStore.setSlides(list)
      ids.splice(0, ids.length, ...slides.value.map(x => x.id))
      pageId = page.id
      helperPages.add(page.id) // not a page the user asked for: it must not satisfy "new content goes on the new page"
    }
    const cols = roomy ? items.length : items.length <= 2 ? items.length : items.length <= 4 ? 2 : 3, rows = Math.ceil(items.length / cols)
    const lines: string[] = []
    let added = 0
    for (const [k, q] of items.entries()) {
      const c = k % cols, r = Math.floor(k / cols)
      const box = roomy ? { x: 5 + k * (90 / items.length), y: 62, w: 90 / items.length - 2, h: 33 }
        : { x: 5 + c * (90 / cols), y: 24 + r * (70 / rows), w: 90 / cols - 3, h: 70 / rows - 4 }
      try {
        // each item is a new picture of its own: never "replace", which overwrote the previous item
        lines.push(await runCall(text, 'add_image', { ...args, items: [], replaces_existing: false, index: roomy ? pg : at + 2, page_named: true, query: q, box }, ids))
        added++
      }
      catch (e: any) {
        if (e instanceof OutOfCredits && added) { lines.push(t("You're out of credits; {p0} of the pictures were added.", { p0: added })); break } // keep what was paid for
        if (!(e instanceof StepFailed)) throw e
        lines.push(e.message)
      }
    }
    if (!added) { // nothing came: no empty picture page, and a failure, not a success
      if (pageId) { slidesStore.setSlides(slides.value.filter(x => x.id !== pageId)); ids.splice(0, ids.length, ...slides.value.map(x => x.id)) }
      fail(lines.join(' ') || t('Could not get a picture.'))
    }
    picturedThisTurn.add(setKey)
    return (pageId ? `${t('Page {p0} was full, so the pictures are on a new page {p1}.', { p0: at + 1, p1: at + 2 })} ` : '') + lines.join(' ')
  }
  if (action === 'add_image') {
    const id = idOf(args.index)
    const { width: W, height: H } = slideSize()
    const b = args.box || { x: 55, y: 20, w: 40, h: 60 }
    const clamp = (v: number) => Math.min(Math.max(v || 0, 0), 100) / 100
    const sl = slides.value[pos(id)]
    const wanted = { left: clamp(b.x) * W, top: clamp(b.y) * H, width: Math.max(clamp(b.w), 0.05) * W, height: Math.max(clamp(b.h), 0.05) * H }
    // at most 4 new pictures on a page per message (a reworded re-plan would otherwise keep adding near-duplicates)
    if ([...picturedThisTurn].filter(k => k.startsWith(`${id}|`)).length >= 4) fail(t('Page {p0} already got 4 new pictures in this message.', { p0: pos(id) + 1 }))
    let box = freeBox(sl, wanted)
    if (!box && roomMadeThisTurn.has(id)) fail(t('Page {p0} has no free space any more.', { p0: pos(id) + 1 })) // made room once: never shrink it again
    if (!box) { // a full page: make room rather than refuse
      roomMadeThisTurn.add(id)
      await editSlide(pos(id), 'Make room for a picture: rearrange and narrow the text so the right 40% of the slide is free of text, keeping all the wording and keeping it readable. Do not change the wording.').catch(() => undefined)
      const room = { left: 0.58 * W, top: 0.2 * H, width: 0.38 * W, height: 0.6 * H }
      box = freeBox(slides.value[pos(id)], room)
      if (!box) { // the edit did not free it: the page's content moves into the left part, scaled by position and width
        const page = slides.value[pos(id)]
        slidesStore.updateSlide({ elements: page.elements.map((e: any) => (e.type === 'image' && isBackdrop(e) ? e : { ...e, left: e.left * 0.56, width: e.width * 0.56 })) }, id)
        box = freeBox(slides.value[pos(id)], room)
      }
    }
    if (!box) fail(t('Page {p0} is full; there is no free space for a picture. Remove or shrink something first, or put the picture on a new page.', { p0: pos(id) + 1 }))
    busy.value = args.source === 'ai' ? t('Drawing the picture…') : t('Finding a picture…')
    const src = await findPicture(args.query, args.source, box.width >= box.height, sl.elements.filter((e: any) => e.type === 'image').map((e: any) => e.src))
    if (!src) fail(t('Could not get a picture.'))
    const now = freeBox(slides.value[pos(id)], box) // re-checked: the page may have changed while the picture loaded
    if (!now) fail(t('Page {p0} has no free space any more.', { p0: pos(id) + 1 }))
    slidesStore.updateSlide({ elements: [...slides.value[pos(id)].elements, { type: 'image', id: nanoid(10), src, fixedRatio: true, rotate: 0, ...now } as any] }, id)
    touched([id])
    picturedThisTurn.add(picKey(id, args.query))
    slidesStore.updateSlideIndex(pos(id))
    const how = pictureSource === 'attached' ? (t('your attached')) : pictureSource === 'ai' ? (t('AI-generated')) : pictureSource === 'wikimedia' ? (t('Wikimedia')) : (t('stock photo'))
    const cx = (now.left + now.width / 2) / W, cy = (now.top + now.height / 2) / H
    const askedX = clamp(b.x) + clamp(b.w) / 2 // where the router wanted it; the free-space search may have moved it
    const moved = (askedX > 2 / 3 && cx <= 2 / 3) || (askedX < 1 / 3 && cx >= 1 / 3)
    const movedNote = moved ? (t(' (no free space on the {p0}, so it went here)', { p0: askedX > 2 / 3 ? t('right') : t('left') })) : ''
    const where = t('{p0} {p1}', { p0: cy < 1 / 3 ? t('top') : cy > 2 / 3 ? t('bottom') : t('middle'), p1: cx < 1 / 3 ? t('left') : cx > 2 / 3 ? t('right') : t('centre') })
    return (t('Added a picture ({p0}) for "{p1}" at the {p2}{p3}.', { p0: how, p1: String(args.query || '').slice(0, 60), p2: where, p3: movedNote })) + (pictureNote ? ` ${pictureNote}` : '')
  }
  if (action === 'remove_images') {
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    // several pages go only after a yes, like deleting pages: an attached picture plus "every slide" removed all 14
    // pictures of the deck (2026-11-12)
    if (targets.length > 1 && removeConfirmable !== targets.join()) {
      const q = t('Remove the pictures from {p0} pages? Nothing is removed yet.', { p0: targets.length })
      s.removeAsked = targets.join()
      s.pending = { request: userText, question: q, options: [t('Yes, remove them'), t('No, keep them')] }
      askedThisTurn = true
      save()
      return q
    }
    const logos = logoSpots()
    let n = 0
    for (const id of targets) {
      // the numbering the router saw (pageText): 0 = the full-slide background picture, 1.. = the content pictures; with
      // no numbers, "delete the pictures" removes content pictures, never the background or a logo
      const sl = slides.value[pos(id)]
      const imgs = pictures(sl), backdrop = firstBackdrop(sl)
      const content = imgs.filter(e => !logos.has(imgSpot(e)) && !isBackdrop(e))
      // "no pictures at all" includes the full-slide one ("no image at all in the last slide" answered "there is no
      // picture to remove", prod 2026-11-13) when it is the only picture there; with content pictures it stays ("remove
      // the last photo" came as no_pictures_at_all and took page 10's background with it, replay 2026-11-13)
      const all = content.length ? content : sl.elements.filter((e: any) => e.type === 'image' && !logos.has(imgSpot(e)))
      const drop = new Set((args.images?.length ? args.images.map((k: number) => (k === 0 ? backdrop : imgs[k - 1])) : args.no_pictures_at_all ? all : content.length ? content : [backdrop]) // only the background is left: that is the picture they see
        .filter(Boolean).map((e: any) => e.id))
      if (!drop.size) continue
      n += drop.size
      slidesStore.updateSlide({ elements: slides.value[pos(id)].elements.filter(e => !drop.has(e.id)) }, id)
      touched([id])
    }
    // nothing but a background photo on the named pages: that is the picture meant ("remove that background from
    // slide 1" answered "there is no picture to remove", 2026-11-11)
    if (!n && args.pages?.length && (!args.images?.length || args.images.includes(0))) {
      // never a page that is one flat picture: that background IS the page
      const flat = targets.filter(id => isFlat(slides.value[pos(id)]))
      const photo = targets.filter(id => slides.value[pos(id)].background?.type === 'image' && !flat.includes(id))
      for (const id of photo) slidesStore.updateSlide({ background: { type: 'solid', color: '#ffffff' } }, id)
      if (photo.length) { touched(photo); return t('Removed {p0} picture(s).', { p0: photo.length }) }
      if (flat.length) fail(t('Page {p0} is one flat picture: its text and pictures are all inside that single image, so I can\'t change a picture in it. Nothing was changed. I can add a new picture on top of it.', { p0: flat.map(id => pos(id) + 1).join(', ') }))
    }
    if (!n) fail(t('There is no picture to remove.'))
    return t('Removed {p0} picture(s).', { p0: n })
  }
  if (action === 'set_transition') {
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    for (const id of targets) slidesStore.updateSlide({ turningMode: args.effect }, id)
    return t('Transition set on {p0} page(s).', { p0: targets.length })
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
    return (args.effect === 'none' ? t('Animations removed.') : t('Animated {p0} page(s).', { p0: targets.length }))
  }
  if (action === 'delete_pages') {
    const del = nums(args.pages)
    if (del.length >= n) fail(t('A deck needs at least one page, so I kept it.'))
    // more than one page goes only after the user says yes: a pasted homework task once deleted 9 of 10 pages
    const repeatedWish = args.yes_to_delete_question === true && args.user_asked_to_cut === true && del.every(id => deleteAnswered.includes(id)) // those pages or some of them
    if (del.length > 1 && !del.every(id => deleteConfirmable.includes(id)) && !repeatedWish) {
      // the counts before and after, in words: "delete 27 slides" on a 28-page deck was confirmed and left 1 (2026-11-12)
      const q = t('Delete {p0} of your {p1} pages ({p2})? {p3} would remain. Nothing is deleted yet.', { p0: del.length, p1: n, p2: spans(del.map(id => pos(id) + 1)), p3: n - del.length })
      s.deleteAsked = del
      s.pending = { request: userText, question: q, options: [t('Yes, delete them'), t('No, keep them')] }
      askedThisTurn = true
      save()
      return q
    }
    slidesStore.deleteSlide(del)
    return t('Deleted {p0} page(s); {p1} remain.', { p0: del.length, p1: slides.value.length })
  }
  if (action === 'read_file') {
    const named = (s.images || []).find(m => m.name === args.file)
    const doc = named ? undefined : s.docs.find(d => d.name === args.file) || (s.docs.length === 1 ? s.docs[0] : undefined)
    // an attached picture: its text, read by the vision model ("write what's in my picture into the slides" was
    // answered "I can't read the attached image", 2026-11-11)
    const pic = named || (!doc && s.images?.length === 1 && !/\.(pdf|docx?|pptx?|xlsx?|csv|txt|md)$/i.test(String(args.file || '')) ? s.images[0] : undefined)
    if (pic) {
      busy.value = t('Reading {p0}…', { p0: pic.name })
      const r = await fetch('/api/tools/ask_json', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ instruction: 'Copy all the text in this picture exactly, in its own language, keeping line breaks and numbering. If it has no text, describe what it shows in one sentence. JSON {"text": "..."}', image: pic.url }) })
      const got = String((await r.json().catch(() => ({})))?.text || '').trim()
      if (!r.ok || !got) fail(t('No text could be read from {p0}{p1} (it may be a scanned image).', { p0: pic.name, p1: '' }))
      memo = `<file-content name="${pic.name.replace(/["<>]/g, '')}">\n${fence(got.slice(0, 12000))}\n</file-content>`
      return t('Read {p0}{p1}.', { p0: pic.name, p1: '' })
    }
    if (!doc) {
      fail(s.docs.length ? (t('There is no attached file "{p0}". Attached: {p1}.', { p0: args.file, p1: s.docs.map(d => d.name).join(', ') }))
        : (t('No file is attached yet: attach it with the 📎 icon in the chat box.')))
    }
    const parts = doc.text.split(/^(?=\[page \d+\]$)/m)
    const want: number[] = args.pages || []
    const total = parts.filter(p => /^\[page \d+\]/.test(p)).length
    const missing = want.filter(n => n < 1 || n > total)
    if (want.length && total && missing.length === want.length) fail(t('{p0} has only {p1} pages.', { p0: doc.name, p1: total }))
    const got = want.filter(n => !missing.includes(n))
    const picked = want.length && total ? parts.filter(p => got.includes(+(p.match(/^\[page (\d+)\]/)?.[1] || 0))).join('\n') : doc.text
    if (!picked.replace(/\[page \d+\]/g, '').trim()) fail(t('No text could be read from {p0}{p1} (it may be a scanned image).', { p0: doc.name, p1: want.length ? t(' page(s) {p0}', { p0: want.join(', ') }) : '' }))
    memo = `<file-content name="${doc.name.replace(/["<>]/g, '')}"${want.length ? ` pages="${want.join(',')}"` : ''}>\n${fence(picked.slice(0, 12000))}${picked.length > 12000 ? '\n[…cut; read fewer pages for the rest]' : ''}\n</file-content>`
    const skipped = missing.length && total ? (t(' (no page {p0}; it has {p1})', { p0: missing.join(', '), p1: total })) : ''
    return (t('Read {p0}{p1}.', { p0: doc.name, p1: want.length ? t(', page(s) {p0}', { p0: got.join(', ') }) : '' })) + skipped
  }
  if (action === 'find_papers') {
    busy.value = t('Finding papers…')
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'scholar', query: args.query }) })).json()
    if (r.error || !r.papers?.length) { // no paper on that: real web sources still answer "add citations"
      const found: string = await runCall(text, 'web_search', { query: args.query }, ids)
      return t('No academic papers on that, so I used web sources. {p0}', { p0: found })
    }
    memo = `<web-results>\n${fence(r.papers.map((p: any, i: number) => `[${i + 1}] ${p.authors} (${p.year}). ${p.title}. ${p.venue}. ${p.url} (cited ${p.cited}x)`).join('\n'))}\n</web-results>`
    return t('Found {p0} relevant paper(s).', { p0: r.papers.length })
  }
  if (action === 'read_link') {
    busy.value = t('Reading the link…')
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'url', query: args.url }) })).json()
    if (r.error || !String(r.text || '').trim()) fail(t('I could not read that link ({p0}).', { p0: r.error || 'no content' }))
    memo = `<file-content name="${String(r.title).replace(/["<>]/g, '').slice(0, 80)}">\n${fence(String(r.text).slice(0, 12000))}\n</file-content>`
    return t('Read the {p0}: {p1}.', { p0: r.kind, p1: String(r.title).slice(0, 60) })
  }
  if (action === 'web_search') {
    busy.value = t('Searching the web…')
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'web', query: args.query }) })).json()
    if (r.error || !r.sources?.length) fail(t('I found no reliable sources online for "{p0}", so I added no citations.', { p0: args.query }))
    memo = `<web-results>\n${fence(`${r.answer}\nSources:\n${r.sources.slice(0, 6).map((x: any, i: number) => `[${i + 1}] ${x.title} ${x.url}`).join('\n')}`)}\n</web-results>`
    return t('Searched online.')
  }
  if (action === 'merge_pages') {
    const ids = nums(args.pages)
    if (ids.length < 2) fail(t('Which pages should be combined?'))
    const [keep, ...rest] = ids
    // the router sees each page in one short line; the merge gets every word of the pages being folded in
    const extra = rest.map(id => readingOrder(slides.value[pos(id)]).map(elText).filter(Boolean).join('\n')).join('\n\n')
    await editSlide(pos(keep), `Combine the content below into this page: keep this page's title, add the points of the other page(s) condensed so everything fits without overflow, and rearrange the layout if needed. ${args.instruction || ''}\n\nContent of the other page(s):\n${extra}`)
    slidesStore.deleteSlide(rest)
    return t('Combined {p0} pages into page {p1}; {p2} remain.', { p0: ids.length, p1: pos(keep) + 1, p2: slides.value.length })
  }
  if (action === 'duplicate_pages') {
    args.pages = (args.pages || []).map((p: number) => (p >= 1 && p <= slides.value.length ? p : slideIndex.value + 1))
    const list = [...slides.value]
    for (const id of nums(args.pages)) {
      const i = list.findIndex(x => x.id === id)
      for (let k = 0; k < Math.min(Math.max(args.copies || 1, 1), 50); k++) list.splice(i + 1, 0, cloneSlide(list[i]))
    }
    slidesStore.setSlides(list)
    return t('Duplicated.')
  }
  if (action === 'move_pages') {
    const move = nums(args.pages)
    const moved = slides.value.filter(x => move.includes(x.id)) // keep deck order
    const rest = slides.value.filter(x => !move.includes(x.id))
    const at = Math.min(Math.max((args.to || 1) - 1, 0), rest.length)
    rest.splice(at, 0, ...moved)
    if (rest.map(x => x.id).join() === slides.value.map(x => x.id).join()) fail(t('Those pages are already there; the order is unchanged.'))
    slidesStore.setSlides(rest)
    return t('Moved to position {p0}.', { p0: at + 1 })
  }
  if (action === 'reorder_pages') {
    const order = nums(args.order)
    const list = [...order, ...slides.value.map(x => x.id).filter(id => !order.includes(id))] // pages left out keep their relative order at the end
      .map(id => slides.value[pos(id)])
    if (list.map(x => x.id).join() === slides.value.map(x => x.id).join()) fail(t('That is already the current order; nothing changed.'))
    slidesStore.setSlides(list)
    return t('Reordered.')
  }
  if (action === 'swap_pages') {
    const list = [...slides.value]
    const i = pos(idOf(args.a)), j = pos(idOf(args.b))
    ;[list[i], list[j]] = [list[j], list[i]]
    slidesStore.setSlides(list)
    return t('Swapped pages {p0} and {p1}.', { p0: i + 1, p1: j + 1 })
  }
  if (action === 'edit_page') {
    // text from other pages ("slide 7: the text of slide 10, points 1-3"): the edit sees only its own page, so it
    // answered "send me slide 10's text" (2026-11-11)
    const editing = Number(args.index)
    const sources = nums((args.source_pages || []).map(Number).filter((n: number) => n >= 1 && n <= slides.value.length && n !== editing))
    if (sources.length) args = { ...args, whole_deck_option: '', instruction: `${args.instruction}\n\n${sources.map((id: string) => `Text of page ${pos(id) + 1} (the source to use):\n${readingOrder(slides.value[pos(id)]).map(elText).filter(Boolean).join('\n')}`).join('\n\n').slice(0, 4000)}` }
    // a named page is not a whole-deck question: the offer also ends the turn, which cut off "and insert pictures"
    // after "slide 5: <text>" (2026-11-11)
    if (args.new_content && !args.page_named) return runCall(text, 'add_page', { instruction: args.instruction, count: 1, after: slideIndex.value + 1, layout_from: slideIndex.value + 1, after_named: false, message_new_pages: 1, topics: [], hard: args.hard }, ids)
    // the router's page, always: replacing it with the page on screen when page_named was false sent "change slide 2"
    // to page 8 (2026-11-13), and earlier pages 1-4 all onto page 1
    const i = pos(idOf(args.index))
    const before = slides.value[i].elements.map(elText).filter(Boolean)
    const was = JSON.parse(JSON.stringify(slides.value[i].elements)) as PPTElement[]
    // the advanced model: a hard request (the router says), or an edit the standard model could not do. It costs
    // credits and is on paid plans; anyone else gets the standard result and, when it mattered, the upgrade offer
    const canAdvanced = !!(chatslide?.canUseAdvanced?.() ?? (window as any).__evalAdvanced) // the eval's paid user
    const offer = () => {
      // once per session: a frustrated user got the line and the paywall on two turns in a row (2026-11-11)
      if (advancedOffered) return ''
      advancedOffered = true
      chatslide?.offerAdvanced?.()
      return chatslide?.offerAdvanced ? (t('Requests like this work better on our advanced model, on paid plans.')) : ''
    }
    // "improve the design" with nothing concrete: the standard model moved boxes a few pixels and said it improved the
    // design (5 of 5 replays of a real user's request, 2026-11-11); the advanced model redesigned the page (2 of 2).
    // So free users are told plainly and offered what visibly works; paid users get the advanced model.
    if (args.redesign && !canAdvanced) {
      s.pending = { request: args.instruction, question: '', options: [t('Apply a new theme to the deck'), t('Make the text bigger')], offer: true, localize: true }
      save()
      offer() // ChatSlide's paywall, once per session
      return t('A redesign of this page needs the advanced model, on paid plans. Without it I can apply a new theme or make the text bigger.')
    }
    let upgradeLine = ''
    const reply = await editSlide(i, args.instruction, (!!args.hard || !!args.redesign) && canAdvanced, !args.redesign).catch((e) => {
      if (!(e instanceof NoChange) || editedThisTurn.has(slides.value[i].id)) throw e
      if (lastEditAdvanced || !(e as any).tooHard) throw e // the advanced model already had its try, or nothing was needed
      if (canAdvanced) return editSlide(i, args.instruction, true)
      const line = offer()
      throw line ? new NoChange(`${e.message} ${line}`) : e
    })
    if (args.hard && !canAdvanced) upgradeLine = offer()
    keepUnaskedSizes(i, was)
    if (JSON.stringify(slides.value[i].elements) === JSON.stringify(was)) throw new NoChange(t('Page {p0} was not changed.', { p0: i + 1 }))
    // a translation that left lines as they were (the small header stayed English): one pass naming them
    if (args.translate_to) {
      const left = slides.value[i].elements.map(elText).filter(t => t && before.includes(t) && /\p{L}{3}/u.test(t))
      if (left.length) await editSlide(i, `Translate these remaining lines into ${args.translate_to}, keeping their style: ${left.map(t => `"${t.slice(0, 160)}"`).join('; ')}`).catch(() => undefined)
      lastEditTerse = t('Translated page {p0} into {p1} (only this page).', { p0: i + 1, p1: args.translate_to })
      slidesStore.updateSlideIndex(i)
      return lastEditTerse
    }
    slidesStore.updateSlideIndex(i)
    // longer text that now runs behind a picture ("the letters are behind the picture", 2026-11-10): one layout pass.
    // Measured after the page renders, since text boxes grow to fit their text then.
    await nextTick(); await new Promise(r => setTimeout(r, 300))
    const hadBefore = textOnPictures(was)
    const paidLayout = lastEditAdvanced // the cheaper model must not rework a layout the user paid for
    if ([...textOnPictures(slides.value[i].elements)].some(k => !hadBefore.has(k))) {
      await editSlide(i, 'Fix the layout only: no text may overlap other text or sit behind a picture. Make the text boxes fit (smaller font or narrower) or shrink/move the pictures. Keep every word.', paidLayout).catch(() => undefined)
    }
    if (upgradeLine) lastEditTerse = `${lastEditTerse} ${upgradeLine}`
    // a named page is not a whole-deck question; nor is an edit of the selected element ("the selected point 1" on
    // every page left most pages "not changed", 2026-11-11)
    if (args.whole_deck_option && args.page_named !== true && !turnSelected) { // (none for an edit using another page's text)
      const others = slides.value.map((_, k) => k + 1).filter(k => k !== i + 1)
      s.pending = { request: args.instruction, question: '', options: [args.whole_deck_option.trim()], offer: true, localize: true,
        repeat: [{ action: 'edit_page', args: { ...args, whole_deck_option: '' } }], pages: others, from: i + 1 }
    }
    return upgradeLine ? `${reply} ${upgradeLine}` : reply
  }
  if (action === 'set_font') { // direct: element default + every inline font-family, no model, no layout change
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    touched(targets)
    const sample: any = args.same_as_page > 0 ? readingOrder(slides.value[args.same_as_page - 1]).find((e: any) => elText(e)) : null
    const sampleHtml = String(sample?.content || sample?.text?.content || '')
    const sampleFont = sampleHtml.match(/font-family:\s*([^;"]+)/)?.[1] || sample?.defaultFontName || sample?.text?.defaultFontName
    const font = String(sampleFont || args.font || '').replace(/[^\p{L}\p{N} ,_-]/gu, '').trim().slice(0, 60) // a font name, nothing CSS can parse further
    if (!font) fail(t('Which font?'))
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
    return t('Set {p0} on {p1} text boxes across {p2} page(s).', { p0: font, p1: n, p2: targets.length })
  }
  if (action === 'set_font_size') { // direct, like set_font: "title bigger", "more" got barely-changed model rewrites (2026-11-10)
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    const marked = (e: any) => e.textType === 'title' || e.text?.type === 'title'
    // the first text is the title only on a page with no element marked as one (a header above it is not the title)
    const isTitle = (sl: any, e: any) => marked(e) || (!sl.elements.some(marked) && e === readingOrder(sl).find((x: any) => elText(x)))
    const inScope = (sl: any, e: any) => (e.type === 'text' || (e.type === 'shape' && e.text)) && elText(e)
      && (args.scope === 'titles' ? isTitle(sl, e) : args.scope === 'body' ? !isTitle(sl, e) : true)
    const html = (e: any) => String(e.type === 'text' ? e.content : e.text?.content || '')
    const sizes = (h: string) => [...h.matchAll(/font-size:\s*([\d.]+)px/g)].map(m => +m[1])
    let to = 0 // "same size as page N": that page's largest size in the same scope
    if (args.same_as_page > 0) {
      const sl = slides.value[args.same_as_page - 1]
      to = Math.max(0, ...(sl?.elements || []).filter((e: any) => inScope(sl, e)).flatMap((e: any) => sizes(html(e)).length ? sizes(html(e)) : [16]))
      if (!to) fail(t('That page has no text to match.'))
    }
    const k = Math.min(3, Math.max(0.3, Number(args.scale) || 1.25))
    const resize = (h: string) => {
      const doc = new DOMParser().parseFromString(h, 'text/html')
      if (!h.includes('font-size')) doc.querySelectorAll('p').forEach(p => { (p as HTMLElement).style.fontSize = '16px' }) // PPTist's default
      // a step the eye can see: 1.1 of a 14px body is 15px, and "bigger", "more", "even bigger!" went on for five
      // messages (prod 2026-11-10/11, b1c, 415, e13). At least 4px up or 2px down per step
      const step = (n: number) => to || (k > 1 ? Math.max(n * k, n + 4) : k < 1 ? Math.min(n * k, n - 2) : n)
      return doc.body.innerHTML.replace(/font-size:\s*([\d.]+)px/g, (_, n) => `font-size: ${Math.max(8, Math.round(step(+n)))}px`)
    }
    let n = 0
    for (const id of targets) {
      const sl = slides.value[pos(id)]
      for (const e of sl.elements as any[]) {
        if (!inScope(sl, e)) continue
        if (e.type === 'text') slidesStore.updateElement({ id: e.id, props: { content: resize(e.content) }, slideId: id })
        else slidesStore.updateElement({ id: e.id, props: { text: { ...e.text, content: resize(e.text.content) } }, slideId: id })
        n++
      }
    }
    if (!n) fail(t('There is no such text to resize there.'))
    touched(targets)
    return t('Resized the text in {p0} text boxes across {p1} page(s).', { p0: n, p1: targets.length })
  }
  if (action === 'edit_all_pages') { // 5 pages in flight, the next starts as soon as one ends; merged by slide id
    const all = slides.value.map(x => x.id)
    const viewing = slides.value[slideIndex.value]?.id // fitting shows each page; come back to this one
    let done = 0, next = 0, stop = false
    const same: number[] = [], failed: number[] = []
    let noChangeWhy = ''
    const worker = async () => {
      while (next < all.length && !stop) {
        const id = all[next++]
        try {
          await editSlide(pos(id), args.instruction)
        }
        catch (e) { // one page never stops the others; an unexpected error stops them all
          if (e instanceof NoChange) {
            same.push(pos(id) + 1)
            noChangeWhy ||= e.message
          }
          else if (e instanceof StepFailed) failed.push(pos(id) + 1)
          else {
            stop = true
            throw e
          }
        }
        busy.value = t('Working on the pages: {p0} of {p1}…', { p0: ++done, p1: all.length })
      }
    }
    // wait for every worker (an edit in flight still lands), then report the first unexpected error
    const settled = await Promise.allSettled(Array.from({ length: Math.min(5, all.length) }, worker))
    const err = settled.find((x): x is PromiseRejectedResult => x.status === 'rejected')
    if (err) throw err.reason
    // the request applies to every page, and some came back unchanged: one more try for those, naming that
    if (args.every_page) {
      const again = same.splice(0).map(p => all[p - 1])
      for (let k = 0; k < again.length; k += 5) {
        // every edit in flight lands before an error is reported: a rollback must not race a page still being written
        const retried = await Promise.allSettled(again.slice(k, k + 5).map(id => editSlide(pos(id), `${args.instruction}\nThis applies to every slide of the deck, this one included: make the change here too.`).catch((e) => {
          if (e instanceof NoChange) same.push(pos(id) + 1)
          else if (e instanceof StepFailed) failed.push(pos(id) + 1)
          else throw e
        })))
        const lost = retried.find((x): x is PromiseRejectedResult => x.status === 'rejected')
        if (lost) throw lost.reason
      }
    }
    await fitQueue
    if (viewing && pos(viewing) >= 0) slidesStore.updateSlideIndex(pos(viewing))
    const earlier = same.filter(p => editedThisTurn.has(all[p - 1])) // done by an earlier step of this turn
    const untouched = same.filter(p => !editedThisTurn.has(all[p - 1]))
    const changedNow = all.length - same.length - failed.length
    if (!changedNow && earlier.length && !untouched.length && !failed.length) return t('No further change: the earlier step already covered these pages.')
    const ok = changedNow + earlier.length
    const list = (xs: number[]) => xs.sort((a, b) => a - b).join(', ')
    const what = String(args.instruction || '').slice(0, 60)
    const why = (failed.length ? (t(' Page(s) {p0} failed.', { p0: list(failed) })) : '')
      + (untouched.length && ok ? (t(' Page(s) {p0} had nothing for that.', { p0: list(untouched) })) : '')
    void earlier
    if (!ok && !failed.length) return t('No page was changed. {p0}', { p0: noChangeWhy })
    if (!ok) fail((t('I could not make that change on any page: "{p0}".', { p0: what })) + why)
    return (t('Number of pages updated: {p0}.', { p0: ok })) + why
  }
  if (action === 'change_image') {
    const id = idOf(args.index)
    const sl = slides.value[pos(id)]
    const content = pictures(sl).filter(e => !isBackdrop(e))
    const said = saidBackground(text, args.said_background)
    let image = pictureNo(sl, args.image)
    // "change the picture(s)" means the pictures on the page, not the background, unless the user said background
    // (a user asked for slide 5's pictures and got its background replaced, 2026-11-10)
    if (image === 0 && !said && content.length) {
      image = content.findIndex((_, k) => !replacedThisTurn.has(`${id}:${k + 1}`)) + 1 || 1
    }
    // a page with only a background picture: "change the image / the cover photo" means that picture (the eval's
    // "Change the image on the slide" got a small extra picture, 2026-11-12); only a request to ADD one keeps it
    // ("insert real photos of X" swapped the background)
    // the page is one flat picture: the pictures the user sees are inside it, and replacing it would replace the page
    if (isFlat(sl) && !args.wants_additional_picture) fail(t('Page {p0} is one flat picture: its text and pictures are all inside that single image, so I can\'t change a picture in it. Nothing was changed. I can add a new picture on top of it.', { p0: pos(id) + 1 }))
    if (!said && !content.length && (args.wants_additional_picture || !firstBackdrop(sl))) {
      return runCall(text, 'add_image', { index: args.index, query: args.query, source: args.source, page_named: true, replaces_existing: false }, ids)
    }
    // "the pictures" / "all of them": every content picture on the page, in one step (the router asked for one per
    // picture replaced only the first of slide 5's three, 2026-11-13)
    // never a logo or an icon: "the pictures" swapped the company logo for a stock photo (review of 49a715d3)
    const logos = logoSpots(), { width: SW, height: SH } = slideSize()
    const photos = content.map((e, k) => ({ e, k: k + 1 })).filter(({ e }) => !logos.has(imgSpot(e)) && e.width * e.height >= 0.005 * SW * SH)
    if (args.all_pictures && photos.length > 1) {
      const left = photos.map(({ k }) => k).filter(k => !replacedThisTurn.has(`${id}:${k}`))
      if (!left.length) return ''
      busy.value = args.source === 'ai' ? t('Drawing the picture…') : t('Finding a picture…')
      slidesStore.updateSlideIndex(pos(id))
      let note = ''
      for (const k of left) {
        note = (await changeImage(text, id, k, args.query, args.source)).note || note
        replacedThisTurn.add(`${id}:${k}`)
      }
      const how = pictureSource === 'attached' ? (t('your attached')) : pictureSource === 'ai' ? (t('AI-generated')) : pictureSource === 'wikimedia' ? (t('Wikimedia')) : (t('stock photo'))
      return (t('Replaced {p0} pictures on page {p1} with {p2} pictures for "{p3}"; the background was not touched. Check they are what you meant.', { p0: left.length, p1: pos(id) + 1, p2: how, p3: String(args.query || '').slice(0, 60) })) + (note ? ` ${note}` : '')
    }
    // several pictures and nothing in the message picks one: ask (Quanlai: "if there are multiple choices you should at least ask")
    if (askedThisTurn) return '' // already asking which picture this message
    if (image !== 0 && content.length > 1 && !args.picture_named) {
      const { width: W, height: H } = slideSize()
      const where = (e: any) => [e.top + e.height / 2 < H / 3 ? t('top') : e.top + e.height / 2 > (2 * H) / 3 ? t('bottom') : t('middle'),
        e.left + e.width / 2 < W / 3 ? t('left') : e.left + e.width / 2 > (2 * W) / 3 ? t('right') : t('center')].join(', ')
      return runCall(text, 'ask_user', {
        question: t('Which picture on page {p0} should I change?', { p0: pos(id) + 1 }),
        options: [...content.slice(0, 3).map((e, k) => `${t('Picture {p0}', { p0: k + 1 })} (${where(e)})`), t('All of them')],
      }, ids)
    }
    busy.value = args.source === 'ai' ? t('Drawing the picture…') : t('Finding a picture…')
    slidesStore.updateSlideIndex(pos(id))
    const { note, backdrop } = await changeImage(text, id, image, args.query, args.source)
    replacedThisTurn.add(`${id}:${image ?? 1}`)
    const how = pictureSource === 'attached' ? (t('your attached')) : pictureSource === 'ai' ? (t('AI-generated')) : pictureSource === 'wikimedia' ? (t('Wikimedia')) : (t('stock photo'))
    // says WHICH picture changed, so a later reply about this step is grounded in what happened
    // the page's only picture is "the picture", full-slide or not: told "the background picture" was replaced, the
    // router undid its own correct change (2026-11-12)
    const which = backdrop && content.length ? (t('the background picture on page {p0}', { p0: pos(id) + 1 })) : (t('the picture on page {p0}', { p0: pos(id) + 1 }))
    return (t('Replaced {p0} with a {p1} picture for "{p2}"; check it is what you meant.', { p0: which, p1: how, p2: String(args.query || '').slice(0, 60) })) + (note ? ` ${note}` : '')
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
    if (!el) fail(t('Could not find that element.'))
    if (!args.copy) slidesStore.updateSlide({ elements: from.elements.filter(e => e.id !== el.id) }, from.id)
    const target = slides.value[pos(to)]
    slidesStore.updateSlide({ elements: [...target.elements, { ...JSON.parse(JSON.stringify(el)), id: nanoid(10) }] }, to)
    slidesStore.updateSlideIndex(pos(to))
    return args.copy ? (t('Copied.')) : (t('Moved.'))
  }
  if (action === 'set_background') {
    // "the same background as the rest": copied from the nearest other page, never a photo search for "same style"
    // (the doctor's slide 5 got a stock photo first, then four more calls to undo it, 2026-11-11)
    if (args.like_other_pages) { // no page named: this page (never every page, which would leave nothing to copy from)
      const pages = args.pages?.length ? args.pages : [slideIndex.value + 1]
      return runCall(text, 'copy_page_style', { from: pages[0], page_named: false, pages }, ids)
    }
    // a page that is one flat picture has its background inside that picture: replacing or fading it wipes the page
    const named = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    const flat = named.filter(id => isFlat(slides.value[pos(id)]))
    const flatNote = flat.length ? ' ' + t('Page(s) {p0} are one flat picture each: their background is part of that picture, so they were not changed.', { p0: flat.map(id => pos(id) + 1).join(', ') }) : ''
    const targets = named.filter(id => !flat.includes(id))
    if (!targets.length) fail(flatNote.trim())
    let background: any = { type: 'solid', color: args.color }
    if (args.photo_query) {
      busy.value = t('Finding a picture…')
      // never the photo already there ("change the background", then "again", gave the same one, 2026-11-12)
      const src = await findPicture(args.photo_query, 'photo', true, [...usedBackgrounds, ...slides.value.map(x => x.background?.image?.src).filter(Boolean) as string[]], true)
      if (!src) fail(t('Could not get a background picture.'))
      usedBackgrounds.add(src)
      background = { type: 'image', image: { src, size: 'cover' } }
    }
    else if (!/^#[0-9a-fA-F]{6}$/.test(args.color || '')) fail(t('Which background colour?'))
    for (const id of targets) slidesStore.updateSlide({ background }, id)
    // a full-slide picture covers the background fill: a new background PHOTO replaces that picture, or nothing shows
    if (args.photo_query) {
      for (const id of targets) {
        for (const e of slides.value[pos(id)].elements.filter((x: any) => x.type === 'image' && isBackdrop(x))) {
          slidesStore.updateElement({ id: e.id, props: { src: background.image.src } as any, slideId: id })
          showBackdrop(id, e.id)
        }
      }
    }
    // same for a colour: the picture has to go, or "make it royal blue" changes nothing anyone can see
    if (!args.photo_query) for (const id of targets) fadeBackdrop(id)
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
    const note = missed.length ? (t(' (the large shapes on page(s) {p0} could not be recoloured and may still cover it)', { p0: missed.join(', ') })) : ''
    const bgHow = args.photo_query
      ? (pictureSource === 'ai' ? (t('an AI-generated picture for "{p0}"', { p0: String(args.photo_query).slice(0, 60) })) : (t('a stock photo for "{p0}"', { p0: String(args.photo_query).slice(0, 60) })))
      : (t('colour {p0}', { p0: args.color }))
    return (t('Background of {p0} page(s) set to {p1}.', { p0: targets.length, p1: bgHow })) + note + flatNote
  }
  if (action === 'write_notes') {
    const hasNotes = (id: string) => !!(slides.value[pos(id)]?.remark || '').replace(/<[^>]+>/g, '').trim()
    const asked = args.pages?.length ? nums(args.pages) : args.said_all ? slides.value.map(x => x.id) : [slides.value[slideIndex.value].id]
    if (args.clear) {
      for (const id of asked) slidesStore.updateSlide({ remark: '' }, id)
      touched(asked)
      return t('Cleared the speaker notes on {p0} page(s).', { p0: asked.length })
    }
    // revising touches only pages that have notes; with no notes anywhere ("make the script funnier") write them that way
    const having = args.existing_only ? asked.filter(hasNotes) : asked
    const targets = having.length ? having : asked
    const written = await writeNotes(targets, args.instruction || '')
    if (!written) fail(t('The speaker notes could not be written.'))
    return (t('Wrote speaker notes for {p0} page(s).', { p0: written })) + (written < targets.length ? (t(' ({p0} page(s) failed)', { p0: targets.length - written })) : '')
  }
  if (action === 'download_deck') {
    if (chatslide && !chatslide.canExport()) return t('Downloading needs a paid plan; I opened the plans for you.')
    if (s.mode === 'template' && !s.edited) { // the engine's own .pptx is exact; after editor changes, export the editor
      const a = document.createElement('a')
      a.href = s.versions[s.versions.length - 1].deckUrl
      a.download = 'deck.pptx'
      a.click()
      return t('Downloading the .pptx now.')
    }
    if (args.format === 'pdf' || args.format === 'images') {
      const n = await downloadRendered(args.format, (i, all) => {
        busy.value = t('Exporting page {p0} of {p1}…', { p0: i, p1: all })
      })
        .catch(() => fail(t('I could not create that file this time; the PowerPoint download still works.')))
      return args.format === 'pdf'
        ? (t('Downloading the PDF ({p0} pages).', { p0: n }))
        : (t('Downloading {p0} images (zip).', { p0: n }))
    }
    busy.value = t('Exporting the .pptx…')
    exportPPTX(slides.value, true, false) // the browser saves the file; no dialog
    const asked = String(args.asked || '').trim()
    const other = asked && !/pptx|powerpoint|ppt/i.test(asked) // a format we cannot export to directly (Google Slides, Keynote…)
    return other
      ? (t('There is no direct export to {p0} here; downloading a .pptx, which {p1} opens directly.', { p0: asked, p1: asked }))
      : (t('Downloading the .pptx now.'))
  }
  if (action === 'share_deck') {
    if (!hostSteps()?.share) fail(t('Sharing is not available here: open this project in ChatSlide to share it.'))
    if (chatslide && !chatslide.canExport()) return t('Sharing needs a paid plan; I opened the plans for you.')
    const { url, copied } = await hostSteps()!.share!()
    return (t('Sharing is on: {p0}', { p0: url })) + (copied ? (t(' (link copied)')) : '')
  }
  if (action === 'open_step') {
    if (!hostSteps()?.goto && args.step === 'scripts') { // no Scripts step outside ChatSlide: the notes here ARE the script
      const targets = slides.value.map(x => x.id)
      const written = await writeNotes(targets, text)
      if (!written) fail(t('The speaker notes could not be written.'))
      return t('Wrote the speaker notes (the script) for {p0} page(s).', { p0: written })
    }
    if (!hostSteps()?.goto) fail(t('That step is not available here: open this project in ChatSlide.'))
    await hostSteps()!.goto!(args.step)
    return t('Opening {p0}.', { p0: ({ video: t('the Video step'), scripts: t('the Scripts step'), outlines: t('the Outlines step'), dashboard: t('your Dashboard, where a new project starts') } as Record<string, string>)[args.step] || args.step })
  }
  if (action === 'open_refund') {
    if (!hostSteps()?.goto) return t('You can request a refund at https://app.chatslide.ai/refund')
    await hostSteps()!.goto!('refund')
    return t('Opening the refund page.')
  }
  if (action === 'apply_theme') {
    const hex = (c: any) => /^#[0-9a-fA-F]{6}$/.test(c || '')
    if (!hex(args.background) || !hex(args.font_color) || !(args.colors || []).every(hex) || !(args.colors || []).length) fail(t('That theme has invalid colours.'))
    for (const sl of slides.value) fadeBackdrop(sl.id) // a full-slide picture would hide it
    applyPresetTheme({ background: args.background, fontColor: args.font_color, fontname: args.font || '', colors: args.colors }, true)
    touched(slides.value.map(x => x.id))
    return t('Applied a new theme to the whole deck.')
  }
  if (action === 'say_unable') {
    // only says "recorded" when it was: outside ChatSlide nothing records it
    const saved = await (hostSteps()?.featureRequest?.(String(args.what || ''), text) ?? Promise.resolve(false)).catch(() => false)
    return [args.message, saved ? (t('I have recorded this as a feature request.')) : ''].filter(Boolean).join(' ')
  }
  if (action === 'set_page_size') {
    if (!(args.width > 0) || !(args.height > 0)) fail(t('Which page size?'))
    slidesStore.setViewportRatio(args.height / args.width)
    touched(slides.value.map(x => x.id))
    return t('Slide shape set to {p0}:{p1}.', { p0: args.width, p1: args.height })
  }
  if (action === 'copy_page_style') {
    // "this slide" is the source only when it is not the page being changed: "make slide 5's background like the rest"
    // on slide 5 copied slide 5 onto itself ("applied to 0 slides") and the router fell back to a stock photo (2026-11-11)
    const here = slideIndex.value + 1
    const targetsAsked = (args.pages || []).map(Number)
    let src = args.page_named === false && !targetsAsked.includes(here) ? slides.value[slideIndex.value] : slides.value[pos(idOf(args.from))]
    if (targetsAsked.includes(pos(src.id) + 1)) { // the source is one of the pages being changed: the nearest page that is not
      const k = slides.value.map((_, j) => j + 1).filter(j => !targetsAsked.includes(j)).sort((x, y) => Math.abs(x - here) - Math.abs(y - here))[0]
      if (k) src = slides.value[k - 1]
    }
    const html = (e: any) => String(e?.content ?? e?.text?.content ?? '')
    const first: any = readingOrder(src).find((e: any) => elText(e))
    const color = html(first).match(/(?<![-\w])color:\s*([^;"]+)/)?.[1] || first?.defaultColor || first?.text?.defaultColor
    const backdrop: any = src.elements.find((e: any) => e.type === 'image' && isBackdrop(e))
    const targets = (args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)).filter((id: string) => id !== src.id)
    const setIn = (h: string) => h.replace(/(?<![-\w])color:\s*[^;"]+/g, `color: ${color}`)
    for (const id of targets) {
      const sl = slides.value[pos(id)]
      const elements = sl.elements.filter((e: any) => !(e.type === 'image' && isBackdrop(e))).map((e: any) => {
        if (!color || !elText(e)) return e
        return e.type === 'text' ? { ...e, defaultColor: color, content: setIn(e.content) } : { ...e, text: { ...e.text, defaultColor: color, content: setIn(e.text.content) } }
      })
      slidesStore.updateSlide({ background: JSON.parse(JSON.stringify(src.background || {})), elements: backdrop ? [{ ...JSON.parse(JSON.stringify(backdrop)), id: `${backdrop.id}-${id}`.slice(0, 40) }, ...elements] : elements }, id)
    }
    touched(targets)
    return t("Applied page {p0}'s background and text colour to {p1} page(s).", { p0: pos(src.id) + 1, p1: targets.length })
  }
  if (action === 'open_design') {
    if (!isPC()) { // phones have no Design panel: change the design for them (a phone user asked and was turned away, 2026-11-10)
      const bg = slides.value[0]?.background?.color
      const pick = PRESET_THEMES.filter(p => p.background !== bg)
      const th = pick[Math.floor(Math.random() * pick.length)]
      for (const sl of slides.value) fadeBackdrop(sl.id)
      applyPresetTheme(th, true)
      touched(slides.value.map(x => x.id))
      return t('Applied a new theme to the whole deck.')
    }
    mainStore.setToolbarState(ToolbarStates.SLIDE_DESIGN)
    chatLeft.value = 8
    return t('I opened the Design panel on the right, where you can change the theme. Template pages are in the arrow next to "Add slide", and they insert as new pages.')
  }
  if (action === 'talk_to_human') {
    if (!hostSteps()?.support) fail(t('Please email contact@chatslide.ai to reach the support team.'))
    busy.value = t('Contacting support…')
    return (await hostSteps()!.support!(args.message || userText)) || (t('Sent to the support team.'))
  }
  if (action === 'go_to_page') {
    const id = idOf(args.page)
    slidesStore.updateSlideIndex(pos(id))
    return t('Showing page {p0}.', { p0: pos(id) + 1 })
  }
  if (action === 'ask_user') {
    const opts = (args.options || []).filter(Boolean).slice(0, 4)
    s.pending = { request: userText, question: args.question || '', options: opts }
    save()
    askedThisTurn = true // the answer comes in the next message
    return args.question || (t('Could you say that more specifically?'))
  }
  // "back to how it was at first": one step to the page as opened, not a chain of undos (a user needed five, 2026-11-11)
  // Version history: the deck's saved copies (ChatSlide keeps the newest 20 and one per hour for 7 days), offered as
  // buttons. A doctor asked for "yesterday's version" and "the version history" after a day's work was replaced, and
  // the editor had neither (2026-11-12).
  if (action === 'version_history') {
    askedThisTurn = true // an answer either way: the turn ends here and nothing changed
    return listVersions(text, /^\d{4}-\d{2}-\d{2}$/.test(String(args.day || '')) ? args.day : '')
  }
  if (action === 'restore_original') {
    // only a tap on the confirmation restores more than one page (the router's own flag is never enough)
    const ok = args.confirmed && tappedRestore
    // "everything back": the deck as opened, its deleted pages back and pages added since gone (Undo reverses it)
    if (args.whole_deck && Object.keys(openedSlides).length) {
      if (!ok) {
        s.pending = { request: text, question: '', options: [t('Restore the whole deck')], offer: true,
          repeat: [{ action: 'restore_original', args: { whole_deck: true, confirmed: true } }], pages: [1] }
        askedThisTurn = true // the turn ends on the confirmation; nothing changed
        return t('This puts back the whole deck as it was when you opened the editor this time: {p0} pages (you have {p1} now). Every change since then is removed; Undo brings it back. Tap to restore.', { p0: Object.keys(openedSlides).length, p1: slides.value.length })
      }
      slidesStore.setSlides(JSON.parse(JSON.stringify(Object.values(openedSlides))))
      ids.splice(0, ids.length, ...slides.value.map(x => x.id))
      touched(slides.value.map(x => x.id))
      slidesStore.updateSlideIndex(0)
      return t('Restored {p0} page(s) to how they were when you opened the deck.', { p0: slides.value.length })
    }
    const targets = (args.pages?.length ? nums(args.pages) : [slides.value[slideIndex.value]?.id]).filter(Boolean)
    const back = targets.filter((id: string) => openedSlides[id])
    if (!back.length) fail(t('This page was added after the deck was opened, so there is no earlier version of it.'))
    if (back.length > 1 && !ok) { // several pages: the same confirmation, then each page in turn
      s.pending = { request: text, question: '', options: [t('Restore these pages')], offer: true,
        repeat: [{ action: 'restore_original', args: { whole_deck: false, confirmed: true } }], pages: back.map((id: string) => pos(id) + 1) }
      askedThisTurn = true
      return t('This puts pages {p0} back as they were when you opened the editor this time. Every change on them since then is removed; Undo brings it back. Tap to restore.', { p0: back.map((id: string) => pos(id) + 1).join(', ') })
    }
    for (const id of back) { const { id: _id, ...page } = JSON.parse(JSON.stringify(openedSlides[id])); void _id; slidesStore.updateSlide(page, id) } // notes, animations and transition too
    touched(back)
    slidesStore.updateSlideIndex(pos(back[0]))
    return t('Restored {p0} page(s) to how they were when you opened the deck.', { p0: back.length })
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
      else return t('There is nothing to undo yet.')
    }
    else await snapshotStore.reDo()
    return (action === 'undo' ? t('Undone.') : t('Redone.'))
  }
  return args.answer || (t('Could you say that more specifically?'))
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
let removeConfirmable = '' // the pages the previous message asked to remove pictures from, answered yes
let deleteConfirmable: string[] = [] // pages the previous message asked about: this message may delete them
let deleteAnswered: string[] = [] // the same, answered in words: deletable when the router asks for exactly them again
let askedThisTurn = false // a question for the user ends the message
let directAllowed = false // true only while add_site_pictures places the pictures it fetched
const usedBackgrounds = new Set<string>() // background photos already tried in this session: "again" never brings one back
let tappedRestore = false // this message is the tap on a restore confirmation
let turnSelected = '' // what was selected on the canvas when the message was sent
let advancedOffered = false // the paid-model offer, shown once per session
let openedSlides: Record<string, any> = {} // slide id -> the page as it was when this project was first opened here
const helperPages = new Set<string>() // pages the editor added for pictures: not "new pages the user asked for"
// each page as it was when the editor opened the deck this time. Never an older session's copy: "go back to the previous
// version" restored a deck from two days earlier over a day of a doctor's work (2026-11-12)
const rememberOriginal = () => {
  openedSlides = Object.fromEntries(slides.value.map(x => [x.id, JSON.parse(JSON.stringify(x))]))
}
const GEOMETRY = ['left', 'top', 'width', 'height', 'rotate'] // numbers only, whatever the model writes
// a number, or a non-blank string that is one ("190"); anything else (":", "", null, false) is NaN
const toNum = (v: any) => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' ? Number(v) : NaN)
// the deck after the last step that finished: a crash in a later step keeps the finished steps (a user lost a whole
// turn's slide edit to one failing picture step, 2026-11-11)
let lastGoodSlides: any = null
const deckTurn = async (text: string) => {
  // a tap on a version-history button: that saved copy becomes the deck (one Undo step), no AI call
  newTurn() // a new message: charged on its own
  replyTurn = currentTurn()
  const versionUrl = s.savedVersions?.[text.trim()]
  s.savedVersions = undefined
  if (versionUrl) {
    s.pending = null
    const saved = await (await fetch(versionUrl, { cache: 'no-store' })).json()
    if (!saved?.slides?.length) fail(t('That version could not be loaded.'))
    // the deck on screen is saved first, so it stays in the version history
    await saveNow().catch(() => fail(t('The deck you have now could not be saved, so nothing was restored. Check your connection and try again.')))
    if (saved.theme) slidesStore.setTheme(saved.theme)
    if (saved.viewportRatio) slidesStore.setViewportRatio(saved.viewportRatio)
    slidesStore.setSlides(saved.slides)
    slidesStore.updateSlideIndex(0)
    deckBeforeCreate = null // Undo steps back to the deck that was on screen, by the snapshot
    s.edited = true
    addHistorySnapshot()
    const line = t('Went back to the version from {p0}: {p1} pages. Undo brings back the deck you had.', { p0: text.trim().split(' · ')[0], p1: saved.slides.length })
    say('assistant', line)
    recordTurn(text, line, ['load_version'], true)
    return
  }
  if (!Object.keys(openedSlides).length) rememberOriginal() // a deck opened another way (?pptx=, a local session): its state before the first change
  // the Yes button confirms; so does a typed reply to the question when the router, reading it, asks again to delete
  // exactly those pages as the user's own wish ("можешь их заменить!" got the same question back, 2026-11-12)
  deleteConfirmable = s.deleteAsked && text.trim() === t('Yes, delete them') ? s.deleteAsked : []
  deleteAnswered = s.deleteAsked && text.trim() !== t('No, keep them') ? s.deleteAsked : []
  s.deleteAsked = undefined
  removeConfirmable = s.removeAsked && text.trim() === t('Yes, remove them') ? s.removeAsked : '' // the button only: any other reply let the router remove them unasked
  s.removeAsked = undefined
  askedThisTurn = false
  userText = text
  editedThisTurn.clear()
  picturedThisTurn.clear()
  roomMadeThisTurn.clear()
  replacedThisTurn.clear()
  const pagesAtStart = slides.value.length
  takeSpend() // spend from an earlier, failed turn never lands in this footnote
  // an offer counts only when this message is its button; otherwise it is dropped (a question holds for any reply)
  if (s.pending?.offer && !s.pending.options?.some(o => o.trim() === text.trim())) { s.pending = null; save() }
  const pendingAtStart = s.pending || null
  const done: string[] = []
  const replies: string[] = []
  const terse: string[] = [] // per step: the edit model's own sentence can contradict the rest of a multi-step turn
  const free: string[] = [] // lines the router wrote itself (any language): the reply step puts them in the user's
  let changed = false
  const turnCalls: string[] = [] // stepKey of each step that ran, for the next message's Undo guard
  const ran: any[] = [] // the calls that ran
  let unfinished = false // a deck-wide step had to wait for the next message
  const turnExact: string[] = [] // action + exact arguments of each step that ran in this message
  const startPageId = slides.value[slideIndex.value]?.id
  // what the user has selected on the canvas when they send: "this picture" / "make it bigger" means it
  const startSl = slides.value[slideIndex.value]
  const selected = turnSelected = startSl ? startSl.elements.filter((e: any) => mainStore.activeElementIdList.includes(e.id)).map((e: any) => {
    // numbered as changeImage resolves them: 0 is the first full-slide picture, the rest by pictures()
    if (e.type === 'image') {
      const n = pictures(startSl).indexOf(e) + 1
      return n ? `image ${n}` : e === firstBackdrop(startSl) ? 'image 0 (the background picture)' : 'a see-through design layer over the page (not a picture to change)'
    }
    if (e.type === 'chart') return `the ${e.chartType} chart`
    if (e.type === 'table') return 'the table'
    const tx = elText(e)
    return tx ? `the text "${tx.slice(0, 80)}"` : `a ${e.type}`
  }).join('; ') : ''
  let pendingInstead = '', nudged = false // what an undo said the user wants instead, until a change does it
  let offerOrder = '' // the page order when a whole-deck offer was made in this message
  let failed = 0 // steps of this message that failed
  currentAtSend = startPageId || ''
  createdThisMessage.clear()
  const idsAtStart = new Set(slides.value.map(x => x.id)) // pages created in this message are the others
  const created = () => slides.value.filter(x => !idsAtStart.has(x.id) && !helperPages.has(x.id)).length
  const ADDS = ['add_page', 'add_chart_page', 'add_table_page', 'expand_page']
  // pages this message animated (an earlier step's arguments, as logged in done); none listed means every page
  const animatedThisTurn = (pages?: number[]) => done.some((d) => {
    if (!d.startsWith('animate_pages ') || d.includes('-> skipped')) return false
    const a = (() => { try { return JSON.parse(d.slice(14, d.indexOf(' -> '))) } catch { return {} } })()
    return a.effect !== 'none' && (!a.pages?.length || !pages?.length || pages.some((p: number) => a.pages.includes(p)))
  })
  let pagesWanted = -1 // declared by the first page-adding step (message_new_pages)
  // ponytail: one step per round; 16 lets a pasted 'Slide 1 … Slide 12' finish in one message (10 cut it off)
  // "Do the same on the other pages": the editor repeats the steps that ran on the first page, page by page, without
  // asking the router again (asked to re-plan it, the router redid page 1 and stopped, 2026-11-11)
  const repeat: any[] = pendingAtStart?.repeat && pendingAtStart.options?.some(o => o.trim() === text.trim())
    ? [...(pendingAtStart.pages || []).flatMap(pg => pendingAtStart.repeat!.map(c => ({ action: c.action,
        args: { ...c.args, index: pg, pages: [pg], page_named: true, whole_deck_option: '' } }))), { action: 'finish', args: { not_done: '' } }]
    : []
  const repeating = repeat.length > 0
  if (repeating && pendingAtStart?.from) {
    const general = await forAnyPage(String(pendingAtStart.repeat![0].args?.instruction || ''))
    for (const c of repeat) if (c.action === 'edit_page') c.args.instruction = general
  }
  tappedRestore = repeat.some((c: any) => c.action === 'restore_original')
  let round = 0 // 16 after the loop: it ran out of steps with the message unfinished
  const rounds = 16 + repeat.length // fixed before the queue drains (re-read each round it ended a long repeat early)
  for (; round < rounds; round++) {
    const { calls, error } = repeat.length ? { calls: [repeat.shift()], error: '' } : await (await fetch('/api/tools/deck_chat', {
      method: 'POST',
      body: JSON.stringify({
        // "this page" stays the page the user was on when they sent the message: steps that show a new page moved it,
        // and "add a page after this one" then re-ran after each new page (30 -> 36)
        instruction: text, pages: deckLines(), current: (slides.value.findIndex(x => x.id === startPageId) + 1) || slideIndex.value + 1, template: !!s.template, done,
        today: localDay(Date.now()), // "yesterday's version", "the version from 11.11"
        pending: pendingAtStart || undefined,
        selected: selected || undefined,
        files: s.docs.map(d => ({ name: d.name, pages: (d.text.match(/^\[page \d+\]$/gm) || []).length || undefined, chars: d.text.length })),
        pictures: (s.images || []).map(m => m.name),
        history: answeredHistory().map(h => ({ role: h.role, content: h.text })),
      }),
      headers: { 'Content-Type': 'application/json' },
    })).json()
    if (error) throw new Error(error)
    // replacing the deck needs a topic in this very message: a complaint ("same problem") rebuilt a deck on a guess
    const newDeck = calls.find((c: any) => c.action === 'create_deck')
    if (newDeck && !String(newDeck.args?.topic_in_this_message || '').trim()) {
      done.push('create_deck -> skipped: the latest message names no topic for a new deck. Do not replace the deck: edit it or ask_user')
      continue
    }
    // after steps that only read material (the attached paper, the web): the deck is built from it, with the router's
    // full instruction ("here is my paper" read the file and stopped, 2026-11-12)
    const onlyRead = done.every(d => /^(read_file|web_search|read_link|find_papers) /.test(d))
    if (newDeck?.args?.from_users_file && !s.docs.length && !(s.images || []).length) { // no file to build it from: ask for it
      const ask = await runCall(text, 'create_deck', newDeck.args, slides.value.map(x => x.id))
      replies.push(ask)
      terse.push(ask)
      break
    }
    if (newDeck && onlyRead) {
      const brief = done.length ? String(newDeck.args?.instruction || text) : text
      deckBeforeCreate = JSON.parse(JSON.stringify(slides.value)) // one Undo click brings this deck back
      openedSlides = {} // "the original" of a new deck is the new deck (taken at the next message)
      busy.value = s.template ? t('Building the deck (about a minute)…') : t('Writing the deck…')
      await (s.template ? templateTurn(brief) : generateNative(brief))
      addHistorySnapshot() // the new deck is a history step of its own
      return
    }
    if (calls.some((c: any) => c.action === 'rebuild_from_template') && s.template) return templateTurn(text)
    // one step per round: when the model batches calls it numbers later ones as if earlier ones had run, so the rest
    // is re-planned against the updated deck in the next round
    const work = calls.filter((c: any) => !['finish', 'rebuild_from_template', 'create_deck'].includes(c.action)).slice(0, 1)
    const left = calls.find((c: any) => c.action === 'finish')?.args?.not_done
    if (!work.length && left && String(left).trim()) { // never pretend the rest is done
      free.push(String(left).trim())
      replies.push(String(left).trim())
      terse.push(String(left).trim())
    }
    // a correction stopped at "Undone." (1 run in 4): the undo said what the user wants instead, so finishing
    // before that is done sends the router back once
    if (!work.length && pendingInstead && !nudged) {
      nudged = true
      done.push(`not finished: after the undo, the user still wants: ${pendingInstead}. Do it now on the right target.`)
      continue
    }
    if (!work.length) break
    const ids = slides.value.map(x => x.id) // page numbers in this round's calls refer to this deck
    if (work.every((c: any) => done.some(d => d.startsWith(`${c.action} ${JSON.stringify(c.args || {})} ->`)))) break // repeating itself
    // the same page edited twice already this turn: a third pass is the router looping (seen: 7 edits of page 2)
    if (work.some((c: any) => c.action === 'edit_page' && done.filter(d => d.startsWith('edit_page ') && new RegExp(`"index":${Number(c.args?.index)}[,}]`).test(d)).length >= 2)) break
    // one new picture per page per message: a second one on the same page was an unasked extra (seen: a turkey).
    // Compared on the page it would really land on; skipped (and told to the router), never ending the turn.
    const landsOn = (c: any) => ids[Number(c.args?.index) - 1]
    const extra = work.find((c: any) => c.action === 'edit_page' && c.args?.new_content && !c.args?.page_named && created() >= 1)
      || work.find((c: any) => c.action === 'delete_pages' && !c.args?.user_asked_to_cut && (c.args?.pages?.length || 0) >= Math.max(2, slides.value.length * 0.3))
      || work.find((c: any) => c.action === 'add_image' && picturedThisTurn.has(picKey(landsOn(c), c.args?.items?.length > 1 ? c.args.items.slice(0, 6).join('|') : c.args?.query)))
      // more new pages than the message asks for: re-plans added "the next" page, or extended the page just added
      // (30 -> 32 for "a page"). The first page-adding step declares the total; steps beyond it are skipped.
      || work.find((c: any) => ADDS.includes(c.action) && pagesWanted >= 0 && created() >= pagesWanted)
      // the very same step again (same action, same arguments): "add a page about TikTok" ran three times (30 -> 33)
      || work.find((c: any) => !['undo', 'redo', 'answer_question', 'ask_user', 'go_to_page'].includes(c.action)
        && turnExact.includes(`${c.action} ${JSON.stringify(c.args || {})}`))
      // a second deck-wide edit is the router re-planning the first ("dark blue, bigger headers" ran twice and ended
      // in "Not done yet … tap Continue", 2026-11-11); what is really left it names in finish
      || work.find((c: any) => c.action === 'edit_all_pages' && done.some(d => d.startsWith('edit_all_pages ')))
      // a request about pictures is never a background change: "replace the photos ... the last one" swapped page 10's
      // background (2026-11-10, replayed 2026-11-13)
      || work.find((c: any) => c.action === 'set_background' && c.args?.photo_query && !saidBackground(text, c.args?.said_background))
      // the same picture replaced again in one message (a re-plan picked a "better" one, three times)
      || work.find((c: any) => {
        const id = ids[Number(c.args?.index) - 1], sl = slides.value.find(x => x.id === id)
        return c.action === 'change_image' && !!sl && replacedThisTurn.has(`${id}:${pictureNo(sl, c.args?.image)}`)
      })
      // after an Undo in a correction, redoing the very change just undone ("not the background" -> background again)
      || work.find((c: any) => done.some(d => d.startsWith('undo ')) && lastTurnCalls.includes(stepKey(c.action, c.args)))
      // Undo reverts what EARLIER messages did; within a message the router once undid its own correct change
      || work.find((c: any) => (c.action === 'undo' || c.action === 'redo') && (changed || done.some(d => /^(undo|redo) /.test(d))))
      // nor removes the animations it just added: "one point per space bar press" ended "Animated 10 pages. Animations removed."
      || work.find((c: any) => c.action === 'animate_pages' && c.args?.effect === 'none' && animatedThisTurn(c.args?.pages))
    // undone for a "correction", then the router wants the very same change back: the undone change was already
    // what the user meant ("I asked for the image, not the background" after the image was changed). Put it back.
    if (extra?.action === 'change_image' && done.some(d => d.startsWith('undo ')) && lastTurnCalls.includes(stepKey(extra.action, extra.args))) {
      await snapshotStore.reDo()
      pendingInstead = ''
      // the "Undone." line no longer describes the deck: the change is back
      const at = done.findIndex(d => d.startsWith('undo '))
      if (at >= 0) {
        const undoneLine = done[at].split(' -> ')[1]
        const k = replies.indexOf(undoneLine)
        if (k >= 0) {
          replies.splice(k, 1)
          terse.splice(k, 1)
        }
      }
      const kept = t('My last change was already to the picture you meant (the background was not touched), so I kept it. Should I also restore the background?')
      replies.push(kept)
      terse.push(kept)
      break
    }
    if (extra) {
      const why = extra.action === 'delete_pages' ? 'the user did not ask to delete or cut these pages. Do not delete: ask_user what they want, with options'
        : extra.action === 'edit_page' && extra.args?.new_content ? 'that new content is already on the new page just added (see Done so far); finish'
          : ADDS.includes(extra.action) && pagesWanted >= 0 && created() >= pagesWanted
            ? `the ${pagesWanted} new page(s) this message asks for are already added (see Done so far); finish`
            : turnExact.includes(`${extra.action} ${JSON.stringify(extra.args || {})}`)
              ? 'this exact step already ran in this message (see Done so far): the request is done, finish'
              : lastTurnCalls.includes(stepKey(extra.action, extra.args)) && done.some(d => d.startsWith('undo '))
                ? 'that is exactly the change just undone because the user said it was wrong; do the request on the target the user means instead'
                : extra.action === 'animate_pages' ? 'the animations this message asks for were just added; keep them and finish'
                : extra.action === 'set_background' ? 'the user\'s words do not ask for the page background, so it stays: for pictures use change_image (all_pictures for every picture on a page), add_image or remove_images'
                : extra.action === 'edit_all_pages' ? 'a deck-wide edit already ran in this message (see Done so far); finish, and put in not_done only a part of the request it did not cover'
                : extra.action === 'add_image' ? 'that picture was already added to that page in this message' : extra.action === 'change_image' ? 'that picture was already replaced in this message' : 'undo only reverts changes from earlier messages, not this one'
      done.push(`${extra.action} ${JSON.stringify(extra.args || {})} -> skipped: ${why}`)
      // a router that keeps proposing what was just refused is done: "khác đi" replanned the same picture for 16 rounds
      if (done.length >= 2 && done.slice(-2).every(d => d.includes(' -> skipped: '))) break // two in a row
      continue
    }
    // one deck-wide pass per turn: after pages are deleted the numbering shifts and a re-plan deleted MORE pages
    // ("delete 25-28" took 6); a second edit_all_pages re-edits every page for nothing (30 pages, twice)
    // (a second duplicate_pages copied the copies again); the reply says the rest was left for another message
    const again = work.find((c: any) => ['delete_pages', 'edit_all_pages', 'reorder_pages', 'set_background', 'duplicate_pages'].includes(c.action) && done.some(d => d.startsWith(`${c.action} `)))
    // the same step on the same pages again is the router repeating itself, not a second request: the turn just ends
    // (a single background change ended in "Not done yet: set background … tap Continue", 2026-11-11)
    // the same step on the same pages, even with other words ("nền khác": a second photo search for the same page)
    // (only actions whose schema requires pages; the others repeat only when identical)
    // or pages that step already did ("slides 1, 2 and 3 yellow", then again for 1 and 2: "Not done yet", 2026-11-12)
    const samePages = (d: string) => {
      try {
        const before: number[] = JSON.parse(d.slice(again!.action.length + 1, d.indexOf(' -> '))).pages || []
        const now: number[] = again!.args?.pages || []
        return !before.length || (now.length > 0 && now.every(p => before.includes(p)))
      }
      catch { return false }
    }
    if (again && done.some(d => d.startsWith(`${again.action} `) && (['set_background', 'delete_pages', 'duplicate_pages'].includes(again.action) ? samePages(d) : d.startsWith(`${again.action} ${JSON.stringify(again.args || {})} ->`)))) break
    if (again) {
      // names what is left, and Continue picks it up (Quanlai: a user with 10 requests should get all of them done)
      const line = t('Not done yet: {p0}.', { p0: String(again.args?.instruction || again.action.replace(/_/g, ' ')).slice(0, 160) })
      unfinished = true
      replies.push(line)
      terse.push(line)
      break
    }
    for (const c of work) {
      // planned in the same batch as the change it would revert, or a second undo/redo ("Undone. Redone." for "undo")
      if ((c.action === 'undo' || c.action === 'redo') && (changed || done.some(d => /^(undo|redo) /.test(d)))) {
        done.push(`${c.action} -> skipped: undo only reverts changes from earlier messages, not this one`)
        continue
      }
      if (ADDS.includes(c.action) && pagesWanted < 0) pagesWanted = Math.max(1, Number(c.args?.message_new_pages) || Number(c.args?.count) || 1)
      // no step adds more than this message still allows
      if (ADDS.includes(c.action) && Number(c.args?.count) > 0) c.args.count = Math.min(Number(c.args.count), Math.max(1, pagesWanted - created()))
      else if (ADDS.includes(c.action) && created() >= pagesWanted) { // planned in the same batch as the pages already added
        done.push(`${c.action} -> skipped: the ${pagesWanted} new page(s) this message asks for are already added`)
        continue
      }
      busy.value = t('Working (step {p0})…', { p0: done.length + 1 })
      let reply = ''
      try {
        reply = await runCall(text, c.action, c.args || {}, ids)
      }
      catch (e) {
        if (!(e instanceof StepFailed)) throw e
        replies.push(e.message)
        terse.push(e.message)
        if (repeating && e instanceof NoChange) continue // one page already as asked: the other pages still get it
        // A failure used to end the message: "change the photo on slide 1 to a plainer one, and the photo on slide 3"
        // stopped at "there is no picture to remove" with slide 3 untouched (2026-11-12). Now the user is told at
        // once, the router gets it as that step's result, and it may go on with the OTHER parts of the request. Still
        // the end of the message: a question to the user ("Which font?"), a page that is already as asked, a repeat
        // over pages, or a second failure.
        failed++
        if (repeating || e instanceof NoChange || /[?？؟]\s*$/.test(e.message) || failed >= 2) { round = 99; break }
        done.push(`${c.action} ${JSON.stringify(c.args || {})} -> failed, and the user was told: ${e.message} (Do not retry it. Go on with any other part of the request.)`)
        turnExact.push(`${c.action} ${JSON.stringify(c.args || {})}`)
        continue
      }
      replies.push(reply)
      if (['answer_question', 'ask_user', 'say_unable'].includes(c.action)) free.push(reply)
      ran.push(c)
      if (!['answer_question', 'ask_user', 'say_unable', 'read_file', 'web_search', 'find_papers', 'read_link', 'go_to_page', 'open_design', 'download_deck'].includes(c.action)) {
        lastGoodSlides = JSON.parse(JSON.stringify(slides.value)) // what a later step's crash rolls back to
      }
      if (c.action === 'undo') pendingInstead = String(c.args?.instead || '').trim()
      else if (!['answer_question', 'ask_user', 'redo', 'go_to_page', 'read_file', 'web_search', 'find_papers', 'read_link'].includes(c.action)) pendingInstead = ''
      turnCalls.push(stepKey(c.action, c.args))
      turnExact.push(`${c.action} ${JSON.stringify(c.args || {})}`)
      for (const x of slides.value) if (!idsAtStart.has(x.id)) createdThisMessage.add(x.id)
      terse.push(c.action === 'edit_page' ? lastEditTerse : reply)
      // the whole-deck offer is a button, not the end of the message: "make the title red and add a picture" stopped
      // after the title (2026-11-12). The router is told, so it neither applies it itself nor stops there.
      const offered = c.action === 'edit_page' && s.pending?.offer && s.pending.options?.[0] === String(c.args?.whole_deck_option || '').trim()
        ? ' (The user now has a button to apply this to the whole deck: never do that yourself. Go on only with a different part of the request, if there is one.)' : ''
      if (offered) offerOrder = slides.value.map(x => x.id).join()
      else if (offerOrder && s.pending?.offer && offerOrder !== slides.value.map(x => x.id).join()) { s.pending = null; offerOrder = '' } // pages added, deleted or moved since: its numbers are stale
      done.push(`${c.action} ${JSON.stringify(c.args || {})} -> ${reply}${offered}${memo ? `\n${memo}` : ''}`)
      memo = ''
      if (!['answer_question', 'ask_user', 'say_unable', 'undo', 'redo', 'download_deck', 'read_file', 'web_search', 'find_papers', 'read_link', 'go_to_page', 'open_design'].includes(c.action) && !askedThisTurn) changed = true
    }
    if (askedThisTurn || work.every((c: any) => ['answer_question', 'ask_user', 'say_unable', 'go_to_page', 'open_design'].includes(c.action))) break // a question is asked/answered once; showing a page ends the turn
  }
  if (pendingAtStart && s.pending === pendingAtStart) { // answered; but never wipe a NEW question asked this turn
    s.pending = null
    save()
  }
  lastTurnCalls = turnCalls
  // Only this page changed and the message named no page: done here, the rest offered as one tap (Quanlai, 2026-11-11:
  // "we can change and ask a follow-up question"). "Add historical images" / "too much text" changed one page silently.
  // edit_page has its own offer (whole_deck_option, set by the router when the request could mean the deck): adding facts
  // to one slide got a pointless 'the same on the other pages?' (2026-11-11). set_background: one pass per message
  const perPage = ['add_image', 'change_image', 'set_font_size', 'set_font', 'remove_images']
  let followUp = ''
  // more to do than one message runs (16 steps, one deck-wide pass of a kind): one tap continues, nothing is dropped
  if ((unfinished || round === rounds) && !askedThisTurn) {
    followUp = t('There is more to do from your message. Tap Continue and I will carry on.')
    // answering an earlier question or offer (Continue included): the request stays the original one, steps accumulate
    s.pending = { request: pendingAtStart ? pendingAtStart.request : text, question: followUp, options: [t('Continue')], localize: true, done: [...(pendingAtStart?.done || []), ...done].slice(-30) } // any reply continues it
    save()
  }
  else if (changed && !turnSelected && !s.pending && !askedThisTurn && slides.value.length > 1 && editedThisTurn.size === 1 && editedThisTurn.has(startPageId)
    && ran.length && ran.every(c => perPage.includes(c.action) && c.args?.page_named !== true && !(c.args?.items?.length > 1))) { // a picture set may add pages: not repeated
    followUp = t('Done on this page. Want the same on the other pages?')
    const here = slides.value.findIndex(x => x.id === startPageId) + 1
    const others = slides.value.map((_, k) => k + 1).filter(k => k !== here)
    // the router learns which pages are left and what was done here, so it neither redoes this page nor skips one
    s.pending = { request: text, question: `${followUp} (the other pages: ${others.join(', ')})`, options: [t('Do the same on the other {p0} pages', { p0: others.length })], done: done.slice(-30), offer: true, localize: true,
      repeat: ran.map(c => ({ action: c.action, args: c.args })), pages: others }
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
  if (slides.value.length !== pagesAtStart) lines.push(t('The deck now has {p0} pages.', { p0: slides.value.length }))
  if (followUp) lines.push(followUp)
  const reply = [...new Set(lines.filter(Boolean))].join('\n') || (t('Done.')) // every step reports: merging repeats hid a doubled page
  const sp = takeSpend()
  const fee = sp.charged > 0 ? (t('\n(Credits used: {p0}{p1})', { p0: sp.charged, p1: sp.left >= 0 && sp.left <= 9000000 ? t(' · {p0} left', { p0: sp.left }) : '' })) : ''
  const said = await finalize(text, reply, free, s.pending !== pendingAtStart && s.pending?.localize ? s.pending : null)
  say('assistant', said + fee)
  recordTurn(text, said, done, true)
}

// ChatSlide keeps a record of every chat turn for review: the message, the reply, and what each step did (step,
// arguments, result), so a complaint like "I don't see a change" can be traced to the edit before it. Never charged.
const recordTurn = (message: string, reply: string, steps: string[], ok: boolean) => {
  if (!chatslide) return
  fetch('/api/record_turn', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ message, reply, ok, page: slideIndex.value + 1, pages: slides.value.length, steps: steps.map(d => d.split('\n')[0].slice(0, 400)).slice(0, 30) }),
  }).catch(() => undefined)
}

// "Apply to the whole deck" repeats an edit_page instruction the router wrote for one page, and its topic went onto
// every other page ("keep its focus on obturation timing" on all 10; a roasting paragraph on "Coffee Today", even when
// told to keep each page's own topic, 2026-11-12). So the instruction is rewritten once, for any page.
const forAnyPage = async (instruction: string) => {
  try {
    const r = await (await fetch('/api/tools/ask_json', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ instruction: `This edit instruction was written for one slide. Rewrite it so it applies to ANY slide of the deck: keep the kind of change, its amount and style, any text meant for every slide (a footer, a reference, a name) and any line in parentheses word for word; drop the topic, facts and wording that describe that one slide's own content (each slide uses its own).\nInstruction: ${instruction.slice(0, 3000)}\nReturn JSON {"instruction": "..."}` }) })).json()
    return typeof r.instruction === 'string' && r.instruction.trim() ? r.instruction.trim() : instruction
  }
  catch {
    return instruction
  }
}

// Every turn's reply in the language of the user's message: the code's lines follow the app language and the
// router's own sentences (say_unable, ask_user) came back in Chinese for English messages without this step.
// The follow-up buttons of this message too: the router's "Apply to the whole deck" label came in the deck's language
// (Spanish for a Vietnamese user) and the code's ones in the app's. A tap sends the label, so the pending keeps the
// translated labels it matches against.
const finalize = async (instruction: string, reply: string, free: string[] = [], offer: NonNullable<typeof s.pending> | null = null) => {
  try {
    const options = offer?.options || []
    const r = await (await fetch('/api/tools/deck_chat', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ localize: reply, instruction, ui_language: lang, free, options, earlier: s.history.filter(h => h.role === 'user').slice(-4, -1).map(h => h.text) }) })).json()
    if (offer && s.pending === offer && Array.isArray(r.options) && r.options.length === options.length && r.options.every((o: unknown) => typeof o === 'string' && o.trim())) {
      offer.options = r.options.map((o: string) => o.trim())
      save()
    }
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
  lastGoodSlides = null
  return run(t('Working…'), async () => {
    try {
      await deckTurn(text)
      save()
    }
    catch (e) { // never leave a half-applied operation (e.g. an inserted page that was never rewritten)
      const kept = !!lastGoodSlides
      slidesStore.setSlides(lastGoodSlides || before)
      if (kept) { (e as any).keptSteps = true; save() }
      // where it broke goes into the record: a user saw a raw "Unexpected non-whitespace character after JSON" and
      // its source could not be found afterwards (2026-11-11)
      const err = e as Error
      const where = err?.stack ? err.stack.split('\n').slice(0, 5).join(' | ') : ''
      recordTurn(text, String(err?.message || e), [`error: ${err?.name || 'Error'}: ${err?.message || e} @ ${where}`], false)
      throw e
    }
  }).finally(() => { replyTurn = '' })
}

// One attach control (or drop files on the box): a .pptx becomes the company template while none is set and no deck
// was built or opened yet (s.edited: a ChatSlide deck, which a template import would replace); everything else is material
const attach = (files?: FileList | File[] | null) => {
  if (busy.value || !files) return
  for (const f of Array.from(files)) {
    if (/^image\//.test(f.type)) uploadPicture(f)
    else if (/\.pptx$/i.test(f.name) && !s.template && !s.mode && !s.edited) uploadTemplate(f)
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
// one off-screen renderer: a sync and an export take turns, or one replaces the pages the other is drawing
let renderer: Promise<unknown> = Promise.resolve()
const withRenderer = <T>(job: () => Promise<T>): Promise<T> => {
  const run = renderer.then(job)
  renderer = run.catch(() => undefined)
  return run
}
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
// PDF / images straight from the chat, using the same off-screen pages: a real file (the Export window's PDF is the
// browser's print dialog, which gives no file). A page that cannot be drawn fails the export, it never goes missing.
const downloadRendered = (format: 'pdf' | 'images', progress: (i: number, n: number) => void) => withRenderer(async () => {
  toRender.value = JSON.parse(JSON.stringify(slides.value))
  try {
    await nextTick()
    await new Promise(r => setTimeout(r, 800)) // images inside the pages load
    const nodes = [...(renderRef.value?.children || [])] as HTMLElement[]
    const fontEmbedCSS = await getFontEmbedCSS(nodes[0]).catch(() => '')
    const shots: JpegPage[] = []
    for (const [i, node] of nodes.entries()) {
      progress(i + 1, nodes.length)
      node.querySelectorAll('foreignObject [xmlns]').forEach(e => e.removeAttribute('xmlns'))
      const url = await toJpeg(node, { quality: 0.92, width: 1600, fontEmbedCSS, backgroundColor: '#ffffff' })
      const img = new Image()
      img.src = url
      await img.decode()
      shots.push({ bytes: new Uint8Array(await (await fetch(url)).arrayBuffer()), width: img.naturalWidth, height: img.naturalHeight })
    }
    const name = slidesStore.title || 'deck'
    if (format === 'pdf') saveAs(jpegsToPdf(shots), `${name}.pdf`)
    else {
      const zip = new JSZip()
      shots.forEach((p, i) => zip.file(`${name}-${String(i + 1).padStart(2, '0')}.jpg`, p.bytes))
      saveAs(await zip.generateAsync({ type: 'blob' }), `${name}-images.zip`)
    }
    return shots.length
  }
  finally {
    toRender.value = []
  }
})

let syncing: Promise<void> | null = null
const syncOnce = async () => {
  const deck: Slide[] = JSON.parse(JSON.stringify(slides.value)) // what is sent is what was rendered
  const keys = pageKeys(deck)
  if (s.synced && s.synced.keys.join() === keys.join()) return
  const urls = s.synced?.urls || {}
  const need = deck.filter((_, i) => !urls[keys[i]])
  const images: Record<string, Blob> = {}
  if (need.length) {
    await withRenderer(async () => {
      toRender.value = need
      try {
        await nextTick()
        await new Promise(r => setTimeout(r, 400)) // images inside the pages load
        const nodes = [...(renderRef.value?.children || [])] as HTMLElement[]
        for (const [j, sl] of need.entries()) {
          nodes[j].querySelectorAll('foreignObject [xmlns]').forEach(e => e.removeAttribute('xmlns'))
          const dataUrl = await toJpeg(nodes[j], { quality: 0.9, width: 1600, fontEmbedCSS: '' })
          images[keys[deck.indexOf(sl)]] = await (await fetch(dataUrl)).blob()
        }
      }
      finally {
        toRender.value = []
      }
    })
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
  }
})()

// Bring the deck back after a page reload.
onMounted(() => {
  if (chatslide) { // ChatSlide project: its saved editor state, or on first open its deck converted from ChatSlide
    if (chatslide.title) slidesStore.setTitle(chatslide.title)
    chatslide.flush = () => syncToChatSlide() // before Scripts: the latest pages first
    run(t('Opening your deck…'), async () => {
      const saved = await chatslide!.load()
      if (saved?.slides?.length) {
        Object.assign(s, blank(), saved.session || {})
        // a restore confirmation left from another session would restore to this session's baseline, not the one it named
        if (s.pending?.repeat?.some((c: any) => c.action === 'restore_original')) s.pending = null
        if (saved.theme) slidesStore.setTheme(saved.theme)
        if (saved.viewportRatio) slidesStore.setViewportRatio(saved.viewportRatio)
        slidesStore.setSlides(saved.slides)
        slidesStore.updateSlideIndex(0)
      }
      else {
        loadingPage = slides.value.map(x => x.id).join() // the empty page shown while it loads; set first, as a
        // failed conversion (the classic backup included) is reported in the chat, and a chat message saves
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
        // saved at once: the saved deck is what makes the project open in this editor next time. Never a failed
        // conversion: a saved blank page would be opened instead of the project's slides from then on
        save()
      }
      // the files the user already gave this project: "use my file" reads them, no re-attach (12 replies asked for
      // the 📎 from users whose file the project already had, 2026-11-10 to 13). Not awaited: the deck opens first.
      if (chatslide!.files && !s.projectFiles) {
        chatslide!.files().then(files => {
          const add = files.filter(f => f.text.trim() && !s.docs.some(d => d.name === f.name))
          s.docs.push(...add.map(f => ({ name: f.name, text: f.text, project: true })))
          if (!files.length) return // none yet (still being read, or not this user's to read): asked again next open
          s.projectFiles = true // kept with the next save
          if (add.length) save()
        }, () => undefined) // unreadable now: asked again next time the deck opens
      }
      // each page as it was when the deck was opened: "back to how it was at first" restores from this
      rememberOriginal()
      // the loaded deck is where Undo stops (not the blank page shown while it loaded)
      await db.snapshots.clear()
      await snapshotStore.initSnapshotDatabase()
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
    run(t('Opening your deck…'), async () => {
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
    run(t('Loading your deck…'), () => showPptx(s.versions[s.versions.length - 1].deckUrl))
  }
})
</script>

<style lang="scss" scoped>
.inline {
  flex: 1;
  min-height: 0;
  padding: 8px;
  background: #fff;

  :deep(.button) {
    min-height: 36px;
    padding: 0 16px;
  }
  .msg {
    font-size: 15px;
  }
  .hint {
    font-size: 14px;
  }
  .attach {
    font-size: 24px;
  }
}
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
  .rate {
    display: flex;
    gap: 8px;
    margin-top: 4px;
    color: #aaa;
    font-size: 13px;

    span {
      cursor: pointer;
      display: flex;

      &:hover, &.on {
        color: $themeColor;
      }
    }
  }
  .why {
    display: block;
    width: 100%;
    margin-top: 4px;
    font-size: 12px;
    border: 1px solid $borderColor;
    border-radius: 4px;
    padding: 2px 6px;
    outline: none;
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
    margin-inline-start: 4px;
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
.paid-strip {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 12px;
  line-height: 1.3;
  color: #6b4eff;
  background: #f3f0ff;
  border-radius: 6px;
  padding: 5px 8px;
  cursor: pointer;

  span { flex: 1; }
  b { white-space: nowrap; }
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
