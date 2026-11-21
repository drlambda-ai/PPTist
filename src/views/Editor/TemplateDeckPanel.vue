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
        <div class="msg" :class="m.role" v-for="(m, i) in s.history" :key="i">{{ m.text }}<div class="rate" v-if="m.turn && chatslide?.feedback && (m.rating !== undefined || i === s.history.length - 1)"><span :class="{ on: m.rating === 1 }" :aria-label="t('Good reply')" @click="rate(m, true)"><i-icon-park-outline:thumbs-up /></span><span :class="{ on: m.rating === 0 }" :aria-label="t('Bad reply')" @click="rate(m, false)"><i-icon-park-outline:thumbs-down /></span></div><input class="why" v-if="m.rating === 0 && m.comment === undefined" v-model="drafts[i]" :placeholder="t('What went wrong? (optional)')" @keydown.enter="e => { if (!e.isComposing && e.keyCode !== 229) comment(m, i) }" @blur="comment(m, i)" /></div>
        <div class="msg assistant loading" v-if="busy">{{ busy }}</div>
        <div class="choices" v-if="!busy && (s.pending?.options?.length || upgradeHint)">
          <span class="choice" v-for="o in s.pending?.options || []" :key="o" @click="input = o; send()">{{ o }}</span>
          <span class="choice" v-if="upgradeHint" @click="hostSteps()?.offerAdvanced?.()">{{ t('Upgrade to the advanced model') }}</span>
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
import { CLIPPATHS } from '@/configs/imageClip'
import { t, direction, lang } from '@/i18n'

defineProps<{ inline?: boolean }>() // on a phone: part of the page, not a floating window
// The chat opens at the right edge, over the Design sidebar: in the middle it covered half the slide on a laptop and
// users could not see what it changed (2026-11-10). Asking for the Design panel moves it to the left edge.
const chatLeft = ref(-8)

import { nextTick, onMounted, reactive, ref, watch } from 'vue'
import { host as chatslide, currentTurn, newTurn, OutOfCredits, takeSpend } from '@/embed'
import { fitCheck } from './turnCheckPrompt'
import { storeToRefs } from 'pinia'
import { nanoid } from 'nanoid'
import tinycolor from 'tinycolor2'
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
import { parse } from 'pptxtojson'
import { normalizePptx } from '@/utils/normalizePptx'
import { fileLook } from '@/utils/fileLook'
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
  projectFiles?: boolean | 2 // the project's files were added once (a removed one stays removed); 2 = the originals
  brief?: { text: string, pages?: number } // what the user asked for when the project was made (ChatSlide), read on each open
  images?: { name: string, url: string }[] // pictures the user attached, placed on slides by add_image source "attached"
  // what each picture the chat added or replaced shows (element id -> its search words): "the police photo" names one.
  // Without it every picture read "large image" and the router changed the page on screen, not page 6 (2026-11-13)
  pics?: Record<string, string>
  // the pictures an image held before the chat replaced it (element id -> oldest first), listed for the router as
  // "earlier pictures here": a user asked six times for "the first nurse image", got an Undo or a new stock photo each
  // time, and uploaded it again herself (2026-11-16)
  // others / background: for a background picture, the page's other full-slide pictures and its background as they
  // were then, so a picture put back puts each of them back too
  was?: Record<string, { src: string, clip: any, shows: string, others?: Record<string, { src: string, clip: any }>, background?: any }[]>
  // the letter sizes a text box had before the chat first resized it (element id -> px, in order): "back as it was"
  // puts them back exactly. Asked after two "bigger" messages, a 0.8 scale left other sizes, on the wrong pages (2026-11-19)
  sizes?: Record<string, number[]>
  // turn: the chat message this reply answers (its x-turn-id); rating 1 up, 0 down; comment: sent with a thumbs-down
  history: { role: 'user' | 'assistant', text: string, turn?: string, rating?: 0 | 1, comment?: string }[]
  versions: { spec: any, deckUrl: string, h: number }[]
  edited?: boolean // template deck changed in the editor since the last build
  synced?: { keys: string[], urls: Record<string, string> } // ChatSlide: pages last sent to the project, and their pictures
  // offer: a button for an optional next step; it reaches the router only when the user taps it (an untouched offer
  // pulled the next, unrelated message onto its page, 2026-11-11). A question (ask_user) holds for any reply.
  pending?: { request: string, question: string, options?: string[], done?: string[], offer?: boolean, localize?: boolean, repeat?: { action: string, args: any }[], pages?: number[], from?: number, like?: number } | null // an ask_user waiting for the user's answer; done: steps already run for it
  shaped?: string[] // pages (ids) whose text look a message changed on that page alone: "the same on the other pages" then carries the page's look (pending.like)
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
;(window as any).__deckPanel = { s, fitOverflow: (...a: Parameters<typeof fitOverflow>) => fitOverflow(...a), resetUndo: async () => {
  await db.snapshots.clear(); await snapshotStore.initSnapshotDatabase()
} }
const input = ref('')
const busy = ref('')
const messagesRef = ref<HTMLElement>()
let blobMode = ''

const SLIDES_KEY = 'ai-deck-slides'
let saveTimer = 0
let savePending = false // a debounced save has not reached ChatSlide yet
let saving: Promise<void> = Promise.resolve() // the save that left its timer (it reads the deck after picture uploads)
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
    saveTimer = window.setTimeout(() => {
      savePending = false
      saving = (async () => {
        await externalizeImages().catch(() => undefined) // pictures saved as files: the deck file stays small and loads fast
        await chatslide!.save({ session: JSON.parse(JSON.stringify(s)), slides: slides.value, theme: theme.value, viewportRatio: viewportRatio.value })
      })().then(() => { saveFailed = false }, () => { // said once (saying saves again): the next edit retries
        // and never under itself: on a line that drops and comes back every save after a good one said it again, 22
        // times in a row in one user's chat (2026-11-18)
        const line = t('Saving failed; your last edits may not be kept. Check your connection.')
        if (!saveFailed && s.history[s.history.length - 1]?.text !== line) say('assistant', line)
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
// The deck on screen saved before a restore replaces it, so it stays in the version history and a restore never loses
// work (a user's 21 pages went to 17 with no way back, 2026-11-12). A save already under way finishes first (it reads
// the deck before this returns, so never the restored one); then the deck is saved again unless that save holds it
// (a save still waiting on its timer, or a failed one, goes now). Throws when it cannot be saved.
// A copied background is the editor's own, never an imported render: the copy drops the import marker isFlat reads
const copiedBackground = (sl: any) => {
  const bg = JSON.parse(JSON.stringify(sl.background || {}))
  if (bg.image) delete bg.image.imported
  return bg
}
const saveNow = async () => {
  if (!chatslide) return
  const waiting = savePending
  clearTimeout(saveTimer)
  savePending = false
  await saving
  if (!waiting && !saveFailed) return
  await externalizeImages().catch(() => undefined)
  try {
    await chatslide.save({ session: JSON.parse(JSON.stringify(s)), slides: JSON.parse(JSON.stringify(slides.value)), theme: theme.value, viewportRatio: viewportRatio.value })
  }
  catch (e) {
    saveFailed = true // the next restore must try this save again, not skip it
    throw e
  }
  saveFailed = false
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
// a new background colour replaces the page's full-slide picture: hidden, kept for undo. Faded to 12% its painted
// patches still showed ("there are still orange patches, remove them", "still not a uniform white", "dark grey on all"
// with the capsules left: 6 turns of 3 users, 2026-11-17..18)
// A page that is only that picture (a pasted screenshot: its text is inside it) keeps it as it is
const fadeBackdrop = (id: string) => {
  const sl = slides.value.find(x => x.id === id)
  if (!sl || onlyFullPicture(sl)) return
  const opacity = '0%'
  for (const e of sl.elements as any[]) {
    if (e.type === 'image' && isBackdrop(e)) slidesStore.updateElement({ id: e.id, props: { filters: { ...e.filters, opacity } }, slideId: id })
  }
}
const hostSteps = () => chatslide || (window as any).__evalHost as typeof chatslide // the eval stands in for ChatSlide's own steps
// a dropped connection on an AI call is retried once: it showed the user "Failed to fetch" for nothing else
if (!(window as any).__deckFetchRetry) {
  (window as any).__deckFetchRetry = true
  const fetch0 = window.fetch.bind(window)
  // the services that answer with one JSON object: an answer that does not parse is asked for once more. A broken
  // answer from the picture search ended a paying user's message with "Unexpected non-whitespace character after JSON
  // at position 4" and she typed the whole request again (2026-11-11). Only the unpaid ones: a paid page edit or AI
  // picture that answered 200 is charged, and asking again would charge the user twice
  const jsonTool = (u: string) => /\/api\/tools\/(research|deck_chat|ask_json|img_search)(\?|$)/.test(u)
  const readable = async (r: Response) => !r.ok || await r.clone().json().then(() => true, () => false)
  window.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String((input as any).url || input)
    try {
      const r = await fetch0(input, init)
      if (!jsonTool(url) || await readable(r)) return r
      await new Promise(ok => setTimeout(ok, 800))
      return fetch0(input, init)
    }
    catch (e) {
      if (!(e instanceof TypeError) || !url.includes('/api/tools/')) throw e
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
    // a bug of ours, or a storage / network failure (prod 2026-11-12: "Vercel Blob: Failed to retrieve the client token"),
    // never shown raw to the user (it is in the turn record with its stack)
    else if (e instanceof SyntaxError || e instanceof TypeError || e instanceof ReferenceError || e instanceof RangeError
      || /^Vercel Blob:|Failed to fetch|NetworkError|Load failed/.test(String((e as Error)?.message))) {
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
// A listed file as the user uploaded it: attached in this session, or one the project already had (asked from ChatSlide)
// The one presentation whose look can be read now: the only .pptx listed (by name: a re-attached file is listed twice),
// a project file or one attached in this session. The router is told the same fact (files[].look).
const lookFile = () => {
  const names = [...new Set(s.docs.filter(d => /\.pptx$/i.test(d.name)).map(d => d.name))]
  return names.length === 1 && (s.docs.some(d => d.name === names[0] && d.project) || docFiles.has(names[0])) ? names[0] : ''
}
const fileOf = async (name: string) => {
  const here = docFiles.get(name) || (docFiles.size === 1 && !s.docs.some(d => d.name === name) ? [...docFiles.values()][0] : undefined)
  if (here) return here
  const doc = s.docs.find(d => d.name === name) || (s.docs.filter(d => d.project).length === 1 ? s.docs.find(d => d.project) : undefined)
  const blob = doc?.project ? await hostSteps()?.file?.(doc.name).catch(() => null) : null
  return blob ? new File([blob], doc!.name) : undefined
}
// A file's text by page or slide, from the markers its reader wrote: "[page N]" on a line of its own (a PDF read here),
// "[slide N] ..." (a .pptx read here), "Slide N:" on a line of its own (a .pptx the project already had, read by
// ChatSlide). "Take the definition from slide 2 of my example" read the whole file, or the deck's own slide 2
// (a paying user, four messages, 2026-11-11): only "[page N]" was known, so a presentation file had no pages at all
const PAGE_MARK = /^(?:\[page (\d+)\]$|\[slide (\d+)\](?= |$)|Slide (\d+):$)/
const filePages = (text: string): { n: number, text: string }[] => {
  const out: { n: number, text: string }[] = []
  for (const line of text.split('\n')) {
    const m = line.match(PAGE_MARK)
    if (m) out.push({ n: Number(m[1] || m[2] || m[3]), text: line })
    else if (out.length) out[out.length - 1].text += `\n${line}`
  }
  // markers a reader wrote count up from 1; anything else is the document's own words ("Slide 3:" inside a text)
  return out.length && out.every((p, k) => p.n === k + 1) ? out : []
}
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

// asWritten: the message is the user's own text for the slides (the router's as_written), laid out section by section
// word for word. Outlined freely, a pasted essay came back regrouped and retold, with claims it never made (Ukrainian
// student, 2026-11-19)
const generateNative = async (text: string, asWritten = false) => {
  const brief = s.docs.length && !asWritten ? `${text}\n\n${materialBrief()}` : text // their pasted text is the whole source
  let outline = ''
  await readStream(await api.AIPPT_Outline({ content: brief, language: requestLanguage(text), model: '', asWritten }), t => {
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
  // their sections whole: left to its 2-4 items a slide, the converter split them under titles of its own and added
  // transition pages with retold lines (replay, 2026-11-21); the editor splits a long slide itself, under the same title
  const keep = asWritten ? '\n\nEvery "##" section is ONE content slide titled with its heading, holding every bullet of it as an item, in order, wording unchanged, however many. No transition slides.' : ''
  await readStream(await api.AIPPT({ content: `User request: ${text}${keep}\n\n${outline}`, language: requestLanguage(text), style: '通用', model: '' }), (chunk) => {
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
// The text look of a page, for its title and for its body: the size of the main text (the longest), whether all of it
// is bold, and its alignment when every paragraph has the same one. It is what "the same on the other pages" carries
// from a page the user shaped over several messages (match_look); text under 9px (a watermark) is no part of it
const lookHtml = (e: any) => String(e.type === 'text' ? e.content : e.text?.content || '')
// text with no size of its own is drawn at the editor's default, 16px
const lookSize = (e: any) => Math.max(...(lookHtml(e).includes('font-size') ? [...lookHtml(e).matchAll(/font-size:\s*([\d.]+)px/g)].map(m => +m[1]) : [16]), 0)
const lookTexts = (sl: Slide, title: boolean) => {
  const marked = (e: any) => e.textType === 'title' || e.text?.type === 'title'
  const first = sl.elements.some(marked) ? null : readingOrder(sl).find((e: any) => elText(e))
  return (sl.elements as any[]).filter(e => (e.type === 'text' || (e.type === 'shape' && e.text)) && elText(e) && lookSize(e) >= 9 && (marked(e) || e === first) === title)
}
const lookMain = (els: any[]) => els.reduce((a, e) => (elText(e).length > elText(a).length ? e : a), els[0])
const textLook = (sl: Slide) => [true, false].map((title) => {
  const els = lookTexts(sl, title)
  if (!els.length) return null
  const heavy = (x: Element | null): boolean => !!x && (/^(strong|b)$/i.test(x.tagName) || /bold|[6-9]00/.test((x as HTMLElement).style?.fontWeight || '') || heavy(x.parentElement))
  const bodies = els.map(e => new DOMParser().parseFromString(lookHtml(e), 'text/html').body)
  const bold = bodies.every((b) => {
    const w = b.ownerDocument.createTreeWalker(b, NodeFilter.SHOW_TEXT)
    for (let n = w.nextNode(); n; n = w.nextNode()) if ((n as Text).data.trim() && !heavy(n.parentElement)) return false
    return true
  })
  const aligns = new Set(bodies.flatMap(b => Array.from(b.querySelectorAll<HTMLElement>('p, li')).map(p => p.style.textAlign)))
  return { size: lookSize(lookMain(els)), bold, align: aligns.size === 1 ? [...aligns][0] : '' }
})
// One numbering everywhere (page list, selection, change_image, remove_images, "already replaced"): content pictures are
// 1.., the full-slide picture 0; on a page with no content picture the lowest full-slide picture is 1, and 0 means it too
// (counted as two pictures, the cover photo was replaced twice in one message).
const pictureNo = (sl: Slide, image: number | undefined) => (image === 0 && !pictures(sl).some(e => !isBackdrop(e)) && firstBackdrop(sl) ? 1 : (image ?? 1))
// A page that is one flat picture: ChatSlide's pages rendered to one image (Blank_layout background_url, which the
// editor's first open imports as the slide's background picture) with nothing on it. Its text and pictures are inside
// that image, so none can be changed on its own, and replacing or repainting the background wipes the page. Only a
// background that came with the import counts: a photo background set in the editor (set_background replaces the whole
// background, marker included) stays an ordinary background, and a photo slide made in the editor is a picture element.
const isFlat = (sl: Slide) => sl.background?.type === 'image' && !!sl.background.image?.imported && !sl.elements.some((e: any) => e.type !== 'shape' || elText(e))
// A page whose one picture is a full-slide picture element and nothing else (a pasted screenshot: its text is inside it)
const onlyFullPicture = (sl: Slide) => {
  const imgs = sl.elements.filter((e: any) => e.type === 'image')
  return imgs.length === 1 && isBackdrop(imgs[0]) && !sl.elements.some((e: any) => e.type === 'chart' || e.type === 'table' || elText(e))
}

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
      // numbered as change_image "earlier" takes them; the picture it shows now is not one to put back
      const earlier = (s.was?.[e.id] || []).map((p, k) => (p.src === e.src ? '' : `${k + 1}${k ? '' : ' (the first)'}${p.shows ? ` "${p.shows.replace(/["\]]/g, '').slice(0, 40)}"` : ''}`)).filter(Boolean)
      const had = earlier.length ? `, earlier pictures here (change_image "earlier" puts one back): ${earlier.join('; ')}` : ''
      const attached = (s.images || []).find(m => m.url && m.url === e.src)
      const shows = !attached && s.pics?.[e.id] ? `, shows "${s.pics[e.id].replace(/["\]]/g, '').slice(0, 60)}"` : ''
      // the background picture says what it shows too: "keep the nurse, change the other image" names it by elimination
      if (!n) return e === backdrop ? `[image 0: background picture covering the whole slide${shows}${had}]` : '[a see-through design layer covering the whole slide, not a picture]'
      const img = n - 1
      const kind = `${area < 0.005 ? 'icon' : area < 0.04 ? 'small image' : 'large image'}${e.width * e.height === biggest ? ', the largest image here' : ''}`
      const where = `${e.top + e.height / 2 < H / 3 ? 'top' : e.top + e.height / 2 > (2 * H) / 3 ? 'bottom' : 'middle'}-${e.left + e.width / 2 < W / 3 ? 'left' : e.left + e.width / 2 > (2 * W) / 3 ? 'right' : 'centre'}`
      // a logo, or the user's own attached picture, keeps its size and place so "move my logo top-right, smaller" can act
      if (logos.has(imgSpot(e))) return `[image ${img + 1}: logo, ${Math.round(area * 100)}% of slide, ${where}]`
      const own = attached ? `, the user's attached picture "${attached.name.replace(/["\]]/g, '')}"` : ''
      return `[image ${img + 1}: ${kind}, ${Math.round(area * 100)}% of slide, ${where}${own}${shows}${had}]`
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
const readThisTurn = new Set<string>() // names of the user's files a read_file step read in the current message
const roomMadeThisTurn = new Set<string>() // pages whose text was moved aside for a picture in the current message
let sitePaywalledPage = '' // the page a free user was just told the brand's own pictures are a paid feature for
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
  return [text, ...(lastChangedBackground() && prev ? [prev] : [])].some(x => norm(x).includes(w))
}
const lastChangedBackground = () => lastTurnCalls.some(k => k.startsWith('set_background ') || (k.startsWith('change_image ') && /"image":0[,}]/.test(k)))
let currentAtSend = '' // id of the page the user was on when they sent the message
const createdThisMessage = new Set<string>() // pages added since the message was sent
let lastTurnCalls: string[] = [] // the previous message's steps, as stepKey: what an Undo just reverted
// a step's TARGET (action + where), not its wording: a redo of the undone change rarely repeats the query word for word
const stepKey = (action: string, args: any) => `${action} ${JSON.stringify(Object.fromEntries(Object.entries(args || {})
  .filter(([k]) => !['answer', 'source'].includes(k)).sort()))}` // what to change stays in: "no, red instead" is a new step
// text boxes that run into a picture or into another plain text box (a paid redesign left a heading on top of the
// card text, 2026-11-11), as "a|b" keys. Text on a card shape is a design, not a collision.
// A shape with words and no fill of its own is a plain text box too: every text box of an imported ChatSlide or
// PowerPoint deck is one, so a heading the paid redesign added on top of the body text went unseen, and "the heading
// overlaps the body text" was answered "Updated page 26" with the page looking the same (2026-11-18)
const bareText = (e: any) => !!elText(e) && (e.type === 'text' || (e.type === 'shape' && !e.gradient && !e.pattern && (!e.fill || e.fill === 'transparent' || /^#[0-9a-f]{6}00$/i.test(e.fill))))
const textOnText = (els: any[]) => {
  const texts = els.filter(bareText)
  // the letters, not the box: a text box keeps 10px of padding above and below its lines (a one-line heading right
  // above its text touches it by that padding and collides with nothing)
  const pad = (e: any) => (e.type === 'text' ? 10 : 4)
  return new Set<string>(texts.flatMap((a: any, i: number) => texts.slice(i + 1).filter((b: any) => {
    const w = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
    const h = Math.min(a.top + a.height - pad(a), b.top + b.height - pad(b)) - Math.max(a.top + pad(a), b.top + pad(b))
    return w > 0 && h > 0 && w * h > 0.15 * Math.min(a.width * a.height, b.width * b.height)
  }).map((b: any) => [a.id, b.id].sort().join('|'))))
}
const textOnPictures = (els: any[]) => {
  const pics = els.filter((e: any) => (e.type === 'image' && !isBackdrop(e)))
  return new Set([...textOnText(els), ...els.filter((e: any) => e.type !== 'image' && elText(e)).flatMap((a: any) => pics.filter((b: any) => {
    const w = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
    const h = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top)
    return w > 0 && h > 0 && w * h > 0.05 * Math.min(a.width * a.height, b.width * b.height)
  }).map((b: any) => `${a.id}|${b.id}`))])
}
// Text whose lines run into the text under it (a heading that wraps onto the body): its letters go 2px smaller, up
// to three times, until the page as shown has no such collision. Only real text boxes, whose height is their lines'.
// Put back when that is not enough: the page model then moves things. True when it fixed the page
const shrinkOntoText = async (id: string) => {
  const at = () => slides.value.findIndex(x => x.id === id)
  const show = async () => {
    slidesStore.updateSlideIndex(at())
    await nextTick()
    await new Promise(r => setTimeout(r, 350)) // the editor measures and stores the new heights
    return slides.value[at()].elements as any[]
  }
  const uppers = (els: any[]) => [...textOnText(els)].map(k => k.split('|').map(x => els.find(e => e.id === x)).sort((a, b) => a.top - b.top)[0]).filter(e => e.type === 'text')
  let els = await show()
  if (!uppers(els).length) return false
  const start = JSON.stringify(els)
  for (let step = 0; step < 3; step++) {
    let any = false
    for (const e of uppers(els)) {
      const content = String(e.content).replace(/font-size:\s*([\d.]+)px/g, (_m: string, v: string) => `font-size: ${Math.max(10, Math.round(+v) - 2)}px`)
      if (content === e.content) continue
      slidesStore.updateElement({ id: e.id, props: { content }, slideId: id })
      any = true
    }
    if (!any) break
    els = await show()
    if (!uppers(els).length) return true
  }
  slidesStore.updateSlide({ elements: JSON.parse(start) }, id)
  return false
}
// Text a page edit left unreadable: every letter in (nearly) the colour of the plain colour right behind it (its own
// box fill, else the nearest opaque shape under at least 30% of it, else nothing: a picture or the page background is
// not judged). "Its background the same colour as the title" put a white full-slide panel over the cover photo, under
// a white title (replay 2026-11-20; the user undid the same request in production, 2026-11-16). Ids of such text
const unreadableText = (els: any[]) => els.filter((e: any, k: number) => {
  if (!elText(e) || !(e.type === 'text' || (e.type === 'shape' && e.text))) return false
  const doc = new DOMParser().parseFromString(lookHtml(e), 'text/html'), w = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT), ink = new Set<string>()
  for (let node = w.nextNode(); node; node = w.nextNode()) {
    if (!(node as Text).data.trim()) continue
    let x = node.parentElement
    while (x && !x.style.color) x = x.parentElement
    ink.add(tinycolor(x?.style.color || (e.type === 'text' ? e.defaultColor : e.text?.defaultColor) || '#333').toHexString())
  }
  const opaque = (c: any) => !!c && tinycolor(c).isValid() && tinycolor(c).getAlpha() >= 0.9
  const under = (x: any) => Math.max(0, Math.min(e.left + e.width, x.left + x.width) - Math.max(e.left, x.left)) * Math.max(0, Math.min(e.top + e.height, x.top + x.height) - Math.max(e.top, x.top)) >= 0.3 * e.width * e.height
  // what lies right behind the letters: the box's own fill, else the nearest thing under it
  const behind = opaque(e.fill) && !e.gradient && !e.pattern ? e.fill : (() => {
    const x = els.slice(0, k).reverse().find((y: any) => under(y) && (y.type === 'image' ? y.filters?.opacity !== '0%' : y.type === 'shape' && (!!y.gradient || !!y.pattern || tinycolor(y.fill || 'transparent').getAlpha() > 0.1)))
    return x && x.type === 'shape' && !x.gradient && !x.pattern && opaque(x.fill) && (x.opacity ?? 1) >= 0.9 ? x.fill : ''
  })()
  return !!behind && ink.size > 0 && [...ink].every(c => tinycolor.readability(c, behind) < 1.25)
}).map((e: any) => e.id as string)
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
// words that need more room than their box has, measured in the page as drawn: a shape's words run out over its edges.
// The page model reads only the boxes, so "make the text visible everywhere" changed a colour or a few pixels on each
// page and said every page was updated, while the words still ran out of the red banners (a replayed real turn,
// 2026-11-17 / 2026-11-20)
const spillNote = (sl: Slide) => withRenderer(async () => {
  try {
    await document.fonts?.ready
    toRender.value = [JSON.parse(JSON.stringify(sl))]
    await nextTick()
    await new Promise(r => setTimeout(r, 400))
    const out: string[] = []
    for (const e of sl.elements as any[]) {
      const html = e.type === 'text' ? e.content : e.text?.content
      if ((e.type !== 'shape' && e.type !== 'text') || e.vertical || !plain(html)) continue
      const node = renderRef.value?.querySelector(`.base-element-${CSS.escape(e.id)} .ProseMirror-static`) as HTMLElement | null
      const box = node?.parentElement
      if (!node || !box) continue
      const cs = getComputedStyle(box)
      const need = Math.round(node.offsetHeight + parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom))
      if (need > e.height + 6) out.push(`${e.id}: its words need about ${need}px of height and its box is ${Math.round(e.height)}px, so they run out of it`)
    }
    return out.length ? `\n(Measured on this page as drawn: ${out.slice(0, 8).join('; ')}. When the request is about layout, readability, or text that is cut off, hidden or not clearly visible, fix these: make the box tall enough (moving what lies under it), or fewer or smaller words; body text not under 14px.)` : ''
  }
  finally {
    toRender.value = []
  }
}).catch(() => '')

// A line break the page model writes as <br> becomes its own paragraph: the editor's text boxes have no line-break
// node, so "<p>Name<br>University<br>Department</p>" is shown as one line. "Write my name, the university and the
// department one under another" was answered "updated" three times with the three still on one line (a paying user,
// 2026-11-15; the same in 1 of 3 replays, 2026-11-20). Each part keeps the paragraph's and the runs' own styles
const brToParagraphs = (html: string) => {
  if (!/<br\b/i.test(html)) return html
  const doc = new DOMParser().parseFromString(html, 'text/html')
  for (let br = doc.body.querySelector('br'); br; br = doc.body.querySelector('br')) {
    const block = br.closest('p, li') as HTMLElement | null
    if (!block) { br.replaceWith(doc.createTextNode(' ')); continue }
    const range = doc.createRange()
    range.setStartBefore(block.firstChild || br)
    range.setEndBefore(br)
    const first = block.cloneNode(false) as HTMLElement
    first.append(range.extractContents()) // the part before the break, inside copies of the runs it sat in
    br.remove()
    if (first.textContent?.trim()) block.before(first)
    // a break at the very end leaves an empty rest: the paragraph that held it goes
    if (!block.textContent?.trim() && !block.querySelector('br')) block.remove()
  }
  return doc.body.innerHTML
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
// A page that is one flat picture (an imported or pasted page) shows its title and text, but they are pixels of that
// image: "rewrite the text", "make the title bigger" moved or stretched the picture, or stacked a second title over the
// painted one, and the reply said the page was updated
const FLAT_NOTE = '\n(This page is ONE flat picture: its title, text, pictures and colours are all painted inside that single image, so none of them can be edited, restyled, resized or moved. Unless the request is to add something new on top (a text box, a shape, a label) or to change that picture as a whole, change nothing, and say in "reply" that the page is one flat picture whose text cannot be edited here.)'
// The words painted in such a picture, each line with the box its letters fill on the canvas. The page model sees no
// picture, so "underline the last sentence", "a check icon before Application" and "an arrow" were answered "this page
// has no elements" three times, and a drawn arrow landed on a heading (a paying user, 2026-11-14). A vision model reads
// the picture once (a fraction of a cent; boxes within a few px on the user's page), and marks go by those boxes.
type PaintedLine = { text: string, left: number, top: number, width: number, height: number, ground?: string, cover?: number[] }
// The colour around a painted line, read from the picture's pixels in a ring just outside its box (x, y, w, h in percent
// of the picture), when that ring is one even colour or one smooth left-to-right gradient. "移除講者名稱" on an imported cover was refused because the page
// model could not see the colour behind the name, and a replay covered it with a guessed white box on a blue cover
// (a paying user, 2026-11-21). Only an even ground can hide words: on a photo or a pattern there is none.
// With it, the box (in percent of the picture) where that ground was read: the cover goes exactly there, since a box
// the model grew by itself ran past the sampled ground and showed (replay 2026-11-22).
const groundOf = (src: string, lines: { x: number, y: number, w: number, h: number }[]) => new Promise<({ ground: string, cover: number[] } | undefined)[]>((ok) => {
  const img = new Image()
  img.crossOrigin = 'anonymous'
  img.onload = () => {
    try {
      const k = Math.min(1, 1280 / img.naturalWidth)
      const c = document.createElement('canvas')
      c.width = Math.round(img.naturalWidth * k)
      c.height = Math.round(img.naturalHeight * k)
      const g = c.getContext('2d', { willReadFrequently: true })!
      g.drawImage(img, 0, 0, c.width, c.height)
      const px = g.getImageData(0, 0, c.width, c.height).data
      ok(lines.map(({ x, y, w, h }) => {
        const [l, t, W, H] = [x, y, w, h].map((v, i) => (v / 100) * (i % 2 ? c.height : c.width))
        const pad = Math.max(3, H * 0.5) // the edge of the box that covers the line: an even ring there means nothing painted shows past it
        const [x0, y0, x1, y1] = [l - pad, t - pad, l + W + pad, t + H + pad]
        const at: number[][] = [] // [x, r, g, b]
        for (let i = 0; i <= 40; i++) {
          const fx = x0 + ((x1 - x0) * i) / 40, fy = y0 + ((y1 - y0) * i) / 40
          for (const [px0, py0] of [[fx, y0], [fx, y1], [x0, fy], [x1, fy]]) {
            const X = Math.round(px0), Y = Math.round(py0)
            if (X >= 0 && Y >= 0 && X < c.width && Y < c.height) at.push([px0, ...px.slice((Y * c.width + X) * 4, (Y * c.width + X) * 4 + 4)])
          }
        }
        if (at.length < 60) return undefined // the line runs off the picture
        if (at.some(p => p[4] < 250)) return undefined // see-through pixels read as black: what shows there is not in the picture
        // left, middle and right thirds of the ring, each colour standing at its third's centre (17/50/83%): a flat patch
        // on a left-to-right gradient showed as a box, so an even ground may run from one colour to another across the
        // line (ponytail: left to right only; a top-to-bottom gradient reads as uneven and the line is not removed)
        const thirds = [0, 1, 2].map(b => at.filter(p => Math.min(2, Math.floor(((p[0] - x0) / (x1 - x0)) * 3)) === b))
        if (thirds.some(t => t.length < 8)) return undefined
        const stat = thirds.map((t) => {
          const mean = [1, 2, 3].map(ch => t.reduce((s, p) => s + p[ch], 0) / t.length)
          return { mean, sd: Math.max(...[1, 2, 3].map((ch, i) => Math.sqrt(t.reduce((s, p) => s + (p[ch] - mean[i]) ** 2, 0) / t.length))) }
        })
        if (stat.some(t => t.sd > 10)) return undefined
        // flat bands of different colours are not a gradient: a smooth ground moves little from one third to the next
        if ([0, 1, 2].some(ch => Math.abs(stat[0].mean[ch] - stat[1].mean[ch]) > 40 || Math.abs(stat[1].mean[ch] - stat[2].mean[ch]) > 40)) return undefined
        const hex = stat.map(t => tinycolor({ r: t.mean[0], g: t.mean[1], b: t.mean[2] }).toHexString())
        const spread = Math.max(...[0, 1, 2].map(ch => Math.max(...stat.map(t => t.mean[ch])) - Math.min(...stat.map(t => t.mean[ch]))))
        const cover = [x0 / c.width, y0 / c.height, (x1 - x0) / c.width, (y1 - y0) / c.height].map(v => v * 100)
        return { cover, ground: spread <= 6 ? hex[1] : `"fill": "${hex[1]}", "gradient": {"type": "linear", "rotate": 0, "colors": [{"pos": 17, "color": "${hex[0]}"}, {"pos": 50, "color": "${hex[1]}"}, {"pos": 83, "color": "${hex[2]}"}]}` }
      }))
    }
    catch {
      ok([]) // a picture whose host does not allow reading its pixels
    }
  }
  img.onerror = () => ok([])
  img.src = src
})
const paintedRead = new Map<string, Promise<PaintedLine[]>>() // by picture, for this session
const paintedKnown = new Map<string, PaintedLine[]>() // the lines already read, by slide picture: where a new picture must not go
const paintedPicture = (sl: Slide): { src: string, box?: { left: number, top: number, width: number, height: number } } => {
  if (sl.background?.type === 'image' && sl.background.image?.imported) return { src: sl.background.image.src || '' }
  const pic = onlyFullPicture(sl) ? sl.elements.find((e: any) => e.type === 'image') as any : null
  return pic ? { src: pic.src || '', box: pic } : { src: '' }
}
const paintedLines = (sl: Slide): Promise<PaintedLine[]> => {
  const { src, box } = paintedPicture(sl)
  // a link, or a picture small enough to send (a request body is capped at 4.5MB)
  if (!/^https:\/\//.test(src) && !(/^data:image\/(png|jpeg|webp);base64,/.test(src) && src.length < 3_000_000)) return Promise.resolve([])
  const { width: W, height: H } = slideSize()
  // keyed by the whole picture (two JPEG data URLs of one length share their first characters, and then one slide
  // was given the other's lines)
  let hash = 0x811c9dc5
  for (let i = 0; i < src.length; i++) hash = Math.imul(hash ^ src.charCodeAt(i), 0x01000193)
  const key = `${(hash >>> 0).toString(36)}|${src.length}|${W}x${H}|${box ? [box.left, box.top, box.width, box.height].join(',') : ''}`
  const known = (read: Promise<PaintedLine[]>) => read.then((lines) => {
    if (lines.length) paintedKnown.set(src, lines)
    else if (paintedRead.get(key) === read) paintedRead.delete(key) // unread (a failed call, no text): asked again next time
  })
  const had = paintedRead.get(key)
  if (had) { known(had); return had }
  const read = (async () => {
    // where the picture lies on the canvas: a background covers the slide (cropped evenly when its shape differs)
    const size = await new Promise<{ w: number, h: number } | null>((ok) => {
      const im = new Image()
      im.onload = () => ok({ w: im.naturalWidth, h: im.naturalHeight })
      im.onerror = () => ok(null)
      im.src = src
    })
    if (!size?.w || !size.h) return []
    const k = box ? 0 : Math.max(W / size.w, H / size.h)
    const at = box ? { left: box.left, top: box.top, width: box.width, height: box.height }
      : { left: (W - size.w * k) / 2, top: (H - size.h * k) / 2, width: size.w * k, height: size.h * k }
    const r = await fetch('/api/tools/ask_json', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ image: src,
      instruction: 'This picture is one presentation slide. List every line of text painted in it, top to bottom, each with the box its letters fill, in percent of the picture (x = left edge, y = top edge, w = width, h = height, each 0-100). JSON {"lines": [{"text": "...", "x": 0, "y": 0, "w": 0, "h": 0}]}' }) })
    const got = r.ok ? (await r.json().catch(() => ({})))?.lines : null
    const lines = (Array.isArray(got) ? got : []).map((l: any) => ({ x: toNum(l?.x), y: toNum(l?.y), w: toNum(l?.w), h: toNum(l?.h), text: String(l?.text || '').replace(/\s+/g, ' ').trim().slice(0, 160) }))
      .filter(l => l.text && [l.x, l.y, l.w, l.h].every(Number.isFinite) && l.w > 0 && l.h > 0).slice(0, 60)
    const grounds = lines.length ? await groundOf(src, lines) : []
    return lines.map((l, i) => ({ text: l.text,
      left: Math.round(at.left + (l.x / 100) * at.width), top: Math.round(at.top + (l.y / 100) * at.height),
      width: Math.round((l.w / 100) * at.width), height: Math.round((l.h / 100) * at.height),
      ...(grounds[i] ? { ground: grounds[i]!.ground, cover: (([x, y, w, h]) => [at.left + x * at.width, at.top + y * at.height, w * at.width, h * at.height].map(Math.round))(grounds[i]!.cover.map(v => v / 100)) } : {}) }))
      .filter(l => l.width > 0 && l.height > 0)
  })().catch(() => [] as PaintedLine[])
  paintedRead.set(key, read)
  known(read)
  return read
}
const paintedNote = async (sl: Slide) => {
  const lines = await paintedLines(sl)
  if (!lines.length) return isFlat(sl) ? FLAT_NOTE : ''
  return '\n(The page itself is ONE flat picture: its title, text, pictures and colours are all painted inside that single image, so none of them can be edited, restyled, resized, moved or deleted. What CAN be done is to put NEW elements on top of it (any element listed above was put there this way and is edited as usual).'
    + ` The lines of text painted in the picture, each with the box its letters fill on this canvas, in px, and the even colour around it when it sits on one ("ground"):\n${lines.map(l => `- "${l.text.replace(/"/g, '\'')}": left ${l.left}, top ${l.top}, width ${l.width}, height ${l.height}${l.ground ? `, ground ${l.ground}, cover box left ${l.cover![0]}, top ${l.cover![1]}, width ${l.cover![2]}, height ${l.cover![3]}` : ''}`).join('\n')}\n`
    + 'Place new elements by these boxes. An underline is a new thin shape (a plain rectangle 2-3px high, in the colour of that text or a dark one) right under the line\'s box, as wide as the words to underline. A check mark, tick, bullet or icon "before" a line is a small new shape just left of that line\'s box, vertically centred on it and about as tall as its letters. A highlight is a new see-through box over the words. An arrow, a label or a new text box goes where no painted line is, unless the user names the place; never cover painted words with a solid shape, except to remove them as below.'
    + ' To remove or hide a painted line ("remove the speaker name", "delete the date") that lists a ground: a new plain rectangle shape filled with exactly that ground (a ground written as "fill" and "gradient" is copied into the shape as written), no outline, its box exactly the line\'s cover box (the area where that ground was read), so the words vanish into the page. A painted line with no ground (it sits on a photo or a pattern) cannot be removed: change nothing and say so in "reply".'
    + ' When the request needs the painted text itself changed (its wording, colour, size or font), change nothing and say in "reply" that the page is one flat picture whose text cannot be edited here, and that marks or new text can be put on top.)'
}
const editSlideWith = async (index: number, instruction: string, advanced: boolean) => {
  const slide = slides.value[index]
  // reading order so "the first point" means the top-left one; the merge below keeps the slide's own stacking order
  // a shape's words (every text box of an imported ChatSlide/PowerPoint deck) are shown like a text box's: the model
  // edits "content" the same way on both, and the merge puts it back inside the shape's text
  // its line and paragraph spacing too: hidden, "the text is too cramped" moved the boxes and left the lines touching
  // (a user asked twice, 2026-11-20)
  const flat = (e: any) => (e.type === 'shape' && e.text ? (({ text, ...r }) => ({ ...r, content: text.content, defaultColor: text.defaultColor, defaultFontName: text.defaultFontName, ...spacing(text) }))(e) : e)
  // the model copies what it sees, so a new box often comes back as a "shape" with "content": put the words where
  // the editor draws them (a shape's text, or a plain text box when it has no outline)
  const unflat = (e0: any) => {
    const e = { ...e0, ...Object.fromEntries(GEOMETRY.filter(k => k in e0).map(k => [k, Number.isFinite(toNum(e0[k])) ? toNum(e0[k]) : 0])) }
    delete e.layer // a stacking instruction, applied by the merge
    // a shape is drawn scaled from its viewBox [w, h]: the model often writes an SVG viewBox string ("0 0 1000 563"),
    // which drew a 2px underline as a black block over half the page (replay 2026-11-20). The string's size is kept,
    // anything else unreadable becomes the shape's own size
    if (e.type === 'shape') {
      const vb = typeof e.viewBox === 'string' ? e.viewBox.trim().split(/[\s,]+/).map(Number).slice(-2) : e.viewBox
      e.viewBox = Array.isArray(vb) && vb.length === 2 && vb.every((v: any) => Number.isFinite(Number(v)) && Number(v) > 0) ? vb.map(Number) : [e.width || 1, e.height || 1]
      // a shape name for a path ("rect") draws nothing: a box covering a painted name stayed invisible (replay
      // 2026-11-22). It becomes the shape it names, an oval or else a rectangle
      if (e.path !== undefined && !/^\s*[Mm]/.test(String(e.path))) {
        const [w, h] = e.viewBox
        e.path = /circle|ellipse|oval/i.test(String(e.path)) ? `M 0 ${h / 2} A ${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2} A ${w / 2} ${h / 2} 0 1 0 0 ${h / 2} Z` : `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`
      }
      // a drawing that paints nothing: "fill": "none" and no outline drew "an outline of a black cat sitting" as nothing
      // at all, and the reply said the page was updated (a paying user's page, replayed 2026-11-20). It gets a dark
      // outline, so what was drawn shows
      const paints = (c: unknown) => typeof c === 'string' && c !== 'none' && tinycolor(c).isValid() && tinycolor(c).getAlpha() > 0.05
      const outlined = toNum(e.outline?.width) > 0 && (e.outline?.color === undefined || paints(e.outline.color))
      if (!plain(String(e.content ?? '')) && !e.gradient && !e.pattern && !paints(e.fill) && !outlined) e.outline = { style: 'solid', width: 3, color: '#222222' }
    }
    if (e.type !== 'shape' || e.content === undefined) return e
    const { content, defaultColor, defaultFontName, lineHeight, paragraphSpace, ...rest } = e
    if (!rest.path) return { ...rest, type: 'text', content, defaultColor: defaultColor || '#333', defaultFontName: defaultFontName || '', rotate: rest.rotate || 0, ...spacing({ lineHeight, paragraphSpace }) }
    return { ...rest, text: { content, defaultColor: defaultColor || '#333', defaultFontName: defaultFontName || '', align: 'middle', ...spacing({ lineHeight, paragraphSpace }) } }
  }
  // a picture says what it shows (the words it was searched or drawn for): "just one arrow" about a stock photo of
  // arrows was answered "there is no arrow on this page" (2026-11-14, replayed 2026-11-20)
  // and whether it is the page's background: "round the corners of the images" rounded the full-page background
  // instead of the photo on it (2026-11-20 replay)
  const elements = readingOrder(slide).map((e: any) => (e.type === 'image' ? { ...e, src: `image:${e.id}`, ...(s.pics?.[e.id] ? { shows: s.pics[e.id].slice(0, 120) } : {}), ...(isBackdrop(e) ? { background: true } : {}) }
    : e.type === 'shape' && e.path?.length > 80 ? { ...flat(e), path: `path:${e.id}` } : flat(e))) // long SVG paths: token, restored below
  const painted = await paintedNote(slide)
  const spill = await spillNote(slide)
  const r = await fetch('/api/tools/chat_edit', {
    method: 'POST', headers: { 'Content-Type': 'application/json', ...(advanced ? { 'X-Quiet-402': '1' } : {}) },
    body: JSON.stringify({ patch: true, advanced, elements, slideSize: slideSize(), instruction: `${instruction}${overlapNote(slide)}${spill}${painted}\n(If you change nothing, the "reply" sentence says why, e.g. what you found on the page; never claim a change you did not make. Add "too_hard": true only when the request applies but you could not do it well; nothing needed is not too hard. Write the "reply" sentence in the language of this message: "${userText.slice(0, 120)}")` }),
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
      const { src: _src, path, layer: _layer, shows: _shows, background: _bg, ...rest } = c // image sources and tokenised paths are never taken from the model
      void _src; void _layer; void _shows; void _bg // layer is a stacking instruction, applied below, not an element field; shows and background are notes for the model
      // a position or size that is not a number is dropped: top ":" from the model threw a heading into the corner
      // of the slide (2026-11-11, a user's "improve the design")
      for (const k of GEOMETRY) {
        if (!(k in rest)) continue
        const v = toNum(rest[k])
        if (Number.isFinite(v)) rest[k] = v
        else delete rest[k]
      }
      if (e.type === 'image') {
        // the editor draws `${radius}px`: a radius of "16px" from the model drew square corners and the reply said the
        // page was updated (2026-11-20, "page 5 does not have a rounder corner image")
        if ('radius' in rest) {
          const v = parseFloat(String(rest.radius))
          if (Number.isFinite(v) && v >= 0) rest.radius = v
          else delete rest.radius
        }
        // a picture shape ("make it round"): the editor's own clip shapes, the visible part of the picture kept
        if ('clip' in rest) {
          const shape = typeof rest.clip === 'string' ? rest.clip : rest.clip?.shape
          if (typeof shape === 'string' && shape in CLIPPATHS) rest.clip = { range: e.clip?.range || [[0, 0], [100, 100]], shape }
          else delete rest.clip
        }
      }
      if (e.type === 'shape' && e.text) {
        const { content, defaultColor, defaultFontName, text, lineHeight, paragraphSpace, ...shape } = rest
        const t = { ...e.text, ...(text || {}), ...(content !== undefined ? { content } : {}), ...(defaultColor ? { defaultColor } : {}), ...(defaultFontName ? { defaultFontName } : {}), ...spacing({ lineHeight, paragraphSpace }) }
        return { ...e, ...shape, text: t, ...(path && !String(path).startsWith('path:') ? { path } : {}) }
      }
      // a text box is as tall as its lines, measured when shown: a height from the model stays in the deck only while
      // the lines do not change, and then says a wrapped heading no longer reaches the text under it
      if (e.type === 'text' && !e.fixedHeight && !e.vertical) delete rest.height
      if (e.type === 'text') { delete rest.lineHeight; delete rest.paragraphSpace; Object.assign(rest, spacing(c)) }
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
  // text the model wrote: its line breaks become paragraphs (only in what it changed or added)
  const wrote = new Set<string>([...Object.keys(changes), ...(ret.patch.added || []).map((e: any) => e?.id)])
  merged.forEach((e: any, k: number) => {
    if (!wrote.has(e.id)) return
    if (e.type === 'text' && typeof e.content === 'string') merged[k] = { ...e, content: brToParagraphs(e.content) }
    else if (e.type === 'shape' && typeof e.text?.content === 'string') merged[k] = { ...e, text: { ...e.text, content: brToParagraphs(e.text.content) } }
  })
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

// Does the message say what a new deck is ABOUT? The router filled the topic with an audience ("une présentation pour un
// congrès scientifique avec des experts") and replaced the user's deck on medicine with a generic one (2026-11-19), the
// very example its tool description gives. Unreadable counts as no: the deck is kept and edited instead
const namesSubject = async (message: string, topic: string) => {
  if (s.docs.length || (s.images || []).length) return true // built from their file
  const r = await fetch('/api/tools/ask_json', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ instruction:
    `A user of a slide editor wrote: ${JSON.stringify(message.slice(0, 2000))}\nA model read its topic as: ${JSON.stringify(topic.slice(0, 300))}\n`
    + 'Does the message name what a NEW presentation should be about (a subject: "climate change", "my company", "photosynthesis")? '
    + 'An audience, purpose, occasion, style or length alone ("for a scientific congress", "for my class", "more professional", "13 slides") is not a subject. '
    + 'Return JSON {"subject": true|false}.' }) }).catch(() => null)
  return r?.ok ? (await r.json().catch(() => ({})))?.subject === true : false
}

// A text box grows to fit its content when its page is shown. After an edit, show the page and shrink the font sizes
// of any changed box that grew past its old height (translations run longer), so it never spills onto other elements.
let fitQueue: Promise<void> = Promise.resolve() // one page shown at a time, even when pages are edited in parallel
// The numbered slides of a text the user wrote ("Slide 1 ... Slide 2 ..."), each copied word for word by a small model;
// [] when its text does not label slides itself or cannot be read
const numberedParts = async (src: string, request = ''): Promise<{ page: number, text: string }[]> => {
  const r = await fetch('/api/tools/ask_json', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ instruction:
    `A user wrote this for a slide editor:\n<message>\n${fence(src.slice(0, 20000))}\n</message>\n`
    + (request === src ? 'Only when this message asks for its text to go on the slides, or is just the text for the slides with nothing else asked, go on, and return only the slides it asks to place (all of them unless it names some); when it asks to review, check, explain or keep the slides as they are, return {"slides": []}. '
      : request ? `Their latest message is: ${JSON.stringify(request.slice(0, 600))}. Only when that message asks to put this text of theirs on the slides (insert it, use it, "the text I wrote"), go on, and return only the slides it asks for (all of them unless it names some); otherwise return {"slides": []}. ` : '')
    + 'If it gives the text of several slides, each marked with its number ("Slide 1", "Слайд 2", "Diapositiva 3", "第4页"), copy each slide\'s text exactly as written, every word, its title first, without the "Slide N" label itself. '
    + 'labels_slides is true only when their text itself says which part goes on which slide ("Slide 1", "Слайд 2", "Diapositiva 3", "第4页"), or their latest message says its numbers are the slide numbers ("put each numbered part on that slide"); false for an essay, report or article whose own sections are numbered or headed ("1. Introduction", "Вступ", "2. ..."): those numbers are its sections, not slides. '
    + 'Return JSON {"labels_slides": true|false, "slides": [{"number": <the number they wrote>, "text": "..."}]}, or {"labels_slides": false, "slides": []} when it is not text for numbered slides.' }) }).catch(() => null)
  // a pasted essay with "1. ... 2. ..." sections had its sections placed as slides 1-4, reworded, over a 14-page deck on
  // another topic (Ukrainian student, 2026-11-19): only slides their text labels are placed; the rest is the router's
  const json = r?.ok ? await r.json().catch(() => ({})) : {}
  const got = json?.labels_slides === true ? (json.slides || []) : []
  return (Array.isArray(got) ? got : []).map((x: any) => ({ page: Number(x?.number), text: String(x?.text || '').trim() }))
    .filter((x: { page: number, text: string }) => Number.isInteger(x.page) && x.page >= 1 && x.text)
}
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
const drafts = reactive<Record<number, string>>({}) // each reply's own comment box (by its place in the chat)
const rate = (m: Session['history'][number], good: boolean) => {
  if (!m.turn || m.rating !== undefined) return
  m.rating = good ? 1 : 0
  save()
  chatslide?.feedback?.(m.turn, good, '', m.text).catch(() => undefined)
}
const comment = (m: Session['history'][number], i: number) => {
  if (m.comment !== undefined) return
  m.comment = (drafts[i] || '').trim().slice(0, 300)
  delete drafts[i]
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
  // every copy of the last hour (the deck saved just before a restore must stay reachable, and the restored deck's own
  // save lands in the same minute), then one per hour, newest first; at most 3 recent ones are offered, so the copies
  // of the last minutes never hide yesterday's version
  const hours = new Set<string>()
  const recent = (v: { at: string }) => Date.now() - +new Date(v.at) < 3600e3
  const perHour = (await chatslide.versions()).filter(v => {
    if (recent(v)) return true
    const h = String(v.at).slice(0, 13)
    if (hours.has(h)) return false
    hours.add(h)
    return true
  })
  const all = (day ? perHour.filter(v => localDay(v.at) === day) : perHour).slice(0, 40)
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
  let recentShown = 0
  for (let k = 0; k < all.length && Object.keys(choices).length < 6; k += 8) {
    const batch = await Promise.all(all.slice(k, k + 8).map(v => fetch(v.url, { cache: 'no-store' }).then(r => r.json()).then(d => ({ v, d })).catch(() => null)))
    for (const x of batch) {
      if (!x?.d?.slides?.length || seen.has(sig(x.d.slides)) || Object.keys(choices).length >= 6) continue
      if (recent(x.v) && recentShown >= 3) continue
      if (recent(x.v)) recentShown++
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

const tableFor = async (instruction: string) => {
  const res = await fetch('/api/tools/ask_json', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ strong: true, instruction: `Title and cells in the deck's language. Return JSON {"title": "...", "rows": [["header 1", "header 2"], ["cell", "cell"]], "note": "one line on the source, or empty"}. At most 8 rows and 6 columns, short cells. Copy every number exactly from the deck or material and keep each value under its own year/column (check direction: which year is higher). Use only facts the user, the deck or the material give; if there are none, return {"error": "<one sentence asking for them>"}.\n\nTable: ${instruction}\n\nDeck: ${deckLines().join(' | ').slice(0, 3000)}\n\nMaterial: ${material().slice(0, 12000)}` }),
  })
  const d = await res.json()
  if (d.error || !d.rows?.length) fail(d.error || (t('Nothing to put in a table.')))
  const { width: W, height: H } = slideSize()
  const cols = Math.max(...d.rows.map((r: any[]) => r.length))
  const style = { fontname: theme.value.fontName, color: theme.value.fontColor, fontsize: '14px' }
  return { title: String(d.title || ''), note: String(d.note || ''), table: {
    type: 'table', id: nanoid(10), left: W * 0.08, top: H * 0.22, width: W * 0.84, height: d.rows.length * 40, rotate: 0,
    colWidths: new Array(cols).fill(1 / cols), cellMinHeight: 36,
    outline: { width: 1, style: 'solid', color: '#cccccc' },
    theme: { color: theme.value.themeColors[0], rowHeader: true, rowFooter: false, colHeader: false, colFooter: false },
    data: d.rows.map((r: any[], i: number) => Array.from({ length: cols }, (_, j) => ({
      id: nanoid(10), colspan: 1, rowspan: 1, text: String(r[j] ?? ''),
      style: i === 0 ? { ...style, color: '#ffffff', bold: true } : style, // header row sits on the accent colour
    }))),
  } as any }
}

const addTablePage = async (text: string, instruction: string, at: number) => {
  const { title, note, table } = await tableFor(instruction)
  insertPage(at, title, note, table)
}

// A table ON an existing page ("put the table on this slide"): in its free space, made on a full page as for a picture.
// It became a new, unstyled page, 2026-11-15.
const addTableOnPage = async (instruction: string, id: string) => {
  const { table } = await tableFor(instruction)
  const { width: W, height: H } = slideSize()
  const box = await roomFor(id, { left: 0.05 * W, top: 0.22 * H, width: 0.9 * W, height: Math.min(table.data.length * 40, 0.64 * H) }, 'a table',
    { width: 0.36 * W, height: Math.min(table.data.length * 24, 0.6 * H) })
  const narrow = box.width < 0.6 * W
  const rowH = Math.max(24, Math.min(40, box.height / table.data.length))
  const data = table.data.map((r: any[]) => r.map((c: any) => ({ ...c, style: { ...c.style, fontsize: narrow ? '11px' : '14px' } })))
  slidesStore.updateSlide({ elements: [...slides.value[slides.value.findIndex(x => x.id === id)].elements,
    { ...table, data, cellMinHeight: rowH - 4, left: box.left, top: box.top, width: box.width, height: rowH * table.data.length }] }, id)
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
  // past the page's background picture is not the page: a ChatSlide deck is drawn 960 wide on the 1000-wide canvas, and
  // a picture asked for page 5 went into the white strip beyond it, half off the design (replay 2026-11-21)
  const bd = firstBackdrop(sl)
  if (bd) {
    for (let r = 0; r < R; r++) {
      for (let c = 0; c < C; c++) {
        if ((c + 1) * cw > bd.left + bd.width + 1 || c * cw < bd.left - 1 || (r + 1) * rh > bd.top + bd.height + 1 || r * rh < bd.top - 1) busyCell[r][c] = true
      }
    }
  }
  // on a page that is one flat picture the text is painted in it: its lines are taken too (two pictures asked for
  // "the space above the page number" covered the page's text, replay 2026-11-20)
  for (const e of [...sl.elements, ...(paintedKnown.get(paintedPicture(sl).src) || [])] as any[]) {
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

// Free space on a page for something new (a picture, a table, a set of pictures): the nearest free box to `wanted`; on a
// full page the text is moved aside once per message so the right 40% is free. Fails when there is still no room.
// `min`: the smallest box that holds it (a set of pictures, a table); a smaller free corner counts as no room.
const roomFor = async (id: string, wanted: { left: number, top: number, width: number, height: number }, what: string, min = { width: 0, height: 0 }) => {
  const { width: W, height: H } = slideSize()
  const at = () => slides.value.findIndex(x => x.id === id)
  const free = (want: typeof wanted) => {
    const b = freeBox(slides.value[at()], want)
    return b && b.width >= min.width && b.height >= min.height ? b : null
  }
  let box = free(wanted)
  // a corner a fraction of the size asked is no room yet: a picture asked for page 5 came out 121x80 px in its top-right
  // corner, too small to see (replay 2026-11-21). The text is moved aside once; that corner stays the fallback
  const corner = box && box.width * box.height < 0.3 * wanted.width * wanted.height && !roomMadeThisTurn.has(id) ? box : null
  if (corner) box = null
  if (!box && !corner && roomMadeThisTurn.has(id)) fail(t('Page {p0} has no free space any more.', { p0: at() + 1 })) // made room once: never shrink it again
  // painted text cannot move aside: no room is no room
  if (!box && !(paintedPicture(slides.value[at()]).src && !slides.value[at()].elements.some((e: any) => elText(e)))) { // a full page: make room rather than refuse
    roomMadeThisTurn.add(id)
    await editSlide(at(), `Make room for ${what}: rearrange and narrow the text so the right 40% of the slide is free of text, keeping all the wording and keeping it readable. Do not change the wording.`).catch(() => undefined)
    const room = { left: 0.58 * W, top: 0.2 * H, width: 0.38 * W, height: 0.6 * H }
    box = free(room)
    if (!box) { // the edit did not free it: the page's content moves into the left part, scaled by position and width
      const page = slides.value[at()]
      slidesStore.updateSlide({ elements: page.elements.map((e: any) => (e.type === 'image' && isBackdrop(e) ? e : { ...e, left: e.left * 0.56, width: e.width * 0.56 })) }, id)
      box = free(room)
    }
  }
  // the corner is measured again on the page as it is now: the moved text may have gone into it
  const now = corner && free(wanted)
  if (now && (!box || box.width * box.height < now.width * now.height)) box = now
  if (!box) fail(t('Page {p0} is full; there is no free space for a picture. Remove or shrink something first, or put the picture on a new page.', { p0: at() + 1 }))
  return box
}

// Fetch a picture: a stock photo first (unless an AI image is asked for), else gpt-image.
// Fetch a picture from the source asked for; never quietly another kind. `credit` is said in the reply.
let pictureNote = ''
let pictureSource: 'photo' | 'wikimedia' | 'ai' | 'attached' = 'photo' // what the last picture really came from (a stock search can fall back to AI)
// The user's own words for the picture service, which looks at its candidates beside them: their last three messages
// ("try a different one" says nothing alone)
const saidLately = () => {
  const said = s.history.filter(m => m.role === 'user').map(m => m.text)
  if (said[said.length - 1] !== userText) said.push(userText)
  return said.slice(-3).join('\n')
}
// Stock photos the picture service has looked at: the first `fits` of them show what was asked. The service failing:
// the plain search in its own order, as before.
const stockPhotos = async (query: string, landscape: boolean, avoid: string[]): Promise<{ data?: { src: string }[], fits?: number }> => {
  const orientation = landscape ? 'landscape' : 'portrait'
  try {
    const r = await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'stock', query, want: saidLately(), orientation, avoid: avoid.filter(u => u.startsWith('http')) }) })
    if (r.ok) return await r.json()
  }
  catch { /* the plain search below */ }
  return api.searchImage({ query, per_page: 5, orientation } as any)
}
// none of the photos looked at shows what was asked: the nearest one goes on the slide, said as such, with the AI picture on offer
const nearestNote = () => t('No stock photo shows exactly that, so this is the nearest one. I can draw it with AI instead (10 credits): just say so.')
const findPicture = async (query: string, source: string, landscape: boolean, avoid: string[] = [], reuse = false) => {
  pictureNote = ''
  if (source === 'direct') { // a picture file already chosen from a website: only the paid add_site_pictures step may place one
    if (!directAllowed) fail(t('Real pictures from a brand or company website are a paid feature. Upgrade your plan and I will take them from the site.'))
    pictureSource = 'photo'
    return query
  }
  const nearest = source === 'medical' // a medical picture may be the nearest real one of the same subject
  const specific = source === 'specific' // one specific, named real thing: only Wikimedia, never a look-alike
  if (nearest || specific) source = 'wikimedia'
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
    const r = await (await fetch('/api/tools/research', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'wikimedia', query, nearest, avoid, want: saidLately() }) })).json()
    if (!r.src && specific) fail(t('Wikimedia Commons has no free picture of {p0}, and a stock photo would show something else. Would you like to upload one? Attach it with the 📎 icon and I will put it on the slide.', { p0: query.slice(0, 60) }))
    if (!r.src) { // no match: a stock photo, said plainly; a question here stopped a turn with three more pictures to add
      const sp = await stockPhotos(query, landscape, avoid).catch(() => ({ data: [] } as Awaited<ReturnType<typeof stockPhotos>>))
      const found = (sp.data || []).map((x: any) => x.src), alt = found.find((u: string) => u && !avoid.includes(u))
      if (!alt) fail(t('Wikimedia Commons has no picture of "{p0}". Should I use a stock photo or generate one instead?', { p0: query }))
      pictureSource = 'photo'
      pictureNote = t('Wikimedia Commons had no picture of "{p0}", so this is a stock photo.', { p0: query.slice(0, 60) })
        + (sp.fits !== undefined && found.indexOf(alt) >= sp.fits ? ` ${nearestNote()}` : '')
      return alt as string
    }
    pictureNote = t('Picture from Wikimedia Commons: {p0} ({p1}{p2}) {p3}', { p0: r.title, p1: r.credit, p2: r.license ? t(', {p0}', { p0: r.license }) : '', p3: r.page })
      + (r.close ? ` ${t('No exact picture exists there, so this is the nearest one of the same subject: check it is right for your slide.')}` : '')
    return r.src as string
  }
  if (source === 'photo') {
    // a few results, so "all the pictures" with similar searches do not get the same photo twice
    const r = await stockPhotos(query, landscape, avoid)
    const found = (r.data || []).map((x: any) => x.src), src = found.find((u: string) => u && !avoid.includes(u))
    if (src && r.fits !== undefined && found.indexOf(src) >= r.fits) pictureNote = nearestNote()
    if (src) return src as string
    // every result already used: the first again, rather than an AI picture nobody asked for (10 credits)
    if (reuse && r.data?.[0]?.src) return r.data[0].src as string
    // photos were found, all already on the page: said so, never a paid AI picture in their place
    if (found.some(Boolean)) fail(t('Every stock photo found for “{p0}” is already on this page.', { p0: query.slice(0, 60) }))
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

// width / height of a picture, 0 when it does not load
const imageRatio = (src: string) => new Promise<number>((ok) => {
  const im = new Image()
  const tm = setTimeout(() => ok(0), 8000)
  im.onload = () => { clearTimeout(tm); ok(im.naturalWidth && im.naturalHeight ? im.naturalWidth / im.naturalHeight : 0) }
  im.onerror = () => { clearTimeout(tm); ok(0) }
  im.src = src
})

// the picture an image holds goes on its list before the chat replaces it
const keepPic = (el: any, sl?: Slide) => {
  const list = [...(s.was?.[el.id] || [])]
  // links only: an embedded picture (a data: URL) kept here could push the saved session past the storage limit
  if (el.src.startsWith('data:')) return
  // a background picture keeps the page's other full-slide pictures and background with it (pass the slide as it is
  // before anything on it changes)
  const page = sl && isBackdrop(el) ? {
    others: Object.fromEntries(sl.elements.filter((e: any) => e.type === 'image' && e.id !== el.id && isBackdrop(e) && !e.src.startsWith('data:'))
      .map((e: any) => [e.id, { src: e.src, clip: e.clip }])),
    background: sl.background && !String(sl.background.image?.src || '').startsWith('data:') ? JSON.parse(JSON.stringify(sl.background)) : undefined,
  } : {}
  const at = list.findIndex(p => p.src === el.src)
  if (at >= 0) {
    if (sl) list[at] = { ...list[at], ...page }
    s.was = { ...s.was, [el.id]: list }
    return
  }
  list.push({ src: el.src, clip: el.clip, shows: s.pics?.[el.id] || '', ...page })
  if (list.length > 5) list.splice(1, 1) // the first one stays: it is the one users ask for
  s.was = { ...s.was, [el.id]: list }
}
const changeImage = async (text: string, slideId: string, image: number, query: string, source: string) => {
  const sl = slides.value.find(x => x.id === slideId)
  // picture 0 = the full-slide backdrop ("change the background picture"); 1.. = the content pictures
  const el = sl && (image === 0 ? firstBackdrop(sl) : pictures(sl)[Math.max(0, (image || 1) - 1)])
  if (!sl || !el) fail(t('That page has no picture.'))
  const src = await findPicture(query, source, el.width >= el.height, sl.elements.filter((e: any) => e.type === 'image').map((e: any) => e.src))
  if (!src) fail(t('Could not get a picture.'))
  // the new picture fills the old frame cropped to it, centred, never stretched (the template's frame stays as designed)
  const ratio = await imageRatio(src), frame = el.width / el.height
  const keep = ratio ? Math.min(1, frame / ratio) * 100 : 100, keepH = ratio ? Math.min(1, ratio / frame) * 100 : 100
  const clip = { shape: el.clip?.shape || 'rect', range: [[(100 - keep) / 2, (100 - keepH) / 2], [(100 + keep) / 2, (100 + keepH) / 2]] }
  // another full-slide picture showing the very same photo (set_background puts it on each of them) is that photo too:
  // replaced alone, the copy above it hid the new one and "Replaced the background picture" showed nothing new (nurse
  // deck, replayed 2026-11-20)
  const old = el.src
  const copies = isBackdrop(el) ? sl.elements.filter((e: any) => e.type === 'image' && e.id !== el.id && isBackdrop(e) && e.src === old) : []
  keepPic(el, sl)
  for (const e of [el, ...copies]) {
    if (e !== el) keepPic(e)
    slidesStore.updateElement({ id: e.id, props: { src, clip } as any, slideId })
    s.pics = { ...s.pics, [e.id]: query }
  }
  // the slide's own background, the same photo, shows at the edges the picture does not reach
  if (isBackdrop(el) && sl.background?.type === 'image' && sl.background.image?.src === old) slidesStore.updateSlide({ background: { ...sl.background, image: { ...sl.background.image, src } } }, slideId)
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
  const written: string[] = [] // the pages whose notes came back
  for (const id of ids) {
    const sl = slides.value.find(x => x.id === id)!
    busy.value = t('Writing notes {p0} of {p1}…', { p0: ids.indexOf(id) + 1, p1: ids.length })
    // a page that is one flat picture (an imported PDF) has its text painted in it: notes from "no separate text"
    // talked about "this picture" instead of the slide's points (replay 2026-11-23)
    const painted = paintedPicture(sl).src ? (await paintedLines(sl)).map(l => l.text).join('\n') : ''
    const r = await fetch('/api/tools/ask_json', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ strong: true, instruction: `Speaker notes for THIS one slide only (never mention page numbers or other slides): what the presenter says aloud, adding context the slide only implies, in the slide's language. Every number must be exactly as on the slide. Default length 3-5 short sentences. Request: ${instruction || 'write them'}\nSlide: ${pageText(sl)}${painted ? `\nThe text painted in the slide's picture, top to bottom:\n${painted}` : ''}\nCurrent notes: ${(sl.remark || '').replace(/<[^>]+>/g, ' ').trim() || '(none)'}\nReturn JSON {"notes": "..."} (paragraphs separated by \\n).` }),
    })
    const notes = String((await r.json()).notes || '').trim()
    if (notes) {
      slidesStore.updateSlide({ remark: String(notes).split('\n').filter(Boolean).map(l => `<p>${l.replace(/</g, '&lt;')}</p>`).join('') }, id)
      written.push(id)
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
  // a new picture or table on a page that is one flat picture: its painted lines are read first, so it is not put on them
  if (['add_image', 'add_table_page', 'use_file_images'].includes(action)) {
    const on = slides.value.find(x => x.id === ids[Number(args.index || args.on_page) - 1])
    if (on && paintedPicture(on).src) await paintedLines(on)
  }
  if (['add_page', 'add_chart_page', 'add_table_page'].includes(action) && Number(args.after) > slides.value.length) args.after = slides.value.length
  if (['add_page', 'add_chart_page', 'add_table_page'].includes(action) && Number(args.layout_from) > slides.value.length) args.layout_from = slides.value.length
  // pages from the user's own file, and none is attached: ask for it (the request waits for the file), never invent
  // the content ("buatkan ppt terkait makalah ini" generated a deck about nothing, 2026-11-12)
  if ((action === 'add_page' || action === 'create_deck') && args.from_users_file && !s.docs.length && !(s.images || []).length) {
    return runCall(text, 'ask_user', { question: t('No file is attached yet: attach it with the 📎 icon in the chat box.'), options: [] }, ids)
  }
  if (action === 'add_page') {
    // "after this one" / no place named: right after the page the user was on (the router once said "after 30")
    // the page the user was on can be gone: "remove slide 2 and add 3 slides" from page 2 put them after page 0 and
    // failed (2026-11-20); the router's own place then stands
    if (args.after_named === false && currentAtSend) {
      const made = slides.value.map((x, i) => (createdThisMessage.has(x.id) ? i : -1)).filter(i => i >= 0)
      const at = made.length ? Math.max(...made) : pos(currentAtSend)
      if (at >= 0) args.after = at + 1 // a 2nd new page follows the 1st
    }
    if (!(Number(args.after) >= 1)) args.after = Math.min(Math.max(slideIndex.value + 1, 1), slides.value.length)
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    // up to 30 new pages in one step: "increase to 30 slides" on a 10-page deck stopped at 20 with "ask again for
    // more" (2026-11-15), and the router's next step claimed a wrong total
    const count = Math.min(Math.max(args.count || 1, 1), 30)
    const capped = (args.count || 1) > 30 ? (t(' ({p0} is the most per request; ask again for more)', { p0: 30 })) : ''
    const n0 = slides.value.length
    // a layout to copy needs text boxes: a page that is one flat picture (an imported PDF) or empty has none, and its
    // copy came back "not changed" and stayed in the deck as a duplicate (three users, 2026-11-14/15). Another page
    // with text lends its layout; a deck without one gets generated pages
    const usable = (x?: Slide) => !!x && !isFlat(x) && x.elements.some((e: any) => elText(e))
    const asked = slides.value[pos(idOf(args.layout_from || args.after))]
    const source = usable(asked) ? asked : slides.value.find(usable)
    if (s.mode === 'native' || !source) {
      const topics = Array.isArray(args.topics) && args.topics.length > 1 ? ` One page per topic, in this order: ${args.topics.join('; ')}.` : ''
      // the generator writes up to 10 pages a call; each further call sees the pages already there (deck so far)
      for (let made = 0; made < count; made += 10) {
        const part = count > 10 ? ` (These are new pages ${made + 1}-${Math.min(made + 10, count)} of ${count}; the deck so far already holds the earlier ones: go on with what is not covered yet, no repeats.)` : ''
        // a later batch that throws keeps the pages already written, and the reply counts them
        if (!await generatePage(text, args.instruction + topics + part, after + (slides.value.length - n0), Math.min(10, count - made)).catch((e) => { if (!made) throw e; return false })) {
          if (!made) fail(t('Could not write that page.'))
          break
        }
      }
      const added = slides.value.length - n0
      return (t('Added {p0} page(s) after page {p1}: {p2}.', { p0: added, p1: after, p2: newTitles(after, added) })) + capped
    }
    // text the user pasted for the pages (in this message or the request this message answers) goes to each page as
    // written: the router's instruction for 21 pasted slides said only "add them after the presentation" (2026-11-15)
    const said = [s.pending?.request, text].filter((x, k, all) => x && all.indexOf(x) === k).join('\n')
    const own = said.length > 600 ? `\nThe user's own text (where it gives this page's content, use that word for word):\n${fence(said.slice(0, 12000))}` : ''
    const written: string[] = [] // titles of the pages written in this step
    let k = 0
    for (; k < count; k++) {
      busy.value = t('Writing page {p0} of {p1}…', { p0: k + 1, p1: count })
      const list = [...slides.value]
      const page = cloneSlide(source)
      list.splice(after + k, 0, page)
      slidesStore.setSlides(list)
      // each new page from its own topic: with one shared instruction "TikTok and Reddit" became two TikTok pages
      // in the slide language, not the topic's: the router wrote "Kawa w Brazylii" for an English request (2026-11-15)
      const own1 = Array.isArray(args.topics) && args.topics.length >= count && args.topics.length === Math.max(args.count || 1, 1) ? String(args.topics[k]) : ''
      const topic = own1 ? `${own1} (part of: ${args.instruction})` : args.instruction
      const before = written.length && !own1 ? ` The pages before it in this batch are: ${written.map(x => `"${x}"`).join(', ')}; this page takes the next part, never one of those again.` : ''
      // the deck's own language, read from its slides and the user's words: the router named German for an English
      // deck and an English request, and 12 German pages came out (e2e, 2026-11-16)
      const write = `This is NEW page ${k + 1} of ${count} on: ${topic}.${before} The page is a copy of another slide, made only to lend its layout: its present text belongs to that slide and is all replaced with this page's own content (no repeats across pages). Keep the layout and styling. Write all its text in ${deckLanguage(said)}.${own}`
      const advanced = !!args.hard && !!(chatslide?.canUseAdvanced?.() ?? (window as any).__evalAdvanced)
      try {
        // "this page contains unrelated coffee text, so I left it unchanged" (page 6 of 12, e2e 2026-11-16): once more, plainly
        await editSlide(after + k, write, advanced).catch((e) => {
          if (!(e instanceof NoChange)) throw e
          return editSlide(after + k, `${write}\nThe text now on this page is NOT its content: replace every line of it.`, advanced)
        })
      }
      catch (e) {
        // the copy that could not be written is taken out again: it is the page it was copied from, a second time
        slidesStore.setSlides(slides.value.filter(x => x.id !== page.id))
        if (!(e instanceof StepFailed) || !k) throw e
        break // the pages written so far stay, and the reply counts them
      }
      // a line left exactly as on the copied page (the small header "Key Takeaway: Match Platform to Your Goal" stayed
      // on a new TikTok page): one more pass names those lines
      const copied = new Set(source.elements.map(elText).filter(Boolean))
      const stale = slides.value[after + k].elements.map(elText).filter(t => t && copied.has(t))
      if (stale.length) await editSlide(after + k, `These lines are still the text of the page this one was copied from; rewrite each for THIS page's topic (${own1 || args.instruction}), same length and style: ${stale.map(t => `"${t.slice(0, 120)}"`).join('; ')}`).catch(() => undefined)
      written.push(titleAt(after + k))
      // pictures still the copied page's (a new "coffee prices" page kept slide 2's barista photos, 2026-11-11): a stock
      // photo for this page's topic in the same frame, so the layout stays and nothing repeats
      const copiedPics = new Set(source.elements.filter((e: any) => e.type === 'image' && !isBackdrop(e)).map((e: any) => e.src))
      const { width: SW, height: SH } = slideSize()
      for (const [n, e] of pictures(slides.value[after + k]).entries()) {
        if (isBackdrop(e) || !copiedPics.has(e.src) || e.width * e.height < 0.02 * SW * SH) continue // keep backdrops and logos
        await changeImage(text, page.id, n + 1, String(own1 || written[k] || args.instruction).slice(0, 80), 'photo').catch(() => undefined)
      }
    }
    return (t('Added {p0} page(s) after page {p1}: {p2}.', { p0: k, p1: after, p2: newTitles(after, k) })) + capped
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
  if (action === 'add_table_page' && Number(args.on_page) > 0) {
    const id = idOf(Number(args.on_page))
    await addTableOnPage(args.instruction, id)
    touched([id])
    slidesStore.updateSlideIndex(pos(id))
    return t('Put the table on page {p0}.', { p0: pos(id) + 1 })
  }
  if (action === 'add_table_page') {
    const after = beforeClosing(pos(idOf(args.after)) + 1)
    await addTablePage(text, args.instruction, after)
    slidesStore.updateSlideIndex(after)
    return t('Added a table page "{p0}".', { p0: titleAt(after) })
  }
  // the colours and the logo of the user's own .pptx (the file the deck was made from, or one attached here)
  if (action === 'use_file_look') {
    const listed = s.docs.filter(d => /\.pptx$/i.test(d.name))
    const file = await fileOf(listed.some(d => d.name === args.file) || listed.length !== 1 ? args.file : listed[0].name)
    if (!file || !/\.pptx$/i.test(file.name)) fail(t('No PowerPoint file of yours is here yet: attach it with the 📎 icon.'))
    busy.value = t('Reading the look of {p0}…', { p0: file!.name })
    const look = await file!.arrayBuffer().then(normalizePptx).then(b => parse(b, { imageMode: 'base64', videoMode: 'blob', audioMode: 'blob' })).then(fileLook)
      .catch(() => fail(t('I could not read {p0} as a PowerPoint file.', { p0: file!.name })))
    const ids = slides.value.map(x => x.id)
    const said: string[] = []
    if (args.colors !== false) {
      // the template's own full-slide decoration (shapes on white or see-through) would hide the new background and
      // is not the user's look: it goes; a full-slide photo (a cover picture) stays as it is
      const decoration = new Map<string, boolean | null>() // null: the picture could not be read
      const isDecoration = (src: string) => new Promise<boolean | null>(res => {
        const img = new Image()
        img.crossOrigin = 'anonymous'
        img.onload = () => {
          try {
            const c = document.createElement('canvas')
            c.width = c.height = 32
            const g = c.getContext('2d')!
            g.drawImage(img, 0, 0, 32, 32)
            const px = g.getImageData(0, 0, 32, 32).data
            let plain = 0
            for (let i = 0; i < px.length; i += 4) if (px[i + 3] < 32 || (px[i] > 235 && px[i + 1] > 235 && px[i + 2] > 235)) plain++
            res(plain >= 32 * 32 * 0.55)
          }
          catch {
            res(null)
          }
        }
        img.onerror = () => res(null)
        img.src = src
      })
      for (const sl of slides.value) {
        for (const e of sl.elements as any[]) if (e.type === 'image' && isBackdrop(e) && !decoration.has(e.src)) decoration.set(e.src, await isDecoration(e.src))
      }
      // a page that is only its picture (a pasted screenshot, an imported flat page) keeps it whatever it looks like;
      // a page with a photo keeps its own text colours too (a white title on a dark photo stays readable)
      const themed: string[] = []
      for (const sl of [...slides.value]) {
        const hasText = (sl.elements as any[]).some(e => plain(e.type === 'text' ? e.content : e.text?.content))
        const back = (sl.elements as any[]).filter(e => e.type === 'image' && isBackdrop(e))
        const goes = hasText ? back.filter(e => decoration.get(e.src) === true) : []
        if (goes.length) slidesStore.updateSlide({ elements: (sl.elements as any[]).filter(e => !goes.includes(e)) }, sl.id)
        const unread = hasText && back.some(e => decoration.get(e.src) === null)
        if (unread) fadeBackdrop(sl.id) // unreadable picture: faded, as a new theme does, so the background shows
        if (unread || back.length === goes.length) themed.push(sl.id)
      }
      applyPresetTheme({ background: look.background, fontColor: look.fontColor, fontname: '', colors: look.colors }, true, themed)
      slidesStore.setTheme({ backgroundColor: look.background, fontColor: look.fontColor, themeColors: look.colors.slice(0, 6), ...(look.font ? { fontName: look.font } : {}) })
      // the file's own type on every page, photo pages too: the template's handwriting title stayed on a surgeon's
      // deck in his colours (2026-11-16)
      said.push(t('Applied the colours of {p0} to the whole deck.', { p0: file!.name }))
      if (look.font) said.push(await runCall(text, 'set_font', { pages: [], font: look.font, same_as_page: 0, scope: 'all' }, ids))
    }
    if (args.logo !== false && !look.logo) said.push(t('{p0} has no logo I can find (a small picture in a corner). Attach the logo with the 📎 icon and I will add it.', { p0: file!.name }))
    else if (args.logo !== false) {
      const logo = look.logo!
      const src = chatslide?.uploadImage ? await chatslide.uploadImage(await (await fetch(logo.src)).blob()) : logo.src
      const { width: W, height: H } = slideSize()
      // under 4% of the slide, so the chat knows it as the logo from then on (logoSpots)
      const shape = await new Promise<number>(res => {
        const i = new Image()
        i.onload = () => res(i.naturalWidth / i.naturalHeight)
        i.onerror = () => res(logo.ratio)
        i.src = logo.src
      })
      let height = H * Math.min(logo.height, 0.14), width = height * shape
      const k = Math.min(1, Math.sqrt(0.035 * W * H / (width * height)))
      width *= k
      height *= k
      const m = W * 0.025
      const right = logo.left + logo.width / 2 > 0.5, bottom = logo.top + logo.height / 2 > 0.5
      // one corner for the whole deck: the file's own corner unless another is free on more pages
      const corners = [[right, bottom], [true, false], [false, false], [true, true], [false, true]]
        .map(([r, b]) => ({ left: r ? W - m - width : m, top: b ? H - m - height : m }))
      const free = (sl: Slide, c: { left: number, top: number }) => !(sl.elements as any[]).some(e => e.width && e.height && e.width * e.height < W * H * 0.6
        && e.left < c.left + width && e.left + e.width > c.left && e.top < c.top + height && e.top + e.height > c.top)
      const at = corners.map(c => ({ c, pages: slides.value.filter(sl => free(sl, c)) })).sort((x, y) => y.pages.length - x.pages.length)[0]
      // a deck that already has a logo keeps its place: the picture is swapped (asking twice gave two logos)
      const spots = logoSpots()
      const has = (sl: Slide) => (sl.elements as any[]).filter(e => e.type === 'image' && spots.has(imgSpot(e)))
      const swap = slides.value.filter(sl => has(sl).length)
      for (const sl of swap) for (const e of has(sl)) slidesStore.updateElement({ id: e.id, props: { src, height: e.width / shape }, slideId: sl.id })
      if (!swap.length) for (const sl of at.pages) slidesStore.updateSlide({ elements: [...sl.elements, { type: 'image', id: nanoid(10), src, fixedRatio: true, rotate: 0, ...at.c, width, height } as any] }, sl.id)
      said.push(t('Put the logo from {p0} on {p1} page(s).', { p0: file!.name, p1: swap.length || at.pages.length }))
    }
    touched(ids)
    return said.join(' ')
  }
  if (action === 'use_file_images') { // figures from the user's own PDF/PPTX/DOCX, extracted by ChatSlide's backend
    const file = await fileOf(args.file)
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
    if (sl?.elements.some((e: any) => e.type === 'image' && !isBackdrop(e))) return runCall(text, 'change_image', { index: pos(sl.id) + 1, image: 1, query: args.query, source: args.source, page_named: true, said_background: '', specific_thing: args.specific_thing }, ids)
  }
  // several listed things: one picture each, in a row along the free bottom or right of the page (asked as separate
  // calls, the router added one picture and stopped, 2026-11-11)
  // the language of new slide text, named by the router and stated in the instruction itself: instructions written in
  // Chinese gave an English user's new pages Chinese text (2026-11-12, "generate more slides upto 30")
  // (a translation is written in its target language, never the deck's: the two would contradict)
  const slideLang = String(args.translate_to || args.language || '').slice(0, 40)
  if (['add_page', 'edit_page', 'edit_all_pages'].includes(action) && slideLang && !String(args.instruction || '').includes('Write all slide text in')) {
    // first, and above the instruction's own language: the router wrote Spanish instructions for English requests
    // and the new slides came out in Spanish under a trailing note (2026-11-15)
    // on a page that already has text, only the text the edit writes: "the text is too cramped" sent in Spanish about an
    // English page came with the language Spanish, and the whole page was translated unasked; the user saw it and wrote
    // "it has to be in English" (2026-11-20). A translation names its language in translate_to, and says so
    args = { ...args, instruction: action !== 'add_page' && !args.translate_to
      ? `Write all slide text in ${slideLang}, whatever language the instruction below is in. This is for the text you write: text already on the page stays in the language it is in unless the instruction below asks to translate it.\n${args.instruction}`
      : `Write all slide text in ${slideLang}, whatever language the instruction below is in.\n${args.instruction}` }
  }
  // a medical picture is a real one (Wikimedia, which falls back to a labelled stock photo): asked in prose, the router
  // still chose stock photos 6 of 6 times for "a picture for each variant of pulmonary hypertension" (2026-11-11)
  // an AI picture only when the user's words asked for one: it costs them 10 credits each ("сделать картинку на каждый
  // вариант" drew 4 AI pictures, 40 credits, unasked, 2026-11-11)
  if ((action === 'add_image' || action === 'change_image') && args.source === 'ai' && !args.asked_for_ai) args = { ...args, source: 'photo' }
  if ((action === 'add_image' || action === 'change_image') && args.medical && ['photo', 'wikimedia'].includes(args.source)) args = { ...args, source: 'medical' }
  // one specific, named real thing (a person, a building, an artwork, a product model; the router decides) comes only from
  // Wikimedia, and with no match the user is asked to upload one: the stock fallback put a stranger on the slide labelled
  // as Daniel Stufflebeam, three times (2026-11-14); Quanlai, 2026-11-15: "not a real person... a specific item"
  if ((action === 'add_image' || action === 'change_image') && String(args.specific_thing || '').trim() && ['photo', 'wikimedia'].includes(args.source) && !(args.items?.length > 1)) {
    args = { ...args, source: 'specific', query: String(args.specific_thing).trim() } // searched by the name, not a description
  }
  // a picture of one brand's own things asked for as an ordinary picture: it comes from that brand's website (the router
  // split "more photos of the brand's products" into generic stock searches when the message had a second wish)
  if ((action === 'add_image' || action === 'change_image') && String(args.own_brand || '').trim() && args.source !== 'specific' && args.source !== 'direct' && args.source !== 'attached' && !args.asked_for_ai && !(args.earlier > 0)) {
    const items = (args.items || []).filter(Boolean)
    return runCall(text, 'add_site_pictures', { index: args.index, page_named: args.page_named, site: '', brand: args.own_brand, what: [args.query, ...items].filter(Boolean).join('; '), count: items.length || 1 }, ids)
  }
  // Real pictures from a brand's or company's own website: a paid feature. Free users are told so and get the plans
  // (Quanlai, 2026-11-12: "say that this is a paid feature and ask them to pay", "an agent reading a certain website").
  if (action === 'add_site_pictures') {
    if (!(chatslide?.canUseAdvanced?.() ?? (window as any).__evalAdvanced)) {
      (chatslide?.offerPaid || chatslide?.offerAdvanced)?.()
      sitePaywalledPage = ids[Number(args.index) - 1] || ''
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
    // the router names the specific thing each picture must show, so each one goes to Wikimedia on its own ("Einstein,
    // Curie and Bohr" drew three stock strangers, review 2026-11-15)
    const specifics: string[] = (Array.isArray(args.item_specifics) ? args.item_specifics : []).filter((_: any, k: number) => args.items[k])
    const items = args.items.filter(Boolean).slice(0, 6) // "6 pictures" got 4 (prod 2026-11-12, b1b)
    const pg = Number(args.index)
    const at = pos(idOf(pg))
    const setKey = picKey(ids[at], items.join('|'))
    if (picturedThisTurn.has(setKey)) fail(t('Those pictures were already added in this message.'))
    const { width: W, height: H } = slideSize()
    const row = { left: 0.05 * W, top: 0.6 * H, width: 0.9 * W, height: 0.3 * H }
    const fb = freeBox(slides.value[at], row)
    // a page that is one flat picture: its text is painted and cannot move aside, and the router was told where its
    // lines are, so the pictures go in the free box nearest to the place the router chose
    const flatPage = !!paintedPicture(slides.value[at]).src && !slides.value[at].elements.some((e: any) => elText(e))
    const roomy = !flatPage && items.length <= 4 && !!fb && fb.width >= row.width * 0.8 && fb.height >= row.height * 0.8 // five or six: a grid page
    // no row free: up to four pictures go on this page in the room made for them ("add photos of A and B on this
    // slide" became a new, unstyled page, 2026-11-15)
    let area: { left: number, top: number, width: number, height: number } | null = null
    if (flatPage) {
      const b = args.box || { x: 5, y: 60, w: 90, h: 30 }
      const pc = (v: number) => Math.min(Math.max(Number(v) || 0, 0), 100) / 100
      area = freeBox(slides.value[at], { left: pc(b.x) * W, top: pc(b.y) * H, width: Math.max(pc(b.w), 0.1) * W, height: Math.max(pc(b.h), 0.08) * H })
    }
    else if (!roomy && items.length <= 4) {
      try {
        area = await roomFor(ids[at], { left: 0.55 * W, top: 0.2 * H, width: 0.4 * W, height: 0.6 * H }, 'pictures', { width: 0.25 * W, height: 0.3 * H })
      }
      catch (e) { if (!(e instanceof StepFailed)) throw e }
    }
    // still no room: the pictures go on a new page right after it, in a grid, in that page's look (its title, header,
    // footer and background; on a full slide four pictures landed in the corners and off the edge, 2026-11-11)
    let pageId = ''
    if (!roomy && !area) {
      const title: any = readingOrder(slides.value[at]).find((e: any) => elText(e))
      insertPage(at + 1, title ? elText(title) : '', '', { type: '__pictures' })
      const page = slides.value[at + 1]
      area = page.elements.find((e: any) => e.type === '__pictures') as any
      slidesStore.updateSlide({ elements: page.elements.filter((e: any) => e.type !== '__pictures') }, page.id)
      ids.splice(0, ids.length, ...slides.value.map(x => x.id))
      pageId = page.id
      helperPages.add(page.id) // not a page the user asked for: it must not satisfy "new content goes on the new page"
    }
    const wide = !!area && area.width >= area.height * 1.5
    const cols = roomy || (wide && items.length <= 4) ? items.length : items.length <= 2 ? (pageId ? items.length : 1) : items.length <= 4 ? 2 : 3
    const rows = Math.ceil(items.length / cols)
    const a = area ? { x: area.left / W * 100, y: area.top / H * 100, w: area.width / W * 100, h: area.height / H * 100 } : null
    const lines: string[] = []
    let added = 0
    for (const [k, q] of items.entries()) {
      const c = k % cols, r = Math.floor(k / cols)
      const box = roomy || !a ? { x: 5 + k * (90 / items.length), y: 62, w: 90 / items.length - 2, h: 33 }
        : { x: a.x + c * (a.w / cols), y: a.y + r * (a.h / rows), w: a.w / cols - 1, h: a.h / rows - 1 }
      try {
        // each item is a new picture of its own: never "replace", which overwrote the previous item
        lines.push(await runCall(text, 'add_image', { ...args, items: [], item_specifics: [], specific_thing: String(specifics[k] || ''), replaces_existing: false, index: pageId ? at + 2 : pg, page_named: true, query: q, box }, ids))
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
    // at most 6 new pictures on a page per message, as a picture set may hold (a reworded re-plan would otherwise keep
    // adding near-duplicates)
    if ([...picturedThisTurn].filter(k => k.startsWith(`${id}|`)).length >= 6) fail(t('Page {p0} already got 6 new pictures in this message.', { p0: pos(id) + 1 }))
    // a specific thing may have no free picture: looked up before a full page is rearranged for it (2026-11-15)
    // never a picture already on this page, or on the page a picture page was added for: "insert related pictures" put
    // the same Wikimedia file of the named place on page 6 and twice on its new picture page, the specific-thing search
    // being sent no pictures to avoid (2026-11-21). Other pages may repeat one (a logo on every page).
    const twin = helperPages.has(id) ? slides.value[pos(id) - 1] : undefined
    const avoid = [sl, twin].flatMap(x => (x?.elements || []).filter((e: any) => e.type === 'image').map((e: any) => e.src))
    if (args.source === 'specific') busy.value = t('Finding a picture…')
    const early = args.source === 'specific' ? await findPicture(args.query, args.source, wanted.width >= wanted.height, avoid) : ''
    const box = await roomFor(id, wanted, 'a picture')
    busy.value = args.source === 'ai' ? t('Drawing the picture…') : t('Finding a picture…')
    const src = early || await findPicture(args.query, args.source, box.width >= box.height, avoid)
    if (!src) fail(t('Could not get a picture.'))
    const now = freeBox(slides.value[pos(id)], box) // re-checked: the page may have changed while the picture loaded
    if (!now) fail(t('Page {p0} has no free space any more.', { p0: pos(id) + 1 }))
    // the picture keeps its own shape, centred in the box: a 3:4 portrait filled a wide box stretched (2026-11-15)
    const ratio = await imageRatio(src)
    if (ratio) {
      const w = Math.min(now.width, now.height * ratio), h = w / ratio
      Object.assign(now, { left: now.left + (now.width - w) / 2, top: now.top + (now.height - h) / 2, width: w, height: h })
    }
    const picId = nanoid(10)
    slidesStore.updateSlide({ elements: [...slides.value[pos(id)].elements, { type: 'image', id: picId, src, fixedRatio: true, rotate: 0, ...now } as any] }, id)
    s.pics = { ...s.pics, [picId]: String(args.query || '') }
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
  // the user's logo moves and resizes directly: the logo is the picture at the same small spot on most pages, or the
  // picture they attached. Routed through edit_page it moved without shrinking, or the reply said "no logo found"
  // without asking them to attach one (eval 2026-11-14)
  if (action === 'move_logo') {
    const logos = logoSpots()
    const own = new Set((s.images || []).map(m => m.url))
    const { width: W, height: H } = slideSize()
    const f = args.size === 'smaller' ? 0.6 : args.size === 'bigger' ? 1.5 : 1
    const corner = String(args.corner || 'same'), pad = W * 0.03
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    const moved: string[] = []
    for (const id of targets) {
      const sl = slides.value[pos(id)]
      for (const e of sl.elements as any[]) {
        if (e.type !== 'image' || isBackdrop(e) || !(logos.has(imgSpot(e)) || own.has(e.src))) continue
        const k = Math.min(f, (W * 0.4) / e.width, (H * 0.4) / e.height) // never past 40% of the slide either way
        const w = e.width * k, h = e.height * k
        const fit = (v: number, room: number) => Math.min(Math.max(v, 0), room) // kept inside the slide
        const left = fit(corner.endsWith('right') ? W - w - pad : corner.endsWith('left') ? pad : e.left + (e.width - w) / 2, W - w)
        const top = fit(corner.startsWith('top') ? pad : corner.startsWith('bottom') ? H - h - pad : e.top + (e.height - h) / 2, H - h)
        slidesStore.updateElement({ id: e.id, props: { width: w, height: h, left, top } as any, slideId: id })
        if (!moved.includes(id)) moved.push(id)
      }
    }
    if (!moved.length) fail(t('There is no logo on the slides yet. Attach your logo with the 📎 icon and tell me where to put it.'))
    touched(moved)
    return t('Moved the logo on {p0} page(s).', { p0: moved.length })
  }
  if (action === 'remove_images') {
    const targets = args.only?.length ? nums(args.only) : args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    // several pages go only after a yes, like deleting pages: an attached picture plus "every slide" removed all 14
    // pictures of the deck (2026-11-12)
    // and on one page only when the router says the user's words asked for it ("..." after "like page 6" removed one,
    // eval 2026-11-13). The yes button runs exactly this removal, page by page, without asking the router again
    const logos = logoSpots()
    if (!args.confirmed && (targets.length > 1 || args.user_asked_to_remove !== true)) {
      // only pages that have a picture to remove: a page with none would end the confirmed run there (eval 2026-11-14)
      const withPictures = targets.filter(id => pictures(slides.value[pos(id)]).some(e => !logos.has(imgSpot(e))))
      if (!withPictures.length) fail(t('There is no picture to remove.'))
      const count = withPictures.reduce((k, id) => k + (args.images?.length || pictures(slides.value[pos(id)]).filter(e => !logos.has(imgSpot(e))).length), 0)
      const q = t('Remove {p0} picture(s) from {p1} page(s)? Nothing is removed yet.', { p0: count, p1: withPictures.length })
      s.pending = { request: userText, question: q, options: [t('Yes, remove them'), t('No, keep them')], offer: true,
        // the yes removes them in one step: page by page, the reply read "Removed 1 picture(s). Removed 3 picture(s)."
        // for 10 pictures on 6 pages (2026-11-16)
        repeat: [{ action: 'remove_images', args: { ...args, confirmed: true, only: withPictures.map(id => pos(id) + 1) } }], pages: [pos(withPictures[0]) + 1] }
      askedThisTurn = true
      save()
      return q
    }
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
      // never the one full-slide picture of a page that is nothing else (a pasted screenshot): that would blank the page
      const all = content.length ? content : onlyFullPicture(sl) ? [] : sl.elements.filter((e: any) => e.type === 'image' && !logos.has(imgSpot(e)))
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
    const pasted = targets.filter(id => onlyFullPicture(slides.value[pos(id)]))
    if (!n && pasted.length) fail(t('Page {p0} is one flat picture: its text and pictures are all inside that single image, so I can\'t change a picture in it. Nothing was changed. I can add a new picture on top of it.', { p0: pasted.map(id => pos(id) + 1).join(', ') }))
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
    let del = nums(args.pages)
    // "10 slides it must be": the pages listed first go, until that many are left
    const keep = Number(args.pages_left) || 0
    if (keep > 0 && n - del.length < keep) del = del.slice(0, Math.max(0, n - keep))
    if (!del.length) fail(t('The deck already has {p0} pages.', { p0: n }))
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
    const gone = del.map(id => pos(id) + 1).sort((x, y) => x - y)
    slidesStore.deleteSlide(del)
    // the router is told which pages went: after "slide 4 hapus" on 4 pages it saw 3 and added "there is no slide 4,
    // nothing deleted" to a reply that had just deleted it (5 of 6 runs, 2026-11-18)
    memo = `(Page(s) ${gone.join(', ')}, numbered as the deck was just before this step, are deleted: done, never "missing" or "not there". The deck now has ${slides.value.length} pages.)`
    return t('Deleted {p0} page(s); {p1} remain.', { p0: del.length, p1: slides.value.length })
  }
  if (action === 'read_file') {
    // what the user asked for when the deck was made, read on request ("put back what I sent at the start"); a short one
    // is not in the router's context with every message (it pulled one-page requests into whole-deck rewrites)
    if (args.file === 'Original request' && s.brief?.text) {
      memo = `<file-content name="Original request">\n${fence(s.brief.text)}\n</file-content>`
      return t('Read {p0}{p1}.', { p0: t('your original request'), p1: '' })
    }
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
    readThisTurn.add(doc.name)
    // one thing from a project file longer than one read: answered from the whole file on the server (a 100-page PDF
    // read from the top showed the chat only its first 4-5 pages, 2026-11-15)
    const about = String(args.about || '').trim()
    if (doc.project && about && doc.text.length > 12000 && !(args.pages || []).length && hostSteps()?.askFiles) {
      busy.value = t('Reading {p0}…', { p0: doc.name })
      const got = await hostSteps()!.askFiles!(about).catch(() => null)
      if (got) {
        memo = `<file-content name="${doc.name.replace(/["<>]/g, '')}" about="${about.replace(/["<>]/g, '').slice(0, 120)}">\n${fence(got.answer.slice(0, 12000))}${got.source ? `\n(from: ${got.source.slice(0, 200)})` : ''}\n</file-content>`
        return t('Read {p0}{p1}.', { p0: doc.name, p1: '' })
      }
    }
    const want: number[] = args.pages || []
    let parts = filePages(doc.text)
    // a PDF the project already had comes from ChatSlide as one text with no page marks: asked for a page, its original
    // is read here once, page by page, like a file attached in the chat
    if (want.length && !parts.length && doc.project && /\.pdf$/i.test(doc.name)) {
      busy.value = t('Reading {p0}…', { p0: doc.name })
      const file = await fileOf(doc.name).catch(() => undefined)
      const again = file ? await deckApi('read', { url: await uploadFile(file), name: file.name }).catch(() => null) : null
      if (again?.text && filePages(String(again.text)).length) {
        doc.text = String(again.text)
        save()
        parts = filePages(doc.text)
      }
    }
    const total = parts.length
    const missing = want.filter(n => n < 1 || n > total)
    if (want.length && total && missing.length === want.length) fail(t('{p0} has only {p1} pages.', { p0: doc.name, p1: total }))
    const got = want.filter(n => !missing.includes(n))
    const picked = want.length && total ? parts.filter(p => got.includes(p.n)).map(p => p.text).join('\n') : doc.text
    if (!picked.replace(/^(\[page \d+\]|\[slide \d+\]|Slide \d+:)/gm, '').trim()) fail(t('No text could be read from {p0}{p1} (it may be a scanned image).', { p0: doc.name, p1: want.length ? t(' page(s) {p0}', { p0: want.join(', ') }) : '' }))
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
    // a long pasted message is material: page by page, its words go on the page, not the router's summary of them
    // (a pasted 11-slide plan edited one page at a time came back reworded, 2026-11-19)
    if (text.length > 600 && !args.translate_to) args = { ...args, instruction: `${args.instruction}\n\nThe user's own message, in full: where it gives text for this page, use that part word for word, every sentence:\n${fence(text.slice(0, 12000))}` }
    // "slide 5: main clinical signs: 1 skin ..." gives the page its topic: the new text was added in a corner of the
    // old page, its old text squeezed aside, and the user asked three more times ("make a NEW slide 5", "redo slide 5
    // completely", 2026-11-11; replayed 2026-11-20: the same on 2 of 3 such messages). The router says the page takes
    // new content; the page edit is then told the old content goes
    // Its pictures, charts and tables go only when the same message brings pictures for it (pictures_next): a pasted
    // "Slide 1 ... Slide 5 ..." placed page by page lost page 2's two pictures, the end check undid the message, and
    // the retry left pages 2-5 unplaced (editor case ed-numbered-text-literal, 3 of 4 runs, 2026-11-20)
    const oldParts = Number(args.pictures_next) > 0 ? 'text box, picture, chart and table' : 'text box'
    if (args.replaces_content === true && !args.translate_to) args = { ...args, instruction: `${args.instruction}\n\nThis page takes NEW content: what the instruction gives replaces what the page says now. Delete every ${oldParts} that belongs to the old content (anything about another subject, or naming the author or the event on a page that is not the cover); keep the page's design (background, decorative shapes, logo, footer, page number) and anything that already is part of the new content. A card, panel or number badge that held old content is used for the new content or deleted, never left empty, and text stays inside its card. Then lay the new content out over the free area: one title, and the body in readable sizes. Never squeeze old text aside and never leave it beside or under the new.${oldParts === 'text box' ? ' Keep the page\'s pictures, charts and tables, moved or resized to fit the new text if needed.' : ''}` }
    // text and pictures for one page in one message: the text filled the page, so the pictures went to a new page
    // after it ("slide 11 ... with pictures" got them on slide 12) or into the corners over the text (2026-11-11). The
    // page edit keeps the part free that the picture step asks for first
    if (Number(args.pictures_next) > 0 && !args.translate_to) {
      const { width: W, height: H } = slideSize()
      args = { ...args, instruction: `${args.instruction}\n\n${Math.min(Number(args.pictures_next), 6)} picture(s) will be placed on this page right after this edit: keep the right part of the page free for them. No text, shape or table may lie right of x = ${Math.round(W * 0.53)}px between y = ${Math.round(H * 0.18)}px and y = ${Math.round(H * 0.84)}px (the title may run across the top above that). Put the text in the left part, in readable sizes, shortened only if it cannot fit.` }
    }
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
      const host = hostSteps()
      upgradeHint.value = !!host?.offerAdvanced
      // once per session: a frustrated user got the line and the paywall on two turns in a row (2026-11-11)
      if (advancedOffered) return ''
      advancedOffered = true
      host?.offerAdvanced?.()
      return host?.offerAdvanced ? (t('Requests like this work better on our advanced model, on paid plans.')) : ''
    }
    // "improve the design" with nothing concrete: the standard model moved boxes a few pixels and said it improved the
    // design (5 of 5 replays of a real user's request, 2026-11-11); the advanced model redesigned the page (2 of 2).
    // So free users are told plainly and offered what visibly works; paid users get the advanced model.
    if (args.redesign && !canAdvanced) {
      s.pending = { request: args.instruction, question: '', options: [t('Apply a new theme to the deck'), t('Make the text bigger')], offer: true, localize: true }
      save()
      offer() // ChatSlide's paywall, once per session
      // the router is told the rest still runs: "cute elements, delete slide 8, the group name on the cover" stopped
      // here, slide 8 and the name not done, and the user asked again (2026-11-17)
      // "write a conclusion, a better design and more text" got neither the conclusion nor the text (2026-11-18): the
      // text the user asked for is its own edit_page, without the design part
      memo = '(The redesign is not done and is never retried in another call. Go on only with the OTHER parts of the request, if any: text the user asked for on this page (a conclusion, more text) is a new edit_page with redesign false whose instruction asks for that text only, no layout or design words; pages to delete; pictures.)'
      return t('A redesign of this page needs the advanced model, on paid plans. Without it I can apply a new theme or make the text bigger.')
    }
    let upgradeLine = ''
    const pageId = slides.value[i].id
    const reply = await editSlide(i, args.instruction, (!!args.hard || !!args.redesign) && canAdvanced, !args.redesign).catch(async (e) => {
      // "the heading overlaps the body text" and the page model changed nothing (a paid user, twice, 2026-11-18): the
      // collision is measured and fixed here
      if (e instanceof NoChange && args.overlap === true && await shrinkOntoText(pageId)) return (lastEditTerse = t('Updated page {p0}.', { p0: i + 1 }))
      // not fixed by smaller letters (text in shapes, a long collision): the page model, told only to separate them
      if (e instanceof NoChange && args.overlap === true && textOnText(slides.value[i].elements).size) {
        return editSlide(i, 'Fix the layout only: no text may overlap other text or sit behind a picture. Make the text boxes fit (smaller font or narrower) or shrink/move the pictures. Keep every word.', canAdvanced)
      }
      // the user pasted the page's text and the page model answered it was already there, untouched: wrong in 6 of 6
      // replayed cases (slides 1, 8, 9 of three decks, 2026-11-11..16). One more pass, told it is not there
      if (e instanceof NoChange && text.length > 300 && !editedThisTurn.has(slides.value[i].id)) {
        return editSlide(i, `${args.instruction}\n\nThe user's message gives text for this page. It is NOT on the page yet: put it there now, in their words. Never answer that it is already there.`)
      }
      if (!(e instanceof NoChange) || editedThisTurn.has(slides.value[i].id)) throw e
      if (lastEditAdvanced || !(e as any).tooHard) throw e // the advanced model already had its try, or nothing was needed
      if (canAdvanced) return editSlide(i, args.instruction, true)
      const line = offer()
      throw line ? new NoChange(`${e.message} ${line}`) : e
    })
    if (args.hard && !canAdvanced) upgradeLine = offer()
    keepUnaskedSizes(i, was)
    if (JSON.stringify(slides.value[i].elements) === JSON.stringify(was)) throw new NoChange(t('Page {p0} was not changed.', { p0: i + 1 }))
    // text the edit made unreadable (it was readable before): one pass that keeps the change and the text readable;
    // still unreadable, the page goes back as it was and the user is asked, never left with invisible text
    const hadHidden = new Set(unreadableText(was as any[]))
    const hidden = () => unreadableText(slides.value[i].elements as any[]).filter(id => !hadHidden.has(id))
    if (hidden().length) {
      const names = hidden().map(id => `"${elText(slides.value[i].elements.find(e => e.id === id)).slice(0, 60)}"`).join(', ')
      await editSlide(i, `After the last edit this text cannot be read: ${names} now has (nearly) the same colour as the plain colour right behind it. Keep what the user asked for (${String(args.instruction).slice(0, 400)}), but every text must stay readable: give the box or panel behind that text a colour that contrasts with it (a new full-page panel must never hide the page's picture or sit under text of its own colour), or leave that text on its picture.`).catch(() => undefined)
      if (hidden().length) {
        slidesStore.updateSlide({ elements: was }, pageId)
        throw new NoChange(t('Page {p0} was not changed.', { p0: i + 1 }) + ' ' + t('That change would have left text in the same colour as what is behind it, so it could not be read. Which colour should it be?'))
      }
    }
    // a translation that left lines as they were (the small header stayed English): one pass naming them
    if (args.translate_to) {
      const left = slides.value[i].elements.map(elText).filter(t => t && before.includes(t) && /\p{L}{3}/u.test(t))
      if (left.length) await editSlide(i, `Translate these remaining lines into ${args.translate_to}, keeping their style: ${left.map(t => `"${t.slice(0, 160)}"`).join('; ')}`).catch(() => undefined)
      // every line still as it was: the page was in that language already ("it has to be in English" about an English
      // page was answered "translated page 7 into English", 2026-11-20)
      const now = slides.value[i].elements.map(elText).filter(Boolean)
      lastEditTerse = before.length && before.every(x => now.includes(x))
        ? t('Page {p0} was already in {p1}.', { p0: i + 1, p1: args.translate_to })
        : t('Translated page {p0} into {p1} (only this page).', { p0: i + 1, p1: args.translate_to })
      slidesStore.updateSlideIndex(i)
      return lastEditTerse
    }
    slidesStore.updateSlideIndex(i)
    // more line or paragraph spacing makes the text taller than the box the page model gave it: the lines ran over the
    // next heading ("el texto se ve muy pegado", replayed 3 of 3, 2026-11-21). Measured: a box whose text runs over
    // widens into the free room beside it, then grows down to the next element
    const gap = (e: any) => { const x = e.type === 'text' ? e : e.text || {}; return [x.lineHeight ?? 1.5, x.paragraphSpace ?? 5] }
    const spaced = slides.value[i].elements.filter((e: any) => { const o = was.find(w => w.id === e.id); return o && elText(e) && gap(e).some((v, k) => v > gap(o)[k]) }).map(e => e.id)
    // where it still has no room its letters go smaller, never under 14px (body text the user reads)
    const floor14 = Object.fromEntries(spaced.map((id) => { const e: any = slides.value[i].elements.find(x => x.id === id); return [id, [...String(e?.content ?? e?.text?.content ?? '').matchAll(/font-size:\s*([\d.]+)px/g)].map(m => Math.min(14, Number(m[1])))] }))
    const spacedFrom = Object.fromEntries(spaced.map(id => [id, gap(was.find(w => w.id === id))]))
    if (spaced.length) await fitOverflow([pageId], floor14, new Set(spaced), spacedFrom).catch(() => undefined)
    // longer text that now runs behind a picture ("the letters are behind the picture", 2026-11-10): one layout pass.
    // Measured after the page renders, since text boxes grow to fit their text then.
    await nextTick(); await new Promise(r => setTimeout(r, 300))
    const hadBefore = textOnPictures(was)
    const paidLayout = lastEditAdvanced // the cheaper model must not rework a layout the user paid for
    // text on text that the edit made, or that the user said overlaps and the edit left: the wrapping box's letters go
    // smaller first (exact, measured); the page model only for what is still there
    const fresh = () => [...textOnPictures(slides.value[i].elements)].some(k => !hadBefore.has(k))
    if (args.overlap === true || fresh()) await shrinkOntoText(slides.value[i].id)
    if (fresh() || (args.overlap === true && textOnText(slides.value[i].elements).size)) {
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
    // a name in quotes is written &quot;Helvetica Neue&quot; once the browser has re-written the style: it read as "quot"
    const sampleFont = sampleHtml.match(/font-family:\s*((?:&quot;|[^;"])+)/)?.[1].replace(/&quot;/g, '') || sample?.defaultFontName || sample?.text?.defaultFontName
    const font = String(sampleFont || args.font || '').replace(/[^\p{L}\p{N} ,_-]/gu, '').trim().slice(0, 60) // a font name, nothing CSS can parse further
    if (!font) fail(t('Which font?'))
    const setIn = (html: string) => html.replace(/font-family:\s*(?:&quot;|[^;"])*/g, `font-family: ${font}`)
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
    // "titles 28pt": the size PowerPoint shows after download (the export's own px-to-pt ratio); the scale factor gave 27px
    let to = Number(args.points) > 0 ? Math.round(Math.min(200, Number(args.points)) * 96 / 72 * slidesStore.viewportSize / 960) : 0
    // "same size as page N": that page's largest size in the same scope
    if (args.same_as_page > 0) {
      const sl = slides.value[args.same_as_page - 1]
      to = Math.max(0, ...(sl?.elements || []).filter((e: any) => inScope(sl, e)).flatMap((e: any) => sizes(html(e)).length ? sizes(html(e)) : [16]))
      if (!to) fail(t('That page has no text to match.'))
    }
    // back with no sizes kept (resized before 2026-11-20, or by hand): one "bigger" step down, on every page
    const k = args.back === true ? 0.77 : Math.min(3, Math.max(0.3, Number(args.scale) || 1.25))
    const resize = (h: string) => {
      const doc = new DOMParser().parseFromString(h, 'text/html')
      if (!h.includes('font-size')) doc.querySelectorAll('p').forEach(p => { (p as HTMLElement).style.fontSize = '16px' }) // PPTist's default
      // a step the eye can see: 1.1 of a 14px body is 15px, and "bigger", "more", "even bigger!" went on for five
      // messages (prod 2026-11-10/11, b1c, 415, e13). At least 4px up or 2px down per step
      const step = (n: number) => to || (k > 1 ? Math.max(n * k, n + 4) : k < 1 ? Math.min(n * k, n - 2) : n)
      return doc.body.innerHTML.replace(/font-size:\s*([\d.]+)px/g, (_, n) => `font-size: ${Math.max(8, Math.round(step(+n)))}px`)
    }
    // "back as it was": each box's sizes from before the chat resized it, wherever the chat did (the pages the router
    // lists are the last message's, not every page that was resized)
    const back = (h: string, was: number[]) => {
      let i = 0
      const doc = new DOMParser().parseFromString(h, 'text/html')
      if (!h.includes('font-size')) doc.querySelectorAll('p').forEach(p => { (p as HTMLElement).style.fontSize = '16px' })
      return doc.body.innerHTML.replace(/font-size:\s*([\d.]+)px/g, () => `font-size: ${was[Math.min(i++, was.length - 1)] ?? 16}px`)
    }
    // sizes kept and all put back already: asked again, nothing is left to restore (never one more step down)
    if (args.back === true && s.sizes && !Object.keys(s.sizes).length) fail(t('I have not resized any text here, so there is no earlier size to go back to.'))
    const restore = args.back === true && !!s.sizes
    const pagesDone = new Set<string>()
    let n = 0
    const was: Record<string, number[]> = {} // sizes before this step, of each box it changed
    const grew: Record<string, number[]> = {} // the boxes this step made bigger (scale, points or another page's size): never smaller than they were
    // ponytail: every resized box goes back, never only the pages named; a per-page restore when someone asks for one
    for (const id of args.back === true ? slides.value.map(x => x.id) : targets) {
      const sl = slides.value[pos(id)]
      for (const e of sl.elements as any[]) {
        if (restore ? !s.sizes?.[e.id] || !html(e) : !inScope(sl, e)) continue
        if (!restore) was[e.id] = sizes(html(e)).length ? sizes(html(e)) : [...html(e).matchAll(/<p[\s>]/g)].map(() => 16)
        const next = restore ? back(html(e), s.sizes![e.id]) : resize(html(e))
        if (!restore && sizes(next).some((v, i) => v > (was[e.id][i] || 16))) grew[e.id] = was[e.id]
        if (restore) s.sizes = Object.fromEntries(Object.entries(s.sizes!).filter(([k]) => k !== e.id))
        else if (args.back !== true && !s.sizes?.[e.id]) s.sizes = { ...s.sizes, [e.id]: sizes(html(e)) }
        if (e.type === 'text') slidesStore.updateElement({ id: e.id, props: { content: next }, slideId: id })
        else slidesStore.updateElement({ id: e.id, props: { text: { ...e.text, content: next } }, slideId: id })
        pagesDone.add(id)
        n++
      }
    }
    if (restore) {
      if (!n) fail(t('I have not resized any text here, so there is no earlier size to go back to.'))
      touched([...pagesDone])
      return t('The text in {p0} text boxes across {p1} page(s) is back to the size it had before I resized it.', { p0: n, p1: pagesDone.size })
    }
    if (!n) fail(t('There is no such text to resize there.'))
    // bigger text takes only the room its box has: x2.5 ran over the next text and off every slide, and the user then
    // asked twice to make it smaller (prod 2026-11-12, 112b)
    // only the boxes this step changed: a "titles bigger" left the body text as it was, byte for byte
    await fitOverflow([...pagesDone], grew, new Set(Object.keys(was)))
    touched([...pagesDone])
    // a box with no room to grow kept its size: said, never counted as resized
    const stuck = [...pagesDone].flatMap(id => (slides.value[pos(id)]?.elements || []) as any[])
      .filter(e => grew[e.id] && !sizes(html(e)).some((v, i) => v >= (grew[e.id][i] || 16) + 1)).length // under 1px is not a change anyone sees
    if (stuck === n) fail(t('{p0} text box(es) had no room to grow and kept their size.', { p0: stuck }))
    return t('Resized the text in {p0} text boxes across {p1} page(s).', { p0: n - stuck, p1: pagesDone.size })
      + (stuck ? ` ${t('{p0} text box(es) had no room to grow and kept their size.', { p0: stuck })}` : '')
  }
  if (action === 'set_text_style') { // direct, like set_font_size: "text black", "bold", "underline 4.5 billion" through edit_page changed nothing or the wrong thing (prod 2026-11-15..17)
    const targets = args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id)
    const marked = (e: any) => e.textType === 'title' || e.text?.type === 'title'
    const isTitle = (sl: any, e: any) => marked(e) || (!sl.elements.some(marked) && e === readingOrder(sl).find((x: any) => elText(x)))
    const inScope = (sl: any, e: any) => (e.type === 'text' || (e.type === 'shape' && e.text)) && elText(e)
      && (args.scope === 'titles' ? isTitle(sl, e) : args.scope === 'body' ? !isTitle(sl, e) : true)
    const color = /^#[0-9a-f]{6}$/i.test(String(args.color || '')) ? String(args.color) : ''
    const phrase = String(args.phrase || '').trim()
    const want = { bold: args.bold, italic: args.italic, underline: args.underline }
    if (!color && !Object.values(want).some(v => v === 'on' || v === 'off')) fail(t('Which colour or style?'))
    // a page that is one flat picture has no text to style: its painted words are underlined or marked on top by the
    // page edit, which is given where they are ("underline the last sentence" was answered "no such text", 2026-11-14)
    const flatOne = targets.length === 1 ? slides.value[pos(targets[0])] : undefined
    if (flatOne && paintedPicture(flatOne).src && !flatOne.elements.some((e: any) => inScope(flatOne, e))) {
      const how = [color && `colour ${color}`, ...Object.entries(want).filter(([, v]) => v !== 'same').map(([k, v]) => (v === 'on' ? k : `not ${k}`))].filter(Boolean).join(', ')
      return runCall(text, 'edit_page', { index: pos(targets[0]) + 1, instruction: `Style asked for (${how}) on ${phrase ? `the words "${phrase}"` : args.scope === 'titles' ? 'the title' : 'the text the user means'}. The user's message: "${text.slice(0, 400)}"`, page_named: true, translate_to: '', hard: false, redesign: false, new_content: false, whole_deck_option: '', user_said_whole_deck: false, source_pages: [], language: '' }, ids)
    }
    // a mark's element goes, its own size or colour stays (on a span)
    const unwrap = (root: Element, sel: string) => root.querySelectorAll(sel).forEach((x) => {
      const keep = x.getAttribute('style') ? Object.assign(root.ownerDocument.createElement('span'), { style: x.getAttribute('style') }) : null
      if (keep) { keep.append(...Array.from(x.childNodes)); x.replaceWith(keep) }
      else x.replaceWith(...Array.from(x.childNodes))
    })
    const styled = (root: HTMLElement) => [root, ...Array.from(root.querySelectorAll<HTMLElement>('[style]'))]
    const noUnderline = (x: HTMLElement) => ['text-decoration', 'text-decoration-line'].forEach((k) => { // a strikethrough stays
      const v = x.style.getPropertyValue(k)
      if (!v.includes('underline')) return
      if (v.includes('line-through')) x.style.setProperty(k, 'line-through')
      else x.style.removeProperty(k)
    })
    const style = (root: HTMLElement, box: HTMLElement) => { // root: where the old marks are cleared; box: the span that carries the new ones
      if (color) { styled(root).forEach(x => x.style.removeProperty('color')); box.style.color = color }
      if (want.bold !== 'same') { unwrap(root, 'strong, b'); styled(root).forEach(x => x.style.removeProperty('font-weight')) }
      if (want.italic !== 'same') { unwrap(root, 'em, i'); styled(root).forEach(x => x.style.removeProperty('font-style')) }
      if (want.underline !== 'same') { unwrap(root, 'u'); styled(root).forEach(noUnderline) }
      if (want.underline === 'on') box.style.textDecoration = 'underline'
      let inner = box.innerHTML
      if (want.bold === 'on') inner = `<strong>${inner}</strong>`
      if (want.italic === 'on') inner = `<em>${inner}</em>`
      box.innerHTML = inner
    }
    // a style the browser re-wrote quotes its font names as &quot;: single quotes keep "font-family: &quot" out of
    // what reads the style as text (the title's font became "&quot" after "chữ nên đậm", 2026-11-20)
    // (an apostrophe in the name, "Rock'n'Roll One", is escaped so the quoted name stays one name)
    const written = (doc: Document) => doc.body.innerHTML.replace(/&quot;([^&]*?)&quot;/g, (_, n: string) => `'${n.replace(/'/g, "\\'")}'`)
    const restyle = (h: string) => {
      const doc = new DOMParser().parseFromString(h, 'text/html')
      if (!phrase) {
        const blocks = doc.body.querySelectorAll<HTMLElement>('p, li')
        ;(blocks.length ? Array.from(blocks).filter(b => !b.querySelector('p, li')) : [doc.body]).forEach(b => {
          const box = doc.createElement('span')
          box.append(...Array.from(b.childNodes))
          b.append(box)
          style(b, box)
        })
        return written(doc)
      }
      // the named words, inside one run of text; split across two runs, or to be un-bolded inside a bold run, they go to the page edit below
      const inside = (x: Element | null) => [want.bold === 'off' && 'strong, b, [style*="font-weight"]', want.italic === 'off' && 'em, i, [style*="font-style"]', want.underline === 'off' && 'u, [style*="underline"]']
        .some(sel => sel && x?.closest(sel))
      const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT)
      for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
        const at = node.data.toLowerCase().indexOf(phrase.toLowerCase())
        if (at < 0) continue
        if (inside(node.parentElement)) return h
        const hit = node.splitText(at)
        hit.splitText(phrase.length)
        const box = doc.createElement('span')
        hit.replaceWith(box)
        box.append(hit)
        style(box, box)
        return written(doc)
      }
      return h
    }
    // the colours a box's letters are drawn in: each run's own, else the box's
    const letters = (e: any) => {
      const doc = new DOMParser().parseFromString(lookHtml(e), 'text/html'), w = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT), out = new Set<string>()
      for (let node = w.nextNode(); node; node = w.nextNode()) {
        if (!(node as Text).data.trim()) continue
        let x = node.parentElement
        while (x && !x.style.color) x = x.parentElement
        out.add(tinycolor(x?.style.color || (e.type === 'text' ? e.defaultColor : e.text.defaultColor) || '#333').toHex())
      }
      return out
    }
    const already: [string, any][] = [] // boxes whose letters were all that colour before: recolouring them shows nothing
    let n = 0
    for (const id of targets) {
      const sl = slides.value[pos(id)]
      for (const e of sl.elements as any[]) {
        if (!inScope(sl, e)) continue
        const was = String(e.type === 'text' ? e.content : e.text.content), now = restyle(was)
        if (now === was) continue
        // (the box that holds the named words too: recolouring words that already are that colour shows nothing either)
        if (color && [...letters(e)].every(c => c === tinycolor(color).toHex())) already.push([id, e])
        const whole = color && !phrase ? { defaultColor: color } : {}
        if (e.type === 'text') slidesStore.updateElement({ id: e.id, props: { ...whole, content: now }, slideId: id })
        else slidesStore.updateElement({ id: e.id, props: { text: { ...e.text, ...whole, content: now } }, slideId: id })
        n++
      }
    }
    // The named words are not one run of text (they span boxes: "the part with the title and my name whiter") and the
    // text there already is that colour: the same band as below, never the page edit's "it is already pure white"
    // (1 of 4 replays of the user's four messages, 2026-11-20)
    let spanned = 0
    if (!n && phrase && targets.length === 1 && color && !Object.values(want).some(v => v === 'on' || v === 'off')) {
      const sl = slides.value[pos(targets[0])]
      // only the boxes the named words are in: the phrase is copied from the page's text, so a box belongs to it when
      // its text is part of the phrase or holds it; a white footer elsewhere is not the title the user named
      const norm = (s: string) => s.replace(/\s+/g, ' ').trim().toLowerCase()
      const named = (e: any) => {
        const box = norm(elText(e)), asked = norm(phrase)
        return box.length >= 4 && (asked.includes(box) || box.includes(asked))
      }
      const boxes = (sl.elements as any[]).filter(e => inScope(sl, e) && named(e))
      // every named box already that colour, or the page edit recolours them
      const same = boxes.every(e => [...letters(e)].every(c => c === tinycolor(color).toHex())) ? boxes : []
      if (same.length) { already.push(...same.map((e): [string, any] => [targets[0], e])); spanned = n = same.length }
    }
    const byPageEdit = () => { // not one plain run of text: the page edit restyles those words
      const how = [color && `colour ${color}`, ...Object.entries(want).filter(([, v]) => v !== 'same').map(([k, v]) => (v === 'on' ? k : `not ${k}`))].filter(Boolean).join(', ')
      return runCall(text, 'edit_page', { index: pos(targets[0]) + 1, instruction: `Restyle only the words "${phrase}": ${how}. Change nothing else on the page.`, page_named: true, translate_to: '', hard: false, redesign: false, new_content: false, whole_deck_option: '', user_said_whole_deck: false, source_pages: [], language: '' }, ids)
    }
    if (!n && phrase && targets.length === 1) return byPageEdit()
    if (!n) fail(phrase ? t('I could not find “{p0}” in the text there.', { p0: phrase.slice(0, 60) }) : t('There is no such text to style there.'))
    if (!spanned) touched(targets)
    // Text that already is the colour asked for and sits on a picture is hard to read because of the picture: "the
    // cover's text is unreadable, change its colour", "whiter", "white I said" on white letters over a pale photo was
    // answered four times with "it is already pure white" (2026-11-15), and then with "restyled" and no visible change.
    // It gets a see-through band of the opposite tone behind it, and the reply says so
    if (already.length === n && !Object.values(want).some(v => v === 'on' || v === 'off')) {
      // on a picture: the nearest thing under at least 30% of the box is a picture that shows, not one that is hidden
      // or lies under an opaque shape; with nothing there, the slide's background picture
      const onPicture = ([id, e]: [string, any]) => {
        const sl = slides.value[pos(id)]
        if (!bareText(e) || (e.type === 'text' && e.fill)) return false
        const under = (x: any) => Math.max(0, Math.min(e.left + e.width, x.left + x.width) - Math.max(e.left, x.left)) * Math.max(0, Math.min(e.top + e.height, x.top + x.height) - Math.max(e.top, x.top)) >= 0.3 * e.width * e.height
        const solid = (x: any) => x.type === 'shape' && (x.opacity ?? 1) === 1 && (!!x.gradient || !!x.pattern || (!!x.fill && tinycolor(x.fill).getAlpha() === 1))
        const behind = sl.elements.slice(0, sl.elements.findIndex(x => x.id === e.id)).reverse()
          .find((x: any) => under(x) && ((x.type === 'image' && x.filters?.opacity !== '0%') || solid(x)))
        return behind ? behind.type === 'image' : sl.background?.type === 'image'
      }
      const backed = already.filter(onPicture)
      if (backed.length) {
        const light = tinycolor(color).isLight()
        for (const [id, e] of backed) slidesStore.updateElement({ id: e.id, props: { fill: light ? '#00000080' : '#FFFFFFB3' }, slideId: id })
        args.backed = true // not a look to repeat on the other pages
        touched(targets)
        return light ? t('The text there already is that colour, so I put a dark see-through band behind it: it now stands out on the picture.')
          : t('The text there already is that colour, so I put a light see-through band behind it: it now stands out on the picture.')
      }
    }
    if (spanned) return byPageEdit() // not on a picture: nothing was changed here, the page edit looks at the words
    return t('Restyled the text in {p0} text boxes across {p1} page(s).', { p0: n, p1: targets.length })
  }
  // "Do the same on the other pages" after a size or style step on a page the user had already shaped in earlier
  // messages: the other pages take that page's text look as it is now, not the last step again. A user centred,
  // bolded and enlarged page 6 over three messages, said "smaller", tapped the button, and the nine other pages only
  // shrank (2026-11-17). Each page's main title and body text go to that page's sizes, the rest in proportion, never
  // past what the box and the free space under it hold; bold and alignment are copied when the whole of it has them
  if (action === 'match_look') {
    const src = slides.value[pos(idOf(args.from))]
    const targets = nums(args.pages).filter((id: string) => id !== src.id)
    const { width: W, height: H } = slideSize()
    const want = textLook(src)
    let n = 0
    for (const id of targets) {
      const sl = slides.value[pos(id)]
      ;[true, false].forEach((title, r) => {
        const look = want[r], els = lookTexts(sl, title)
        if (!look || !els.length) return
        // how far a box's text can grow: the room down to the next thing under it (as fitOverflow reads it), at about
        // half an em a letter and a line and a half of height
        const fits = (e: any) => {
          const under = (sl.elements as any[]).filter(o => o.id !== e.id && o.top > e.top + 4 && o.left < e.left + e.width && o.left + o.width > e.left
            && !(o.type === 'shape' && !elText(o) && o.width * o.height > W * H * 0.2) && !(o.type === 'image' && isBackdrop(o)))
          const room = Math.max(e.height, Math.min(H - 8, ...under.map(o => o.top - 6)) - e.top)
          const letters = [...elText(e)].reduce((a, ch) => a + (ch > '\u2e7f' ? 2 : 1), 0)
          return Math.min(room / 1.5, Math.sqrt(Math.max(40, e.width - 20) * room / (0.9 * letters)))
        }
        // one factor for the page's title, one for its body, so its own text keeps its proportions: up to the other
        // page's size, as far as every box holds it. A page with a title only (a cover, a closing page) keeps its sizes
        const k = look.size / lookSize(lookMain(els))
        const f = !lookTexts(sl, false).length ? 1 : k > 1 ? Math.min(k, ...els.map(e => Math.max(1, fits(e) / lookSize(e)))) : k
        for (const e of els) {
          const doc = new DOMParser().parseFromString(lookHtml(e), 'text/html')
          if (!lookHtml(e).includes('font-size')) doc.querySelectorAll('p').forEach((p) => { (p as HTMLElement).style.fontSize = '16px' }) // as set_font_size does
          if (look.align) doc.body.querySelectorAll<HTMLElement>('p, li').forEach((p) => { p.style.textAlign = look.align })
          const content = doc.body.innerHTML.replace(/font-size:\s*([\d.]+)px/g, (_, v) => `font-size: ${Math.max(9, Math.round(+v * f))}px`)
          if (content === lookHtml(e)) continue
          if (e.type === 'text') slidesStore.updateElement({ id: e.id, props: { content }, slideId: id })
          else slidesStore.updateElement({ id: e.id, props: { text: { ...e.text, content } }, slideId: id })
          n++
        }
      })
    }
    for (const [r, scope] of ['titles', 'body'].entries()) {
      if (want[r]?.bold) await runCall(text, 'set_text_style', { pages: targets.map((id: string) => pos(id) + 1), scope, phrase: '', color: '', bold: 'on', italic: 'same', underline: 'same' }, ids).then(() => n++).catch(() => undefined)
    }
    // the other pages already look like it: the steps queued after this one (a font, a colour) still run
    if (!n && args.more) return ''
    if (!n) fail(t('There is no such text to style there.'))
    await fitOverflow().catch(() => undefined) // a box whose text still runs over grows into the free space or shrinks
    touched(targets)
    return t("Gave {p1} page(s) the text look of page {p0}: its sizes as far as the text fits, its bold and its alignment.", { p0: pos(src.id) + 1, p1: targets.length })
  }
  // The user's own text, page by page: each page gets only its part. Given the whole text, each page picked a part
  // itself and slide 3 took slide 2's (replayed 2026-11-18, 4 decks); one page edit at a time stopped after slide 1
  if (action === 'place_texts') {
    let items = (Array.isArray(args.items) ? args.items : []).map((x: any) => ({ page: Number(x?.page), text: String(x?.text || '').trim() }))
      .filter((x: { page: number, text: string }) => x.text && x.page >= 1)
    // the router copies the parts itself and left out the last slide in 1 run in 4 ("slide 11 is not in the deck"): the
    // user's own long text is read once more, and its full set is used when it has more parts
    const source = [text, ...s.history.filter(h => h.role === 'user' && h.text.length > 600).slice(-3).reverse().map(h => h.text), s.brief?.text || '']
      .filter(x => x.length > 600 && items.some((it: { text: string }) => x.includes(it.text.slice(0, 40))))[0]
    if (source && source !== text) {
      // what this message asks for from that text: "only slide 2" stays slide 2; a full set is taken only when it holds
      // every page the router chose (Codex review: a deliberate subset was widened)
      const full = await numberedParts(source, text)
      const chosen = new Set(items.map((it: { page: number }) => it.page))
      if (full.length > items.length && [...chosen].every(n => full.some(f => f.page === n))) items = full
    }
    if (!items.length) fail(t('Which text goes on which page?'))
    const ids0 = slides.value.map(x => x.id)
    const onPages = items.filter((x: { page: number }) => x.page <= ids0.length)
    const past = items.filter((x: { page: number }) => x.page > ids0.length)
    const same: number[] = [], failed: number[] = []
    let next = 0, why = ''
    const worker = async () => {
      while (next < onPages.length) {
        const x = onPages[next++]
        try {
          const i = pos(ids0[x.page - 1])
          await editSlide(i, `Put this text on this page, exactly as written, replacing the page's current text (its first line is the title when it reads as one). Every sentence of it goes on the page, none dropped or shortened, and none of the page's old text stays (the last page's "thank you" included). Keep the pictures and colours; split it over the page's text boxes, and add or resize text boxes when it does not fit. Never refuse for lack of room:\n${fence(x.text)}`)
          // their own lines, copied: a list layout dropped the opening sentence in 2 of 3 runs (2026-11-19). A line that
          // is not on the page verbatim gets one more pass naming it
          // decoded text ("&amp;" is "&"), and both ends of each line: a paragraph that kept its opening but lost its
          // later sentences passed a start-only check (Codex review)
          // compared without any spaces (inline tags add some), sentence by sentence (a middle sentence was lost while
          // both ends of its line stayed; Codex verify)
          const flat = (v: string) => (new DOMParser().parseFromString(v, 'text/html').documentElement.textContent || '').replace(/\s+/g, '').toLowerCase()
          // every sentence, short ones and a one-word heading included; what is missing is the sentence, never its whole line
          const sentences = String(x.text).split('\n').flatMap((l: string) => l.split(/(?<=[.!?…。！？])\s+/)).map((l: string) => l.trim()).filter((l: string) => flat(l).length >= 2)
          const has = (page: string, l: string) => page.includes(flat(l))
          const onPage = flat(slides.value[i].elements.map(elText).join(' '))
          const missing = sentences.filter((l: string) => !has(onPage, l))
          if (missing.length) await editSlide(i, `These lines of the user's text are missing from this page. Add each one, word for word, where it belongs (add or resize text boxes when needed); change nothing else:\n${fence(missing.join('\n'))}`).catch(() => undefined)
          // still not there (a list layout keeps dropping an opening paragraph, 2026-11-19): their words go at the top
          // of the page's main text box, in its own style, and the box is fitted
          const now = flat(slides.value[i].elements.map(elText).join(' '))
          const still = missing.filter((l: string) => !has(now, l))
          const sl = slides.value[i]
          // a text box or a shape holding text (list layouts keep their body in shapes)
          const box = (sl.elements as any[]).filter(e => (e.type === 'text' || (e.type === 'shape' && e.text)) && elText(e) && e.textType !== 'title' && e.text?.type !== 'title')
            .sort((a, b) => b.width * b.height - a.width * a.height)[0]
          if (still.length && box) {
            const html = String(box.type === 'text' ? box.content : box.text.content)
            const open = html.match(/<p[^>]*>/)?.[0] || '<p>'
            const content = still.map((l: string) => `${open}${l.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</p>`).join('') + html
            slidesStore.updateElement({ id: box.id, props: box.type === 'text' ? { content } : { text: { ...box.text, content } }, slideId: sl.id })
            await fitText(sl.id, [box.id], { [box.id]: box.height })
          }
        }
        catch (e) {
          if (e instanceof NoChange) { same.push(x.page); why ||= e.message } // the page's own reason: "one flat picture" is not "already there"
          else if (e instanceof StepFailed) failed.push(x.page)
          else throw e
        }
      }
    }
    const settled = await Promise.allSettled(Array.from({ length: Math.min(5, onPages.length) }, worker))
    const err = settled.find((x): x is PromiseRejectedResult => x.status === 'rejected')
    if (err) throw err.reason
    touched(onPages.filter((x: { page: number }) => !same.includes(x.page) && !failed.includes(x.page)).map((x: { page: number }) => ids0[x.page - 1]))
    const lines = [t('Placed your text on {p0} page(s).', { p0: onPages.length - same.length - failed.length })]
    if (same.length) lines.push(`${t('Page(s) {p0} were not changed.', { p0: same.join(', ') })} ${why}`.trim())
    if (failed.length) lines.push(t(' Page(s) {p0} failed.', { p0: failed.join(', ') }).trim())
    // the deck grows with each new page: its page numbers are read from it now
    for (const x of past) lines.push(await runCall(text, 'add_page', { instruction: `Exactly this text, as written, every sentence, nothing added:\n${x.text}`, count: 1, after: slides.value.length, layout_from: slides.value.length, after_named: true, message_new_pages: past.length, topics: [], hard: false, language: '', from_users_file: false }, slides.value.map(sl => sl.id)))
    await fitQueue
    return lines.join('\n')
  }
  if (action === 'edit_all_pages') { // 5 pages in flight, the next starts as soon as one ends; merged by slide id
    const all = slides.value.map(x => x.id)
    const todo: string[] = args.pages?.length ? nums(args.pages) : all // a whole-presentation page edit leaves out the cover and the closing page
    const viewing = slides.value[slideIndex.value]?.id // fitting shows each page; come back to this one
    let done = 0, next = 0, stop = false, landed = 0
    const same: number[] = [], failed: number[] = []
    let noChangeWhy = ''
    // "add references of the articles to each slide so it matches where you got the info": each page matches its own
    // content against the files; given only a list of citations, pages guessed and wrote "no source can be linked
    // with certainty" onto the slides (PRO user cancelled, "inaccurate", 2026-11-14)
    const files = args.cite_files ? s.docs.filter(d => d.text.trim()) : []
    // a file read for this request is its material on every page: the router's instruction said "as in the document"
    // and each page answered that no document was attached (a teacher asked twice, 2026-11-15)
    const read = files.length ? [] : s.docs.filter(d => readThisTurn.has(d.name) && d.text.trim())
    const given = (docs: typeof s.docs) => docs.map(d => `<file name="${d.name.replace(/["<>]/g, '')}">\n${fence(d.text.slice(0, Math.floor(40000 / docs.length)))}\n</file>`).join('\n')
    const instruction: string = files.length ? `${args.instruction}\n\nThe user's files. Cite on this page only the ones its content comes from (author, title, journal, year when the file gives them); a page none of them informed gets no reference:\n${given(files)}` : args.instruction
    // a long message is material too (a pasted outline "fill in exactly these infos"): the router's instruction only
    // describes it, and the pages wrote their own version of it (replayed conversation, 2026-11-16)
    // "insert the text I wrote on slides 1-10": that text was an earlier message or the request the deck was made from;
    // given only the message, slide 1 was edited and called done (2026-11-11, replayed 2026-11-18)
    const own = args.users_text ? [...s.history.filter(h => h.role === 'user' && h.text.length > 300).slice(-3).map(h => h.text), s.brief?.text || ''] : []
    const said = [s.pending?.request, text, ...own].filter((x, k, arr) => x && arr.indexOf(x) === k).join('\n\n')
    const forPage = (id: string) => instruction
      + (read.length ? `\n\nThis is page ${pos(id) + 1} of ${all.length}. The user's file, read for this request: take from it what belongs on THIS page (its own slide or section when the file is laid out by slide), as the request says; "no changes" means its words exactly:\n${given(read)}` : '')
      + ((said.length > 600 || args.users_text) && !files.length ? `\n\nThis is page ${pos(id) + 1} of ${all.length}. The user's own message, in full: where it gives text for this page, use that part word for word:\n${fence(said.slice(0, 12000))}` : '')
    const worker = async () => {
      while (next < todo.length && !stop) {
        const id = todo[next++]
        try {
          await editSlide(pos(id), forPage(id))
          landed++
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
        busy.value = t('Working on the pages: {p0} of {p1}…', { p0: ++done, p1: todo.length })
      }
    }
    // wait for every worker (an edit in flight still lands), then report the first unexpected error
    const settled = await Promise.allSettled(Array.from({ length: Math.min(5, all.length) }, worker))
    const err = settled.find((x): x is PromiseRejectedResult => x.status === 'rejected')
    if (err) {
      // the pages already rewritten are kept (the turn's crash handling keeps lastGoodSlides): a crash on page 9 of 12
      // never takes back the eight done
      if (landed) lastGoodSlides = JSON.parse(JSON.stringify(slides.value))
      throw err.reason
    }
    // the request applies to every page, and some came back unchanged: one more try for those, naming that (not when
    // citing the files: a page none of them informed rightly stays as it was)
    if (args.every_page && !files.length) {
      const again = same.splice(0).map(p => all[p - 1])
      for (let k = 0; k < again.length; k += 5) {
        // every edit in flight lands before an error is reported: a rollback must not race a page still being written
        const retried = await Promise.allSettled(again.slice(k, k + 5).map(id => editSlide(pos(id), `${forPage(id)}\nThis applies to every slide of the deck, this one included: make the change here too.`).catch((e) => {
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
    const changedNow = todo.length - same.length - failed.length
    if (!changedNow && earlier.length && !untouched.length && !failed.length) return t('No further change: the earlier step already covered these pages.')
    const ok = changedNow + earlier.length
    const list = (xs: number[]) => xs.sort((a, b) => a - b).join(', ')
    const what = String(args.instruction || '').replace(/^Write all slide text in [^\n]*\n/, '').slice(0, 60) // the user's words, not our language note
    const why = (failed.length ? (t(' Page(s) {p0} failed.', { p0: list(failed) })) : '')
      + (untouched.length && ok ? (t(' Page(s) {p0} had nothing for that.', { p0: list(untouched) })) : '')
    void earlier
    if (!ok && !failed.length) return t('No page was changed. {p0}', { p0: noChangeWhy })
    if (!ok) fail((t('I could not make that change on any page: "{p0}".', { p0: what })) + why)
    // the pages by number: "Обновлены слайды: 10" (a count) was read back by the router as "only slide 10 was updated;
    // slides 1-9 were not" (2026-11-12), and a count never says which pages were left out
    const updated = todo.map(id => pos(id) + 1).filter(p => p > 0 && !untouched.includes(p) && !failed.includes(p))
    return (t('Updated pages {p0}.', { p0: spans(updated) })) + why
  }
  if (action === 'change_image') {
    const id = idOf(args.index)
    const sl = slides.value[pos(id)]
    const content = pictures(sl).filter(e => !isBackdrop(e))
    const said = saidBackground(text, args.said_background)
    let image = pictureNo(sl, args.image)
    // pictures the user said to keep are never the one replaced: "keep the nurse picture, change the other image" replaced
    // the nurse, three messages in a row (2026-11-16). The other one is the next picture they did not keep, else the
    // background picture
    // numbered as the image is (pictureNo): on a page with only a background picture, its 0 and 1 are the same picture
    const kept = new Set((Array.isArray(args.keep_pictures) ? args.keep_pictures : []).map((n: unknown) => pictureNo(sl, Number(n))))
    const free = (k: number) => !kept.has(k) && !replacedThisTurn.has(`${id}:${k}`)
    // "all the pictures" never reaches the background picture
    const other = () => content.findIndex((_, k) => free(k + 1)) + 1 || (firstBackdrop(sl) && content.length && free(0) && !args.all_pictures ? 0 : -1)
    const askWhich = (background: boolean) => {
      const { width: W, height: H } = slideSize()
      const where = (e: any) => [e.top + e.height / 2 < H / 3 ? t('top') : e.top + e.height / 2 > (2 * H) / 3 ? t('bottom') : t('middle'),
        e.left + e.width / 2 < W / 3 ? t('left') : e.left + e.width / 2 > (2 * W) / 3 ? t('right') : t('center')].join(', ')
      return runCall(text, 'ask_user', {
        question: t('Which picture on page {p0} should I change?', { p0: pos(id) + 1 }),
        options: [...(background ? [t('The background picture')] : []), ...content.slice(0, 3).map((e, k) => `${t('Picture {p0}', { p0: k + 1 })} (${where(e)})`), ...(background ? [] : [t('All of them')])],
      }, ids)
    }
    // the router named the background picture (image 0) but the user's words do not say background. It stands when the
    // words pick it (picture_named) or the previous message changed the background ("another one, like a city slum");
    // otherwise the one content picture not kept is meant ("change the picture"; slide 5's pictures got its background
    // replaced, 2026-11-10), the background when that one is kept ("keep the nurse, change the other image" replaced the
    // nurse, 2026-11-16), and with several content pictures and nothing saying which, the user is asked
    if (image === 0 && !said && !lastChangedBackground() && content.length && !args.picture_named && !(args.earlier > 0)) {
      const left = content.map((_, k) => k + 1).filter(free)
      if (left.length > 1 && !args.all_pictures) return askedThisTurn ? '' : askWhich(true)
      image = left[0] ?? (content.every((_, k) => kept.has(k + 1)) ? 0 : 1)
    }
    if (kept.has(image)) {
      image = other()
      if (image < 0) return t('You asked to keep every picture on page {p0}, so none was replaced.', { p0: pos(id) + 1 })
    }
    // "all the pictures" means the content pictures, never the background picture; none left: nothing to replace
    if (args.all_pictures && image === 0 && content.length) {
      image = content.findIndex((_, k) => free(k + 1)) + 1
      if (!image) return kept.size ? t('You asked to keep every picture on page {p0}, so none was replaced.', { p0: pos(id) + 1 }) : ''
    }
    // a page with only a background picture: "change the image / the cover photo" means that picture (the eval's
    // "Change the image on the slide" got a small extra picture, 2026-11-12); only a request to ADD one keeps it
    // ("insert real photos of X" swapped the background)
    // the page is one flat picture: the pictures the user sees are inside it, and replacing it would replace the page
    if (isFlat(sl) && !args.wants_additional_picture) fail(t('Page {p0} is one flat picture: its text and pictures are all inside that single image, so I can\'t change a picture in it. Nothing was changed. I can add a new picture on top of it.', { p0: pos(id) + 1 }))
    // an earlier picture of this image, back as it was: nothing is searched
    if (args.earlier > 0) {
      const el = image === 0 ? firstBackdrop(sl) : pictures(sl)[Math.max(0, image - 1)]
      const was = el && s.was?.[el.id]?.[args.earlier - 1]
      if (el && was) {
        // the copies above it and the slide background that were replaced with it (they show its picture now) go back
        // to what each of them held then, or they hide the picture put back
        const now = el.src
        const copies = isBackdrop(el) ? sl.elements.filter((e: any) => e.type === 'image' && e.id !== el.id && isBackdrop(e) && e.src === now && was.others?.[e.id]) : []
        const bgBack = isBackdrop(el) && was.background && sl.background?.type === 'image' && sl.background.image?.src === now
        keepPic(el, sl)
        slidesStore.updateElement({ id: el.id, props: { src: was.src, clip: was.clip } as any, slideId: id })
        s.pics = { ...s.pics, [el.id]: was.shows }
        for (const e of copies) {
          keepPic(e)
          slidesStore.updateElement({ id: e.id, props: { ...was.others![e.id] } as any, slideId: id })
        }
        if (bgBack) slidesStore.updateSlide({ background: was.background }, id)
        if (isBackdrop(el)) showBackdrop(id, el.id)
        touched([id])
        replacedThisTurn.add(`${id}:${image}`)
        slidesStore.updateSlideIndex(pos(id))
        return t('Put the earlier picture back on page {p0}.', { p0: pos(id) + 1 })
      }
      // not kept (replaced before the editor kept them, or by hand): a search for what it showed, as before; with
      // no words for that, said, never a guess passed off as it
      if (!String(args.query || '').trim()) {
        askedThisTurn = true
        return t('I no longer have the earlier picture of page {p0}. Attach it and I will put it back, or tell me what it shows and I will find one like it.', { p0: pos(id) + 1 })
      }
    }
    // "replace the pictures on all slides" has nothing to replace on a page without one: no new picture there (eval
    // 2026-11-15 added pictures to the two pages that had none)
    if (args.every_page && !said && !content.length && !args.wants_additional_picture) return t('Page {p0} has no picture to replace.', { p0: pos(id) + 1 })
    if (!said && !content.length && (args.wants_additional_picture || !firstBackdrop(sl))) {
      return runCall(text, 'add_image', { index: args.index, query: args.query, source: args.source, page_named: true, replaces_existing: false }, ids)
    }
    // "the pictures" / "all of them": every content picture on the page, in one step (the router asked for one per
    // picture replaced only the first of slide 5's three, 2026-11-13)
    // never a logo or an icon: "the pictures" swapped the company logo for a stock photo (review of 49a715d3)
    const logos = logoSpots(), { width: SW, height: SH } = slideSize()
    const photos = content.map((e, k) => ({ e, k: k + 1 })).filter(({ e }) => !logos.has(imgSpot(e)) && e.width * e.height >= 0.005 * SW * SH)
    if (args.all_pictures && photos.length > 1) {
      const left = photos.map(({ k }) => k).filter(free)
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
    if (image !== 0 && content.filter((_, k) => !kept.has(k + 1)).length > 1 && !args.picture_named) return askWhich(false)
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
    // a picture of the page as its background is about one page: the one named, else the one the user is on
    const onePage = Number(args.page_picture) > 0 && !args.pages?.length ? [slides.value.find(x => x.id === currentAtSend)?.id || slides.value[slideIndex.value].id] : null
    const named = onePage || (args.pages?.length ? nums(args.pages) : slides.value.map(x => x.id))
    const flat = named.filter(id => isFlat(slides.value[pos(id)]))
    const flatNote = flat.length ? ' ' + t('Page(s) {p0} are one flat picture each: their background is part of that picture, so they were not changed.', { p0: flat.map(id => pos(id) + 1).join(', ') }) : ''
    const targets = named.filter(id => !flat.includes(id))
    if (!targets.length) fail(flatNote.trim())
    let background: any = { type: 'solid', color: args.color }
    // A picture already on the page becomes its background: the same picture, behind everything, covering the page.
    // "Put this photo as the background of page 1" and "make the photo you put on slide 1 the cover" were asked
    // "which background colour?" twice, then given another stock photo (a paying user, 2026-11-15)
    const own = Number(args.page_picture) > 0 ? (() => {
      const sl = slides.value[pos(targets[0])]
      const list = pictures(sl)
      return { sl, pic: (list[Number(args.page_picture) - 1] || (list.length === 1 ? list[0] : undefined)) as any }
    })() : null
    if (own) {
      const { sl, pic } = own
      if (!pic) fail(t('That page has no picture.'))
      if (isBackdrop(pic)) fail(t('That picture already is the background of page {p0}: it covers the whole page.', { p0: pos(sl.id) + 1 }))
      for (const e of sl.elements.filter((x: any) => x.type === 'image' && isBackdrop(x))) keepPic(e, sl)
      // the page's full-slide pictures (a template draws its background as one) take the picture, or it would lie
      // under them unseen; the picture itself leaves its place
      slidesStore.updateSlide({ background: { type: 'image', image: { src: pic.src, size: 'cover' } }, elements: sl.elements.filter(e => e.id !== pic.id) }, sl.id)
      for (const e of slides.value[pos(sl.id)].elements.filter((x: any) => x.type === 'image' && isBackdrop(x))) {
        slidesStore.updateElement({ id: e.id, props: { src: pic.src, clip: undefined } as any, slideId: sl.id })
        if (s.pics?.[pic.id]) s.pics = { ...s.pics, [e.id]: s.pics[pic.id] }
        showBackdrop(sl.id, e.id)
      }
      replacedThisTurn.add(`${sl.id}:${pictureNo(slides.value[pos(sl.id)], 0)}`)
      touched([sl.id])
      slidesStore.updateSlideIndex(pos(sl.id))
      return t('The picture on page {p0} now is its background: it covers the whole page, behind the text.', { p0: pos(sl.id) + 1 })
    }
    if (args.photo_query) {
      busy.value = t('Finding a picture…')
      // never the photo already there ("change the background", then "again", gave the same one, 2026-11-12)
      const src = await findPicture(args.photo_query, 'photo', true, [...usedBackgrounds, ...slides.value.map(x => x.background?.image?.src).filter(Boolean) as string[]], true)
      if (!src) fail(t('Could not get a background picture.'))
      usedBackgrounds.add(src)
      background = { type: 'image', image: { src, size: 'cover' } }
    }
    else if (!/^#[0-9a-fA-F]{6}$/.test(args.color || '')) fail(t('Which background colour?'))
    // each full-slide picture keeps what it and the rest of the page held, before any of it changes
    if (args.photo_query) for (const id of targets) for (const e of slides.value[pos(id)].elements.filter((x: any) => x.type === 'image' && isBackdrop(x))) keepPic(e, slides.value[pos(id)])
    for (const id of targets) slidesStore.updateSlide({ background }, id)
    // a full-slide picture covers the background fill: a new background PHOTO replaces that picture, or nothing shows
    if (args.photo_query) {
      for (const id of targets) {
        for (const e of slides.value[pos(id)].elements.filter((x: any) => x.type === 'image' && isBackdrop(x))) {
          slidesStore.updateElement({ id: e.id, props: { src: background.image.src } as any, slideId: id })
          s.pics = { ...s.pics, [e.id]: String(args.photo_query) }
          showBackdrop(id, e.id)
        }
        // the background picture is replaced: a change_image on image 0 after it in this message is that same picture,
        // never moved onto a content picture ("find a picture of homeless people" then replaced the nurse, 2026-11-16)
        replacedThisTurn.add(`${id}:${pictureNo(slides.value[pos(id)], 0)}`)
      }
    }
    // same for a colour: the picture has to go, or "make it royal blue" changes nothing anyone can see
    if (!args.photo_query) for (const id of targets) fadeBackdrop(id)
    touched(targets)
    const missed: number[] = []
    // a page whose background is mostly painted over by big shapes would look unchanged: recolour those shapes too
    if (!args.photo_query) {
      const { width: W, height: H } = slideSize()
      const bandsOf = (id: string) => slides.value[pos(id)].elements
        .filter((e: any) => e.type === 'shape' && !elText(e) && e.width >= W * 0.9 && e.width * (e.height || 0) >= W * H * 0.25) // full-width bands only, never cards or diagrams
      // a page whose bands all have that colour already has nothing to recolour: a full-slide white shape under "uniform
      // white" came back as "the large shapes on page 1 could not be recoloured and may still cover it", and the turn was
      // run again (2026-11-18, replayed 2026-11-20). One band of another colour and the page is recoloured as before
      const already = (e: any) => !e.gradient && !e.pattern && String(e.fill || '').toLowerCase() === String(args.color).toLowerCase()
      const covered = targets.filter(id => bandsOf(id).reduce((a: number, e: any) => a + Math.min(e.width, W) * Math.min(e.height, H), 0) >= W * H * 0.45
        && !bandsOf(id).every(already))
      for (let i = 0; i < covered.length; i += 3) {
        await Promise.all(covered.slice(i, i + 3).map(id => editSlide(pos(id), `Change the page's BACKGROUND to ${args.color}: recolour only the large decorative background shapes (full-width bands, full-page panels) to ${args.color} or a lighter/darker tint of it; keep cards, text boxes, pictures and their colours unchanged and keep text readable.`)
          .catch((err) => {
            // the page model failed: bands with no text over them are recoloured directly, the user asked for this colour
            // everywhere ("still not a uniform white", three messages in a row, 2026-11-18). A band under text keeps its
            // colour (white text on a now-white band), and so does every band when the model found nothing to change
            const sl = slides.value[pos(id)]
            const over = (a: any, b: any) => a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + (b.height || 0) && b.top < a.top + (a.height || 0)
            // charts, tables and formulas carry text too; a rotated box can reach a band its rectangle does not touch
            const texts = sl.elements.filter((e: any) => elText(e) || ['chart', 'table', 'latex'].includes(e.type))
            const bands = err instanceof NoChange || texts.some((x: any) => x.rotate) ? [] : sl.elements.filter((e: any) => e.type === 'shape' && !elText(e) && e.width >= W * 0.9 && e.width * (e.height || 0) >= W * H * 0.25
              && !texts.some((x: any) => over(x, e)))
            for (const e of bands) slidesStore.updateElement({ id: e.id, props: { fill: args.color, gradient: undefined, pattern: undefined } as any, slideId: id })
            if (!bands.length) missed.push(pos(id) + 1)
          })))
      }
    }
    const note = missed.length ? (t(' (the large shapes on page(s) {p0} could not be recoloured and may still cover it)', { p0: missed.join(', ') })) : ''
    const bgHow = args.photo_query
      ? (pictureSource === 'ai' ? (t('an AI-generated picture for "{p0}"', { p0: String(args.photo_query).slice(0, 60) })) : (t('a stock photo for "{p0}"', { p0: String(args.photo_query).slice(0, 60) })))
      : (t('colour {p0}', { p0: args.color }))
    // a count read as a page number: "Фон 1 слайдів" ("background of 1 slide(s)") was taken for page 1 (replay 2026-11-18)
    const where = targets.length <= 3 ? t('Background of page(s) {p0} set to {p1}.', { p0: targets.map(id => pos(id) + 1).sort((a, b) => a - b).join(', '), p1: bgHow }) : t('Background of {p0} page(s) set to {p1}.', { p0: targets.length, p1: bgHow })
    return where + note + flatNote + (args.photo_query && pictureNote ? ` ${pictureNote}` : '')
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
    if (!written.length) fail(t('The speaker notes could not be written.'))
    // notes live in a one-line bar under the slide, for the page on screen: notes for pages 4-5 written while on page 1
    // were never seen, and the user asked "Mana?" ("where?", 2026-11-22). Show the first page and open the bar.
    if (!written.includes(slides.value[slideIndex.value].id)) slidesStore.updateSlideIndex(pos(written[0]))
    mainStore.remarkHeight = Math.max(mainStore.remarkHeight, 200)
    const where = written.length <= 5 ? written.map(id => pos(id) + 1).sort((a, b) => a - b).join(', ') : ''
    return (where ? t('Wrote speaker notes for page(s) {p0}: they are in the notes bar under the slide.', { p0: where }) : t('Wrote speaker notes for {p0} page(s): they are in the notes bar under each slide.', { p0: written.length })) + (written.length < targets.length ? (t(' ({p0} page(s) failed)', { p0: targets.length - written.length })) : '')
  }
  if (action === 'download_deck') {
    if (chatslide && !chatslide.canExport()) return t('Downloading needs a paid plan; I opened the plans for you.')
    if (s.mode === 'template' && !s.edited) { // the engine's own .pptx is exact; after editor changes, export the editor
      const a = document.createElement('a')
      a.href = s.versions[s.versions.length - 1].deckUrl
      a.download = 'deck.pptx'
      a.click()
      chatslide?.exported?.('pptx') // ChatSlide counts every saved file (refund eligibility)
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
      if (!written.length) fail(t('The speaker notes could not be written.'))
      return t('Wrote the speaker notes (the script) for {p0} page(s).', { p0: written.length })
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
    // "only the first page, the rest as it was": the theme went on the whole deck twice, then a paywalled redesign (2026-11-14)
    const only = args.pages?.length ? nums(args.pages) : []
    const ids = only.length && only.length < slides.value.length ? only : slides.value.map(x => x.id)
    for (const id of ids) fadeBackdrop(id) // a full-slide picture would hide it
    applyPresetTheme({ background: args.background, fontColor: args.font_color, fontname: args.font || '', colors: args.colors }, true, ids.length < slides.value.length ? ids : undefined)
    touched(ids)
    return ids.length < slides.value.length ? t('Applied a new theme to {p0} page(s).', { p0: ids.length }) : t('Applied a new theme to the whole deck.')
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
      slidesStore.updateSlide({ background: copiedBackground(src), elements: backdrop ? [{ ...JSON.parse(JSON.stringify(backdrop)), id: `${backdrop.id}-${id}`.slice(0, 40) }, ...elements] : elements }, id)
    }
    touched(targets)
    return t("Applied page {p0}'s background and text colour to {p1} page(s).", { p0: pos(src.id) + 1, p1: targets.length })
  }
  if (action === 'open_design') {
    // "open the text settings for me" (after five "whiter" messages changed nothing, 2026-11-15) opened the Design
    // panel: the text of the page on screen is selected, so the toolbar on the right shows its style panel
    if (args.panel === 'text') {
      if (!isPC()) return t('A phone has no text settings panel: tell me the font, size or colour you want and I will set it.')
      const sl = slides.value[slideIndex.value]
      const first = sl && (sl.elements.find((e: any) => (e.textType === 'title' || e.text?.type === 'title') && elText(e)) || readingOrder(sl).find((e: any) => elText(e)))
      if (!first) fail(t('This page has no text box to open the text settings for. Click a text on any page and its settings open on the right.'))
      mainStore.setActiveElementIdList([first.id])
      mainStore.setToolbarState(ToolbarStates.EL_STYLE)
      chatLeft.value = 8
      return t('I selected the text "{p0}" and opened its settings on the right: font, size, colour, bold and spacing. Click another text to change that one.', { p0: elText(first).slice(0, 40) })
    }
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
      else if (snapshotStore.canUndo) {
        // a correction goes back to before the previous message's change, not one editor step: a title dragged after
        // that change took the Undo, so the replaced picture stayed ("Keep the image of the nurse" lost the nurse,
        // 2026-11-16). A bare "undo" stays one step, like the toolbar's
        const snaps = await db.snapshots.orderBy('id').toArray()
        undoSteps = 1
        for (let k = snapshotStore.snapshotCursor - 1; k >= 0 && beforeLastChange && args.instead; k--) {
          if (JSON.stringify(snaps[k].slides) !== beforeLastChange) continue
          undoSteps = snapshotStore.snapshotCursor - k
          break
        }
        beforeLastChange = ''
        // the steps in between are the user's own edits by hand: the deck with them is saved first, so it stays in
        // the version history (the next change drops them from Redo); when it cannot be saved, one step as before
        if (undoSteps > 1) await saveNow().catch(() => { undoSteps = 1 })
        for (let n = undoSteps; n > 0; n--) await snapshotStore.unDo()
      }
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
let beforeLastChange = '' // the deck before the last message that changed it (JSON): where a chat Undo goes back to
let undoSteps = 1 // editor steps that Undo took, for putting the change back
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
let deleteConfirmable: string[] = [] // pages the previous message asked about: this message may delete them
let deleteAnswered: string[] = [] // the same, answered in words: deletable when the router asks for exactly them again
let askedThisTurn = false // a question for the user ends the message
let directAllowed = false // true only while add_site_pictures places the pictures it fetched
const usedBackgrounds = new Set<string>() // background photos already tried in this session: "again" never brings one back
let tappedRestore = false // this message is the tap on a restore confirmation
let turnSelected = '' // what was selected on the canvas when the message was sent
let advancedOffered = false // the paid-model offer, shown once per session
// the upgrade button: after every turn the advanced model would have done better, not only the one the paywall opened on
const upgradeHint = ref(false)
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
// line spacing (a multiple) and paragraph spacing (px) of a text, as numbers the editor draws, or left out: "8px" from
// the page model would draw as "8pxpx"
const spacing = (o: any) => {
  const lh = toNum(o?.lineHeight), ps = toNum(o?.paragraphSpace)
  return { ...(Number.isFinite(lh) && lh >= 0.8 && lh <= 3 ? { lineHeight: lh } : {}), ...(Number.isFinite(ps) && ps >= 0 && ps <= 80 ? { paragraphSpace: ps } : {}) }
}
// the deck after the last step that finished: a crash in a later step keeps the finished steps (a user lost a whole
// turn's slide edit to one failing picture step, 2026-11-11)
let lastGoodSlides: any = null
const deckTurn = async (text: string, again = ''): Promise<void> => {
  // for a second attempt: the deck, its theme and page size, and the chat state
  const startSlides = JSON.parse(JSON.stringify(slides.value)), startS = JSON.parse(JSON.stringify(s))
  const startTheme = JSON.parse(JSON.stringify(slidesStore.theme)), startRatio = slidesStore.viewportRatio
  // a tap on a version-history button: that saved copy becomes the deck (one Undo step), no AI call
  if (!again) newTurn() // a new message: charged on its own (a second attempt stays the same message)
  replyTurn = currentTurn()
  const versionUrl = s.savedVersions?.[text.trim()]
  s.savedVersions = undefined
  if (versionUrl) {
    s.pending = null
    const saved = await (await fetch(versionUrl, { cache: 'no-store' })).json()
    if (!saved?.slides?.length) fail(t('That version could not be loaded.'))
    // the deck on screen is saved first, so it stays in the version history
    await saveNow().catch(() => fail(t('The deck you have now could not be saved, so nothing was restored. Check your connection and try again.')))
    // one Undo step, theme and page size included: the deck on screen is recorded first, then the restored one
    await snapshotStore.addSnapshot()
    if (saved.theme) slidesStore.setTheme(saved.theme)
    if (saved.viewportRatio) slidesStore.setViewportRatio(saved.viewportRatio)
    slidesStore.setSlides(saved.slides)
    slidesStore.updateSlideIndex(0)
    deckBeforeCreate = null // Undo steps back to the deck that was on screen, by the snapshot
    s.edited = true
    await snapshotStore.addSnapshot()
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
  askedThisTurn = false
  userText = text
  editedThisTurn.clear()
  picturedThisTurn.clear()
  readThisTurn.clear()
  roomMadeThisTurn.clear()
  sitePaywalledPage = ''
  replacedThisTurn.clear()
  const pagesAtStart = slides.value.length
  if (!again) takeSpend() // spend from an earlier, failed turn never lands in this footnote (a first attempt does)
  // an offer counts only when this message is its button; otherwise it is dropped (a question holds for any reply)
  if (s.pending?.offer && !s.pending.options?.some(o => o.trim() === text.trim())) { s.pending = null; save() }
  const pendingAtStart = s.pending || null
  const done: string[] = again ? [`note: a first attempt at this message was undone because: ${again} Do the message right this time.`] : []
  const replies: string[] = []
  // what each find_papers / web_search of this message found, by its query: a page writing references is given the
  // results of ITS topic only, never those of another search of the same message
  const found = new Map<string, string>()
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
  // the page on screen is one flat picture: the router is told the lines painted in it and where they are, in the
  // percent of the slide its picture boxes use ("in the space above the page number 02" had no place, 2026-11-13)
  const paintedAtSend = startSl && paintedPicture(startSl).src ? (await paintedLines(startSl)).map((l) => {
    const { width: W, height: H } = slideSize()
    return `"${l.text.replace(/"/g, '\'').slice(0, 80)}" (${[l.left / W, l.top / H, l.width / W, l.height / H].map(v => Math.round(v * 100)).join(', ')})`
  }).join('; ').slice(0, 2500) : ''
  let pendingInstead = '', nudged = false // what an undo said the user wants instead, until a change does it
  let offerOrder = '' // the page order when a whole-deck offer was made in this message
  let failed = 0 // steps of this message that failed
  let leftNudged = false // a not_done the router was sent back over, once per message
  let wordsNudged = false // an answer that only offered the change asked for was sent back, once per message
  currentAtSend = startPageId || ''
  const pagesAtSend = deckLines() // the deck as the router is shown it, for the check of the finished message
  createdThisMessage.clear()
  const idsAtStart = new Set(slides.value.map(x => x.id)) // pages created in this message are the others
  const created = () => slides.value.filter(x => !idsAtStart.has(x.id) && !helperPages.has(x.id)).length
  const picturedPages = new Set<string>() // pages that got a picture for an every-page request in this message
  let picNudgedAt = 0 // picturedPages.size when the router was last sent back for the pages left
  let everyPageAction = 'add_image' // which every-page picture step this message runs: a replace lists only pages with a picture
  const ADDS = ['add_page', 'add_chart_page', 'add_table_page', 'expand_page']
  const adds = (c: any) => ADDS.includes(c.action) && !(Number(c.args?.on_page) > 0) // a table ON a page adds no page
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
  // a font change takes every page in one step: page by page, identical lines merged and the reply listed 4 changes
  // for 9 pages, and the user asked again (2026-11-17)
  const allAtOnce = ['set_font_size', 'set_font', 'set_text_style']
  const repeat: any[] = pendingAtStart?.repeat && pendingAtStart.options?.[0]?.trim() === text.trim()
    // from a page shaped over earlier messages (like): its look as it is now, in place of the last size step again
    ? [...(pendingAtStart.like ? [{ action: 'match_look', args: { from: pendingAtStart.like, pages: pendingAtStart.pages || [], more: pendingAtStart.repeat.some(c => c.action !== 'set_font_size') } }] : []),
      ...pendingAtStart.repeat.filter(c => allAtOnce.includes(c.action) && !(pendingAtStart.like && c.action === 'set_font_size')).map(c => ({ action: c.action,
        args: { ...c.args, pages: pendingAtStart.pages || [], page_named: true, whole_deck_option: '' } })),
      ...(pendingAtStart.pages || []).flatMap(pg => pendingAtStart.repeat!.filter(c => !allAtOnce.includes(c.action)).map(c => ({ action: c.action,
        args: { ...c.args, index: pg, pages: [pg], page_named: true, whole_deck_option: '' } }))), { action: 'finish', args: { not_done: '' } }]
    : []
  const repeating = repeat.length > 0
  if (repeating && pendingAtStart?.from) {
    const general = await forAnyPage(String(pendingAtStart.repeat![0].args?.instruction || ''))
    for (const c of repeat) if (c.action === 'edit_page') c.args.instruction = general
  }
  tappedRestore = repeat.some((c: any) => c.action === 'restore_original')
  // a pasted slide-by-slide text ("Slide 1 ... Slide 2 ..."): its parts go on the pages with those numbers in one step,
  // before the router runs. Left to the router, 1 run in 4 edited page by page and inserted a missing slide mid-deck,
  // shifting the rest (replayed on users' decks, 2026-11-19). Anything else in the message the router does after
  const pre: any[] = []
  if (!repeating && !pendingAtStart && text.length > 600) {
    // "review these slides, change nothing" with numbered slides in it is not a request to place them (Codex review)
    const items = await numberedParts(text, text)
    if (items.length) pre.push({ action: 'place_texts', args: { items } }) // one slide is fine when that is what they asked for
  }
  // "insert the text I wrote into slides 1 to 10" points back at that text: edited slide 1 only, then stopped (1 run in
  // 4, 2026-11-19). With a long text of theirs on hand, the same small model decides whether this message asks for it
  // only their own long message in this chat, and never with a file attached: "only text" answered a question about the
  // file they sent, and their original request was placed instead (backtest 2026-11-19)
  else if (!repeating && !pendingAtStart && !s.docs.length) {
    const theirs = s.history.slice(0, -1).filter(h => h.role === 'user' && h.text.length > 600).map(h => h.text)
    const src = theirs[theirs.length - 1]
    if (src) {
      const items = await numberedParts(src, text)
      if (items.length) pre.push({ action: 'place_texts', args: { items } })
    }
  }
  let round = 0 // 16 after the loop: it ran out of steps with the message unfinished
  const rounds = 16 + repeat.length // fixed before the queue drains (re-read each round it ended a long repeat early)
  for (; round < rounds; round++) {
    const { calls, error } = pre.length ? { calls: [pre.shift()], error: '' } : repeat.length ? { calls: [repeat.shift()], error: '' } : await (await fetch('/api/tools/deck_chat', {
      method: 'POST',
      body: JSON.stringify({
        // "this page" stays the page the user was on when they sent the message: steps that show a new page moved it,
        // and "add a page after this one" then re-ran after each new page (30 -> 36)
        instruction: text, pages: deckLines(), current: (slides.value.findIndex(x => x.id === startPageId) + 1) || slideIndex.value + 1, template: !!s.template, done,
        today: localDay(Date.now()), // "yesterday's version", "the version from 11.11"
        // how the deck was when the message came: told only the size it has now, the router took the 12 pages it had
        // just added for pages that were always there ("it already had 16 slides", then an undo and a delete; e2e 2026-11-16)
        before: pagesAtStart, added: slides.value.map((x, k) => (idsAtStart.has(x.id) ? 0 : k + 1)).filter(Boolean),
        pending: pendingAtStart || undefined,
        selected: selected || undefined,
        painted: paintedAtSend || undefined,
        files: s.docs.map(d => ({ name: d.name, pages: filePages(d.text).length || undefined, chars: d.text.length, look: d.name === lookFile() || undefined })),
        pictures: (s.images || []).map(m => m.name),
        brief: s.brief || undefined,
        history: answeredHistory().map(h => ({ role: h.role, content: h.text })),
      }),
      headers: { 'Content-Type': 'application/json' },
    })).json()
    if (error) throw new Error(error)
    // replacing the deck needs a topic in this very message: a complaint ("same problem") rebuilt a deck on a guess
    const newDeck = calls.find((c: any) => c.action === 'create_deck')
    if (newDeck && (!String(newDeck.args?.topic_in_this_message || '').trim() || !(await namesSubject(text, String(newDeck.args.topic_in_this_message))))) {
      done.push('create_deck -> skipped: the latest message names no topic for a new deck. Do not replace the deck: edit it or ask_user')
      continue
    }
    // after steps that only read material (the attached paper, the web): the deck is built from it, with the router's
    // full instruction ("here is my paper" read the file and stopped, 2026-11-12)
    // nor is the note a second attempt starts with, or a create_deck skipped in an earlier round: with either in done the
    // deck the router asked for again was dropped, nothing ran, and a pasted 6,000-character text got "Готово." (2026-11-19)
    const onlyRead = done.every(d => /^(read_file|web_search|read_link|find_papers) /.test(d) || d.startsWith('not sent: ') || d.startsWith('note: ') || d.startsWith('create_deck -> skipped')) // the sent-back answer is no step
    if (newDeck?.args?.from_users_file && !s.docs.length && !(s.images || []).length) { // no file to build it from: ask for it
      const ask = await runCall(text, 'create_deck', newDeck.args, slides.value.map(x => x.id))
      replies.push(ask)
      terse.push(ask)
      break
    }
    if (newDeck && onlyRead) {
      // their own pasted text is the deck: their message, never the router's retelling of it
      // the router says which is the source: a file still attached to the chat does not turn a pasted text into a summary
      const asWritten = newDeck.args?.as_written === true && newDeck.args?.from_users_file !== true
      const brief = done.length && !asWritten ? String(newDeck.args?.instruction || text) : text
      deckBeforeCreate = JSON.parse(JSON.stringify(slides.value)) // one Undo click brings this deck back
      openedSlides = {} // "the original" of a new deck is the new deck (taken at the next message)
      busy.value = s.template ? t('Building the deck (about a minute)…') : t('Writing the deck…')
      await (s.template ? templateTurn(brief) : generateNative(brief, asWritten))
      addHistorySnapshot() // the new deck is a history step of its own
      return
    }
    if (calls.some((c: any) => c.action === 'rebuild_from_template') && s.template) return templateTurn(text)
    // A question asking for the presentation the chat already has becomes the step itself, here, so it is recorded,
    // undone and continued like any use_file_look (with a surgeon's file listed the router still asked him to attach
    // it, production 2026-11-16). Once per message: after a look transfer ran or failed, the question is asked.
    const look = lookFile()
    for (const c of calls) {
      const want = c.action === 'ask_user' ? String(c.args?.presentation_look || '') : ''
      if (!look || !['colors', 'logo', 'both'].includes(want) || done.some((d: string) => d.startsWith('use_file_look'))) continue
      c.action = 'use_file_look'
      c.args = { file: look, colors: want !== 'logo', logo: want !== 'colors' }
    }
    // one step per round: when the model batches calls it numbers later ones as if earlier ones had run, so the rest
    // is re-planned against the updated deck in the next round
    const work = calls.filter((c: any) => !['finish', 'rebuild_from_template', 'create_deck'].includes(c.action)).slice(0, 1)
    // "make the presentation more detailed" edited pages 2 and 1 and offered the rest as a tap (3 of 3 eval runs,
    // 2026-11-13): an edit the router itself says is for the whole presentation runs on every page, once. Not a
    // redesign or a hard edit: those keep edit_page, its paywall and the advanced model (with the offer for the rest)
    for (const c of work) {
      if (c.action === 'edit_page' && c.args?.user_said_whole_deck === true && c.args?.page_named !== true && !c.args?.source_pages?.length
        && !c.args?.redesign && !c.args?.hard) {
        Object.assign(c, { action: 'edit_all_pages', args: { instruction: String(c.args.instruction || ''), every_page: false, language: c.args.language || '', from_page_edit: c.args } })
      }
    }
    // "pictures on every slide": a finish with pages still without one goes back to the router, once per page done
    // (it stopped after page 1 in 1 of 5 runs although told which pages were left); no progress, no second nudge
    const picLeft = picturedPages.size ? slides.value.map((x, k) => (idsAtStart.has(x.id) && !picturedPages.has(x.id) && (everyPageAction === 'add_image' || pictures(x).length) ? k + 1 : 0)).filter(Boolean) : []
    if (!work.length && picLeft.length && picturedPages.size > picNudgedAt) {
      picNudgedAt = picturedPages.size
      done.push(`not finished: every page was asked for a picture; pages still to do: ${picLeft.join(', ')}. Call ${everyPageAction} for the next one, on that page's own topic.`)
      continue
    }
    const left = calls.find((c: any) => c.action === 'finish')?.args?.not_done
    // a tool the router itself says it skipped for an asked part goes back to it once: "make it nicer, and delete slide
    // 4" ended with "slide 4 still needs to be deleted" (3 of 16 runs, 2026-11-19; the old build 0 of 16). Not after a
    // failed step, a question or an offer; never for a not_done that is only a part that does not apply
    // pages only (a deletion, a new page): a picture or text named as skipped sent the router back to undo the right one
    let skipped = (calls.find((c: any) => c.action === 'finish')?.args?.skipped_tools || []).map(String).filter((x: string) => ['delete_pages', 'add_page'].includes(x))
    // the router also finished with nothing listed while a part was never done ("... and delete slide 4": 2 of 16 runs,
    // 2026-11-19): a separate small check reads the message and the steps once, before the message ends
    if (!work.length && !skipped.length && !leftNudged && done.length && !failed && !askedThisTurn && (!s.pending || s.pending.offer)) {
      const r = await fetch('/api/tools/ask_json', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ instruction:
        `A user asked a slide editor: ${JSON.stringify(text.slice(0, 3000))}\nThe steps the editor ran: ${JSON.stringify(done.map((d: string) => d.slice(0, 300)))}\n`
        + 'Only two kinds of part count: deleting a page the user named, or a new page they asked for (a conclusion, a sources page, "add a slide about X"). List each such part that no step did. Nothing else: pictures, texts and designs on existing pages are never listed. '
        + 'JSON {"missing": [{"tool": "delete_pages" or "add_page", "part": "the part, in the user\'s words"}]}; [] when there is none.' }) }).catch(() => null)
      // pages only: asked about every kind of part, it read "remove the barista photo, put a café photo" as two and the
      // router undid the right picture (2 of 4 runs, 2026-11-19)
      const missing = r?.ok ? ((await r.json().catch(() => ({})))?.missing || []).filter((m: any) => ['delete_pages', 'add_page'].includes(m?.tool) && m?.part) : []
      skipped = missing.map((m: any) => `${m.tool} (${String(m.part).slice(0, 120)})`)
    }
    if (!work.length && skipped.length && !leftNudged && done.length && !failed && !askedThisTurn && (!s.pending || s.pending.offer)) {
      leftNudged = true
      // a page the message asks for and no step added yet is room for one more: "add 2-3 more pages ... and at last a
      // conclusion & bibliography" declared 3, and the conclusion was refused as one page too many (2026-11-20)
      const morePages = skipped.filter((x: string) => x.startsWith('add_page')).length
      if (morePages && pagesWanted >= 0) pagesWanted = Math.max(pagesWanted, created() + morePages)
      done.push(`not finished: you said you have not yet called ${skipped.slice(0, 5).join(', ')} for a part the user asked for. Call it now for that part (call delete_pages itself, it shows its own confirmation; never ask_user about a deletion); never for anything they did not ask for.`)
      continue
    }
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
    // an answer that only offers or promises the change the message asks for is sent back once, to be made: "there are
    // 6 topics, why so little, is 10 slides the limit?" got "25 were asked, 15 are missing, shall I add them?" (4 of 4
    // replays, 2026-11-20), and "I thought you would help" got "I will tidy the layout" and no change (2026-11-11)
    if (!wordsNudged && !askedThisTurn && work.some((c: any) => c.action === 'answer_question' && c.args?.change_not_made === true)) {
      wordsNudged = true
      done.push(`not sent: your answer ("${String(work[0].args?.answer || '').slice(0, 300)}") only offers or promises a change the message asks for, and words leave the deck as it was. Make that change now with its tool, on the most likely reading (the user can undo). Answer in words only if no tool here can make it.`)
      continue
    }
    const ids = slides.value.map(x => x.id) // page numbers in this round's calls refer to this deck
    if (work.every((c: any) => done.some(d => d.startsWith(`${c.action} ${JSON.stringify(c.args || {})} ->`)))) break // repeating itself
    // the same page edited twice already this turn: a third pass is the router looping (seen: 7 edits of page 2)
    if (work.some((c: any) => c.action === 'edit_page' && done.filter(d => d.startsWith('edit_page ') && new RegExp(`"index":${Number(c.args?.index)}[,}]`).test(d)).length >= 2)) break
    // one new picture per page per message: a second one on the same page was an unasked extra (seen: a turkey).
    // Compared on the page it would really land on; skipped (and told to the router), never ending the turn.
    const landsOn = (c: any) => ids[Number(c.args?.index) - 1]
    const stockForBrand = (c: any) => !!sitePaywalledPage && ['add_image', 'change_image'].includes(c.action) && landsOn(c) === sitePaywalledPage
      && !['attached', 'wikimedia', 'specific'].includes(c.args?.source)
    // pages this message added are what it asked for: after "increase to 16 slides" added 12, the router asked to
    // delete those 12 (e2e, 2026-11-16)
    const dropsNew = (c: any) => c.action === 'delete_pages' && (c.args?.pages || []).some((p: number) => !!ids[p - 1] && !idsAtStart.has(ids[p - 1]))
    // the count the user gave is the count left: "10 slides it must be" on a 52-page deck asked to delete 49 and leave 3
    // (2026-11-20). Too long a list is cut in the step itself (the router miscounts long decks); too short is sent back.
    const leftWrong = (c: any) => c.action === 'delete_pages' && Number(c.args?.pages_left) > 0
      && slides.value.length - new Set((c.args?.pages || []).filter((p: number) => p >= 1 && p <= slides.value.length)).size > Number(c.args.pages_left)
    const extra = work.find((c: any) => c.action === 'edit_page' && c.args?.new_content && !c.args?.page_named && created() >= 1)
      || work.find(dropsNew)
      || work.find(leftWrong)
      || work.find((c: any) => c.action === 'delete_pages' && !c.args?.user_asked_to_cut && (c.args?.pages?.length || 0) >= Math.max(2, slides.value.length * 0.3))
      || work.find((c: any) => c.action === 'add_image' && picturedThisTurn.has(picKey(landsOn(c), c.args?.items?.length > 1 ? c.args.items.slice(0, 6).join('|') : c.args?.query)))
      // more new pages than the message asks for: re-plans added "the next" page, or extended the page just added
      // (30 -> 32 for "a page"). The first page-adding step declares the total; steps beyond it are skipped.
      || work.find((c: any) => adds(c) && pagesWanted >= 0 && created() >= pagesWanted)
      // the very same step again (same action, same arguments): "add a page about TikTok" ran three times (30 -> 33)
      || work.find((c: any) => !['undo', 'redo', 'answer_question', 'ask_user'].includes(c.action)
        && turnExact.includes(`${c.action} ${JSON.stringify(c.args || {})}`))
      // a second deck-wide edit is the router re-planning the first ("dark blue, bigger headers" ran twice and ended
      // in "Not done yet … tap Continue", 2026-11-11); what is really left it names in finish
      || work.find((c: any) => c.action === 'edit_all_pages' && done.some(d => d.startsWith('edit_all_pages ')))
      // stock photos are not the brand's products: after the paid-feature answer the router added three anyway (2 of 4 eval runs)
      || work.find(stockForBrand)
      // a request about pictures is never a background change: "replace the photos ... the last one" swapped page 10's
      // background (2026-11-10, replayed 2026-11-13)
      // right after the previous message changed the background, a new background photo is the router's call: "do you
      // have another one, like a city slum" after two background photos was refused and the nurse picture replaced
      // instead (2026-11-16)
      || work.find((c: any) => c.action === 'set_background' && c.args?.photo_query && !saidBackground(text, c.args?.said_background) && !lastChangedBackground())
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
      // nor recolours text again after its colour turned out to be there already and it got a band instead
      // (the very same call only: another colour, scope or page is another part of the message)
      || work.find((c: any) => c.action === 'set_text_style' && ran.some(r => r.args?.backed && JSON.stringify({ ...r.args, backed: undefined }) === JSON.stringify(c.args || {})))
    // undone for a "correction", then the router wants the very same change back: the undone change was already
    // what the user meant ("I asked for the image, not the background" after the image was changed). Put it back.
    if (extra?.action === 'change_image' && done.some(d => d.startsWith('undo ')) && lastTurnCalls.includes(stepKey(extra.action, extra.args))) {
      for (let n = undoSteps; n > 0; n--) await snapshotStore.reDo()
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
      const keep = Number(extra.args?.pages_left), n = slides.value.length
      const why = dropsNew(extra) ? 'those pages were added by this message (see Done so far): they are what the user asked for and they stay; finish'
        : leftWrong(extra) ? `that would leave ${n - new Set(extra.args.pages.filter((p: number) => p >= 1 && p <= n)).size} pages, but the user wants ${keep} in the end: call delete_pages with exactly ${n - keep} pages, the least essential first`
        : extra.action === 'delete_pages' ? 'the user did not ask to delete or cut these pages. Do not delete: ask_user what they want, with options'
        : extra.action === 'edit_page' && extra.args?.new_content ? 'that new content is already on the new page just added (see Done so far); finish'
          : adds(extra) && pagesWanted >= 0 && created() >= pagesWanted
            ? `the ${pagesWanted} new page(s) this message asks for are already added (see Done so far); finish`
            : turnExact.includes(`${extra.action} ${JSON.stringify(extra.args || {})}`)
              ? 'this exact step already ran in this message (see Done so far): the request is done, finish'
              : lastTurnCalls.includes(stepKey(extra.action, extra.args)) && done.some(d => d.startsWith('undo '))
                ? 'that is exactly the change just undone because the user said it was wrong; do the request on the target the user means instead'
                : extra.action === 'animate_pages' ? 'the animations this message asks for were just added; keep them and finish'
                : extra.action === 'set_text_style' ? 'the text already has that colour and now stands on a band that makes it readable (see Done so far); finish'
                : extra.action === 'set_background' ? 'the user\'s words do not ask for the page background, so it stays: for pictures use change_image (all_pictures for every picture on a page), add_image or remove_images'
                : extra.action === 'edit_all_pages' ? 'a deck-wide edit already ran in this message (see Done so far); finish, and put in not_done only a part of the request it did not cover'
                : stockForBrand(extra) ? 'the brand\'s own pictures are a paid feature the user was just told about; a stock photo is not their product: finish'
                : extra.action === 'add_image' ? 'that picture was already added to that page in this message' : extra.action === 'change_image' ? 'that picture was already replaced in this message' : 'undo only reverts changes from earlier messages, not this one: the steps already run for this message are what the user asked for and they stay, so never redo one of them on another page or picture. Finish, unless a DIFFERENT part of the message is still to do'
      done.push(`${extra.action} ${JSON.stringify(extra.args || {})} -> skipped: ${why}`)
      // a router that keeps proposing what was just refused is done: "khác đi" replanned the same picture for 16 rounds
      if (done.length >= 2 && done.slice(-2).every(d => d.includes(' -> skipped: '))) break // two in a row
      continue
    }
    // one deck-wide pass per turn: after pages are deleted the numbering shifts and a re-plan deleted MORE pages
    // ("delete 25-28" took 6); a second edit_all_pages re-edits every page for nothing (30 pages, twice)
    // (a second duplicate_pages copied the copies again); the reply says the rest was left for another message
    // (a pass that failed did nothing: "delete pages 9-11" on a 10-page deck failed, and the right one was refused as
    // "Not done yet", 2026-11-20)
    // nor did a skipped one: a delete sent back for leaving the wrong count must still run with the right pages
    const passed = (d: string, action: string) => d.startsWith(`${action} `) && !d.includes(' -> failed') && !d.includes(' -> skipped: ')
    const again = work.find((c: any) => ['delete_pages', 'edit_all_pages', 'reorder_pages', 'set_background', 'duplicate_pages'].includes(c.action) && done.some(d => passed(d, c.action)))
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
    if (again && done.some(d => passed(d, again.action) && (['set_background', 'delete_pages', 'duplicate_pages'].includes(again.action) ? samePages(d) : d.startsWith(`${again.action} ${JSON.stringify(again.args || {})} ->`)))) break
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
        done.push(`${c.action} -> skipped: undo only reverts changes from earlier messages, not this one: the steps already run for this message are what the user asked for and they stay, so never redo one of them on another page or picture. Finish, unless a DIFFERENT part of the message is still to do`)
        continue
      }
      if (adds(c) && pagesWanted < 0) pagesWanted = Math.max(1, Number(c.args?.message_new_pages) || Number(c.args?.count) || 1)
      // no step adds more than this message still allows
      if (adds(c) && Number(c.args?.count) > 0) c.args.count = Math.min(Number(c.args.count), Math.max(1, pagesWanted - created()))
      else if (adds(c) && created() >= pagesWanted) { // planned in the same batch as the pages already added
        done.push(`${c.action} -> skipped: the ${pagesWanted} new page(s) this message asks for are already added`)
        continue
      }
      if (c.action === 'edit_all_pages' && c.args?.from_page_edit) {
        const n = slides.value.length
        const one = c.args.from_page_edit
        if (n > 12) { // a long deck: the page count first, one tap runs it page by page (each page kept as it lands)
          const q = t('Edit all {p0} pages? Nothing is changed yet.', { p0: n })
          s.pending = { request: userText, question: q, options: [t('Yes, edit all {p0} pages', { p0: n })], offer: true, localize: true,
            repeat: [{ action: 'edit_page', args: { ...one, whole_deck_option: '', user_said_whole_deck: false } }], pages: slides.value.map((_, k) => k + 1).slice(1, -1), from: Number(one.index) || 1 } // cover and closing kept, as below
          askedThisTurn = true
          save()
          replies.push(q)
          terse.push(q)
          break
        }
        c.args = { instruction: await forAnyPage(c.args.instruction), every_page: false, language: c.args.language,
          pages: n > 3 ? slides.value.map((_, k) => k + 1).slice(1, -1) : [] } // the cover and the closing page keep their few words
      }
      // "replace the pictures on all slides": a page with none has nothing to replace. After replacing, the router added
      // new pictures to those pages in 5 of 16 runs, old build too (2026-11-19)
      if (c.action === 'add_image' && c.args?.every_page && everyPageAction === 'change_image' && picturedPages.size) {
        const line = t('Page {p0} has no picture to replace.', { p0: Number(c.args.index) || 1 })
        picturedPages.add(ids[Number(c.args.index) - 1])
        done.push(`add_image ${JSON.stringify(c.args)} -> skipped: the user asked to replace pictures, and this page has none to replace`)
        replies.push(line)
        terse.push(line)
        continue
      }
      busy.value = t('Working (step {p0})…', { p0: done.length + 1 })
      let reply = ''
      try {
        lastEditTerse = '' // set only by a page edit that ran: a paywalled redesign of slide 5 was reported as the earlier "Slide 8 updated." (2026-11-17)
        // a reference list with no search behind it: the page wrote "fill in the sources here" as its literature list
        // (a teacher's "a literature list at the end, as I asked in my request", 2026-11-15, replayed 2026-11-20). The
        // papers found for this topic go to the page itself (the router's instruction said "include the verified
        // references" without them, and the page wrote that it had none); none found yet: they are found first, and
        // none found anywhere fails the step and the user is told so. The router leaves the topic empty when the user
        // supplied the references (pasted, in their file, on the deck): theirs are used, nothing is searched
        const topic = ['edit_page', 'add_page'].includes(c.action) ? String(c.args?.references_topic || '').trim() : ''
        if (topic && !found.has(topic)) {
          const got = await runCall(text, 'find_papers', { query: topic }, ids)
          found.set(topic, memo)
          memo = ''
          replies.push(got)
          terse.push(got)
          done.push(`find_papers ${JSON.stringify({ query: topic })} -> ${got}\n${found.get(topic)}`)
        }
        if (topic && found.get(topic)) c.args = { ...c.args, instruction: `${c.args.instruction}\n\nThe sources found for it. The reference list holds only these, each written out in full (authors, year, title, journal, DOI or link), in the format the user asked for; never a placeholder or a note about the list:\n${found.get(topic)}` }
        reply = await runCall(text, c.action, c.args || {}, ids)
      }
      catch (e0) {
        // a service answered this step with something unreadable, twice (the call is asked again once): the step
        // failed, said plainly, and the other parts of the message go on; it never ends the message with a parser error
        const e = e0 instanceof SyntaxError ? new StepFailed(t('That step failed on our side and was not done. Please ask for it again.')) : e0
        if (!(e instanceof StepFailed)) throw e
        memo = '' // what a failed step read never rides along with a later step
        replies.push(e.message)
        terse.push(e.message)
        if (repeating && e instanceof NoChange) continue // one page already as asked: the other pages still get it
        // A failure used to end the message: "change the photo on slide 1 to a plainer one, and the photo on slide 3"
        // stopped at "there is no picture to remove" with slide 3 untouched (2026-11-12). Now the user is told at
        // once, the router gets it as that step's result, and it may go on with the OTHER parts of the request. Still
        // the end of the message: a question to the user ("Which font?"), a page that is already as asked, a repeat
        // over pages, or a second failure.
        // pasted text for several pages: one page already holding its part is no reason to stop; the message ended
        // there and the other pages never got theirs (replayed 2026-11-18, slides 1 and 8 of two decks)
        // nor one page already as asked in a message with other parts: "eight slides ... the last a thank-you slide"
        // stopped at page 7 "already explains it", the rest never done (2026-11-20). The router goes on; a page it
        // edits twice still ends the message
        const goOn = e instanceof NoChange && c.action === 'edit_page'
        if (!goOn) failed++
        // the step that ended the message is kept too: the turn record showed no step at all for "Page 6 was not
        // changed", so the cause could not be read from it (2026-11-14)
        done.push(goOn ? `${c.action} ${JSON.stringify(c.args || {})} -> no change: ${e.message} (Do not retry this page. Go on with any other part of the request, or finish if none is left.)`
          : `${c.action} ${JSON.stringify(c.args || {})} -> failed, and the user was told: ${e.message} (Do not retry it. Go on with any other part of the request.)`)
        if (repeating || (e instanceof NoChange && !goOn) || /[?？؟]\s*$/.test(e.message) || failed >= 2) { round = 99; break }
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
      terse.push(c.action === 'edit_page' && lastEditTerse ? lastEditTerse : reply)
      // the whole-deck offer is a button, not the end of the message: "make the title red and add a picture" stopped
      // after the title (2026-11-12). The router is told, so it neither applies it itself nor stops there.
      const offered = c.action === 'edit_page' && s.pending?.offer && s.pending.options?.[0] === String(c.args?.whole_deck_option || '').trim()
        ? ' (The user now has a button to apply this to the whole deck: never do that yourself. Go on only with a different part of the request, if there is one.)' : ''
      if (offered) offerOrder = slides.value.map(x => x.id).join()
      else if (offerOrder && s.pending?.offer && offerOrder !== slides.value.map(x => x.id).join()) { s.pending = null; offerOrder = '' } // pages added, deleted or moved since: its numbers are stale
      // "pictures on every slide": the router added one and stopped, listing the rest for the user (2026-11-13). It is told
      // which of the deck's pages still have no new picture, so it goes on page by page (each step its usual checks)
      // the same for "replace the pictures on all slides": change_image touched page 2 and stopped (2026-11-14); only
      // pages that have a picture to replace are listed
      const everyPage = (c.action === 'add_image' || c.action === 'change_image') && c.args?.every_page
      if (everyPage) { picturedPages.add(ids[Number(c.args?.index) - 1]); everyPageAction = c.action }
      const left = everyPage
        ? slides.value.map((x, k) => (idsAtStart.has(x.id) && !picturedPages.has(x.id) && (c.action === 'add_image' || pictures(x).length) ? k + 1 : 0)).filter(Boolean) : []
      const rest = left.length ? `\n(Every page was asked for: pages still without a ${c.action === 'add_image' ? 'new' : 'replaced'} picture: ${left.join(', ')}. Call ${c.action} for the next one, on that page's own topic.)` : ''
      if ((c.action === 'find_papers' || c.action === 'web_search') && memo) found.set(String(c.args?.query || '').trim(), memo)
      done.push(`${c.action} ${JSON.stringify(c.args || {})} -> ${reply}${offered}${memo ? `\n${memo}` : ''}${rest}`)
      memo = ''
      if (!['answer_question', 'ask_user', 'say_unable', 'undo', 'redo', 'download_deck', 'read_file', 'web_search', 'find_papers', 'read_link', 'go_to_page', 'open_design'].includes(c.action) && !askedThisTurn) changed = true
    }
    // showing a page goes on with the rest of the message: "show the picture on page 4 and enlarge the text in the whole
    // presentation" showed page 4 and stopped (2026-11-15); the router finishes when nothing is left
    if (askedThisTurn || work.every((c: any) => ['answer_question', 'ask_user', 'say_unable', 'open_design'].includes(c.action))) break // a question is asked/answered once
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
  const perPage = ['add_image', 'change_image', 'set_font_size', 'set_font', 'set_text_style', 'remove_images']
  let followUp = ''
  // a message that changed the text look of the page it was sent on, and no other page, shaped that page
  const startNow = slides.value.find(x => x.id === startPageId), startWas = startSlides.find((x: Slide) => x.id === startPageId)
  const shapedBefore = !!startPageId && !!s.shaped?.includes(startPageId)
  if (changed && startNow && startWas && editedThisTurn.size === 1 && editedThisTurn.has(startNow.id) && JSON.stringify(textLook(startNow)) !== JSON.stringify(textLook(startWas))) {
    s.shaped = [...new Set([...(s.shaped || []), startNow.id])].slice(-50)
  }
  // more to do than one message runs (16 steps, one deck-wide pass of a kind): one tap continues, nothing is dropped
  if ((unfinished || round === rounds) && !askedThisTurn) {
    followUp = t('There is more to do from your message. Tap Continue and I will carry on.')
    // answering an earlier question or offer (Continue included): the request stays the original one, steps accumulate
    s.pending = { request: pendingAtStart ? pendingAtStart.request : text, question: followUp, options: [t('Continue')], localize: true, done: [...(pendingAtStart?.done || []), ...done].slice(-30) } // any reply continues it
    save()
  }
  else if (changed && !turnSelected && !s.pending && !askedThisTurn && slides.value.length > 1 && editedThisTurn.size === 1 && editedThisTurn.has(startPageId)
    && ran.length && ran.every(c => perPage.includes(c.action) && c.args?.page_named !== true && !(c.args?.items?.length > 1) && !c.args?.phrase && !c.args?.backed)) { // a picture set may add pages: not repeated; named words are on one page
    followUp = t('Done on this page. Want the same on the other pages?')
    const here = slides.value.findIndex(x => x.id === startPageId) + 1
    const others = slides.value.map((_, k) => k + 1).filter(k => k !== here)
    // the router learns which pages are left and what was done here, so it neither redoes this page nor skips one
    s.pending = { request: text, question: `${followUp} (the other pages: ${others.join(', ')})`, options: [t('Do the same on the other {p0} pages', { p0: others.length })], done: done.slice(-30), offer: true, localize: true,
      repeat: ran.map(c => ({ action: c.action, args: c.args })), pages: others,
      like: shapedBefore && ran.every(c => allAtOnce.includes(c.action)) ? here : undefined }
    save()
  }
  const lines = terse.length > 1 ? replies.map((r, i) => terse[i] ?? r) : replies // one step: the model's sentence is fine
  // counts quoted by earlier steps go stale when a later step adds or removes pages: the final size closes the turn
  if (slides.value.length !== pagesAtStart) lines.push(t('The deck now has {p0} pages.', { p0: slides.value.length }))
  if (followUp) lines.push(followUp)
  const reply = [...new Set(lines.filter(Boolean))].join('\n') || (t('Done.')) // every step reports: merging repeats hid a doubled page
  // a check of the result, once: the same build passed a replayed message in one run and missed it in the next (60-72
  // of 100 real conversations, 2026-11-19). A miss is undone and the message runs again, told what was wrong. Not after
  // a question, a step ChatSlide opened or confirms, or a new deck
  const own = ['talk_to_human', 'download_deck', 'share_deck', 'open_refund', 'open_step', 'open_design', 'version_history', 'undo', 'redo', 'restore_original', 'delete_pages', 'create_deck']
  // nor when putting an earlier picture back was all the message did: the check sees a page's texts and picture
  // count, so "no change"
  // nor a band put behind text that already had the colour asked for: the check sees no colours and no band, and
  // undid it as "a band instead of whiter letters" (4 of 4 replays, 2026-11-20)
  if (!again && !askedThisTurn && !unfinished && round < rounds && !ran.some(c => own.includes(c.action)) && !(ran.length && ran.every(c => c.args?.earlier > 0 || c.args?.backed))) {
    // the question is passed only when the message is one of its own buttons: any other reply may be a new request
    // ("never mind, make the background white"), and the check would hold it to the request it dropped
    const problem = await checkTurn(text, startSlides, reply, pendingAtStart?.options?.some(o => o.trim() === text.trim()) ? pendingAtStart : null, pagesAtSend, done)
    if (problem) {
      slidesStore.setSlides(startSlides)
      slidesStore.setTheme(startTheme)
      slidesStore.setViewportRatio(startRatio)
      lastGoodSlides = null // a failed retry must not bring the rejected attempt back
      for (const k of Object.keys(s)) if (!(k in startS)) delete (s as any)[k]
      Object.assign(s, startS)
      slidesStore.updateSlideIndex(Math.max(0, startSlides.findIndex((x: Slide) => x.id === startPageId)))
      return deckTurn(text, problem)
    }
  }
  if (changed) {
    deckBeforeCreate = null // after another edit, Undo steps back normally
    beforeLastChange = JSON.stringify(startSlides)
    renumber()
    addHistorySnapshot()
    if (s.mode !== 'native') s.edited = true // a template deck or a deck opened in the editor: keep the edited copy
  }
  const sp = takeSpend()
  const fee = sp.charged > 0 ? (t('\n(Credits used: {p0}{p1})', { p0: sp.charged, p1: sp.left >= 0 && sp.left <= 9000000 ? t(' · {p0} left', { p0: sp.left }) : '' })) : ''
  const said = await finalize(text, reply, free, s.pending !== pendingAtStart && s.pending?.localize ? s.pending : null)
  say('assistant', said + fee)
  recordTurn(text, said, done, true)
}

// What a turn did, read by a small model against the message: '' when it did what was asked (or the check could not
// run), else what was wrong, in one sentence for the second attempt
// It is given what the router had. Without the slide count asked for when the deck was made it called "25 slides were
// asked, 15 are missing" an invented claim; without the pasted 21-slide plan that a question's answer continued it called
// that plan's speaker notes unasked; and it could not see a background picture being hidden ("the orange patches are
// not confirmed removed"). Each undid a right turn and ran it again: 9 of its 12 rejections in two replays of 11 real
// turns (2026-11-20), one of them a 21-page build written twice
// It rejected 25 of the 97 turns it read in one run of the 106 editor cases (2026-11-20), 20 of them right turns: a named
// picture changed on the page that has it ("the user was viewing page 2"), "every slide" where the other pages had no
// picture, a shape, a colour or a layout it is not shown, a page count off by a page that needed no change. So it also
// gets the deck as the router saw it (what each picture shows, the logo, the pages with nothing to change) and the
// steps with their results (the page each one worked on and why), and it fails a turn only for a ground it can show
// from that record; its rules come first, so a long record is cut at its end, never at the question
const checkTurn = async (text: string, was: Slide[], reply: string, pending: typeof s.pending, pagesWere: string[], steps: string[]) => {
  // a page's text is cut to the room there is, and the cut is marked: cut bare, at 700 characters a page and 7,000 in
  // all, 21 new pages came back as "the content is truncated on slides 19 and 21" and were written again (2026-11-20)
  const differ = slides.value.filter(x => JSON.stringify(was.find(o => o.id === x.id)) !== JSON.stringify(x)).length + was.filter(o => !slides.value.some(x => x.id === o.id)).length
  const room0 = Math.max(120, Math.min(700, Math.floor(6000 / Math.max(1, differ))))
  // every changed page keeps its line, whatever else the record holds (fitCheck): at 9,000 characters a ten-page update
  // showed pages 1-9, and the check undid it as "the reply claims 10 pages, the change log shows 9" (replay of a real
  // turn, 2026-11-20). Under 120 characters a side the mark is the short one, said in the rules
  const cut = (t: string, n: number) => (t.length > n ? `${t.slice(0, n)}… ${n < 120 ? '(cut)' : '(cut here for length: the rest is there, not shown)'}` : t)
  const texts = (sl: Slide, room: number) => cut(sl.elements.map(elText).filter(Boolean).join(' | '), room)
  const shown = (sl: Slide, room: number) => `${texts(sl, room)} [${sl.elements.filter(e => e.type === 'image').length} picture(s)${sl.elements.some(e => e.type === 'chart') ? ', a chart' : ''}${sl.elements.some(e => e.type === 'table') ? ', a table' : ''}]`
  const old = new Map(was.map((x, k) => [x.id, { sl: x, n: k + 1 }]))
  const now = new Set(slides.value.map(x => x.id))
  const hidden = (sl: Slide) => sl.elements.filter((e: any) => e.type === 'image' && isBackdrop(e) && e.filters?.opacity === '0%').length
  const changes = (room: number) => {
    const lines: string[] = []
    slides.value.forEach((x, k) => {
      const o = old.get(x.id)
      if (!o) lines.push(`page ${k + 1} NEW: ${shown(x, room)}`)
      else if (o.n !== k + 1 && JSON.stringify(o.sl) === JSON.stringify(x)) lines.push(`page ${k + 1} MOVED (was page ${o.n}): ${cut(texts(x, room), 120)}`)
      else if (JSON.stringify(o.sl) !== JSON.stringify(x)) lines.push(`page ${k + 1} CHANGED${o.n !== k + 1 ? ` (was page ${o.n})` : ''}\n  before: ${shown(o.sl, room)}\n  after: ${shown(x, room)}${JSON.stringify(o.sl.background) !== JSON.stringify(x.background) ? '\n  (its background changed)' : ''}${texts(o.sl, 100000) === texts(x, 100000) ? '\n  (the same words: what changed is what you are not shown, such as positions, sizes, line spacing or colours)' : ''}${hidden(x) > hidden(o.sl) ? '\n  (its full-slide background picture, patterns and patches included, is hidden, so the background colour is what shows)' : ''}`)
    })
    was.forEach((x, k) => { if (!now.has(x.id)) lines.push(`page ${k + 1} DELETED: ${shown(x, room)}`) })
    return lines
  }
  const earlier = (userMax: number) => s.history.slice(-5, -1).map(m => `${m.role}: ${m.text.slice(0, m.role === 'user' ? userMax : 300)}`).join('\n')
  // only the count: given the request's text too, it held "there are no pictures" to the schematic drawings that text
  // asked for, and the second attempt put AI pictures on all ten pages (replay 2026-11-20)
  // and only while the deck is shorter than that, as the router is told it: a project recorded as 1 slide made the check
  // call a ten-page update of a ten-page deck "instead of the one-slide deck asked for"; the update was undone and the
  // second attempt added one page (replay of a real turn, 2026-11-20)
  const made = Number(s.brief?.pages) > was.length ? `When the deck was made the user asked for ${s.brief!.pages} slides.\n` : ''
  const answers = (max: number) => (pending?.question ? `The message answers the editor's question ${JSON.stringify(pending.question.slice(0, 400))} about the user's request ${JSON.stringify(cut(String(pending.request || ''), max))}: that request is what was asked.\n` : '')
  const flat = (t: string) => t.replace(/\s+/g, ' ').trim()
  const ran = steps.filter(d => !d.startsWith('note: ') && !d.startsWith('not sent: ')).map((d) => {
    const at = d.indexOf(' -> ')
    return at < 0 ? cut(flat(d), 500) : `${cut(flat(d.slice(0, at)), 300)} -> ${cut(flat(d.slice(at + 4)), 500)}`
  })
  const deckWas = (per: number) => was.map((_, k) => `page ${k + 1}: ${cut(pagesWere[k] || '', Math.max(60, Math.min(per, Math.floor(6000 / Math.max(1, was.length)))))}`)
  // a picture asked for ON a page is right there: "new pictures for a page with no room go on a new page" read as a
  // rule made the check undo a picture put on the page the user named, and the second attempt added a page (2026-11-20)
  const head = 'You check an AI slide editor\'s work on one chat message, from the record below. Return JSON {"ok": true|false, "problem": "<what is wrong, one sentence, \\"\\" when ok>"}.\n'
    + 'ok is false only when the record itself shows one of these:\n'
    + '1. A part of the message was not done at all, or was only offered or promised, and no step result and no line of the reply says why (nothing of that kind on the page, no room, a paid plan, something it cannot do). Look at the deck as it was first: a page with nothing of the kind asked (no picture to replace or remove) is no missed part, and a thing drawn as a shape (an arrow, a box) where a picture of it was asked is done.\n'
    + '2. Something was changed that the message did not ask for and the request does not need: the words already on a page put into another language though the message asks for no translation, a page added or removed and left there, pages rewritten that were not meant. What a step reports as part of its own work (the text colour that goes with a background, the font that goes with a file\'s look) is not this, and the language of the reply never is.\n'
    + '3. The work went to a page other than the one the message means, and neither the deck as it was nor a step result explains it. The page the user was looking at is only where "this slide" and "here" point. The editor picks pages by its own rules with the whole deck in view, and these are right: a picture or thing the message names is changed on the page that has it (each picture\'s line in the deck as it was says what it shows); new pictures for a page with no room may go on a new page after it, and a picture put on the very page the message names is right however full that page was; the logo, a change for an audience and a change to "the presentation" or "the slides" go on every page; a request for every page skips the pages that have nothing of that kind.\n'
    + '4. The reply says a change was made, of a kind you are shown, and the list of changes has nothing of it. The deck is what is checked: a reply that is clumsy, or wrong in a number or a side remark while the deck is right, is no reason.\n'
    + 'Never a reason: anything you are not shown (you see each page\'s words and how many pictures, charts and tables it has; positions, sizes, colours, shapes, the layout, the theme and what a new picture shows are not shown, and for those the step result is the record: a picture found for the words asked is that picture); the wording, quality or taste of what was made; a page count in the reply that is off by a page that needed no change; something that only "may" be wrong; anything marked as cut, shortened or left out for length (it is there, only not shown).\n'
    + 'Also fine: a needed question, a note that something needs a paid plan, an honest "cannot" for something it cannot do, an answer to a question that asked for no change, and making what the user complains is missing or asks why it is missing (fewer slides than they asked for, their text not on the slides) rather than only offering it.\n'
    + `${made}${s.docs.length ? `Files the user gave (their "document", "file", "text"): ${s.docs.map(d => d.name).join(', ').slice(0, 600)}\n` : ''}`
    + `The user's message (they were looking at page ${was.findIndex(x => x.id === currentAtSend) + 1} of ${was.length}): ${JSON.stringify(text.slice(0, 3000))}\n`
  const instruction = fitCheck({ head, reply, request: answers, earlier, deck: deckWas, steps: max => cut(ran.join('\n'), max), changes }, room0)
  const r = await fetch('/api/tools/ask_json', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ instruction }) }).catch(() => null)
  const v = r?.ok ? await r.json().catch(() => null) : null
  return v?.ok === false && String(v.problem || '').trim() ? String(v.problem).trim().slice(0, 400) : ''
}

// ChatSlide keeps a record of every chat turn for review: the message, the reply, and what each step did (step,
// arguments, result), so a complaint like "I don't see a change" can be traced to the edit before it. Never charged.
const recordTurn = (message: string, reply: string, steps: string[], ok: boolean) => {
  if (!chatslide) return
  fetch('/api/record_turn', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ message, reply, ok, page: slideIndex.value + 1, start: slides.value.findIndex(x => x.id === currentAtSend) + 1, pages: slides.value.length, steps: steps.map(d => d.split('\n')[0].slice(0, 400)).slice(0, 30) }),
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
  upgradeHint.value = false
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
    chatslide?.exported?.('pptx')
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
// On first open the pages come from ChatSlide's PowerPoint build, whose boxes have the template's sizes: longer text ran
// over the picture under it (6 of a surgeon's 26 pages, 2026-11-13). A box whose text is taller than it grows down to
// the next element; when that is not room enough its type shrinks (never under 60%). ids: only those pages; floor: per
// box, sizes its type never shrinks under (in order of the box's sizes); only: those boxes, the rest stay as they are
// spacedFrom (boxes an edit gave more line/paragraph spacing, with the spacing they had): such a box first takes the
// free width right of it, up to the next element beside it (fewer lines, same type size); still over its room, its
// spacing steps back toward what it was, then its type shrinks, measured each round, to the floor and if need be under
// it: less spacing or smaller letters, never text over the element below or off the slide (Codex review, 2026-11-21)
const fitOverflow = (ids?: string[], floor: Record<string, number[]> = {}, only?: Set<string>, spacedFrom?: Record<string, number[]>) => withRenderer(async () => {
  const { width: W, height: H } = slideSize()
  const widen = !!spacedFrom
  const widened = new Set<string>()
  const released = new Set<string>() // boxes already let below their size floor: the floor never comes back for them
  try {
    await document.fonts?.ready // measured in the deck's fonts, not a fallback's
    let again = new Set(ids || slides.value.map(x => x.id)) // later rounds draw only the pages a round changed
    for (let round = 0; round < (widen ? 8 : 3) && again.size; round++) {
      const pages = slides.value.filter(x => again.has(x.id))
      toRender.value = JSON.parse(JSON.stringify(pages))
      await nextTick()
      await new Promise(r => setTimeout(r, 400))
      again = new Set()
      for (const sl of pages) {
        for (const e of sl.elements as any[]) {
          const html = e.type === 'text' ? e.content : e.text?.content
          if ((e.type !== 'shape' && e.type !== 'text') || !plain(html) || e.rotate || (only && !only.has(e.id))) continue
          const node = renderRef.value?.querySelector(`.base-element-${CSS.escape(e.id)} .ProseMirror-static`) as HTMLElement | null
          const box = node?.parentElement
          if (!node || !box) continue
          const pad = parseFloat(getComputedStyle(box).paddingTop) + parseFloat(getComputedStyle(box).paddingBottom)
          const need = node.offsetHeight + pad
          // what it would run into: anything under it in its columns, except a band that sits behind other elements
          // (one that starts below this box's top: a picture already under its lower edge counts)
          const under = (sl.elements as any[]).filter(o => o.id !== e.id && o.top > e.top + 4 && o.left < e.left + e.width && o.left + o.width > e.left
            && !(o.type === 'shape' && !plain(o.text?.content) && o.width * o.height > W * H * 0.2) && !(o.type === 'image' && isBackdrop(o)))
          const room = Math.min(H - 8, ...under.map(o => o.top - 6)) - e.top
          // a text box on the page shown has already grown to its text (the editor stores it): checked against the room
          const was = spacedFrom?.[e.id]
          // a spaced box is judged by its room, not its box: the page model leaves boxes reaching into the next one
          if (need <= e.height + 2 && (e.type !== 'text' || need <= room + 2) && (!was || need <= room + 2)) continue
          if (widen && !widened.has(e.id)) {
            widened.add(e.id)
            const bottom = e.top + Math.max(e.height, need)
            const beside = (sl.elements as any[]).filter(o => o.id !== e.id && o.left >= e.left + e.width - 2 && o.top < bottom && o.top + o.height > e.top
              && !(o.type === 'image' && isBackdrop(o)) && !(o.type === 'shape' && !plain(o.text?.content) && o.width * o.height > W * H * 0.2))
            const width = Math.min(W - 24, ...beside.map(o => o.left - 12)) - e.left
            if (width > e.width + 20) {
              slidesStore.updateElement({ id: e.id, props: { width }, slideId: sl.id })
              again.add(sl.id)
              continue
            }
          }
          if (was && need > room && room > 8) {
            const x = e.type === 'text' ? e : e.text || {}
            const lh = x.lineHeight ?? 1.5, ps = x.paragraphSpace ?? 5
            if (lh > was[0] + 0.01 || ps > was[1] + 0.5) { // halfway back, then all the way
              const back = { lineHeight: lh - was[0] > 0.15 ? Math.round((lh + was[0]) * 50) / 100 : was[0], paragraphSpace: ps - was[1] > 3 ? Math.round((ps + was[1]) / 2) : was[1] }
              slidesStore.updateElement({ id: e.id, props: e.type === 'text' ? back : { text: { ...e.text, ...back } }, slideId: sl.id })
              again.add(sl.id)
              continue
            }
          }
          if (room <= 8 || (room <= e.height + 2 && round && !was)) continue // off the slide, or nowhere to grow and already shrunk once
          const props: any = { height: Math.max(e.height, Math.min(need, room)) }
          if (need > room) {
            const k = Math.max(0.6, Math.sqrt(room / need)) // smaller type also wraps less: height falls with its square
            const shrink = (min: number[]) => { let i = 0; return String(html).replace(/font-size:\s*([\d.]+)px/g, (_m: string, v: string) => `font-size: ${Math.max(9, min[i++] || 0, Math.round(Number(v) * k * 10) / 10)}px`) }
            let content = shrink(released.has(e.id) ? [] : floor[e.id] || [])
            if (was && content === html && !released.has(e.id)) { released.add(e.id); content = shrink([]) } // at the floor and still over: smaller still, never over the next element
            if (was && content === html) continue // 9px: nothing smaller to do
            if (e.type === 'text') props.content = content
            else props.text = { ...e.text, content }
          }
          slidesStore.updateElement({ id: e.id, props, slideId: sl.id })
          again.add(sl.id)
        }
      }
    }
  }
  finally {
    toRender.value = []
  }
})
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
    chatslide?.exported?.(format)
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
        await fitOverflow().catch(() => undefined)
        s.edited = true // the editor copy is the deck from now on
        s.synced = { keys: pageKeys(slides.value), urls: {} } // unchanged: the project keeps its own slides until an edit
        URL.revokeObjectURL(link)
        // saved at once: the saved deck is what makes the project open in this editor next time. Never a failed
        // conversion: a saved blank page would be opened instead of the project's slides from then on
        save()
      }
      // the files the user already gave this project: "use my file" reads them, no re-attach (12 replies asked for
      // the 📎 from users whose file the project already had, 2026-11-10 to 13). Not awaited: the deck opens first.
      // 2: each original file by its own name; a deck opened before (true) held the outline answers under the joined
      // file names, replaced once (2026-11-15)
      // what the user asked for when the project was made: "why only 10 slides?" (they asked for 20), "I sent much more
      // text" (their slide-by-slide text was the topic) had no answer, the chat never saw it (2026-11-10..17)
      chatslide!.brief?.().then((b) => {
        if (!b?.text?.trim()) return
        s.brief = { text: b.text.slice(0, 16000), pages: Number(b.pages) > 0 ? Number(b.pages) : undefined }
        save()
      }, () => undefined)
      if (chatslide!.files && s.projectFiles !== 2) {
        chatslide!.files().then(files => {
          if (s.projectFiles === true && files.length) s.docs = s.docs.filter(d => !d.project)
          const add = files.filter(f => f.text.trim() && !s.docs.some(d => d.name === f.name))
          s.docs.push(...add.map(f => ({ name: f.name, text: f.text, project: true })))
          if (!files.length) return // none yet (still being read, or not this user's to read): asked again next open
          s.projectFiles = 2 // kept with the next save
          save()
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
