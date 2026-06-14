<template>
  <MoveablePanel
    class="chat-panel"
    :width="320"
    :height="480"
    title="AI Assistant"
    :left="-340"
    :top="60"
    @close="close()"
  >
    <div class="chat-body">
      <div class="messages" ref="messagesRef">
        <div class="hint" v-if="!messages.length">
          Tell me how to edit this slide. Examples:
          <ul>
            <li>"Make the title bigger and purple"</li>
            <li>"Add a subtitle under the title"</li>
            <li>"Move the image to the right"</li>
          </ul>
          Or type a description and hit <b>Image</b> to generate a picture.
        </div>
        <div
          class="msg"
          :class="msg.role"
          v-for="(msg, index) in messages"
          :key="index"
        >{{ msg.content }}</div>
        <div class="msg assistant loading" v-if="loading">…</div>
      </div>

      <div class="composer">
        <Input
          v-model:value="input"
          placeholder="Edit this slide…"
          @enter="send()"
        />
        <div class="actions">
          <Button :disabled="loading" @click="genImage()">Image</Button>
          <Button type="primary" :disabled="loading" @click="send()">Send</Button>
        </div>
      </div>
    </div>
  </MoveablePanel>
</template>

<script lang="ts" setup>
import { nextTick, ref } from 'vue'
import { storeToRefs } from 'pinia'
import { useMainStore, useSlidesStore } from '@/store'
import useCreateElement from '@/hooks/useCreateElement'
import useHistorySnapshot from '@/hooks/useHistorySnapshot'
import api from '@/services'
import Button from '@/components/Button.vue'
import Input from '@/components/Input.vue'
import MoveablePanel from '@/components/MoveablePanel.vue'

interface ChatMessage {
  role: 'user' | 'assistant'
  content: string
}

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const { currentSlide } = storeToRefs(slidesStore)
const { viewportSize, viewportRatio } = storeToRefs(slidesStore)
const { addHistorySnapshot } = useHistorySnapshot()
const { createImageElement } = useCreateElement()

const messages = ref<ChatMessage[]>([])
const input = ref('')
const loading = ref(false)
const messagesRef = ref<HTMLElement>()

const close = () => mainStore.setChatPanelState(false)

const scrollBottom = () => {
  nextTick(() => {
    if (messagesRef.value) messagesRef.value.scrollTop = messagesRef.value.scrollHeight
  })
}

// Merge AI-returned elements onto the current slide: keep all original fields for elements
// the AI kept (matched by id), use new ones as-is, drop ones the AI omitted (= deletions).
const applyEdits = (aiElements: any[]) => {
  const original = currentSlide.value?.elements || []
  const byId: Record<string, any> = {}
  for (const el of original) byId[el.id] = el
  const merged = aiElements.map(ae => (byId[ae.id] ? { ...byId[ae.id], ...ae } : ae))
  slidesStore.updateSlide({ elements: merged })
  addHistorySnapshot()
}

const send = async () => {
  const text = input.value.trim()
  if (!text || loading.value) return
  messages.value.push({ role: 'user', content: text })
  input.value = ''
  loading.value = true
  scrollBottom()
  try {
    const ret = await api.AI_Edit({
      elements: currentSlide.value?.elements || [],
      instruction: text,
      history: messages.value.slice(-6),
      slideSize: {
        width: viewportSize.value,
        height: Math.round(viewportSize.value * viewportRatio.value),
      },
    })
    messages.value.push({ role: 'assistant', content: ret.reply || 'Done.' })
    if (ret.elements) applyEdits(ret.elements)
  }
  catch (e) {
    messages.value.push({ role: 'assistant', content: 'Sorry, that failed: ' + e })
  }
  loading.value = false
  scrollBottom()
}

const genImage = async () => {
  const text = input.value.trim()
  if (!text || loading.value) return
  messages.value.push({ role: 'user', content: '🎨 ' + text })
  input.value = ''
  loading.value = true
  scrollBottom()
  try {
    const ret = await api.AI_Image({ prompt: text })
    if (ret.src) {
      const sel = mainStore.handleElement
      if (sel && sel.type === 'image') {
        slidesStore.updateElement({ id: sel.id, props: { src: ret.src } })
        addHistorySnapshot()
        messages.value.push({ role: 'assistant', content: 'Replaced the selected image.' })
      }
      else {
        createImageElement(ret.src)
        messages.value.push({ role: 'assistant', content: 'Added the generated image.' })
      }
    }
  }
  catch (e) {
    messages.value.push({ role: 'assistant', content: 'Image generation failed: ' + e })
  }
  loading.value = false
  scrollBottom()
}
</script>

<style lang="scss" scoped>
.chat-body {
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

  ul {
    margin: 6px 0;
    padding-left: 16px;
  }
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
  justify-content: flex-end;
  gap: 8px;
}
</style>
