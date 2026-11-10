<template>
  <div class="mobile-chat">
    <div class="deck">
      <div class="stage" ref="stageRef" @touchstart="swipeStart" @touchend="swipeEnd">
        <ThumbnailSlide v-if="size" :slide="currentSlide" :size="size" />
      </div>
      <div class="bar">
        <span class="nav" :class="{ disable: slideIndex === 0 }" :aria-label="t('上一页')" @click="go(slideIndex - 1)"><i-icon-park-outline:left /></span>
        <span class="count">{{ slideIndex + 1 }} / {{ slides.length }}</span>
        <span class="nav" :class="{ disable: slideIndex === slides.length - 1 }" :aria-label="t('下一页')" @click="go(slideIndex + 1)"><i-icon-park-outline:right /></span>
        <span class="grow"></span>
        <span class="action" :aria-label="t('编辑')" @click="changeMode('editor')"><i-icon-park-outline:edit /> <span class="label">{{ t('编辑') }}</span></span>
        <span class="action" :aria-label="t('播放')" @click="changeMode('player')"><i-icon-park-outline:full-screen-play /> <span class="label">{{ t('播放') }}</span></span>
        <!-- phones had no way to download or share but asking the chat (2026-11-10) -->
        <span class="nav" :aria-label="t('导出')" @click="download()"><i-icon-park-outline:download /></span>
        <span class="nav" v-if="host?.share" :aria-label="t('分享')" @click="shareDeck(t, m => message.success(m))"><i-icon-park-outline:share /></span>
        <a class="nav" href="https://github.com/drlambda-ai/PPTist" target="_blank" rel="noopener" aria-label="Source code (AGPL-3.0)"><i-icon-park-outline:github /></a>
      </div>
    </div>
    <TemplateDeckPanel inline />
  </div>
</template>

<script lang="ts" setup>
import { onMounted, onUnmounted, ref, useTemplateRef } from 'vue'
import { storeToRefs } from 'pinia'
import { useSlidesStore } from '@/store'
import { t } from '@/i18n'
import { host, shareDeck } from '@/embed'
import message from '@/utils/message'
import useExport from '@/hooks/useExport'
import type { Mode } from '@/types/mobile'

import ThumbnailSlide from '@/views/components/ThumbnailSlide/index.vue'
import TemplateDeckPanel from '../Editor/TemplateDeckPanel.vue'

defineProps<{
  changeMode: (mode: Mode) => void
}>()

const slidesStore = useSlidesStore()
const { slides, slideIndex, currentSlide, viewportRatio } = storeToRefs(slidesStore)

const go = (i: number) => {
  if (i >= 0 && i < slides.value.length) slidesStore.updateSlideIndex(i)
}

let startX = 0
const swipeStart = (e: TouchEvent) => startX = e.changedTouches[0].clientX
const swipeEnd = (e: TouchEvent) => {
  const dx = e.changedTouches[0].clientX - startX
  if (Math.abs(dx) > 40) go(slideIndex.value + (dx < 0 ? 1 : -1))
}

// the slide fits the space it has, in either orientation
const stageRef = useTemplateRef<HTMLElement>('stageRef')
const size = ref(0)
const fit = new ResizeObserver(([e]) => {
  size.value = Math.floor(Math.min(e.contentRect.width, e.contentRect.height / viewportRatio.value))
})
onMounted(() => fit.observe(stageRef.value!))
onUnmounted(() => fit.disconnect())

const { exportPPTX } = useExport()
const download = () => {
  if (host && !host.canExport()) return // ChatSlide opens its upgrade prompt
  exportPPTX(slides.value, true, false)
}
</script>

<style lang="scss" scoped>
.mobile-chat {
  height: 100%;
  display: flex;
  flex-direction: column;
  background: #fff;
}
.deck {
  flex: 0 0 42%;
  min-height: 0;
  display: flex;
  flex-direction: column;
  background: #f4f4f5;
  padding: 8px 8px 0;
}
.stage {
  flex: 1;
  min-height: 0;
  display: flex;
  justify-content: center;
  align-items: center;
  overflow: hidden;
}
.bar {
  height: 44px;
  flex-shrink: 0;
  display: flex;
  align-items: center;
  gap: 4px;
  font-size: 14px;
}
.nav, .action {
  color: inherit;
  min-width: 44px;
  height: 44px;
  display: inline-flex;
  justify-content: center;
  align-items: center;
  gap: 4px;
}
.action {
  padding: 0 6px;
}
.disable {
  opacity: .3;
}
.grow {
  flex: 1;
}
.count {
  white-space: nowrap;
}
// with download and share the bar holds 7 buttons: narrow phones show icons only
@media (max-width: 480px) {
  .bar {
    gap: 0;
  }
  .nav, .action {
    min-width: 40px;
  }
  .action .label {
    display: none;
  }
}
@media (orientation: landscape) {
  .mobile-chat {
    flex-direction: row;
  }
  .deck {
    flex: 0 0 55%;
  }
}
</style>
