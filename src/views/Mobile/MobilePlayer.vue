<template>
  <div class="mobile-player" 
    :style="{
      width: playerSize.width + 'px',
      height: playerSize.height + 'px',
      transform: landscape ? 'none' : `rotate(90deg) translateY(-${playerSize.height}px)`,
    }"
  >
    <div 
      class="screen-slide-list" 
      @click="toolVisible = !toolVisible"
      @touchstart="$event => touchStartListener($event)"
      @touchend="$event => touchEndListener($event)"
    >
      <div 
        :class="[
          'slide-item', 
          `turning-mode-${slide.turningMode || 'slideY'}`,
          {
            'current': index === slideIndex,
            'before': index < slideIndex,
            'after': index > slideIndex,
            'hide': (index === slideIndex - 1 || index === slideIndex + 1) && slide.turningMode !== slidesWithTurningMode[slideIndex].turningMode,
            'last': index === slideIndex - 1,
            'next': index === slideIndex + 1,
          }
        ]"
        v-for="(slide, index) in slidesWithTurningMode" 
        :key="slide.id"
      >
        <div 
          class="slide-content" 
          :style="{
            width: slideSize.width + 'px',
            height: slideSize.height + 'px',
          }"
          v-if="Math.abs(slideIndex - index) < 2"
        >
          <ThumbnailSlide 
            :slide="slide" 
            :size="slideSize.width" 
          />
        </div>
      </div>
    </div>

    <template v-if="toolVisible">
      <div class="header">
        <div class="back" @click="changeMode('preview')"><i-icon-park-outline:logout /> {{ t('退出播放') }}</div>
      </div>
      <MobileThumbnails class="thumbnails" />
    </template>
  </div>
</template>

<script lang="ts" setup>
import { host } from '@/embed'
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { storeToRefs } from 'pinia'
import { useSlidesStore } from '@/store'
import type { Mode } from '@/types/mobile'
import useSlidesWithTurningMode from '../Screen/hooks/useSlidesWithTurningMode'

import ThumbnailSlide from '@/views/components/ThumbnailSlide/index.vue'
import MobileThumbnails from './MobileThumbnails.vue'

defineProps<{
  changeMode: (mode: Mode) => void
}>()

const slidesStore = useSlidesStore()
const { slides, slideIndex, viewportRatio } = storeToRefs(slidesStore)

const { slidesWithTurningMode } = useSlidesWithTurningMode()

const toolVisible = ref(false)

const playerSize = ref({ width: 0, height: 0 })
const landscape = ref(false) // a phone held sideways already shows the slide the right way up

const measure = () => {
  const { clientWidth: w, clientHeight: h } = document.body
  landscape.value = w > h
  playerSize.value = landscape.value ? { width: w, height: h } : { width: h, height: w }
}

onMounted(() => {
  if (slideIndex.value !== 0 && !host) slidesStore.updateSlideIndex(0) // inside ChatSlide: stay on the page being worked on

  measure()
  window.addEventListener('resize', measure)
})
onUnmounted(() => window.removeEventListener('resize', measure))

const slideSize = computed(() => {
  const playerRatio = playerSize.value.height / playerSize.value.width

  let slideWidth = 0
  let slideHeight = 0

  if (playerRatio >= viewportRatio.value) {
    slideWidth = playerSize.value.width
    slideHeight = slideWidth * viewportRatio.value
  }
  else {
    slideHeight = playerSize.value.height
    slideWidth = slideHeight / viewportRatio.value
  }

  return {
    width: slideWidth,
    height: slideHeight,
  }
})

const touchInfo = ref<{ x: number; y: number; } | null>(null)
const touchStartListener = (e: TouchEvent) => {
  touchInfo.value = {
    x: e.changedTouches[0].pageX,
    y: e.changedTouches[0].pageY,
  }
}
const touchEndListener = (e: TouchEvent) => {
  if (!touchInfo.value) return

  const offsetX = e.changedTouches[0].pageX - touchInfo.value.x
  const offsetY = e.changedTouches[0].pageY - touchInfo.value.y
  const offsetAbsX = Math.abs(offsetX)
  const offsetAbsY = Math.abs(offsetY)

  if ( offsetAbsX > offsetAbsY && offsetAbsX > 50 ) {
    const toLeft = landscape.value ? offsetX > 0 : offsetX < 0
    if (toLeft && slideIndex.value > 0) slidesStore.updateSlideIndex(slideIndex.value - 1)
    if (!toLeft && slideIndex.value < slides.value.length - 1) slidesStore.updateSlideIndex(slideIndex.value + 1)
  }

  if ( offsetAbsY > offsetAbsX && offsetAbsY > 50 ) {
    if (offsetY > 0 && slideIndex.value > 0) slidesStore.updateSlideIndex(slideIndex.value - 1)
    if (offsetY < 0 && slideIndex.value < slides.value.length - 1) slidesStore.updateSlideIndex(slideIndex.value + 1)
  }
}
</script>

<style lang="scss" scoped>
@import '@/assets/styles/slideAnimation.scss';

.mobile-player {
  transform-origin: 0 0;
  background-color: #1d1d1d;
  position: relative;
}
.screen-slide-list {
  position: relative;
  width: 100%;
  height: 100%;
}
.slide-item {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;

  &:not(.last, .next) {
    z-index: -1;
  }

  &.current {
    z-index: 2;
  }

  &.hide {
    opacity: 0;
  }

  @include slide-animation();
}
.slide-content {
  background-color: #fff;
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  display: flex;
  justify-content: center;
  align-items: center;
}

.header {
  width: 100%;
  height: 40px;
  line-height: 40px;
  padding: 0 15px;
  position: absolute;
  top: 0;
  left: 0;
  z-index: 99;
  background-color: rgba($color: #1d1d1d, $alpha: .7);
  text-align: right;
  font-size: 13px;
  color: #fff;
  animation: slideInDown .15s;

  .back {
    height: 100%;
  }
}
.thumbnails {
  width: 100%;
  position: absolute;
  bottom: 0;
  left: 0;
  z-index: 99;
  background-color: rgba($color: #1d1d1d, $alpha: .7);
  overflow: auto !important;
  animation: slideInUp .15s;
}

@keyframes slideInUp {
  from {
    transform: translateY(100%);
  }
  to {
    transform: translateY(0);
  }
}
@keyframes slideInDown {
  from {
    transform: translateY(-100%);
  }
  to {
    transform: translateY(0);
  }
}
</style>