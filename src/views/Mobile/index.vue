<template>
  <div class="mobile">
    <component 
      v-if="!host || mode !== 'preview'"
      :is="currentComponent" 
      :changeMode="changeMode"
    ></component>
    <!-- inside ChatSlide it stays mounted: its chat panel loads, saves and syncs the deck -->
    <MobileChat v-if="host" v-show="mode === 'preview'" :changeMode="changeMode" />
  </div>
</template>

<script lang="ts" setup>
import { computed, ref } from 'vue'
import type { Mode } from '@/types/mobile'

import MobileEditor from './MobileEditor/index.vue'
import MobilePlayer from './MobilePlayer.vue'
import MobilePreview from './MobilePreview.vue'
import MobileChat from './MobileChat.vue'
import { host } from '@/embed'

const mode = ref<Mode>('preview')
    
const changeMode = (_mode: Mode) => mode.value = _mode

const currentComponent = computed(() => {
  const componentMap = {
    'editor': MobileEditor,
    'player': MobilePlayer,
    'preview': MobilePreview,
  }
  return componentMap[mode.value] || null
})
</script>

<style lang="scss" scoped>
.mobile {
  height: 100%;

  :deep(textarea), :deep(input) {
    font-size: 16px; // below 16px iOS zooms the page on focus
  }
}
</style>