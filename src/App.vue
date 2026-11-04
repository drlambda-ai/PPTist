<template>
  <template v-if="slides.length">
    <Screen v-if="screening || viewDeck" />
    <Editor v-else-if="_isPC" />
    <Mobile v-else />
  </template>
  <FullscreenSpin :tip="t('数据初始化中，请稍等 ...')" v-else  loading :mask="false" />
</template>

<script lang="ts" setup>
import { onMounted } from 'vue'
import { storeToRefs } from 'pinia'
import { nanoid } from 'nanoid'
import { useScreenStore, useMainStore, useSnapshotStore, useSlidesStore } from '@/store'
import { LOCALSTORAGE_KEY_DISCARDED_DB } from '@/configs/storage'
import { deleteDiscardedDB } from '@/utils/database'
import { isPC } from '@/utils/common'
import api from '@/services'
import { host } from '@/embed'

import Editor from './views/Editor/index.vue'
import Screen from './views/Screen/index.vue'
import Mobile from './views/Mobile/index.vue'
import FullscreenSpin from '@/components/FullscreenSpin.vue'

const _isPC = isPC()

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const snapshotStore = useSnapshotStore()
const screenStore = useScreenStore()
const { databaseId } = storeToRefs(mainStore)
const { slides } = storeToRefs(slidesStore)
const { screening } = storeToRefs(screenStore)

const isAudienceMode = new URLSearchParams(window.location.search).get('mode') === 'audience'
// ?mode=view&deck=<saved deck link>: a shared ChatSlide deck, read-only, in the presenter (no editor)
const viewParams = new URLSearchParams(window.location.search)
const viewDeck = viewParams.get('mode') === 'view' ? viewParams.get('deck') || '' : ''

if (import.meta.env.MODE !== 'development' && !host) { // inside ChatSlide edits are saved continuously
  window.onbeforeunload = () => false
}

// the generated template library (tools/templates/generate.mjs) joins the built-in templates in the picker
const loadGeneratedTemplates = async () => {
  const list: { id: string, name: string }[] = await api.getMockData('gen/index').catch(() => [])
  slidesStore.setTemplates([...slidesStore.templates, ...list.map(x => ({ id: `gen/${x.id}`, name: x.name, cover: `./mocks/gen/${x.id}.webp`, origin: 'ChatSlide' }))])
}

onMounted(async () => {
  loadGeneratedTemplates()
  if (isAudienceMode) {
    slidesStore.setSlides([{
      id: nanoid(10),
      elements: [],
    }])
    screenStore.setScreening(true)
  }
  else if (viewDeck) {
    // only our own deck storage, so the viewer can't be pointed at an arbitrary URL
    const url = new URL(viewDeck)
    if (url.protocol !== 'https:' || !url.hostname.endsWith('.public.blob.vercel-storage.com')) return
    const saved = await (await fetch(url)).json()
    if (saved.theme) slidesStore.setTheme(saved.theme)
    slidesStore.setSlides(saved.slides || [])
    screenStore.setScreening(true)
  }
  else {
    // inside ChatSlide: a blank page so the editor mounts at once; it then loads the project's deck (no demo-deck fetch)
    slidesStore.setSlides(host ? [{ id: nanoid(10), elements: [] }] : await api.getMockData('slides'))

    await deleteDiscardedDB()
    snapshotStore.initSnapshotDatabase()
  }
})

// 应用注销时向 localStorage 中记录下本次 indexedDB 的数据库ID，用于之后清除数据库
window.addEventListener('beforeunload', () => {
  const discardedDB = localStorage.getItem(LOCALSTORAGE_KEY_DISCARDED_DB)
  const discardedDBList: string[] = discardedDB ? JSON.parse(discardedDB) : []

  discardedDBList.push(databaseId.value)

  const newDiscardedDB = JSON.stringify(discardedDBList)
  localStorage.setItem(LOCALSTORAGE_KEY_DISCARDED_DB, newDiscardedDB)
})
</script>

<style lang="scss">
#app {
  height: 100%;
}
</style>