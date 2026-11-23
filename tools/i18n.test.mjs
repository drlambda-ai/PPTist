import { test } from 'node:test'
import { Buffer } from 'node:buffer'
import assert from 'node:assert/strict'
import { build } from 'esbuild'
const built = await build({ entryPoints: ['src/i18n-core.ts'], bundle: true, write: false, format: 'esm', platform: 'node' })
const { resolveLocale, createTranslator, pluralIndex } = await import(`data:text/javascript;base64,${Buffer.from(built.outputFiles[0].text).toString('base64')}`)

test('resolves host locale variants and falls back to English for unsupported locales', () => {
  for (const [input, expected] of Object.entries({ 'vi-VN': 'vi', 'es-419': 'es', 'fa-IR': 'fa', 'id-ID': 'id', 'ru-RU': 'ru', 'ar-EG': 'ar', 'uk-UA': 'uk', 'pt-PT': 'pt-BR', 'zh-Hant': 'zh-TW', 'zh_HK': 'zh-TW', 'zh-Hans': 'zh-CN', 'fr-FR': 'fr', 'de-DE': 'en', '': 'en' })) assert.equal(resolveLocale(input), expected)
})
test('uses catalog text, preserves supplied values, and falls back per missing key', () => {
  const t = createTranslator({ done: 'Saved {p0}', fallback: 'Retry' }, { done: 'Đã lưu {p0}' })
  assert.equal(t('done', { p0: 'title {p1} $&' }), 'Đã lưu title {p1} $&')
  assert.equal(t('fallback'), 'Retry')
  assert.equal(t('unknown'), 'unknown')
  assert.equal(t('done'), 'Đã lưu {p0}')
})
test('plural messages pick the form of the locale by the count', () => {
  const uk = createTranslator({ gone: 'Deleted {p0} page(s).' }, { gone: 'Видалено {p0} слайд. | Видалено {p0} слайди. | Видалено {p0} слайдів.' }, 'uk')
  assert.deepEqual([1, 2, 5, 11, 21, 22].map(n => uk('gone', { p0: n }, n)), ['Видалено 1 слайд.', 'Видалено 2 слайди.', 'Видалено 5 слайдів.', 'Видалено 11 слайдів.', 'Видалено 21 слайд.', 'Видалено 22 слайди.'])
  const en = createTranslator({ gone: 'Deleted {p0} page. | Deleted {p0} pages.' }, {}, 'vi')
  assert.deepEqual([1, 2].map(n => en('gone', { p0: n }, n)), ['Deleted 1 page.', 'Deleted 2 pages.'])
  assert.equal(en('gone', { p0: 3 }), 'Deleted 3 pages.')
  assert.deepEqual([0, 1, 2, 3, 11, 100].map(n => pluralIndex('ar', n, 6)), [0, 1, 2, 3, 4, 5])
  assert.equal(pluralIndex('ar', 2, 2), 1)
  assert.equal(pluralIndex('zh-CN', 1, 1), 0)
})
