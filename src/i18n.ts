import host from '@/embed'
import { EN } from '@/i18n-en'
import { createTranslator, resolveLocale } from './i18n-core'

export let lang = resolveLocale(host?.locale || navigator.language || 'en')
export let direction = ['ar', 'fa'].includes(lang) ? 'rtl' : 'ltr'
let translate = createTranslator(EN, EN)
const catalogs = import.meta.glob<{ default: Record<string, string> }>(['./locales/*.json', '!./locales/en.json'])

// The app's strings in another of its locales, by a language code ("vi", "uk", "zh-CN"); null for a language it has
// no strings in. The editor chat says its own lines in the language of the user's message with it.
export async function translatorIn(code: string): Promise<typeof t | null> {
  const locale = resolveLocale(code)
  if (locale === 'en') return code.toLowerCase().startsWith('en') ? createTranslator(EN, EN) : null
  const load = catalogs[`./locales/${locale}.json`]
  if (!load) return null
  try {
    return createTranslator(EN, (await load()).default, locale)
  }
  catch {
    return null
  }
}

export async function initLocale() {
  if (lang !== 'en') {
    const load = catalogs[`./locales/${lang}.json`]
    if (load) {
      try {
        const messages = await load()
        translate = createTranslator(EN, messages.default, lang)
      }
      catch {
        lang = 'en'
        direction = 'ltr'
      }
    }
  }
  document.documentElement.lang = lang
}

// n: the count a plural message ("a | b | c") picks its form by
export function t(key: string, parameters?: Record<string, string | number>, n?: number): string {
  return translate(key, parameters, n)
}
