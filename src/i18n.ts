import host from '@/embed'
import { EN } from '@/i18n-en'

export const lang = (host?.locale || navigator.language || 'en').toLowerCase().startsWith('zh') ? 'zh' : 'en'

export function t(zh: string): string {
  return lang === 'en' ? (EN[zh] ?? zh) : zh
}
