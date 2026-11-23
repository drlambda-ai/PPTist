import supported from './locales.json'

export function resolveLocale(value: string): string {
  const tag = value.toLowerCase().replace(/_/g, '-')
  if (tag === 'zh-tw' || tag === 'zh-hk' || tag.startsWith('zh-hant')) return 'zh-TW'
  if (tag === 'zh' || tag.startsWith('zh-')) return 'zh-CN'
  if (tag === 'pt' || tag.startsWith('pt-')) return 'pt-BR'
  return supported.find(locale => locale.toLowerCase() === tag || locale === tag.split('-')[0]) || 'en'
}

// vue-i18n's plural messages: "Deleted {p0} page; … | Deleted {p0} pages; …", one form per CLDR category of the locale,
// picked by the count the call passes (t(key, parameters, n)). A message with 2 forms is one | other in any locale; one
// with more lists the locale's categories in CLDR order (uk, ru: one | few | many; ar: zero | one | two | few | many |
// other); one form serves every count (zh, vi, id). The editor chat said "Видалено 1 слайдів" from a "page(s)" string
// that had one form for every count (replay, 2026-11-23).
const ORDER = ['zero', 'one', 'two', 'few', 'many', 'other']
export function pluralIndex(locale: string, n: number, forms: number): number {
  if (forms < 2) return 0
  const rules = new Intl.PluralRules(locale)
  const category = rules.select(n)
  if (forms === 2) return category === 'one' ? 0 : 1
  const categories = ORDER.filter(c => rules.resolvedOptions().pluralCategories.includes(c as Intl.LDMLPluralRule))
  return Math.min(Math.max(categories.indexOf(category), 0), forms - 1)
}

export function createTranslator(english: Record<string, string>, messages: Record<string, string>, locale = 'en') {
  return (key: string, parameters: Record<string, string | number> = {}, n?: number): string => {
    const own = messages[key]
    let value = own ?? english[key] ?? key
    const forms = value.split(' | ')
    // no count: the last form, the plural (a call without one is a bug the i18n check reports)
    if (forms.length > 1) value = forms[n === undefined ? forms.length - 1 : pluralIndex(own === undefined ? 'en' : locale, Math.abs(n), forms.length)]
    return value.replace(/\{(p\d+)\}/g, (placeholder, name: string) =>
      Object.prototype.hasOwnProperty.call(parameters, name) ? String(parameters[name]) : placeholder)
  }
}
