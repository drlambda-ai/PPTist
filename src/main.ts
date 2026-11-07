import '@/embed'
import { initLocale } from './i18n'

// Catalogs must load before components and module-level panel labels evaluate t().
initLocale().then(() => import('./bootstrap'))
