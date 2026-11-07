// A file, not inline in index.html: the viewer runs under script-src 'self', which blocks inline scripts.
let locale = navigator.language || 'en'
try {
  const chatslide = window.parent !== window && window.parent.__chatslide
  if (chatslide?.locale) locale = chatslide.locale
}
catch { /* A standalone or cross-origin embed uses the browser locale. */ }

document.documentElement.lang = locale
document.title = 'ChatSlide'
document.oncontextmenu = e => e.preventDefault()
