// A file, not inline in index.html: the viewer runs under script-src 'self', which blocks inline scripts.
let locale = navigator.language || 'en'
try {
  const chatslide = window.parent !== window && window.parent.__chatslide
  if (chatslide?.locale) locale = chatslide.locale
}
catch {}

const isChinese = locale.toLowerCase().startsWith('zh')
document.documentElement.lang = isChinese ? 'zh-CN' : 'en'
document.title = isChinese ? 'PPTist - 在线演示文稿' : 'PPTist - Presentation Editor'
document.querySelector('.first-screen-loading-text').textContent = isChinese ? '正在加载中，请稍等 ...' : 'Loading, please wait ...'
document.oncontextmenu = e => e.preventDefault()
