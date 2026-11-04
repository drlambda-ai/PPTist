"""Cover thumbnails for the generated templates, rendered by PPTist itself (fonts and shapes exactly as in the editor).

  PPTIST_URL=https://pptist-five.vercel.app PPTIST_BYPASS=... python3 tools/templates/render_covers.py [--only gen_001,...]
  -> public/mocks/gen/<id>.webp (480x270)
"""
import io, json, os, sys
from pathlib import Path
from PIL import Image
from playwright.sync_api import sync_playwright

GEN = Path(__file__).resolve().parents[2] / 'public' / 'mocks' / 'gen'
PINIA = "document.querySelector('#app').__vue_app__.config.globalProperties.$pinia"
url = os.environ.get('PPTIST_URL', 'https://pptist-five.vercel.app')
bypass = os.environ.get('PPTIST_BYPASS', '')
only = sys.argv[sys.argv.index('--only') + 1].split(',') if '--only' in sys.argv else None
ids = [x['id'] for x in json.loads((GEN / 'index.json').read_text()) if not only or x['id'] in only]

with sync_playwright() as pw:
    b = pw.chromium.launch()
    page = b.new_page(viewport={'width': 1600, 'height': 950})
    page.goto(f'{url}/?x-vercel-protection-bypass={bypass}&x-vercel-set-bypass-cookie=true' if bypass else url, timeout=90000)
    page.wait_for_function(f"() => {PINIA}._s.get('slides')?.slides.length > 0", timeout=90000)
    page.add_style_tag(content='.template-deck-panel { display: none !important }')
    for t in ids:
        d = json.loads((GEN / f'{t}.json').read_text())
        page.evaluate(f"d => {{ const st = {PINIA}._s.get('slides'); st.setTheme(d.theme); st.setSlides(d.slides); st.updateSlideIndex(0) }}", d)
        page.wait_for_timeout(600)  # fonts and shapes settle
        png = page.locator('.viewport-wrapper').first.screenshot()
        Image.open(io.BytesIO(png)).convert('RGB').resize((480, 270), Image.LANCZOS).save(GEN / f'{t}.webp', 'WEBP', quality=80)
    b.close()
print(f'{len(ids)} covers -> {GEN}')
