# Restyle a page with a ChatSlide template

**Result, 2026-11-20: built, measured, not shipped (reverted in 3e584398; app side 05f58a0cc).** Design asks are
about 25 of 271 real rejections (10%), not "most" as written below. On 10 real design turns replayed as a paid user,
the build without restyle passed 22 of 30 and the build with it 10 of 30 (before its last three fixes); restyled
pages lost the deck's own colours and chart pages were skipped. Those 10 turns all came from free users.

Why: 46% of judged real editor turns are accepted by the user (accept.py, 2026-11-10..19). The rejections are spread
thin; most are design asks ("more creative", "match the other slides", "more atmospheric", "vertical") on decks whose
design is one full-slide picture under text boxes. The editor can only fade, hide or recolour that picture. Prompt
fixes and model swaps moved 41 real rejected turns by a few points at most (luna 20.3, terra 21.7 of 41, 3 runs each).

## The step

`restyle_page` (router tool): rebuild the named pages with a ChatSlide template, palette and layout, keeping each
page's current text.

1. PPTist `src/hooks/useImport.ts`: split `importPPTXFile` (line ~651) so a parse returns `Slide[]` without inserting
   (`pptxToSlides(buf)`); the panel swaps the page in, keeping its id.
2. PPTist `src/embed.ts`: host hook `restyle?(page, { template, palette, layout, texts }) => Promise<string | null>`
   (a .pptx URL of that one page).
3. app.chatslide.ai `components/slides/AdvancedEditorFrame.tsx`: implement `restyle`. Take the page's classic slide
   from `projects/<id>/classic.json` (saved by `pptx()`, line ~285; the project rows are `Blank_layout` after `sync()`),
   or a fresh classic slide for a page added in the editor; set `template` / `palette` / `layout`; put the editor
   page's title and body text into its `head` / `content`; `buildSlidesPptx([slide])` (`components/slides/ExportButton.tsx:46`).
   The design picture comes from `getTemplateBackground` (`components/slides/slideTemplates.tsx:1376`), templates from
   `availableTemplates` (line 1239).
4. PPTist `api/tools/deck_chat.ts` + `TemplateDeckPanel.vue`: the tool, its rule (design asks on picture-designed
   pages), and the handler; pictures the user added on the page are carried over.
5. Credits: a restyle is a costed edit like edit_page; paid-model gates stay as they are.

## Measuring it

Real turns only. The editor's replay harness has no ChatSlide host, so this step is measured with app.chatslide.ai
`scripts/deck-eval/e2e_app.py` on dev.chatslide.ai: copy each real rejected design turn's saved deck into a test
project, replay the message, grade with `regrade.py`, 3 runs against the shipped build. Then accept.py on production
for a week after the deploy.
