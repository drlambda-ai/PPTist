# Editor localization

The host's `locale` selects the interface language. Standalone sessions use the browser language.
`src/locales.json` lists the supported locales in this prioritized rollout; other locales fall back to English.
Presentation content continues to follow the presentation/request language independently.

`src/locales/en.json` is the source catalog. Existing Chinese identifiers are retained for editor controls;
chat messages use their English source as identifiers. Use `t(key, { p0: value })` for interpolation.
Catalog loading finishes before component imports, so formula, animation, and panel labels use the selected locale.
Only the selected locale bundle is downloaded. Interface direction is scoped to panels to preserve canvas coordinates.

Translate catalog values with Codex and preserve keys and placeholders. Add a supported locale to the registry
only with a complete catalog. `npm run check:i18n` verifies every registered catalog, interpolation parameters,
and literal translation calls; `npm run build-only` runs it before Vite. Run `node --test tools/i18n.test.mjs`
for locale resolution and interpolation behavior.

For local visual checks, open `/tools/fixtures/locale.html?locale=vi` (also `es` and `ar`).
The fixture exercises the real editor with the same host locale boundary and disables paid API calls.
It does not replace an authenticated app `/slides` integration check.
Check the toolbar, design/animation/formula panels, and both chat panels, including long labels and RTL text.

Prioritize from identified editor users in PostHog: exact `app.chatslide.ai` host, `$pageview`,
`$pathname = /slides`, last 90 days, and a non-null person email. Group by `person_id`, take the
latest `$browser_language`, then aggregate by locale family. This excludes public tools traffic.
Keep current counts in query results, not this document.
