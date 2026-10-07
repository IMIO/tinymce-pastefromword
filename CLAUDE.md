# tinymce-pastefromword

TinyMCE plugin that cleans content pasted from Microsoft Word. It is a port of the CKEditor 4.22.1
`pastefromword` and `pastetools` plugins (the last open-source CKEditor 4 release). It is the free
counterpart of the premium TinyMCE PowerPaste plugin for Word content.

Published as `@imiobe/tinymce-pastefromword` on npm.

## Commands

- `make install` / `npm install` — Install dependencies
- `make dev` / `npm run dev` — Vite dev server with the integration page in `dev/` (port 5173)
- `make build` / `npm run build` — Build the library to `dist/` (ES + UMD)
- `make test` / `npm test` — Run the Node test runner (`node --test`, jsdom provides `DOMParser`)
- `make test-corpus` / `npm run test:corpus` — Run the invariant checks over the Word fixtures.
  Set `PFW_CORPUS_DIR` to a directory of HTML files to run them over another corpus
  (for example all 897 inputs of the CKEditor repository under `tests/plugins/pastefromword/generated/_fixtures`)
- `make lint` / `npm run lint` — ESLint

## Architecture

- `src/plugin.js` — TinyMCE plugin: options, paste capture, `PastePreProcess` hook, blob cache images, button, events
- `src/i18n.js` — Translations of the plugin strings
- `src/filter/index.js` — Framework-free entry: `cleanWord()`, `isWordContent()`, `embedRtfImages()`
- `src/filter/tree.js` — Node tree with the shape of the CKEditor `htmlParser` nodes, built from a `DOMParser` document
- `src/filter/filter.js` — Rule-based filter engine (port of `htmlParser.filter`)
- `src/filter/word.js` — Word rules and `cleanWord()` pipeline (port of `pastefromword/filter/default.js`)
- `src/filter/lists.js` — List rebuilding and heuristics
- `src/filter/common.js` — Shared rules and style helpers (port of `pastetools/filter/common.js`)
- `src/filter/inliner.js` — Style sheet inliner with its own small CSS parser
- `src/filter/image.js`, `src/filter/rtf.js` — RTF parsing and image extraction (port of `pastetools/filter/image.js`)
- `src/filter/css.js` — CSS text, unit and border helpers (port of the needed `CKEDITOR.tools` functions)
- `test/` — Node tests. `test/fixtures/word` and `test/fixtures/rtf` come from the CKEditor repository
- `dev/` — Integration page with a real TinyMCE instance

## Pipeline

1. The `paste` listener (registered on `PreInit`, before the core listener) stores the `text/html` and `text/rtf`
   clipboard data when the HTML comes from Office or when the toggle is on.
2. The `PastePreProcess` listener (registered on `PreInit`, after the core WebKit style filter) replaces
   `e.content` with the cleaned HTML. It matches the event with the stored clipboard data by comparing the
   trimmed HTML, so plain text pastes are left alone.
3. `cleanWord()` inlines the style sheets, fixes the conditional comments, builds the node tree, applies the
   common and Word rules, rebuilds the lists and runs a final cleanup.
4. `embedRtfImages()` replaces the `file://` image sources with the images found in the RTF clipboard.

## Conventions

- Vanilla ES modules, no build step needed to run the filter in Node
- Keep the CKEditor license headers on the ported files. The package is tri-licensed (GPL, LGPL, MPL) like CKEditor 4
- Keep the port close to the CKEditor code. Deviations are listed in the README under "Differences from CKEditor 4"
  and must stay documented there
- Comments in Simplified Technical English, no ticket numbers and no history in comments
- Tests compare normalized HTML (`normalizeHtml()` in `test/setup.js`): lower-case, sorted attributes and styles,
  no whitespace at block boundaries
