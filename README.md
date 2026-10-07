# @imiobe/tinymce-pastefromword

TinyMCE plugin that cleans up content pasted from Microsoft Word: lists, tables, styles and images.
It is a port of the open-source [CKEditor 4](https://github.com/ckeditor/ckeditor4) "Paste from Word"
plugin (version 4.22.1, the last release under the GPL/LGPL/MPL licenses) and a free alternative to the
premium TinyMCE PowerPaste plugin for Word content.

## Features

- Detects Office content in the clipboard (`Generator` meta tag, `Mso*` classes, `mso-*` styles).
- Rebuilds Word lists as `<ol>` and `<ul>`: nesting levels, bullet, decimal, alpha and roman numbering,
  `start` values, explicit item `value`s, lists interrupted by other lists.
- Cleans tables: borders, widths, heights, backgrounds and alignment are unified as CSS.
- Inlines the Word style sheets, then removes all Word specific markup: `mso-*` styles, `<o:p>`, VML shapes,
  field codes, bookmarks, conditional comments, `_GoBack` anchors, Word classes.
- Keeps the text formatting as nested `<span>` styles (font family, size, color, line height, alignment)
  or drops the font styles with one option.
- Embeds the images of the RTF clipboard (PNG, JPEG, GIF). The HTML clipboard of Word only points to local
  files, the pixels are in the RTF clipboard.
- Converts Word page breaks to the placeholder of the TinyMCE `pagebreak` plugin.
- Toggle button, menu item and command to force the cleanup on any pasted HTML.
- Events to inspect, cancel or change the result.

## Installation

```bash
npm install @imiobe/tinymce-pastefromword tinymce
```

## Quick start

```js
import tinymce from 'tinymce';
import '@imiobe/tinymce-pastefromword';

tinymce.init({
  selector: '#editor',
  plugins: 'lists table pastefromword',
  toolbar: 'undo redo | pastefromword | bold italic | bullist numlist',
});
```

The plugin registers itself as `pastefromword`. Content pasted from Word is cleaned automatically.
The `pastefromword` toolbar button toggles the cleanup for all pasted HTML, for content copied from
Word-like applications that is not detected.

The UMD build (`dist/tinymce-pastefromword.umd.cjs`) expects a global `tinymce` and can be loaded with a
`<script>` tag after TinyMCE.

## Options

All options are set in the TinyMCE `init()` configuration.

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `pastefromword_remove_font_styles` | `boolean` | `false` | Drop font family, font size, color, background color, line height and text decoration. |
| `pastefromword_keep_zero_margins` | `boolean` | `false` | Keep `margin-*: 0` styles instead of removing them. |
| `pastefromword_inline_images` | `boolean` | `true` | Embed the images found in the RTF clipboard. |
| `pastefromword_prompt_cleanup` | `boolean` | `false` | Ask the user with a confirm dialog before cleaning. |
| `pastefromword_heuristics_edge_list` | `boolean` | `false` | Detect lists in content pasted without Word list metadata, from the text structure and indentation. |

`font_family_formats` is also read: a `<font face>` of the pasted content is replaced by the matching font
stack of the editor.

## Toolbar button, menu item and command

- `pastefromword` toggle button and menu item: when active, every pasted HTML is cleaned, even when it is
  not detected as Office content. The state persists until it is toggled again.
- `mceTogglePasteFromWord` command toggles the state. `editor.queryCommandState('mceTogglePasteFromWord')`
  returns it.

## Events

| Event | Properties | Description |
|-------|------------|-------------|
| `PasteFromWord` | `content`, `rtf` | Fired before the cleanup with the raw clipboard HTML and RTF. Changing `content` changes the input of the filter. `preventDefault()` cancels the cleanup unless the toggle is active. |
| `AfterPasteFromWord` | `content` | Fired after the cleanup. Changing `content` changes the pasted HTML. |
| `PasteFromWordToggle` | `state` | Fired when the toggle changes. |

```js
editor.on('AfterPasteFromWord', (e) => {
  e.content = e.content.replace(/<table/g, '<table class="table"');
});
```

## API

```js
const plugin = editor.plugins.pastefromword;

plugin.cleanWord(html, rtf); // Clean Word HTML with the options of this editor. rtf is optional.
plugin.isWordContent(html); // Detect Office content.
```

The filter does not depend on TinyMCE. It only needs a `DOMParser`, so it runs in a browser or in Node
with jsdom:

```js
import { cleanWord, embedRtfImages, isWordContent } from '@imiobe/tinymce-pastefromword/filter';

let html = cleanWord(clipboardHtml, { removeFontStyles: true });
html = embedRtfImages(html, clipboardRtf); // Images become data URLs.
```

## Images

Word puts the images in the RTF clipboard. Chrome, Edge and Safari expose it as `text/rtf`; Firefox does
not, so images pasted from Word in Firefox stay `file://` references that cannot be shown.

Embedded images go through the blob cache of the editor, like dropped image files. They are uploaded with
the configured `images_upload_url` or `images_upload_handler`, and otherwise saved as data URLs. The TinyMCE
`paste_data_images` option does not apply to them. Images with `blob:` sources, as pasted by some browsers,
are left to TinyMCE.

EMF and WMF images cannot be shown by browsers and are reported with a `console.warn`.

## Interaction with the TinyMCE paste pipeline

- The cleanup runs in `PastePreProcess`, after the WebKit style filter of TinyMCE, so the styles kept by the
  filter survive `paste_remove_styles_if_webkit`.
- A `paste_preprocess` callback runs before this plugin and sees the raw clipboard HTML. Use the
  `AfterPasteFromWord` event to post-process the cleaned HTML.
- `paste_as_text` and plain text pastes (`Ctrl+Shift+V`) are not affected.
- `PastePostProcess` receives the cleaned content.
- The result still goes through the TinyMCE schema (`valid_elements`, `valid_styles`, `invalid_styles`).
  Use them to restrict the styles further, for example `valid_styles: { '*': 'text-align,list-style-type' }`.

## Differences from CKEditor 4

The filter rules are a close port of the CKEditor code, with these changes for TinyMCE, which has no content
filter equivalent to the CKEditor Advanced Content Filter:

- Word (`Mso*`) and Excel (`xl\d+`, `el\d+`, `font\d+`) classes are removed.
- Word-only CSS properties that browsers do not support (`tab-stops`, `text-autospace`, `text-underline`,
  `layout-grid-mode`, `punctuation-wrap`, `-ms-*`) are removed, and the style sheet inliner skips unknown
  properties like a browser does.
- The left margin of list items is reduced by the Word default indentation (0.5in per level) instead of the
  browser default list padding (40px per level). Default Word lists have no residual item margins.
- Vertical margins of list items move to the list element only. CKEditor kept them on both.
- Empty inline elements left by the style stacks are removed, and Office namespaced elements
  (`o:*`, `st1:*`, `w:*`) are unwrapped.
- `<object>` elements are kept when they have a `data` attribute.
- The Microsoft Edge (EdgeHTML) list heuristics are opt-in through `pastefromword_heuristics_edge_list`.
- No custom filter file (`pasteFromWordCleanupFile`), no `enterMode` BR support, no asynchronous `blob:`
  image conversion. Use the events or the filter API instead.

## Development

```bash
make install   # npm install
make dev       # integration page on http://localhost:5173 with a real TinyMCE
make test      # node --test (jsdom provides the DOM parser)
make lint
make build     # dist/ (ES + UMD)
```

The fixtures in `test/fixtures` come from the CKEditor 4 repository. `npm run test:corpus` checks the output
invariants over them. Set `PFW_CORPUS_DIR` to run the same checks over another directory of Word HTML files.

## License

This package is a derived work of CKEditor 4, Copyright (c) 2003-2023, CKSource Holding sp. z o.o.
Like CKEditor 4, it is licensed under the terms of any of the GPL 2+, LGPL 2.1+ or MPL 1.1+ licenses, at
your choice. See [LICENSE.md](LICENSE.md).
