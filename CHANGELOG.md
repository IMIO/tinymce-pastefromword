# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

Initial release, ported from the CKEditor 4.22.1 `pastefromword` and `pastetools` plugins.

### Added
- Detection of Microsoft Office clipboard content (`meta generator`, `Mso` classes, `mso-*` styles)
- Rebuilding of Word lists into `<ol>`/`<ul>`: nesting, numbering types, start values, explicit item values
- Table cleanup: borders, widths, backgrounds and alignment unified as CSS
- Style sheet inlining and removal of all Word specific markup (`mso-*` styles, `o:p`, VML, field codes, bookmarks, conditional comments)
- Images embedded from the RTF clipboard (PNG, JPEG, GIF) through the TinyMCE blob cache
- Page breaks converted to the TinyMCE `pagebreak` placeholder
- `pastefromword_*` options, `PasteFromWord` / `AfterPasteFromWord` events, toggle button, menu item and command
- Framework-free filter export (`@imiobe/tinymce-pastefromword/filter`)
