/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * TinyMCE Paste from Word plugin, ported from the CKEditor 4 Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

import tinymce from 'tinymce';
import { cleanWord, embedRtfImages, isWordContent, hexToBytes, bytesToBase64, SUPPORTED_IMAGE_TYPES } from './filter/index.js';
import { Element } from './filter/tree.js';
import { registerTranslations } from './i18n.js';

const PLUGIN_NAME = 'pastefromword';

let blobCounter = 0;

function registerOptions(editor) {
  const register = editor.options.register;

  register('pastefromword_remove_font_styles', { processor: 'boolean', default: false });
  register('pastefromword_keep_zero_margins', { processor: 'boolean', default: false });
  register('pastefromword_inline_images', { processor: 'boolean', default: true });
  register('pastefromword_heuristics_edge_list', { processor: 'boolean', default: false });
  register('pastefromword_prompt_cleanup', { processor: 'boolean', default: false });
}

/**
 * Read the font stacks of the editor, so that `font[face]` can be mapped to them.
 */
function getFontFamilies(editor) {
  const formats = editor.options.isRegistered('font_family_formats') ? editor.options.get('font_family_formats') : '';

  if (!formats || typeof formats !== 'string') {
    return false;
  }

  const families = formats
    .split(';')
    .map((format) => format.trim())
    .filter(Boolean)
    .map((format) => (format.indexOf('=') === -1 ? format : format.split('=')[1]));

  return families.length ? families : false;
}

/**
 * Create the page break placeholder of the TinyMCE pagebreak plugin.
 */
function createPageBreakFactory(editor) {
  if (!editor.plugins.pagebreak) {
    return null;
  }

  return () =>
    new Element('img', {
      src: tinymce.Env.transparentSrc,
      class: 'mce-pagebreak',
      'data-mce-resize': 'false',
      'data-mce-placeholder': '',
    });
}

/**
 * Store an image extracted from the RTF clipboard in the blob cache of the editor.
 * The editor then uploads it like a dropped image. Without a blob cache a data URL is used.
 */
function createImageSrcFactory(editor) {
  const blobCache = editor.editorUpload && editor.editorUpload.blobCache;

  return (img) => {
    if (!SUPPORTED_IMAGE_TYPES.includes(img.type) || !img.hex) {
      return null;
    }

    const bytes = hexToBytes(img.hex);
    const base64 = bytesToBase64(bytes);

    if (!blobCache || typeof Blob === 'undefined') {
      return 'data:' + img.type + ';base64,' + base64;
    }

    const existing = blobCache.getByData(base64, img.type);
    if (existing) {
      return existing.blobUri();
    }

    const id = 'pastefromword' + Date.now().toString(36) + (blobCounter++).toString(36);
    const blobInfo = blobCache.create(id, new Blob([bytes], { type: img.type }), base64);
    blobCache.add(blobInfo);

    return blobInfo.blobUri();
  };
}

function getFilterConfig(editor) {
  return {
    removeFontStyles: editor.options.get('pastefromword_remove_font_styles'),
    keepZeroMargins: editor.options.get('pastefromword_keep_zero_margins'),
    heuristicsEdgeList: editor.options.get('pastefromword_heuristics_edge_list'),
    fontFamilies: getFontFamilies(editor),
    createPageBreak: createPageBreakFactory(editor),
  };
}

function getClipboardData(dataTransfer, type) {
  try {
    return dataTransfer.getData(type) || '';
  } catch {
    return '';
  }
}

/**
 * Replicate the trimming TinyMCE applies to the clipboard HTML before PastePreProcess.
 * It is used to match a PastePreProcess event with the paste event that produced it.
 */
function trimHtml(html) {
  return html
    .replace(/^[\s\S]*<body[^>]*>\s*|\s*<\/body[^>]*>[\s\S]*$/gi, '')
    .replace(/<!--StartFragment-->|<!--EndFragment-->/g, '')
    .replace(/( ?)<span class="Apple-converted-space">\u00a0<\/span>( ?)/g, (all, s1, s2) => (!s1 && !s2 ? ' ' : '\u00a0'))
    .replace(/<br class="Apple-interchange-newline">/g, '')
    .replace(/<br>$/i, '');
}

function confirmCleanup(editor) {
  if (!editor.options.get('pastefromword_prompt_cleanup')) {
    return true;
  }

  return globalThis.confirm(editor.translate('The text you want to paste seems to be copied from Word. Do you want to clean it before pasting?'));
}

tinymce.PluginManager.add(PLUGIN_NAME, (editor) => {
  registerOptions(editor);
  registerTranslations();

  let forceFromWord = false;
  // Clipboard content of the paste event, consumed by the PastePreProcess event it triggers.
  let pending = null;

  const onError = (code, data) => {
    console.warn('[' + PLUGIN_NAME + '] ' + code, data);
  };

  /**
   * Clean Word HTML with the options of this editor. Images come from the RTF clipboard.
   *
   * @param {string} html
   * @param {string} [rtf]
   * @returns {string}
   */
  const clean = (html, rtf) => {
    let content = cleanWord(html, getFilterConfig(editor));

    if (editor.options.get('pastefromword_inline_images') && rtf && editor.schema.isValid('img')) {
      content = embedRtfImages(content, rtf, { createSrc: createImageSrcFactory(editor), onError });
    }

    return content;
  };

  const processPaste = (e) => {
    const captured = pending;
    pending = null;

    if (!captured || e.internal) {
      return;
    }

    // Plain text paste produces other content: leave it alone.
    if (e.content !== trimHtml(captured.html) && !isWordContent(e.content)) {
      return;
    }

    const pfwEvent = editor.dispatch('PasteFromWord', { content: captured.html, rtf: captured.rtf });

    if (pfwEvent.isDefaultPrevented() && !forceFromWord) {
      return;
    }

    if (!forceFromWord && !confirmCleanup(editor)) {
      return;
    }

    const content = clean(pfwEvent.content, captured.rtf);
    const afterEvent = editor.dispatch('AfterPasteFromWord', { content });

    e.content = afterEvent.content;
  };

  editor.on('PreInit', () => {
    // Runs before the paste handler of the editor core.
    editor.on('paste', (e) => {
      pending = null;

      const dataTransfer = e.clipboardData;
      if (e.isDefaultPrevented() || !dataTransfer) {
        return;
      }

      const html = getClipboardData(dataTransfer, 'text/html');
      if (!html || !(forceFromWord || isWordContent(html))) {
        return;
      }

      pending = { html, rtf: getClipboardData(dataTransfer, 'text/rtf') };

      // The core handles the paste in the same event loop turn.
      setTimeout(() => {
        pending = null;
      }, 0);
    });

    // Registered after the core filters, so that the core does not strip the styles of the result.
    editor.on('PastePreProcess', processPaste);
  });

  editor.addCommand('mceTogglePasteFromWord', () => {
    forceFromWord = !forceFromWord;
    editor.dispatch('PasteFromWordToggle', { state: forceFromWord });
  });

  editor.addQueryStateHandler('mceTogglePasteFromWord', () => forceFromWord);

  const onSetup = (api) => {
    api.setActive(forceFromWord);
    const handler = (e) => api.setActive(e.state);
    editor.on('PasteFromWordToggle', handler);
    return () => editor.off('PasteFromWordToggle', handler);
  };

  editor.ui.registry.addToggleButton(PLUGIN_NAME, {
    icon: 'paste',
    tooltip: 'Paste from Word',
    onAction: () => editor.execCommand('mceTogglePasteFromWord'),
    onSetup,
  });

  editor.ui.registry.addToggleMenuItem(PLUGIN_NAME, {
    icon: 'paste',
    text: 'Paste from Word',
    onAction: () => editor.execCommand('mceTogglePasteFromWord'),
    onSetup,
  });

  return {
    cleanWord: clean,
    isWordContent,
    getMetadata: () => ({
      name: 'Paste from Word',
      url: 'https://github.com/IMIO/tinymce-pastefromword',
    }),
  };
});
