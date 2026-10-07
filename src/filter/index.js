/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

// Framework-free entry point. It only needs a `DOMParser` global.

export { cleanWord, wordRules } from './word.js';
export { commonRules, Style } from './common.js';
export { List, Heuristics } from './lists.js';
export {
  embedRtfImages,
  extractFromRtf,
  extractTagsFromHtml,
  createSrcWithBase64,
  getImageType,
  getImageTypeFromSignature,
  hexToBytes,
  bytesToBase64,
  SUPPORTED_IMAGE_TYPES,
} from './image.js';
export { getGroup, getGroups, removeGroups, extractGroupContent } from './rtf.js';
export { inlineStyles, parseStyleSheet, sortRules } from './inliner.js';
export { parseCssText, writeCssText, convertToPx, Border, parseMargin, parseBackground } from './css.js';
export { Filter } from './filter.js';
export { Element, Text, Comment, Fragment, NODE_ELEMENT, NODE_TEXT, NODE_COMMENT } from './tree.js';

const WORD_REGEXP = /(class="?Mso|style=["'][^"]*?\bmso-|w:WordDocument|<o:\w+>|<\/font>)/;

/**
 * Get the name of the generator of an HTML document from its meta tag.
 *
 * @param {string} content
 * @returns {'microsoft'|'libreoffice'|'unknown'|undefined} `undefined` when there is no generator meta tag.
 */
export function getContentGeneratorName(content) {
  const result = /<meta\s+name=["']?generator["']?\s+content=["']?(\w+)/i.exec(content || '');

  if (!result) {
    return undefined;
  }

  const generatorName = result[1].toLowerCase();

  if (generatorName.indexOf('microsoft') === 0) {
    return 'microsoft';
  }

  if (generatorName.indexOf('libreoffice') === 0) {
    return 'libreoffice';
  }

  return 'unknown';
}

/**
 * Check if an HTML string comes from Microsoft Office.
 *
 * @param {string} html
 * @returns {boolean}
 */
export function isWordContent(html) {
  if (!html) {
    return false;
  }

  const generatorName = getContentGeneratorName(html);

  // The regular expression is only used without a generator meta tag.
  return generatorName ? generatorName === 'microsoft' : WORD_REGEXP.test(html);
}
