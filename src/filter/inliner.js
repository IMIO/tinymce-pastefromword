/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (plugins/pastetools/filter/common.js, styles.inliner) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

import { parseCssText, parseShorthandMargins, writeCssText } from './css.js';

// Word puts many styles in a <style> sheet instead of the elements.
// The inliner copies the matching sheet styles to the `style` attribute of each element.

/** Styles the inliner never copies. */
export const FILTERED_STYLES = [
  'break-before',
  'break-after',
  'break-inside',
  'page-break',
  'page-break-before',
  'page-break-after',
  'page-break-inside',
];

const KNOWN_PROPERTIES = new Set([
  'bottom', 'caption-side', 'clear', 'color', 'content', 'cursor', 'direction', 'display',
  'empty-cells', 'float', 'height', 'left', 'letter-spacing', 'line-height', 'max-height',
  'max-width', 'min-height', 'min-width', 'opacity', 'orphans', 'position', 'quotes', 'right',
  'table-layout', 'text-align', 'text-align-last', 'text-indent', 'text-shadow',
  'text-transform', 'top', 'unicode-bidi', 'vertical-align', 'visibility', 'white-space',
  'widows', 'width', 'word-break', 'word-spacing', 'word-wrap', 'overflow-wrap', 'z-index',
  'box-sizing', 'tab-size', 'hyphens', 'writing-mode', 'text-orientation', 'border-collapse',
  'border-spacing', 'text-justify',
]);

const KNOWN_PREFIXES = /^(?:margin|padding|border|font|background|list-style|text-decoration|page-break|break|column|outline|overflow|text-emphasis|text-underline-position|grid|flex)(?:-|$)/;

/**
 * A browser drops the properties it does not know when it parses a style sheet.
 * This check emulates that for the Word specific `mso-*` and other proprietary properties.
 *
 * @param {string} name
 * @returns {boolean}
 */
export function isKnownCssProperty(name) {
  const lower = name.toLowerCase();

  if (lower.startsWith('-')) {
    return false;
  }

  return KNOWN_PROPERTIES.has(lower) || KNOWN_PREFIXES.test(lower);
}

/**
 * Remove the styles that must not be inlined.
 *
 * @param {Object<string, string>} styles
 * @returns {Object<string, string>}
 */
export function filterInlinerStyles(styles) {
  const result = {};

  for (const name of Object.keys(styles)) {
    if (!FILTERED_STYLES.includes(name) && isKnownCssProperty(name)) {
      result[name] = styles[name];
    }
  }

  return result;
}

/**
 * Parse the text of a style sheet into flat rules. At-rules are skipped.
 *
 * @param {string} cssText
 * @returns {{selector: string, styles: Object<string, string>, index: number}[]}
 */
export function parseStyleSheet(cssText) {
  const rules = [];
  const text = cssText.replace(/<!--|-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '');
  const ruleRegex = /([^{}]+)\{([^{}]*)\}/g;
  let match;

  while ((match = ruleRegex.exec(text))) {
    const selector = match[1].trim();

    // The selector of an at-rule ends with the at-rule prelude, so skip the whole match.
    if (!selector || selector.includes('@')) {
      continue;
    }

    const styles = filterInlinerStyles(parseCssText(match[2], true));

    // A browser serializes `11.0pt` as `11pt`.
    for (const name of Object.keys(styles)) {
      styles[name] = styles[name].replace(/(\d)\.0+(?=[a-z%]|\s|,|$)/g, '$1');
    }

    rules.push({ selector, styles, index: rules.length });
  }

  return rules;
}

function isClassSelector(selector) {
  return String(selector).indexOf('.') !== -1;
}

/**
 * Sort the rules: class selectors first, then later rules first.
 * The rules are applied in this order and never overwrite a style already present.
 *
 * @param {{selector: string, styles: Object, index: number}[]} rules
 * @returns {{selector: string, styles: Object, index: number}[]}
 */
export function sortRules(rules) {
  return rules.slice().sort((rule1, rule2) => {
    const value1 = isClassSelector(rule1.selector) ? 1 : 0;
    const value2 = isClassSelector(rule2.selector) ? 1 : 0;
    const result = value2 - value1;

    return result !== 0 ? result : rule2.index - rule1.index;
  });
}

function applyRule(doc, rule) {
  let elements;

  try {
    elements = doc.querySelectorAll(rule.selector);
  } catch {
    // Word writes selectors such as `v\:*` that are not valid CSS.
    return;
  }

  const style = rule.styles;
  parseShorthandMargins(style);

  for (const element of elements) {
    const oldStyle = parseCssText(element.getAttribute('style'));
    parseShorthandMargins(oldStyle);

    // Styles already present win over the sheet styles.
    const newStyle = { ...oldStyle };
    for (const name of Object.keys(style)) {
      if (newStyle[name] == null) {
        newStyle[name] = style[name];
      }
    }
    element.setAttribute('style', writeCssText(newStyle));
  }
}

/**
 * Inline the `<style>` sheets of a document into its elements.
 *
 * @param {Document} doc
 */
export function inlineDocumentStyles(doc) {
  let rules = [];

  for (const styleElement of doc.querySelectorAll('style')) {
    const parsed = parseStyleSheet(styleElement.textContent || '');
    for (const rule of parsed) {
      rule.index = rules.length;
      rules.push(rule);
    }
  }

  for (const rule of sortRules(rules)) {
    applyRule(doc, rule);
  }
}

/**
 * Parse an HTML document, inline its style sheets and return the body content.
 *
 * @param {string} html Full clipboard HTML.
 * @returns {string} The body inner HTML.
 */
export function inlineStyles(html) {
  const doc = new globalThis.DOMParser().parseFromString(html, 'text/html');

  inlineDocumentStyles(doc);

  return doc.body ? doc.body.innerHTML : '';
}
