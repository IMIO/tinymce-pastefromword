/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (plugins/pastetools/filter/common.js) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

import { Element, Text, NODE_ELEMENT, NODE_TEXT } from './tree.js';
import {
  Border,
  convertToPx,
  parseBackground,
  parseCssText,
  parseShorthandMargins,
  splitMarginShorthand,
  writeCssText,
} from './css.js';

/**
 * @typedef {Object} FilterConfig
 * @property {boolean} [removeFontStyles=false] Drop font family, size, color, background and line height.
 * @property {boolean} [keepZeroMargins=false] Keep `margin-*: 0` styles.
 * @property {boolean} [heuristicsEdgeList=false] Detect lists pasted without Word list metadata.
 * @property {string[]|false} [fontFamilies=false] Font stacks of the editor. `font[face]` is mapped to them.
 * @property {(() => Element)|null} [createPageBreak=null] Factory for the page break element of the editor.
 */

const TABLE_CELLS = { td: 1, th: 1 };

// Word writes these non-standard properties. No browser supports them.
const WORD_ONLY_STYLE = /^(?:tab-stops|text-autospace|text-underline|layout-grid-mode|punctuation-wrap|-ms-)/;

function fixValue(value) {
  // Add 'px' only to values that do not end with '%'.
  return /%$/.test(value) ? value : value + 'px';
}

export const Style = {
  setStyle(element, key, value, dontOverwrite) {
    const styles = parseCssText(element.attributes.style);

    if (dontOverwrite && styles[key]) {
      return;
    }

    if (value === '') {
      delete styles[key];
    } else {
      styles[key] = value;
    }

    element.attributes.style = writeCssText(styles);
  },

  convertStyleToPx(element) {
    const style = element.attributes.style;

    if (!style) {
      return;
    }

    element.attributes.style = style.replace(/\d+(\.\d+)?pt/g, (match) => convertToPx(match) + 'px');
  },

  // Map attributes to styles.
  mapStyles(element, attributeStyleMap) {
    for (const attribute of Object.keys(attributeStyleMap)) {
      if (element.attributes[attribute]) {
        if (typeof attributeStyleMap[attribute] === 'function') {
          attributeStyleMap[attribute](element.attributes[attribute]);
        } else {
          Style.setStyle(element, attributeStyleMap[attribute], element.attributes[attribute]);
        }
        delete element.attributes[attribute];
      }
    }
  },

  // Map common attributes to styles.
  mapCommonStyles(element) {
    return Style.mapStyles(element, {
      vAlign(value) {
        Style.setStyle(element, 'vertical-align', value);
      },
      width(value) {
        Style.setStyle(element, 'width', fixValue(value));
      },
      height(value) {
        Style.setStyle(element, 'height', fixValue(value));
      },
    });
  },

  /**
   * Filter the Word specific styles of an element.
   *
   * @param {Element} element
   * @param {FilterConfig} config
   * @returns {string} The normalized CSS text.
   */
  normalizedStyles(element, config) {
    // Redundant styles and style values.
    const resetStyles = [
      'background-color:transparent',
      'border-image:none',
      'color:windowtext',
      'direction:ltr',
      'mso-',
      'visibility:visible',
      'div:border:none',
    ];
    const textStyles = ['font-family', 'font', 'font-size', 'color', 'background-color', 'line-height', 'text-decoration'];
    const matchStyle = (...args) => {
      const keys = args.filter(Boolean);
      return resetStyles.includes(keys.join(':'));
    };
    const removeFontStyles = config.removeFontStyles === true;

    const styles = parseCssText(element.attributes.style);

    if (element.name === 'cke:li') {
      // IE8 emulates list indentation with text-indent and a left margin. Normalize this.
      if (styles['TEXT-INDENT'] && styles.MARGIN) {
        element.attributes['cke-indentation'] = getElementIndentation(element);
        styles.MARGIN = styles.MARGIN.replace(/(([\w.]+ ){3,3})[\d.]+(\w+$)/, '$10$3');
      } else {
        // Text indent works differently with lists in HTML than in Word.
        delete styles['TEXT-INDENT'];
      }
      delete styles['text-indent'];
    }

    for (const key of Object.keys(styles)) {
      const styleName = key.toLowerCase();
      const styleValue = styles[key];
      const toBeRemoved = removeFontStyles && textStyles.includes(styleName);

      if (
        toBeRemoved ||
        WORD_ONLY_STYLE.test(styleName) ||
        matchStyle(null, styleName, styleValue) ||
        matchStyle(null, styleName.replace(/-.*$/, '-')) ||
        matchStyle(null, styleName) ||
        matchStyle(element.name, styleName, styleValue) ||
        matchStyle(element.name, styleName.replace(/-.*$/, '-')) ||
        matchStyle(element.name, styleName) ||
        matchStyle(styleValue)
      ) {
        delete styles[key];
      }
    }

    const keepZeroMargins = config.keepZeroMargins;

    // Some elements have shorthand margins or longhand margins with zero values.
    parseShorthandMargins(styles);

    for (const side of ['top', 'right', 'bottom', 'left']) {
      const key = 'margin-' + side;
      if (!(key in styles)) {
        continue;
      }

      const value = convertToPx(styles[key]);

      if (typeof value === 'string') {
        // Percentages stay as they are.
        continue;
      }

      if (value || keepZeroMargins) {
        styles[key] = value ? value + 'px' : 0;
      } else {
        delete styles[key];
      }
    }

    return writeCssText(styles);
  },

  /**
   * Surround the children of the element with a stack of `<span>` elements.
   * Each span carries one style that belonged to the element.
   *
   * @param {Element} element
   * @param {import('./filter.js').Filter} filter
   * @param {FilterConfig} config
   * @param {RegExp} [skipStyles] Styles that stay on the element.
   */
  createStyleStack(element, filter, config, skipStyles) {
    const children = [];

    element.filterChildren(filter);

    // Detach the children.
    for (let i = element.children.length - 1; i >= 0; i--) {
      children.unshift(element.children[i]);
      element.children[i].remove();
    }

    Style.sortStyles(element);

    const styles = parseCssText(Style.normalizedStyles(element, config));
    let innermostElement = element;
    // A span root keeps at least one style.
    let styleTopmost = element.name === 'span';

    for (const style of Object.keys(styles)) {
      if (style.match(skipStyles || /margin((?!-)|-left|-top|-bottom|-right)|text-indent|text-align|width|border|padding/i)) {
        continue;
      }

      if (styleTopmost) {
        styleTopmost = false;
        continue;
      }

      const newElement = new Element('span');
      newElement.attributes.style = style + ':' + styles[style];

      innermostElement.add(newElement);
      innermostElement = newElement;

      delete styles[style];
    }

    if (Object.keys(styles).length) {
      element.attributes.style = writeCssText(styles);
    } else {
      delete element.attributes.style;
    }

    // Put the children into the innermost span.
    for (const child of children) {
      innermostElement.add(child);
    }
  },

  // Some styles must be stacked in a given order to work.
  sortStyles(element) {
    const orderedStyles = ['border', 'border-bottom', 'font-size', 'background'];
    const style = parseCssText(element.attributes.style);
    const keys = Object.keys(style);
    const sortedKeys = [];
    const nonSortedKeys = [];

    for (const key of keys) {
      if (orderedStyles.includes(key.toLowerCase())) {
        sortedKeys.push(key);
      } else {
        nonSortedKeys.push(key);
      }
    }

    sortedKeys.sort((a, b) => orderedStyles.indexOf(a.toLowerCase()) - orderedStyles.indexOf(b.toLowerCase()));

    const sortedStyles = {};
    for (const key of [...sortedKeys, ...nonSortedKeys]) {
      sortedStyles[key] = style[key];
    }

    element.attributes.style = writeCssText(sortedStyles);
  },

  /**
   * Move the styles of the element to its element children.
   *
   * @param {Element} element
   * @param {Object<string, boolean>} [exceptions] Style names that stay on the element.
   * @param {boolean} [wrapText=false] Wrap direct text children in a `<span>` to receive the styles.
   * @returns {boolean} `true` when styles were moved.
   */
  pushStylesLower(element, exceptions, wrapText) {
    if (!element.attributes.style || element.children.length === 0) {
      return false;
    }

    exceptions = exceptions || {};

    // Entries that end with a dash match all styles with that prefix.
    const retainedStyles = {
      'list-style-type': true,
      width: true,
      height: true,
      border: true,
      'border-': true,
    };

    const styles = parseCssText(element.attributes.style);

    for (const style of Object.keys(styles)) {
      const lower = style.toLowerCase();
      if (lower in retainedStyles || retainedStyles[lower.replace(/-.*$/, '-')] || lower in exceptions) {
        continue;
      }

      let pushed = false;

      for (let i = 0; i < element.children.length; i++) {
        let child = element.children[i];

        if (child.type === NODE_TEXT && wrapText) {
          const wrapper = new Element('span');
          wrapper.add(new Text(child.value));
          child.replaceWith(wrapper);
          child = wrapper;
        }

        if (child.type !== NODE_ELEMENT) {
          continue;
        }

        pushed = true;

        Style.setStyle(child, style, styles[style]);
      }

      if (pushed) {
        delete styles[style];
      }
    }

    element.attributes.style = writeCssText(styles);

    return true;
  },
};

/**
 * Get the left indentation of an element in pixels.
 *
 * @param {Element} element
 * @returns {number}
 */
export function getElementIndentation(element) {
  const style = parseCssText(element.attributes.style);

  if (style.margin || style.MARGIN) {
    style.margin = style.margin || style.MARGIN;
    const fake = { margin: style.margin };
    splitMarginShorthand(fake);
    style['margin-left'] = fake['margin-left'];
  }

  return parseInt(convertToPx(style['margin-left'] || '0px'), 10);
}

/**
 * Move the children of an element next to it. The caller removes the element.
 *
 * @param {Element} element
 */
export function replaceWithChildren(element) {
  for (let i = element.children.length - 1; i >= 0; i--) {
    element.children[i].insertAfter(element);
  }
}

/**
 * Like createStyleStack, but stack attributes instead of styles.
 *
 * @param {Element} element
 * @param {import('./filter.js').Filter} filter
 */
export function createAttributeStack(element, filter) {
  const children = [];

  element.filterChildren(filter);

  for (let i = element.children.length - 1; i >= 0; i--) {
    children.unshift(element.children[i]);
    element.children[i].remove();
  }

  const attributes = element.attributes;
  let innermostElement = element;
  let topmost = true;

  for (const attribute of Object.keys(attributes)) {
    if (topmost) {
      topmost = false;
      continue;
    }

    const newElement = new Element(element.name);
    newElement.attributes[attribute] = attributes[attribute];

    innermostElement.add(newElement);
    innermostElement = newElement;

    delete attributes[attribute];
  }

  for (const child of children) {
    innermostElement.add(child);
  }
}

function removeSuperfluousStyles(element) {
  const resetStyles = [
    'background-color:transparent',
    'background:transparent',
    'background-color:none',
    'background:none',
    'background-position:initial initial',
    'background-repeat:initial initial',
    'caret-color',
    'font-family:-webkit-standard',
    'font-variant-caps',
    'letter-spacing:normal',
    'orphans',
    'widows',
    'text-transform:none',
    'word-spacing:0px',
    '-webkit-text-size-adjust:auto',
    '-webkit-text-stroke-width:0px',
    'text-indent:0px',
    'margin-bottom:0in',
  ];

  const styles = parseCssText(element.attributes.style);

  for (const styleName of Object.keys(styles)) {
    const styleString = styleName + ':' + styles[styleName];

    if (resetStyles.some((val) => styleString.substring(0, val.length).toLowerCase() === val)) {
      delete styles[styleName];
    }
  }

  const cssText = writeCssText(styles);

  if (cssText !== '') {
    element.attributes.style = cssText;
  } else {
    delete element.attributes.style;
  }
}

function normalizeAttributesName(element) {
  if (element.attributes.bgcolor) {
    const styles = parseCssText(element.attributes.style);

    if (!styles['background-color']) {
      styles['background-color'] = element.attributes.bgcolor;
      element.attributes.style = writeCssText(styles);
    }
  }
}

function replaceWithMatchingFont(fontValue, availableFonts) {
  const fontParts = fontValue.split(',');
  const matchingFont = availableFonts.find((font) => fontParts.every((part) => font.indexOf(part.trim()) !== -1));

  return matchingFont || fontValue;
}

/**
 * Rules shared by all Office paste filters.
 *
 * @param {string} html
 * @param {FilterConfig} config
 * @param {import('./filter.js').Filter} filter
 */
export function commonRules(html, config, filter) {
  const availableFonts = config.fontFamilies && config.fontFamilies.length ? config.fontFamilies : false;

  return {
    elements: {
      '^': (element) => {
        removeSuperfluousStyles(element);
        // Normalize the attributes before any other rule runs.
        normalizeAttributesName(element);
      },

      span: (element) => {
        if (element.hasClass('Apple-converted-space')) {
          return new Text(' ');
        }
      },

      table: (element) => {
        element.filterChildren(filter);

        const parent = element.parent;
        const root = parent && parent.parent;

        // A parent div that only aligns the table: move the alignment to the table.
        if (
          parent.name &&
          parent.name === 'div' &&
          parent.attributes.align &&
          Object.keys(parent.attributes).length === 1 &&
          parent.children.length === 1
        ) {
          element.attributes.align = parent.attributes.align;

          const parentChildren = parent.children.splice(0);

          element.remove();
          for (let i = parentChildren.length - 1; i >= 0; i--) {
            root.add(parentChildren[i], parent.getIndex());
          }
          parent.remove();
        }

        Style.convertStyleToPx(element);
      },

      tr: (element) => {
        // The attributes are moved to the cells.
        element.attributes = {};
      },

      td: (element) => {
        const ascendant = element.getAscendant('table');
        const ascendantStyle = parseCssText(ascendant ? ascendant.attributes.style : '', true);

        // A background set on the whole table is moved to the cells.
        const background = ascendantStyle.background;
        if (background) {
          Style.setStyle(element, 'background', background, true);
        }

        const backgroundColor = ascendantStyle['background-color'];
        if (backgroundColor) {
          Style.setStyle(element, 'background-color', backgroundColor, true);
        }

        const styles = parseCssText(element.attributes.style, true);
        const borderStyles = styles.border ? Border.fromCssRule(styles.border) : {};
        const borders = Border.splitCssValues(styles, borderStyles);
        const tmpStyles = { ...styles };

        // Drop the old border styles so that they do not conflict with the new ones.
        for (const key of Object.keys(tmpStyles)) {
          if (key.indexOf('border') === 0) {
            delete tmpStyles[key];
          }
        }

        element.attributes.style = writeCssText(tmpStyles);

        // Unify the background color property.
        if (styles.background) {
          const bg = parseBackground(styles.background);

          if (bg.color) {
            Style.setStyle(element, 'background-color', bg.color, true);
            Style.setStyle(element, 'background', '');
          }
        }

        // Unify the border properties.
        for (const border of Object.keys(borders)) {
          const borderStyle = styles[border] ? Border.fromCssRule(styles[border]) : borders[border];

          if (borderStyle.style === 'none') {
            Style.setStyle(element, border, 'none');
          } else {
            Style.setStyle(element, border, borderStyle.toString());
          }
        }

        Style.mapCommonStyles(element);

        Style.convertStyleToPx(element);

        Style.createStyleStack(
          element,
          filter,
          config,
          /margin|text-align|padding|list-style-type|width|height|border|white-space|vertical-align|background/i,
        );
      },

      font: (element) => {
        if (element.attributes.face && availableFonts) {
          element.attributes.face = replaceWithMatchingFont(element.attributes.face, availableFonts);
        }
      },
    },
  };
}

export { TABLE_CELLS, parseShorthandMargins };
