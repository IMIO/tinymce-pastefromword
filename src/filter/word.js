/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (plugins/pastefromword/filter/default.js) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

import { Element, Fragment, Text, NODE_ELEMENT, NODE_TEXT } from './tree.js';
import { Filter } from './filter.js';
import { parseCssText } from './css.js';
import { Style, TABLE_CELLS, commonRules, createAttributeStack, replaceWithChildren } from './common.js';
import { Heuristics, List } from './lists.js';
import { inlineStyles } from './inliner.js';

const INVALID_TAGS = ['o:p', 'xml', 'script', 'meta', 'link'];

const SHAPE_TAGS = ['v:arc', 'v:curve', 'v:line', 'v:oval', 'v:polyline', 'v:rect', 'v:roundrect', 'v:group'];

const DEFAULT_CONFIG = {
  removeFontStyles: false,
  keepZeroMargins: false,
  heuristicsEdgeList: false,
  fontFamilies: false,
  createPageBreak: null,
};

const remove = () => false;

const falseIfEmpty = (value) => (value === '' ? false : value);

/**
 * Rules of the Word filter.
 *
 * @param {string} html
 * @param {import('./common.js').FilterConfig} config
 * @param {Filter} filter
 */
export function wordRules(html, config, filter) {
  const msoListsDetected = /mso-list:\s*l\d+\s+level\d+\s+lfo\d+/.test(html);
  const shapesIds = [];
  const links = {};
  let inComment = 0;

  // Keep the ids of the drawn shapes and of the canvas shapes.
  const shapeTagging = (element) => {
    if (element.attributes['o:gfxdata'] || element.parent.name === 'v:group') {
      shapesIds.push(element.attributes.id);
    }
  };

  const listItemOrBlock = (element) => {
    if (List.thisIsAListItem(config, element)) {
      List.convertToFakeListItem(config, element);
    }

    Style.createStyleStack(element, filter, config);
  };

  const rules = {
    root: (element) => {
      element.filterChildren(filter);

      List.cleanup(List.createLists(element, config));
    },

    elementNames: [
      [/^\?xml:namespace$/, ''],
      [/^v:shapetype/, ''],
      [new RegExp(INVALID_TAGS.join('|')), ''],
    ],

    elements: {
      a: (element) => {
        if (element.attributes.name) {
          // Redundant anchor created by Word.
          if (element.attributes.name === '_GoBack') {
            delete element.name;
            return;
          }

          // Garbage links that go nowhere.
          if (element.attributes.name.match(/^OLE_LINK\d+$/)) {
            delete element.name;
            return;
          }
        }

        if (element.attributes.href && element.attributes.href.match(/#.+$/)) {
          const name = element.attributes.href.match(/#(.+)$/)[1];
          links[name] = element;
        }

        if (element.attributes.name && links[element.attributes.name]) {
          const link = links[element.attributes.name];
          link.attributes.href = link.attributes.href.replace(/.*#(.*)$/, '#$1');
        }
      },

      div: (element) => {
        Style.createStyleStack(element, filter, config);
      },

      img: (element) => {
        // The parent can be the fragment, which has no attributes.
        if (element.parent && element.parent.attributes) {
          const attrs = element.parent.attributes;
          const style = attrs.style || attrs.STYLE;
          if (style && style.match(/mso-list:\s?Ignore/)) {
            element.attributes['cke-ignored'] = true;
          }
        }

        Style.mapCommonStyles(element);

        if (
          element.attributes.src &&
          element.attributes.src.match(/^file:\/\//) &&
          element.attributes.alt &&
          element.attributes.alt.match(/^https?:\/\//)
        ) {
          element.attributes.src = element.attributes.alt;
        }

        const imgShapesIds = element.attributes['v:shapes'] ? element.attributes['v:shapes'].split(' ') : [];
        // An image that only renders drawn shapes is removed.
        const isShapeFromList = imgShapesIds.every((shapeId) => shapesIds.indexOf(shapeId) > -1);
        if (imgShapesIds.length && isShapeFromList) {
          return false;
        }
      },

      p: (element) => {
        element.filterChildren(filter);

        if (element.attributes.style && element.attributes.style.match(/display:\s*none/i)) {
          return false;
        }

        if (List.thisIsAListItem(config, element)) {
          if (Heuristics.isEdgeListItem(config, element)) {
            Heuristics.cleanupEdgeListItem(element);
          }

          List.convertToFakeListItem(config, element);

          // A list item can contain several paragraphs. Separate them with a BR.
          let paragraphsReplaced = 0;
          for (let i = 0; i < element.children.length; i++) {
            const node = element.children[i];
            if (node.name === 'p') {
              if (paragraphsReplaced > 0) {
                new Element('br').insertBefore(node);
              }

              node.replaceWithChildren();
              paragraphsReplaced += 1;
            }
          }
        } else {
          // The list level information can be in a paragraph inside a real list item.
          const container = element.getAscendant((el) => el.name === 'ul' || el.name === 'ol');
          const style = parseCssText(element.attributes.style);
          if (container && !container.attributes['cke-list-level'] && style['mso-list'] && style['mso-list'].match(/level/)) {
            container.attributes['cke-list-level'] = style['mso-list'].match(/level(\d+)/)[1];
          }
        }

        Style.createStyleStack(element, filter, config);
      },

      pre: listItemOrBlock,
      h1: listItemOrBlock,
      h2: listItemOrBlock,
      h3: listItemOrBlock,
      h4: listItemOrBlock,
      h5: listItemOrBlock,
      h6: listItemOrBlock,

      font: (element) => {
        if (element.getHtml().match(/^\s*$/)) {
          // An empty font element is replaced by a space, so that the words stay separated.
          if (element.parent.type === NODE_ELEMENT) {
            new Text(' ').insertAfter(element);
          }
          return false;
        }

        if (config.removeFontStyles === true && element.attributes.size) {
          delete element.attributes.size;
        }

        // Excel writes td > font with only class and style attributes.
        const keys = Object.keys(element.attributes);
        if (TABLE_CELLS[element.parent.name] && keys.length === 2 && keys[0] === 'class' && keys[1] === 'style') {
          Style.createStyleStack(element, filter, config);
        } else {
          createAttributeStack(element, filter);
        }
      },

      ul: (element) => {
        // Lists are only processed when the content comes from Word.
        if (!msoListsDetected) {
          return;
        }

        // An unusual way to create a level 2 list.
        if (element.parent.name === 'li' && element.parent.children.indexOf(element) === 0) {
          Style.setStyle(element.parent, 'list-style-type', 'none');
        }

        List.dissolveList(element);
        return false;
      },

      li: (element) => {
        Heuristics.correctLevelShift(element);

        if (!msoListsDetected) {
          return;
        }

        element.attributes.style = Style.normalizedStyles(element, config);

        Style.pushStylesLower(element);
      },

      ol: (element) => {
        if (!msoListsDetected) {
          return;
        }

        // A list that skips a level is surrounded by a li element.
        if (element.parent.name === 'li' && element.parent.children.indexOf(element) === 0) {
          Style.setStyle(element.parent, 'list-style-type', 'none');
        }

        List.dissolveList(element);
        return false;
      },

      span: (element) => {
        element.filterChildren(filter);

        element.attributes.style = Style.normalizedStyles(element, config);

        if (
          !element.attributes.style ||
          // Garbage bookmarks that disrupt the content structure.
          element.attributes.style.match(/^mso-bookmark:OLE_LINK\d+$/) ||
          element.getHtml().match(/^(\s|&nbsp;)+$/)
        ) {
          replaceWithChildren(element);
          return false;
        }

        if (element.attributes.style.match(/FONT-FAMILY:\s*Symbol/i)) {
          element.forEach(
            (node) => {
              node.value = node.value.replace(/\u00a0/g, '');
            },
            NODE_TEXT,
            true,
          );
        }

        Style.createStyleStack(element, filter, config);
      },

      'v:imagedata': remove,

      // Images drawn as VML shapes.
      'v:shape': (element) => {
        // 1. A regular shape without v:imagedata: tag it.
        // 2. A shape with v:imagedata and a matching img tag: remove the shape.
        // 3. A shape without a matching img tag:
        //    a) in a v:group it belongs to a canvas: tag it;
        //    b) otherwise convert the shape to an img tag.
        let duplicate = false;
        const child = element.getFirst('v:imagedata');

        if (child === null) {
          shapeTagging(element);
          return;
        }

        element.parent.find((node) => {
          if (node.name === 'img' && node.attributes && node.attributes['v:shapes'] === element.attributes.id) {
            duplicate = true;
          }
        }, true);

        if (duplicate) {
          return false;
        }

        if (element.parent.name === 'v:group') {
          shapeTagging(element);
          return;
        }

        let src = '';

        element.forEach(
          (node) => {
            if (node.attributes && node.attributes.src) {
              src = node.attributes.src;
            }
          },
          NODE_ELEMENT,
          true,
        );

        element.filterChildren(filter);

        element.name = 'img';
        element.attributes.src = element.attributes.src || src;

        delete element.attributes.type;
      },

      style: () => {
        // No style sheets. Firefox tends to add some.
        return false;
      },

      object: (element) => {
        // An object needs a data attribute.
        if (!(element.attributes && element.attributes.data)) {
          return false;
        }
      },

      br: (element) => {
        if (!config.createPageBreak) {
          return;
        }

        const styles = parseCssText(element.attributes.style, true);

        // Safari uses break-before instead of page-break-before.
        if (styles['page-break-before'] === 'always' || styles['break-before'] === 'page') {
          return config.createPageBreak();
        }
      },
    },

    attributes: {
      style: (styles, element) => {
        // A false value deletes the attribute.
        return Style.normalizedStyles(element, config) || false;
      },
      class: (classes) => {
        // el\d+, font\d+ and xl\d+ are the default Excel classes. Mso* are the Word style classes.
        return falseIfEmpty(
          classes
            .replace(/(^|\s)(el\d+|font\d+|xl\d+|mso\w*)(?=\s|$)/gi, ' ')
            .replace(/\s+/g, ' ')
            .trim(),
        );
      },
      cellspacing: remove,
      cellpadding: remove,
      border: remove,
      'v:shapes': remove,
      'o:spid': remove,
    },

    comment: (value) => {
      // Field codes are between a supportFields comment and an endif comment.
      if (value.match(/\[if.* supportFields.*\]/)) {
        inComment++;
      }
      if (value === '[endif]') {
        inComment = inComment > 0 ? inComment - 1 : 0;
      }
      return false;
    },

    text: (content, node) => {
      if (inComment) {
        return '';
      }

      const grandparent = node.parent && node.parent.parent;

      // The spaces after a list marker are regular spaces.
      if (grandparent && grandparent.attributes && grandparent.attributes.style && grandparent.attributes.style.match(/mso-list:\s*ignore/i)) {
        return content.replace(/\u00a0/g, ' ');
      }

      return content;
    },
  };

  for (const shapeTag of SHAPE_TAGS) {
    rules.elements[shapeTag] = shapeTagging;
  }

  return rules;
}

const EMPTY_INLINE_ELEMENTS = new Set(['span', 'font', 'b', 'i', 'u', 'em', 'strong', 's', 'strike', 'sub', 'sup', 'small', 'big']);

function removeHelperAttributes(fragment) {
  fragment.forEach((element) => {
    for (const name of Object.keys(element.attributes)) {
      if (name.startsWith('cke-') || (name === 'style' && element.attributes.style === '')) {
        delete element.attributes[name];
      }
    }
  }, NODE_ELEMENT);
}

// Office namespaced elements such as o:smarttagtype or st1:place only wrap content.
function unwrapNamespacedElements(fragment) {
  const namespaced = [];

  fragment.forEach((element) => {
    if (element.name.indexOf(':') !== -1) {
      namespaced.push(element);
    }
  }, NODE_ELEMENT);

  for (const element of namespaced) {
    if (element.parent) {
      element.replaceWithChildren();
    }
  }
}

// The style stacks can leave empty formatting elements behind.
function removeEmptyInlineElements(fragment) {
  let removed;

  do {
    removed = false;
    const empty = [];

    fragment.forEach((element) => {
      if (EMPTY_INLINE_ELEMENTS.has(element.name) && element.children.length === 0) {
        empty.push(element);
      }
    }, NODE_ELEMENT);

    for (const element of empty) {
      if (element.parent) {
        element.remove();
        removed = true;
      }
    }
  } while (removed);
}

function finalCleanup(fragment) {
  unwrapNamespacedElements(fragment);
  removeEmptyInlineElements(fragment);
  removeHelperAttributes(fragment);
}

/**
 * Clean up HTML pasted from Microsoft Word.
 *
 * @param {string} html The full clipboard HTML. The head and body are handled.
 * @param {import('./common.js').FilterConfig} [config]
 * @returns {string} The cleaned body HTML.
 */
export function cleanWord(html, config = {}) {
  const cfg = { ...DEFAULT_CONFIG, ...config };

  // Some styles only exist in the style sheets: inline them first.
  let bodyHtml = inlineStyles(html);

  // Word malforms the conditional comments. Fix them so that the content inside is parsed.
  bodyHtml = bodyHtml.replace(/<!\[/g, '<!--[').replace(/\]>/g, ']-->');

  const fragment = Fragment.fromHtml(bodyHtml);
  const filter = new Filter();

  filter.addRules(commonRules(bodyHtml, cfg, filter));
  filter.addRules(wordRules(bodyHtml, cfg, filter));
  filter.applyTo(fragment);

  finalCleanup(fragment);

  return fragment.getHtml();
}
