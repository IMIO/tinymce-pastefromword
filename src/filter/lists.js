/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (plugins/pastefromword/filter/default.js, lists and heuristics) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

import { Element, NODE_ELEMENT, NODE_TEXT } from './tree.js';
import { convertToPx, parseCssText, writeCssText } from './css.js';
import { Style, getElementIndentation } from './common.js';

// Word writes list items as paragraphs with a `mso-list: l0 level1 lfo1` style.
// The list marker is a text such as `1.` or `·` inside a `mso-list: Ignore` span.
// These helpers rebuild real `<ol>` and `<ul>` structures from those paragraphs.

const NBSP = '\u00a0';

/** Indentation of one list level in Word, in pixels (0.5in). */
export const WORD_LIST_INDENT_PX = 48;

/**
 * Methods that use heuristics on the pasted content.
 */
export const Heuristics = {
  /**
   * Decide if an item looks like a list item pasted without any list metadata.
   * Only used when the `heuristicsEdgeList` option is enabled.
   *
   * @param {import('./common.js').FilterConfig} config
   * @param {Element} item
   * @returns {boolean}
   */
  isEdgeListItem(config, item) {
    if (!config.heuristicsEdgeList) {
      return false;
    }

    let innerText = '';

    if (item.forEach) {
      item.forEach((text) => {
        innerText += text.value;
      }, NODE_TEXT);
    }

    if (new RegExp('^(?: |' + NBSP + ')*\\(?[a-zA-Z0-9]+?[.)](?: |' + NBSP + '){2,}').test(innerText)) {
      return true;
    }

    return Heuristics.isDegenerateListItem(config, item);
  },

  /**
   * Remove the indentation written before the list marker as `&nbsp;` characters.
   *
   * @param {Element} item
   */
  cleanupEdgeListItem(item) {
    let textOccurred = false;

    item.forEach((node) => {
      if (!textOccurred) {
        node.value = node.value.replace(new RegExp('^(?:' + NBSP + '|[\\s])+'), '');

        if (node.value.length) {
          textOccurred = true;
        }
      }
    }, NODE_TEXT);
  },

  /**
   * A degenerate list item has some list specific styles, but not the ones that
   * give its list level.
   *
   * @param {import('./common.js').FilterConfig} config
   * @param {Element} item
   * @returns {boolean}
   */
  isDegenerateListItem(config, item) {
    return (
      !!item.attributes['cke-list-level'] ||
      (item.attributes.style &&
        !item.attributes.style.match(/mso-list/) &&
        !!item.find((child) => {
          // A heading can be a list item when a child contains only a list symbol.
          if (child.type === NODE_ELEMENT && item.name.match(/h\d/i) && child.getHtml().match(/^[a-zA-Z0-9]+?[.)]$/)) {
            return true;
          }

          const css = parseCssText(child.attributes && child.attributes.style, true);

          if (!css) {
            return false;
          }
          const fontSize = css.font || css['font-size'] || '';
          const fontFamily = css['font-family'] || '';

          return (fontSize.match(/7pt/i) && !!child.previous) || fontFamily.match(/symbol/i);
        }, true).length)
    );
  },

  /**
   * Assign list levels to the item and the items that follow it, based on
   * the smallest indentation difference between them.
   *
   * @param {import('./common.js').FilterConfig} config
   * @param {Element} item
   * @returns {{indents: number[], levels: number[], diffs: number[]}|undefined}
   */
  assignListLevels(config, item) {
    if (item.attributes && item.attributes['cke-list-level'] !== undefined) {
      return;
    }

    const indents = [getElementIndentation(item)];
    const items = [item];

    while (
      item.next &&
      item.next.attributes &&
      !item.next.attributes['cke-list-level'] &&
      Heuristics.isDegenerateListItem(config, item.next)
    ) {
      item = item.next;
      indents.push(getElementIndentation(item));
      items.push(item);
    }

    const indentationDiffs = indents.map((curIndent, i) => (i === 0 ? 0 : curIndent - indents[i - 1]));
    const indentationPerLevel = Heuristics.guessIndentationStep(indents.filter((val) => val !== 0));

    let levels = indents.map((val) => Math.round(val / indentationPerLevel));

    // A level cannot be 0. Bump all the levels in that case.
    if (levels.indexOf(0) !== -1) {
      levels = levels.map((val) => val + 1);
    }

    items.forEach((curItem, index) => {
      curItem.attributes['cke-list-level'] = levels[index];
    });

    return {
      indents,
      levels,
      diffs: indentationDiffs,
    };
  },

  /**
   * Guess the indentation difference between two list levels.
   *
   * @param {number[]} indentations
   * @returns {number|null}
   */
  guessIndentationStep(indentations) {
    return indentations.length ? Math.min(...indentations) : null;
  },

  /**
   * Move list items that are one level too deep up by one level.
   *
   * @param {Element} element
   */
  correctLevelShift(element) {
    const isShiftedList = (list) => list.children && list.children.length === 1 && Heuristics.isShifted(list.children[0]);

    if (Heuristics.isShifted(element)) {
      const lists = element.children.filter((child) => child.name === 'ul' || child.name === 'ol');

      const listChildren = lists.reduce((acc, list) => {
        const preceding = isShiftedList(list) ? [list] : list.children;
        return preceding.concat(acc);
      }, []);

      for (const list of lists) {
        list.remove();
      }

      for (const child of listChildren) {
        element.add(child);
      }

      delete element.name;
    }
  },

  /**
   * A list item is shifted when it only contains lists or empty paragraphs.
   *
   * @param {Element} element
   * @returns {boolean}
   */
  isShifted(element) {
    if (element.name !== 'li') {
      return false;
    }

    return (
      element.children.filter((child) => {
        if (child.name) {
          if (child.name === 'ul' || child.name === 'ol') {
            return false;
          }

          if (child.name === 'p' && child.children.length === 0) {
            return false;
          }
        }
        return true;
      }).length === 0
    );
  },
};

/**
 * List helpers.
 */
export const List = {
  /**
   * Check if an element is a list item.
   *
   * @param {import('./common.js').FilterConfig} config
   * @param {Element} element
   * @returns {boolean}
   */
  thisIsAListItem(config, element) {
    if (Heuristics.isEdgeListItem(config, element)) {
      return true;
    }

    // A `mso-list: l0 level1 lfo1` style marks a list item. The same style can
    // appear in a paragraph inside a real `<li>`.
    if (
      (element.attributes.style && element.attributes.style.match(/mso-list:\s?l\d/) && element.parent.name !== 'li') ||
      element.attributes['cke-dissolved'] ||
      element.getHtml().match(/<!--\[if !supportLists\]-->/)
    ) {
      return true;
    }

    return false;
  },

  /**
   * Convert an element to a `cke:li` element, with its list data stored in `cke-*` attributes.
   *
   * @param {import('./common.js').FilterConfig} config
   * @param {Element} element
   */
  convertToFakeListItem(config, element) {
    if (Heuristics.isDegenerateListItem(config, element)) {
      Heuristics.assignListLevels(config, element);
    }

    // Cache the parsed list info in the cke-list-* attributes.
    List.getListItemInfo(element);

    if (!element.attributes['cke-dissolved']) {
      // The symbol is the first text descendant that does not start with a space.
      let symbol;

      element.forEach((child) => {
        // Custom markers can be images with a distinctive alt attribute.
        if (!symbol && child.name === 'img' && child.attributes['cke-ignored'] && child.attributes.alt === '*') {
          symbol = '·';
          child.remove();
        }
      }, NODE_ELEMENT);

      element.forEach((child) => {
        if (!symbol && !child.value.match(/^ /)) {
          symbol = child.value;
        }
      }, NODE_TEXT);

      // Without a symbol this is not a list item.
      if (typeof symbol === 'undefined') {
        return;
      }

      element.attributes['cke-symbol'] = symbol.replace(new RegExp('(?: |' + NBSP + ').*$'), '');

      List.removeSymbolText(element);
    }

    const styles = element.attributes && parseCssText(element.attributes.style);

    // Word indents each list level by 0.5in (48px), which matches the default
    // list padding of the browser. Only the indentation beyond that is kept.
    // The margins of the parents are removed later, when the list tree exists.
    if (styles['margin-left']) {
      const level = element.attributes['cke-list-level'];
      let margin = styles['margin-left'];

      // Negative margins are ignored.
      margin = Math.max(convertToPx(margin) - WORD_LIST_INDENT_PX * level, 0);

      if (margin) {
        styles['margin-left'] = margin + 'px';
      } else {
        delete styles['margin-left'];
      }

      element.attributes.style = writeCssText(styles);
    }

    element.name = 'cke:li';
  },

  /**
   * Convert all the `cke:li` elements in `root` to real `<li>` elements.
   *
   * @param {import('./tree.js').Fragment|Element} root
   * @returns {Element[]}
   */
  convertToRealListItems(root) {
    const listElements = [];

    root.forEach(
      (element) => {
        if (element.name === 'cke:li') {
          element.name = 'li';
          listElements.push(element);
        }
      },
      NODE_ELEMENT,
      false,
    );

    return listElements;
  },

  removeSymbolText(element) {
    const symbol = element.attributes['cke-symbol'];
    // The symbol can contain special characters, so use indexOf instead of a RegExp.
    const node = element.findOne((child) => child.value && child.value.indexOf(symbol) > -1, true);

    if (node) {
      node.value = node.value.replace(symbol, '');
      const parent = node.parent;

      if (parent.getHtml().match(/^(\s|&nbsp;)*$/) && parent !== element) {
        parent.remove();
      } else if (!node.value) {
        node.remove();
      }
    }
  },

  setListSymbol(list, symbol, level) {
    level = level || 1;

    const style = parseCssText(list.attributes.style);

    if (list.name === 'ol') {
      if (list.attributes.type || style['list-style-type']) return;

      const typeMap = {
        '[ivx]': 'lower-roman',
        '[IVX]': 'upper-roman',
        '[a-z]': 'lower-alpha',
        '[A-Z]': 'upper-alpha',
        '\\d': 'decimal',
      };

      for (const type of Object.keys(typeMap)) {
        if (List.getSubsectionSymbol(symbol).match(new RegExp(type))) {
          style['list-style-type'] = typeMap[type];
          break;
        }
      }

      list.attributes['cke-list-style-type'] = style['list-style-type'];
    } else {
      const symbolMap = {
        '·': 'disc',
        o: 'circle',
        '§': 'square',
      };

      if (!style['list-style-type'] && symbolMap[symbol]) {
        style['list-style-type'] = symbolMap[symbol];
      }
    }

    List.removeRedundancies(style, level);

    list.attributes.style = writeCssText(style);
    if (!list.attributes.style) {
      delete list.attributes.style;
    }
  },

  // 'disc' and 'decimal' are the defaults in some cases: remove the redundancy.
  removeRedundancies(style, level) {
    if ((level === 1 && style['list-style-type'] === 'disc') || style['list-style-type'] === 'decimal') {
      delete style['list-style-type'];
    }
  },

  setListStart(list) {
    const symbols = [];
    let offset = 0;

    for (const child of list.children) {
      symbols.push(child.attributes['cke-symbol'] || '');
    }

    // A list that starts with a sublist: use the next element as the start indicator.
    if (!symbols[0]) {
      offset++;
    }

    const symbol = symbols[offset] || '';

    switch (list.attributes['cke-list-style-type']) {
      case 'lower-roman':
      case 'upper-roman':
        list.attributes.start = List.toArabic(List.getSubsectionSymbol(symbol)) - offset;
        break;
      case 'lower-alpha':
      case 'upper-alpha':
        list.attributes.start = List.getSubsectionSymbol(symbol).replace(/\W/g, '').toLowerCase().charCodeAt(0) - 96 - offset;
        break;
      case 'decimal':
        list.attributes.start = parseInt(List.getSubsectionSymbol(symbol), 10) - offset || 1;
        break;
      default:
        break;
    }

    if (list.attributes.start == '1' || Number.isNaN(list.attributes.start)) {
      delete list.attributes.start;
    }

    delete list.attributes['cke-list-style-type'];
  },

  numbering: {
    /**
     * Convert a list marker to a number.
     *
     * @param {string} marker
     * @param {string} markerType A CSS `list-style-type` value.
     * @returns {number}
     */
    toNumber(marker, markerType) {
      const fromAlphabet = (str) => {
        const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        str = str.toUpperCase();
        let retVal = 1;
        for (let x = 1; str.length > 0; x *= alphabet.length) {
          retVal += alphabet.indexOf(str.charAt(str.length - 1)) * x;
          str = str.substr(0, str.length - 1);
        }
        return retVal;
      };

      const fromRoman = (str) => {
        const romans = [
          [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
          [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
        ];
        str = str.toUpperCase();
        let retVal = 0;
        for (const [value, numeral] of romans) {
          while (str.substr(0, numeral.length) === numeral) {
            retVal += value;
            str = str.substr(numeral.length);
          }
        }
        return retVal;
      };

      if (markerType === 'decimal') {
        return Number(marker);
      } else if (markerType === 'upper-roman' || markerType === 'lower-roman') {
        return fromRoman(marker.toUpperCase());
      } else if (markerType === 'lower-alpha' || markerType === 'upper-alpha') {
        return fromAlphabet(marker);
      }
      return 1;
    },

    /**
     * Get the list style of a Word marker.
     *
     * The characters `c` and `d` are not roman numerals here. They are 100 and 500,
     * so a list rarely gets there, while alpha lists often start with them.
     *
     * @param {string} marker
     * @returns {string}
     */
    getStyle(marker) {
      const typeMap = {
        i: 'lower-roman',
        v: 'lower-roman',
        x: 'lower-roman',
        l: 'lower-roman',
        m: 'lower-roman',
        I: 'upper-roman',
        V: 'upper-roman',
        X: 'upper-roman',
        L: 'upper-roman',
        M: 'upper-roman',
      };
      const firstCharacter = marker.slice(0, 1);
      let type = typeMap[firstCharacter];

      if (!type) {
        type = 'decimal';

        if (firstCharacter.match(/[a-z]/)) {
          type = 'lower-alpha';
        }
        if (firstCharacter.match(/[A-Z]/)) {
          type = 'upper-alpha';
        }
      }

      return type;
    },
  },

  // For markers such as "1.1.2." get the last part.
  getSubsectionSymbol(symbol) {
    return (symbol.match(/([\da-zA-Z]+).?$/) || ['placeholder', '1'])[1];
  },

  setListDir(list) {
    const dirs = { ltr: 0, rtl: 0 };

    list.forEach((child) => {
      if (child.name === 'li') {
        const dir = child.attributes.dir || child.attributes.DIR || '';
        if (dir.toLowerCase() === 'rtl') {
          dirs.rtl++;
        } else {
          dirs.ltr++;
        }
      }
    }, NODE_ELEMENT);

    if (dirs.rtl > dirs.ltr) {
      list.attributes.dir = 'rtl';
    }
  },

  createList(element) {
    // "o" is a circle marker in unordered lists.
    if ((element.attributes['cke-symbol'].match(/([\da-np-zA-NP-Z]).?/) || [])[1]) {
      return new Element('ol');
    }
    return new Element('ul');
  },

  /**
   * Group the `cke:li` items of the root into real list structures.
   *
   * @param {import('./tree.js').Fragment|Element} root
   * @param {import('./common.js').FilterConfig} config
   * @returns {Element[]} The list items.
   */
  createLists(root, config) {
    const listElements = List.convertToRealListItems(root);

    if (listElements.length === 0) {
      return [];
    }

    const getMargin = (element, margin) => {
      const parsedStyles = parseCssText(element.attributes.style);
      const searchedMargin = 'margin-' + margin;

      if (!(searchedMargin in parsedStyles)) {
        return '';
      }

      const value = convertToPx(parsedStyles[searchedMargin]);

      if (value === 0 && config.keepZeroMargins) {
        return searchedMargin + ': ' + value + '; ';
      }

      if (value > 0) {
        return searchedMargin + ': ' + value + 'px; ';
      }

      return '';
    };

    // Chop the data into continuous lists.
    const lists = List.groupLists(listElements);

    // Create the nested list structures.
    for (const list of lists) {
      let firstLevel1Element = list[0];

      // A level 1 element gives the type of the top-level list.
      for (const item of list) {
        if (item.attributes['cke-list-level'] == 1) {
          firstLevel1Element = item;
          break;
        }
      }

      const containerStack = [List.createList(firstLevel1Element)];
      let innermostContainer = containerStack[0];
      const allContainers = [containerStack[0]];
      const marginTop = getMargin(list[0], 'top');
      const marginBottom = getMargin(list[list.length - 1], 'bottom');

      // Insert the list wrapper before the first known list item.
      innermostContainer.insertBefore(list[0]);
      innermostContainer.attributes.style = marginTop + marginBottom;

      for (const element of list) {
        const level = element.attributes['cke-list-level'];

        while (level > containerStack.length) {
          const content = List.createList(element);

          const children = innermostContainer.children;
          if (children.length > 0) {
            children[children.length - 1].add(content);
          } else {
            const container = new Element('li', { style: 'list-style-type:none' });
            container.add(content);
            innermostContainer.add(container);
          }

          containerStack.push(content);
          allContainers.push(content);
          innermostContainer = content;

          if (level == containerStack.length) {
            List.setListSymbol(content, element.attributes['cke-symbol'], level);
          }
        }

        while (level < containerStack.length) {
          containerStack.pop();
          innermostContainer = containerStack[containerStack.length - 1];

          if (level == containerStack.length) {
            List.setListSymbol(innermostContainer, element.attributes['cke-symbol'], level);
          }
        }

        // This is where the list items go into the lists.
        element.remove();
        innermostContainer.add(element);

        // The list wrapper carries the vertical margins. Remove them from the items.
        const itemStyles = parseCssText(element.attributes.style);
        if ('margin-top' in itemStyles || 'margin-bottom' in itemStyles) {
          delete itemStyles['margin-top'];
          delete itemStyles['margin-bottom'];
          element.attributes.style = writeCssText(itemStyles);
        }
      }

      // Set the symbol of the root (level 1) list.
      if (containerStack[0].children.length) {
        let level1Symbol = containerStack[0].children[0].attributes['cke-symbol'];

        if (!level1Symbol && containerStack[0].children.length > 1) {
          level1Symbol = containerStack[0].children[1].attributes['cke-symbol'];
        }

        if (level1Symbol) {
          List.setListSymbol(containerStack[0], level1Symbol);
        }
      }

      // This can be done only when all the list elements are in place.
      for (const container of allContainers) {
        List.setListStart(container);
      }

      // Apply li[value] when needed. This needs the final ol elements.
      for (const item of list) {
        List.determineListItemValue(item);
      }

      if (!innermostContainer.attributes.style) {
        delete innermostContainer.attributes.style;
      }
      if (!containerStack[0].attributes.style) {
        delete containerStack[0].attributes.style;
      }
    }

    const getParentListItems = (element) => {
      const parents = [];
      let parent = element.parent;

      while (parent) {
        if (parent.name === 'li') {
          parents.push(parent);
        }
        parent = parent.parent;
      }

      return parents;
    };

    const getTotalMarginLeft = (elements) =>
      elements.reduce((total, element) => {
        let marginLeft;
        if (element.attributes && element.attributes.style) {
          marginLeft = parseCssText(element.attributes.style)['margin-left'];
        }
        return marginLeft ? total + parseInt(marginLeft, 10) : total;
      }, 0);

    // Reduce the left margin by the sum of the left margins of the parents.
    for (const element of listElements) {
      const leftOffset = getTotalMarginLeft(getParentListItems(element));

      if (!leftOffset) {
        continue;
      }

      element.attributes = element.attributes || {};

      const styles = parseCssText(element.attributes.style);
      let marginLeft = styles['margin-left'] || 0;
      marginLeft = Math.max(parseInt(marginLeft, 10) - leftOffset, 0);

      if (marginLeft) {
        styles['margin-left'] = marginLeft + 'px';
      } else {
        delete styles['margin-left'];
      }

      element.attributes.style = writeCssText(styles);
    }

    return listElements;
  },

  /**
   * Remove all the `cke-*` helper attributes.
   *
   * @param {Element[]} listElements
   */
  cleanup(listElements) {
    const tempAttributes = ['cke-list-level', 'cke-symbol', 'cke-list-id', 'cke-indentation', 'cke-dissolved'];

    for (const element of listElements) {
      for (const attribute of tempAttributes) {
        delete element.attributes[attribute];
      }
      if (element.attributes.style === '') {
        delete element.attributes.style;
      }
    }
  },

  /**
   * Set the `li[value]` attribute of a list item when its marker does not match its position.
   *
   * @param {Element} element
   */
  determineListItemValue(element) {
    if (element.parent.name !== 'ol') {
      return;
    }

    const assumedValue = List.calculateValue(element);
    let cleanSymbol = (element.attributes['cke-symbol'] || '').match(/[a-z0-9]+/gi);

    if (cleanSymbol) {
      // Markers such as "1.1.4" or "1.A.a.IV" use the last part.
      cleanSymbol = cleanSymbol[cleanSymbol.length - 1];

      const listType = element.parent.attributes['cke-list-style-type'] || List.numbering.getStyle(cleanSymbol);
      const computedValue = List.numbering.toNumber(cleanSymbol, listType);

      if (computedValue !== assumedValue) {
        element.attributes.value = computedValue;
      }
    }
  },

  /**
   * Compute the value of a list item from the preceding items and the list start.
   *
   * @param {Element} element
   * @returns {number}
   */
  calculateValue(element) {
    if (!element.parent) {
      return 1;
    }

    const list = element.parent;
    const elementIndex = element.getIndex();
    let valueFound = null;
    let valueElementIndex;

    for (let i = elementIndex; i >= 0 && valueFound === null; i--) {
      const curElement = list.children[i];

      if (curElement.attributes && curElement.attributes.value !== undefined) {
        valueElementIndex = i;
        valueFound = parseInt(curElement.attributes.value, 10);
      }
    }

    if (valueFound === null) {
      valueFound = list.attributes.start !== undefined ? parseInt(list.attributes.start, 10) : 1;
      valueElementIndex = 0;
    }

    return valueFound + (elementIndex - valueElementIndex);
  },

  /**
   * Replace a real list by paragraphs marked as list items, so that they go
   * through the same list rebuilding as the Word paragraphs.
   *
   * @param {Element} element
   */
  dissolveList(element) {
    const nameIs = (name) => (el) => el.name === name;
    const isList = (el) => nameIs('ul')(el) || nameIs('ol')(el);
    const elements = [];

    element.forEach(
      (child) => {
        elements.push(child);
      },
      NODE_ELEMENT,
      false,
    );

    const items = elements.filter(nameIs('li'));
    const lists = elements.filter(isList);

    const toRoman = (number) => {
      if (number >= 50) return 'l' + toRoman(number - 50);
      if (number >= 40) return 'xl' + toRoman(number - 40);
      if (number >= 10) return 'x' + toRoman(number - 10);
      if (number === 9) return 'ix';
      if (number >= 5) return 'v' + toRoman(number - 5);
      if (number === 4) return 'iv';
      if (number >= 1) return 'i' + toRoman(number - 1);
      return '';
    };

    const countParents = (condition, el) => {
      const count = (parent, number) => {
        if (!parent || !parent.parent) {
          return number;
        }
        return count(parent.parent, condition(parent.parent) ? number + 1 : number);
      };
      return count(el, 0);
    };

    for (const list of lists) {
      let type = list.attributes.type;
      const start = parseInt(list.attributes.start, 10) || 1;
      const level = countParents(isList, list) + 1;

      if (!type) {
        const style = parseCssText(list.attributes.style);
        type = style['list-style-type'];
      }

      list.children.filter(nameIs('li')).forEach((child, index) => {
        let symbol;

        switch (type) {
          case 'disc':
            symbol = '·';
            break;
          case 'circle':
            symbol = 'o';
            break;
          case 'square':
            symbol = '§';
            break;
          case '1':
          case 'decimal':
            symbol = start + index + '.';
            break;
          case 'a':
          case 'lower-alpha':
            symbol = String.fromCharCode('a'.charCodeAt(0) + start - 1 + index) + '.';
            break;
          case 'A':
          case 'upper-alpha':
            symbol = String.fromCharCode('A'.charCodeAt(0) + start - 1 + index) + '.';
            break;
          case 'i':
          case 'lower-roman':
            symbol = toRoman(start + index) + '.';
            break;
          case 'I':
          case 'upper-roman':
            symbol = toRoman(start + index).toUpperCase() + '.';
            break;
          default:
            symbol = list.name === 'ul' ? '·' : start + index + '.';
        }

        child.attributes['cke-symbol'] = symbol;
        child.attributes['cke-list-level'] = level;
      });
    }

    const children = items.reduce((acc, listElement) => {
      const child = listElement.children[0];

      if (child && child.name && child.attributes.style && child.attributes.style.match(/mso-list:/i)) {
        Style.pushStylesLower(listElement, {
          'list-style-type': true,
          display: true,
        });

        const childStyle = parseCssText(child.attributes.style, true);

        Style.setStyle(listElement, 'mso-list', childStyle['mso-list'], true);
        Style.setStyle(child, 'mso-list', '');
        // mso-list wins when the level is determined.
        delete listElement['cke-list-level'];

        // A display style is usually "none". It marks the list element for deletion.
        const styleName = childStyle.display ? 'display' : childStyle.DISPLAY ? 'DISPLAY' : '';
        if (styleName) {
          Style.setStyle(listElement, 'display', childStyle[styleName], true);
        }
      }

      // Skip the elements that only contain another list.
      if (listElement.children.length === 1 && isList(listElement.children[0])) {
        return acc;
      }

      listElement.name = 'p';
      listElement.attributes['cke-dissolved'] = true;
      acc.push(listElement);

      return acc;
    }, []);

    for (let i = children.length - 1; i >= 0; i--) {
      children[i].insertAfter(element);
    }
    for (let i = lists.length - 1; i >= 0; i--) {
      delete lists[i].name;
    }
  },

  groupLists(listElements) {
    const lists = [[listElements[0]]];
    let lastList = lists[0];
    let element = listElements[0];

    element.attributes['cke-indentation'] = element.attributes['cke-indentation'] || getElementIndentation(element);

    for (let i = 1; i < listElements.length; i++) {
      element = listElements[i];
      const previous = listElements[i - 1];

      element.attributes['cke-indentation'] = element.attributes['cke-indentation'] || getElementIndentation(element);

      if (element.previous !== previous) {
        List.chopDiscontinuousLists(lastList, lists);
        lists.push((lastList = []));
      }

      lastList.push(element);
    }

    List.chopDiscontinuousLists(lastList, lists);

    return lists;
  },

  /**
   * Split a flat array of list items into separate lists when the list type
   * or the list id changes.
   *
   * @param {Element[]} list
   * @param {Element[][]} lists All the lists. Modified in place.
   */
  chopDiscontinuousLists(list, lists) {
    const levelSymbols = {};
    const choppedLists = [[]];
    let lastListInfo;

    for (const item of list) {
      const lastSymbol = levelSymbols[item.attributes['cke-list-level']];
      let currentListInfo = List.getListItemInfo(item);
      let currentSymbol;
      let forceType;

      if (lastSymbol) {
        // An "h" before an "i".
        forceType = lastSymbol.type.match(/alpha/) && lastSymbol.index == 7 ? 'alpha' : forceType;
        // An "n" before an "o".
        forceType = item.attributes['cke-symbol'] === 'o' && lastSymbol.index == 14 ? 'alpha' : forceType;

        currentSymbol = List.getSymbolInfo(item.attributes['cke-symbol'], forceType);
        currentListInfo = List.getListItemInfo(item);

        // Chop the list when the type changes, or when the id changes and the
        // list is not a continuation.
        if (
          lastSymbol.type !== currentSymbol.type ||
          (lastListInfo && currentListInfo.id != lastListInfo.id && !List.isAListContinuation(item))
        ) {
          choppedLists.push([]);
        }
      } else {
        currentSymbol = List.getSymbolInfo(item.attributes['cke-symbol']);
      }

      // Reset all the higher levels.
      for (let j = parseInt(item.attributes['cke-list-level'], 10) + 1; j < 20; j++) {
        if (levelSymbols[j]) {
          delete levelSymbols[j];
        }
      }

      levelSymbols[item.attributes['cke-list-level']] = currentSymbol;
      choppedLists[choppedLists.length - 1].push(item);

      lastListInfo = currentListInfo;
    }

    lists.splice(lists.indexOf(list), 1, ...choppedLists);
  },

  /**
   * Check if a list item continues a list that was interrupted by a nested list
   * with another id.
   *
   * @param {Element} listElement
   * @returns {boolean}
   */
  isAListContinuation(listElement) {
    let prev = listElement;

    do {
      prev = prev.previous;

      if (prev && prev.type === NODE_ELEMENT) {
        if (prev.attributes['cke-list-level'] === undefined) {
          // Not a list, so the list is interrupted.
          return false;
        }

        if (prev.attributes['cke-list-level'] === listElement.attributes['cke-list-level']) {
          return prev.attributes['cke-list-id'] === listElement.attributes['cke-list-id'];
        }
      }
    } while (prev);

    return false;
  },

  toArabic(symbol) {
    if (!symbol.match(/[ivxl]/i)) return 0;
    if (symbol.match(/^l/i)) return 50 + List.toArabic(symbol.slice(1));
    if (symbol.match(/^lx/i)) return 40 + List.toArabic(symbol.slice(1));
    if (symbol.match(/^x/i)) return 10 + List.toArabic(symbol.slice(1));
    if (symbol.match(/^ix/i)) return 9 + List.toArabic(symbol.slice(2));
    if (symbol.match(/^v/i)) return 5 + List.toArabic(symbol.slice(1));
    if (symbol.match(/^iv/i)) return 4 + List.toArabic(symbol.slice(2));
    if (symbol.match(/^i/i)) return 1 + List.toArabic(symbol.slice(1));
    // Other characters are ignored.
    return List.toArabic(symbol.slice(1));
  },

  /**
   * Describe a list marker.
   *
   * @param {string} symbol
   * @param {string} [type]
   * @returns {{index: number, type: string}}
   */
  getSymbolInfo(symbol, type) {
    symbol = symbol || '';
    const symbolCase = symbol.toUpperCase() === symbol ? 'upper-' : 'lower-';
    const symbolMap = {
      '·': ['disc', -1],
      o: ['circle', -2],
      '§': ['square', -3],
    };

    if (symbol in symbolMap || (type && type.match(/(disc|circle|square)/))) {
      const entry = symbolMap[symbol] || [type, -1];
      return {
        index: entry[1],
        type: entry[0],
      };
    }

    if (symbol.match(/\d/)) {
      return {
        index: symbol ? parseInt(List.getSubsectionSymbol(symbol), 10) : 0,
        type: 'decimal',
      };
    }

    symbol = symbol.replace(/\W/g, '').toLowerCase();

    if ((!type && symbol.match(/[ivxl]+/i)) || (type && type !== 'alpha') || type === 'roman') {
      return {
        index: List.toArabic(symbol),
        type: symbolCase + 'roman',
      };
    }

    if (symbol.match(/[a-z]/i)) {
      return {
        index: symbol.charCodeAt(0) - 97,
        type: symbolCase + 'alpha',
      };
    }

    return {
      index: -1,
      type: 'disc',
    };
  },

  /**
   * Get the list id and level of an item from its `mso-list` style.
   *
   * @param {Element} list
   * @returns {{id: string, level: string}}
   */
  getListItemInfo(list) {
    if (list.attributes['cke-list-id'] !== undefined) {
      return {
        id: list.attributes['cke-list-id'],
        level: list.attributes['cke-list-level'],
      };
    }

    let propValue = parseCssText(list.attributes.style)['mso-list'];
    const ret = {
      id: '0',
      level: '1',
    };

    if (propValue) {
      // Add a whitespace so that all the values are separated with \s.
      propValue += ' ';

      const levelMatch = propValue.match(/level(.+?)\s+/);
      const idMatch = propValue.match(/l(\d+?)\s+/);
      if (levelMatch) ret.level = levelMatch[1];
      if (idMatch) ret.id = idMatch[1];
    }

    // The list level is reused when present.
    list.attributes['cke-list-level'] =
      list.attributes['cke-list-level'] !== undefined ? list.attributes['cke-list-level'] : ret.level;
    list.attributes['cke-list-id'] = ret.id;

    return ret;
  },
};
