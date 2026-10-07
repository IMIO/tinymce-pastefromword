/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (core/htmlparser) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

// Lightweight node tree with the same shape as the CKEditor 4 htmlParser nodes.
// The paste filter rules mutate `name`, `attributes` and `children` directly,
// so the tree keeps those as plain properties instead of a live DOM.

export const NODE_ELEMENT = 1;
export const NODE_TEXT = 3;
export const NODE_COMMENT = 8;
export const NODE_DOCUMENT_FRAGMENT = 11;

const VOID_ELEMENTS = new Set([
  'area', 'base', 'basefont', 'br', 'col', 'command', 'embed', 'frame', 'hr',
  'img', 'input', 'isindex', 'keygen', 'link', 'meta', 'param', 'source',
  'track', 'wbr',
]);

const BLOCK_LIKE_ELEMENTS = new Set([
  'address', 'article', 'aside', 'blockquote', 'body', 'caption', 'center',
  'col', 'colgroup', 'dd', 'details', 'dialog', 'dir', 'div', 'dl', 'dt',
  'fieldset', 'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3',
  'h4', 'h5', 'h6', 'head', 'header', 'hgroup', 'hr', 'html', 'legend', 'li',
  'main', 'menu', 'meta', 'nav', 'noscript', 'ol', 'p', 'pre', 'script',
  'section', 'style', 'summary', 'table', 'tbody', 'td', 'tfoot', 'th',
  'thead', 'title', 'tr', 'ul', 'br',
]);

function isBlockLike(name) {
  const realName = (name || '').replace(/^cke:/, '');
  return BLOCK_LIKE_ELEMENTS.has(realName);
}

function escapeText(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\u00a0/g, '&nbsp;');
}

function escapeAttribute(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Base class for all nodes. */
export class Node {
  constructor() {
    this.parent = null;
    this.previous = null;
    this.next = null;
  }

  remove() {
    const children = this.parent.children;
    const index = children.indexOf(this);
    const previous = this.previous;
    const next = this.next;

    if (previous) previous.next = next;
    if (next) next.previous = previous;
    children.splice(index, 1);
    this.parent = null;
  }

  replaceWith(node) {
    const children = this.parent.children;
    const index = children.indexOf(this);
    const previous = (node.previous = this.previous);
    const next = (node.next = this.next);

    if (previous) previous.next = node;
    if (next) next.previous = node;

    children[index] = node;

    node.parent = this.parent;
    this.parent = null;
  }

  insertAfter(node) {
    const children = node.parent.children;
    const index = children.indexOf(node);
    const next = node.next;

    children.splice(index + 1, 0, this);

    this.next = node.next;
    this.previous = node;
    node.next = this;
    if (next) next.previous = this;

    this.parent = node.parent;
  }

  insertBefore(node) {
    const children = node.parent.children;
    const index = children.indexOf(node);

    children.splice(index, 0, this);

    this.next = node;
    this.previous = node.previous;
    if (node.previous) node.previous.next = this;
    node.previous = this;

    this.parent = node.parent;
  }

  getAscendant(condition) {
    const check =
      typeof condition === 'function'
        ? condition
        : typeof condition === 'string'
          ? (el) => el.name === condition
          : (el) => el.name in condition;

    let parent = this.parent;

    while (parent && parent.type === NODE_ELEMENT) {
      if (check(parent)) return parent;
      parent = parent.parent;
    }

    return null;
  }

  wrapWith(wrapper) {
    this.replaceWith(wrapper);
    wrapper.add(this);
    return wrapper;
  }

  getIndex() {
    return this.parent.children.indexOf(this);
  }
}

export class Text extends Node {
  constructor(value) {
    super();
    this.type = NODE_TEXT;
    this.value = value;
    this._ = { isBlockLike: false };
  }

  filter(filter) {
    const value = filter.onText(this.value, this);
    if (!value) {
      this.remove();
      return false;
    }
    this.value = value;
    return true;
  }

  writeHtml(out) {
    out.push(escapeText(this.value));
  }

  getOuterHtml() {
    const out = [];
    this.writeHtml(out);
    return out.join('');
  }
}

export class Comment extends Node {
  constructor(value) {
    super();
    this.type = NODE_COMMENT;
    this.value = value;
    this._ = { isBlockLike: false };
  }

  filter(filter) {
    const comment = filter.onComment(this.value, this);

    if (!comment) {
      this.remove();
      return false;
    }

    if (typeof comment !== 'string') {
      this.replaceWith(comment);
      return false;
    }

    this.value = comment;
    return true;
  }

  writeHtml(out) {
    out.push('<!--', this.value, '-->');
  }

  getOuterHtml() {
    const out = [];
    this.writeHtml(out);
    return out.join('');
  }
}

/** Methods shared by Fragment and Element. */
const containerMethods = {
  add(node, index) {
    if (Number.isNaN(Number(index)) || index === undefined) {
      index = this.children.length;
    }

    const previous = index > 0 ? this.children[index - 1] : null;

    if (previous) {
      // A block that follows text: trim the spaces at the right of the text.
      if (node._.isBlockLike && previous.type === NODE_TEXT) {
        previous.value = previous.value.replace(/[ \t\n\r]+$/, '');

        if (previous.value.length === 0) {
          this.children.pop();
          this.add(node);
          return;
        }
      }

      previous.next = node;
    }

    node.previous = previous;
    node.parent = this;
    this.children.splice(index, 0, node);
  },

  filterChildren(filter) {
    // Children filtered twice by the same filter would be damaged.
    if (this.childrenFilteredBy === filter.id) {
      return;
    }

    this.childrenFilteredBy = filter.id;

    // The children array may change while the rules run, so never cache the length.
    for (let i = 0; i < this.children.length; i++) {
      if (this.children[i].filter(filter) === false) {
        i--;
      }
    }
  },

  forEach(callback, type, skipRoot) {
    let ret;
    if (!skipRoot && (!type || this.type === type)) {
      ret = callback(this);
    }

    if (ret === false) {
      return;
    }

    const children = this.children;

    // The callback may change the children list, so never cache the length.
    for (let i = 0; i < children.length; i++) {
      const node = children[i];
      if (node.type === NODE_ELEMENT) {
        node.forEach(callback, type);
      } else if (!type || node.type === type) {
        callback(node);
      }
    }
  },

  writeChildrenHtml(out) {
    for (const child of this.children) {
      child.writeHtml(out);
    }
  },

  getHtml() {
    const out = [];
    this.writeChildrenHtml(out);
    return out.join('');
  },

  getFirst(condition) {
    if (!condition) {
      return this.children.length ? this.children[0] : null;
    }

    if (typeof condition !== 'function') {
      condition = nameCondition(condition);
    }

    for (const child of this.children) {
      if (condition(child)) return child;
    }
    return null;
  },

  find(criteria, recursive = false) {
    let ret = [];

    for (const child of this.children) {
      if (typeof criteria === 'function' && criteria(child)) {
        ret.push(child);
      } else if (typeof criteria === 'string' && child.name === criteria) {
        ret.push(child);
      }

      if (recursive && child.find) {
        ret = ret.concat(child.find(criteria, recursive));
      }
    }

    return ret;
  },

  findOne(criteria, recursive) {
    let nestedMatch = null;
    const match = this.children.find((child) => {
      const isMatching = typeof criteria === 'function' ? criteria(child) : child.name === criteria;

      if (isMatching || !recursive) {
        return isMatching;
      }

      if (child.children && child.findOne) {
        nestedMatch = child.findOne(criteria, true);
      }

      return !!nestedMatch;
    });

    return nestedMatch || match || null;
  },
};

function nameCondition(condition) {
  return (el) =>
    el.type === NODE_ELEMENT &&
    (typeof condition === 'string' ? el.name === condition : el.name in condition);
}

export class Fragment {
  constructor() {
    this.type = NODE_DOCUMENT_FRAGMENT;
    this.parent = null;
    this.children = [];
    this._ = { isBlockLike: true };
  }

  filter(filter) {
    filter.onRoot(this);
    this.filterChildren(filter);
  }

  writeHtml(out) {
    this.writeChildrenHtml(out);
  }

  /**
   * Build a fragment from an HTML string with the DOM parser of the environment.
   *
   * @param {string} html
   * @returns {Fragment}
   */
  static fromHtml(html) {
    const doc = new globalThis.DOMParser().parseFromString(html, 'text/html');
    return Fragment.fromDom(doc.body);
  }

  /**
   * Build a fragment from the children of a DOM node.
   *
   * @param {globalThis.Node} domNode
   * @returns {Fragment}
   */
  static fromDom(domNode) {
    const fragment = new Fragment();
    appendDomChildren(fragment, domNode, false);
    return fragment;
  }
}

Object.assign(Fragment.prototype, containerMethods);

export class Element extends Node {
  constructor(name, attributes) {
    super();
    this.type = NODE_ELEMENT;
    this.name = name;
    this.attributes = attributes || {};
    this.children = [];
    this._ = { isBlockLike: isBlockLike(name) };
  }

  get isEmpty() {
    return VOID_ELEMENTS.has(this.name);
  }

  clone() {
    return new Element(this.name, { ...this.attributes });
  }

  filter(filter) {
    let element = this;
    let originalName;
    let name;

    if (!element.parent) {
      filter.onRoot(element);
    }

    for (;;) {
      originalName = element.name;

      if (!(name = filter.onElementName(originalName))) {
        this.remove();
        return false;
      }

      element.name = name;

      if (!(element = filter.onElement(element))) {
        this.remove();
        return false;
      }

      // A new node replaces this one. The new node is filtered by the caller.
      if (element !== this) {
        this.replaceWith(element);
        return false;
      }

      // The name did not change: stop. Otherwise apply the rules of the new name.
      if (element.name === originalName) {
        break;
      }

      if (element.type !== NODE_ELEMENT) {
        this.replaceWith(element);
        return false;
      }

      // The rule dropped the element but kept the children.
      if (!element.name) {
        this.replaceWithChildren();
        return false;
      }
    }

    const attributes = element.attributes;

    for (const originalAttrName of Object.keys(attributes)) {
      let attrName = originalAttrName;
      let value = attributes[attrName];
      let newAttrName;

      for (;;) {
        newAttrName = filter.onAttributeName(attrName);
        if (!newAttrName) {
          delete attributes[attrName];
          break;
        } else if (newAttrName !== attrName) {
          delete attributes[attrName];
          attrName = newAttrName;
          continue;
        } else {
          break;
        }
      }

      if (newAttrName) {
        value = filter.onAttribute(element, newAttrName, value);
        if (value === false) {
          delete attributes[newAttrName];
        } else {
          attributes[newAttrName] = value;
        }
      }
    }

    if (!element.isEmpty) {
      this.filterChildren(filter);
    }

    return true;
  }

  writeHtml(out) {
    const name = this.name;

    out.push('<', name);

    for (const attrName of Object.keys(this.attributes)) {
      out.push(' ', attrName, '="', escapeAttribute(this.attributes[attrName]), '"');
    }

    out.push('>');

    if (!this.isEmpty) {
      this.writeChildrenHtml(out);
      out.push('</', name, '>');
    }
  }

  getOuterHtml() {
    const out = [];
    this.writeHtml(out);
    return out.join('');
  }

  replaceWithChildren() {
    const children = this.children;

    for (let i = children.length; i; ) {
      children[--i].insertAfter(this);
    }

    this.remove();
  }

  setHtml(html) {
    const fragment = Fragment.fromHtml(html);
    this.children = fragment.children;

    for (const child of this.children) {
      child.parent = this;
    }
  }

  hasClass(className) {
    const classes = this.attributes.class;

    if (!classes) return false;

    return new RegExp('(?:^|\\s)' + className + '(?=\\s|$)').test(classes);
  }

  addClass(className) {
    if (this.hasClass(className)) return;

    const c = this.attributes.class || '';
    this.attributes.class = c + (c ? ' ' : '') + className;
  }

  removeClass(className) {
    let classes = this.attributes.class;

    if (!classes) return;

    classes = classes.replace(new RegExp('(?:\\s+|^)' + className + '(?:\\s+|$)'), ' ').trim();

    if (classes) {
      this.attributes.class = classes;
    } else {
      delete this.attributes.class;
    }
  }
}

Object.assign(Element.prototype, containerMethods);

const DOM_ELEMENT = 1;
const DOM_TEXT = 3;
const DOM_CDATA = 4;
const DOM_COMMENT = 8;

const WHITESPACE_ONLY = /^[ \t\n\r]*$/;

function rtrim(text) {
  return text.replace(/[ \t\n\r]+$/, '');
}

// The CKEditor parser drops the whitespace between blocks and collapses the
// whitespace runs inside text. The DOM parser keeps them, so do the same here.
function appendDomChildren(container, domNode, inPre) {
  const isBlockContainer = container.type === NODE_DOCUMENT_FRAGMENT || container._.isBlockLike;
  let hasInlineStarted = !isBlockContainer;
  let previous = null;

  for (const domChild of domNode.childNodes) {
    let node;

    if (domChild.nodeType === DOM_TEXT || domChild.nodeType === DOM_CDATA) {
      let text = domChild.data;

      if (!inPre) {
        // Trim the spaces at the start of a block.
        if (!hasInlineStarted) {
          text = text.replace(/^[ \t\n\r]+/, '');
          if (text.length === 0) {
            continue;
          }
        }
        text = text.replace(/[\t\r\n ]{2,}|[\t\r\n]/g, ' ');
      }

      node = new Text(text);
    } else {
      node = fromDomNode(domChild, inPre);
      if (!node) continue;
    }

    // A block that follows text: trim the spaces at the right of the text.
    if (node._.isBlockLike && previous && previous.type === NODE_TEXT && !inPre) {
      previous.value = rtrim(previous.value);
      if (previous.value.length === 0) {
        container.children.pop();
        previous = container.children.length ? container.children[container.children.length - 1] : null;
        if (previous) previous.next = null;
      }
    }

    node.parent = container;
    node.previous = previous;
    if (previous) previous.next = node;
    container.children.push(node);
    previous = node;

    if (node.type === NODE_TEXT || (node.type === NODE_ELEMENT && !node._.isBlockLike)) {
      hasInlineStarted = true;
    }
  }

  // Trim the spaces at the end of a block.
  if (isBlockContainer && !inPre) {
    const last = container.children[container.children.length - 1];
    if (last && last.type === NODE_TEXT) {
      last.value = rtrim(last.value);
      if (WHITESPACE_ONLY.test(last.value)) {
        container.children.pop();
        const newLast = container.children[container.children.length - 1];
        if (newLast) newLast.next = null;
      }
    }
  }
}

function fromDomNode(domNode, inPre) {
  switch (domNode.nodeType) {
    case DOM_ELEMENT: {
      const attributes = {};
      for (const attr of domNode.attributes) {
        attributes[attr.name] = attr.value;
      }
      const name = domNode.localName.toLowerCase();
      const element = new Element(name, attributes);
      if (!element.isEmpty) {
        appendDomChildren(element, domNode, inPre || name === 'pre' || name === 'textarea');
      }
      return element;
    }
    case DOM_TEXT:
    case DOM_CDATA:
      return new Text(domNode.data);
    case DOM_COMMENT:
      return new Comment(domNode.data);
    default:
      return null;
  }
}
