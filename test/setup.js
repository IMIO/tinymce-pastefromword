// Give the filter a DOMParser in Node with jsdom.
import { JSDOM } from 'jsdom';

const dom = new JSDOM('<!doctype html><html><body></body></html>');

globalThis.DOMParser = dom.window.DOMParser;
globalThis.Node = dom.window.Node;

/**
 * Normalize HTML for comparisons: lower-case, no whitespace between tags,
 * sorted attributes and sorted style declarations.
 *
 * @param {string} html
 * @returns {string}
 */
export function normalizeHtml(html) {
  const doc = new dom.window.DOMParser().parseFromString('<body>' + html, 'text/html');
  const BLOCKS = /^(p|div|h[1-6]|ul|ol|li|table|thead|tbody|tfoot|tr|td|th|blockquote|pre|hr|br|img|body)$/;
  const isBlock = (node) => node.nodeType === 1 && BLOCKS.test(node.localName.toLowerCase());

  const serialize = (node) => {
    if (node.nodeType === 3) {
      return node.data.replace(/\s+/g, ' ');
    }
    if (node.nodeType === 8) {
      return '<!--' + node.data + '-->';
    }
    if (node.nodeType !== 1) {
      return '';
    }

    const name = node.localName.toLowerCase();
    const attrs = Array.from(node.attributes)
      .map((attr) => {
        let value = attr.value;
        if (attr.name === 'style') {
          value = value
            .split(';')
            .map((decl) => decl.trim().replace(/\s*:\s*/, ':').replace(/"/g, ''))
            .filter(Boolean)
            .sort()
            .join('; ');
        }
        return attr.name + '="' + value + '"';
      })
      .sort()
      .join(' ');

    const childNodes = Array.from(node.childNodes).filter((child, i, all) => {
      // Whitespace next to a block is not content.
      if (child.nodeType === 3 && !child.data.trim()) {
        const prev = all[i - 1];
        const next = all[i + 1];
        return !(!prev || !next || isBlock(prev) || isBlock(next));
      }
      return true;
    });

    let children = childNodes.map(serialize).join('');
    if (isBlock(node)) {
      children = children.replace(/^\s+|\s+$/g, '');
    }
    const open = '<' + name + (attrs ? ' ' + attrs : '') + '>';

    if (/^(br|img|hr|col|input|meta|link)$/.test(name)) {
      return open;
    }
    return open + children + '</' + name + '>';
  };

  return (
    serialize(doc.body)
      .replace(/^<body>|<\/body>$/g, '')
      // Whitespace at the start or end of a block, even inside inline elements, is not content.
      .replace(/(<(?:p|li|h[1-6]|td|th|div)(?:\s[^>]*)?>(?:<[^/][^>]*>)*)\s+/g, '$1')
      .replace(/\s+((?:<\/[^>]+>)*<\/(?:p|li|h[1-6]|td|th|div)>)/g, '$1')
      .replace(/<p><br><\/p>/g, '<p></p>')
      .toLowerCase()
  );
}
