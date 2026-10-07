/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (core/htmlparser/filter.js) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

import { Comment, Node, Text, Fragment, NODE_COMMENT, NODE_ELEMENT, NODE_TEXT } from './tree.js';

let nextFilterId = 1;

class RulesGroup {
  constructor() {
    this.rules = [];
  }

  add(rule, priority) {
    this.rules.splice(this.findIndex(priority), 0, { value: rule, priority });
  }

  addMany(rules, priority) {
    const items = rules.map((rule) => ({ value: rule, priority }));
    this.rules.splice(this.findIndex(priority), 0, ...items);
  }

  findIndex(priority) {
    const rules = this.rules;
    let i = rules.length - 1;

    while (i >= 0 && priority < rules[i].priority) {
      i--;
    }

    return i + 1;
  }

  exec(currentValue, ...rest) {
    const isNode = currentValue instanceof Node || currentValue instanceof Fragment;
    const args = [currentValue, ...rest];
    let orgType;
    let orgName;

    for (const rule of this.rules) {
      if (isNode) {
        orgType = currentValue.type;
        orgName = currentValue.name;
      }

      const ret = rule.value(...args);

      if (ret === false) {
        return ret;
      }

      // A node that is no longer of the same kind must not go through the next rules.
      if (isNode && ret && (ret.name !== orgName || ret.type !== orgType)) {
        return ret;
      }

      if (ret != null) {
        args[0] = currentValue = ret;
      }
    }

    return currentValue;
  }

  execOnName(currentName) {
    for (const rule of this.rules) {
      if (!currentName) break;
      currentName = currentName.replace(rule.value[0], rule.value[1]);
    }

    return currentName;
  }
}

function addNamedRules(groups, newRules, priority) {
  for (const ruleName of Object.keys(newRules)) {
    if (!groups[ruleName]) {
      groups[ruleName] = new RulesGroup();
    }
    groups[ruleName].add(newRules[ruleName], priority);
  }
}

/**
 * Rule based filter for the node tree.
 *
 * Rules object shape:
 * - `elementNames`: array of `[RegExp, replacement]` pairs applied to element names.
 * - `attributeNames`: array of `[RegExp, replacement]` pairs applied to attribute names.
 * - `elements`: map of element name to `(element) => Element | false | undefined`.
 *   The `^` key runs for every element before the named rule.
 * - `attributes`: map of attribute name to `(value, element) => string | false`.
 * - `text`: `(value, node) => string`.
 * - `comment`: `(value, node) => string | false`.
 * - `root`: `(fragment) => void`.
 */
export class Filter {
  constructor(rules) {
    this.id = nextFilterId++;
    this.elementNameRules = new RulesGroup();
    this.attributeNameRules = new RulesGroup();
    this.elementsRules = {};
    this.attributesRules = {};
    this.textRules = new RulesGroup();
    this.commentRules = new RulesGroup();
    this.rootRules = new RulesGroup();

    if (rules) {
      this.addRules(rules, 10);
    }
  }

  addRules(rules, priority = 10) {
    if (rules.elementNames) this.elementNameRules.addMany(rules.elementNames, priority);
    if (rules.attributeNames) this.attributeNameRules.addMany(rules.attributeNames, priority);
    if (rules.elements) addNamedRules(this.elementsRules, rules.elements, priority);
    if (rules.attributes) addNamedRules(this.attributesRules, rules.attributes, priority);
    if (rules.text) this.textRules.add(rules.text, priority);
    if (rules.comment) this.commentRules.add(rules.comment, priority);
    if (rules.root) this.rootRules.add(rules.root, priority);
  }

  applyTo(node) {
    node.filter(this);
  }

  onElementName(name) {
    return this.elementNameRules.execOnName(name);
  }

  onAttributeName(name) {
    return this.attributeNameRules.execOnName(name);
  }

  onText(text, node) {
    return this.textRules.exec(text, node);
  }

  onComment(commentText, comment) {
    return this.commentRules.exec(commentText, comment);
  }

  onRoot(element) {
    return this.rootRules.exec(element);
  }

  onElement(element) {
    const groups = [this.elementsRules['^'], this.elementsRules[element.name], this.elementsRules.$];

    for (const group of groups) {
      if (!group) continue;

      const ret = group.exec(element, this);

      if (ret === false) {
        return null;
      }

      if (ret && ret !== element) {
        return this.onNode(ret);
      }

      // A rule dismissed the element but kept the children.
      if (element.parent && !element.name) {
        break;
      }
    }

    return element;
  }

  onNode(node) {
    const type = node.type;

    if (type === NODE_ELEMENT) return this.onElement(node);
    if (type === NODE_TEXT) return new Text(this.onText(node.value, node));
    if (type === NODE_COMMENT) return new Comment(this.onComment(node.value, node));
    return null;
  }

  onAttribute(element, name, value) {
    const group = this.attributesRules[name];

    if (group) {
      return group.exec(value, element, this);
    }
    return value;
  }
}
