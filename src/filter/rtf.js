/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (plugins/pastetools/filter/common.js) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

// Minimal RTF helpers. A group is `{\name ...}`. Groups can contain subgroups.

/**
 * @typedef {Object} GroupInfo
 * @property {number} start Index where the group starts.
 * @property {number} end Index after the group end.
 * @property {string} content The whole group, with control words and subgroups.
 */

function getNonWhitespaceChar(content, startIndex, direction) {
  let index = startIndex + direction;
  let current = content[index];

  while (current && /\s/.test(current)) {
    index += direction;
    current = content[index];
  }

  return current;
}

function getGroupName(group) {
  const match = group.match(/^\{\\(\w+)/);
  return match ? match[1] : null;
}

/**
 * Find the first group with the given name.
 *
 * @param {string} content RTF content.
 * @param {string} groupName Group name. It can be a regex-like string.
 * @param {{start?: number}} [options]
 * @returns {GroupInfo|null}
 */
export function getGroup(content, groupName, options = {}) {
  let open = 0;
  const startRegex = new RegExp('\\{\\\\' + groupName, 'g');

  startRegex.lastIndex = options.start || 0;
  const group = startRegex.exec(content);

  if (!group) {
    return null;
  }

  let i = group.index;
  let current = content[i];

  do {
    // A group starts with `{\`. Whitespace between the two characters is allowed.
    // Escaped curly braces are not group delimiters.
    const isValidGroupStart =
      current === '{' &&
      getNonWhitespaceChar(content, i, -1) !== '\\' &&
      getNonWhitespaceChar(content, i, 1) === '\\';
    const isValidGroupEnd = current === '}' && getNonWhitespaceChar(content, i, -1) !== '\\' && open > 0;

    if (isValidGroupStart) {
      open++;
    } else if (isValidGroupEnd) {
      open--;
    }

    current = content[++i];
  } while (current && open > 0);

  return {
    start: group.index,
    end: i,
    content: content.substring(group.index, i),
  };
}

/**
 * Find all groups with the given name.
 *
 * @param {string} rtfContent
 * @param {string} groupName Group name. It can be a regex-like string.
 * @returns {GroupInfo[]}
 */
export function getGroups(rtfContent, groupName) {
  const groups = [];
  let from = 0;
  let current;

  while ((current = getGroup(rtfContent, groupName, { start: from }))) {
    from = current.end;
    groups.push(current);
  }

  return groups;
}

/**
 * Remove all groups with the given name.
 *
 * @param {string} rtfContent
 * @param {string} groupName Group name. It can be a regex-like string.
 * @returns {string}
 */
export function removeGroups(rtfContent, groupName) {
  let current;

  while ((current = getGroup(rtfContent, groupName))) {
    rtfContent = rtfContent.substring(0, current.start) + rtfContent.substring(current.end);
  }

  return rtfContent;
}

/**
 * Get the content of a group without its control words and subgroups.
 *
 * @param {string} group The whole group string.
 * @returns {string}
 */
export function extractGroupContent(group) {
  const groupName = getGroupName(group);
  const controlWordsRegex = /^\{(\\[\w-]+\s*)+/g;
  // Content can follow the last subgroup without a space. Add one to parse it.
  const subgroupWithoutSpaceRegex = /\}([^{\s]+)/g;

  group = group.replace(subgroupWithoutSpaceRegex, '} $1');
  group = removeGroups(group, '(?!' + groupName + ')');
  group = group.replace(controlWordsRegex, '').trim();

  return group.replace(/}$/, '');
}
