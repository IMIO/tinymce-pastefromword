import { test } from 'node:test';
import assert from 'node:assert/strict';
import { extractGroupContent, getGroup, getGroups, removeGroups } from '../src/filter/rtf.js';

test('getGroup extracts a group', () => {
  const expectedGroup = '{\\comandeer test}';
  const rtf = '{\\somegroup\\withcontrol\\words' + expectedGroup + '}';
  const startIndex = rtf.indexOf(expectedGroup);

  assert.deepEqual(getGroup(rtf, 'comandeer'), {
    start: startIndex,
    end: startIndex + expectedGroup.length,
    content: expectedGroup,
  });
});

test('getGroup accepts a regex-like name', () => {
  const expectedGroup = '{\\comandeer1 valid}';
  const unexpectedGroup = '{\\comandeer\\d invalid}';
  const rtf = '{\\somegroup\\withcontrol\\words' + unexpectedGroup + expectedGroup + '}';

  assert.equal(getGroup(rtf, 'comandeer\\d').content, expectedGroup);
});

test('getGroup starts the search at the given position', () => {
  const expectedGroup = '{\\comandeer test}';
  const rtf = '{\\somegroup\\withcontrol\\words{\\comandeer fail}}' + expectedGroup + '{\\comandeer fail}';

  assert.equal(getGroup(rtf, 'comandeer', { start: rtf.indexOf(expectedGroup) }).content, expectedGroup);
  assert.equal(getGroup(rtf, 'missing'), null);
});

test('getGroup handles nested groups and escaped braces', () => {
  const rtf = '{\\rtf1{\\pict{\\*\\picprop{\\sp{\\sn x}{\\sv \\}y\\{}}}\\pngblip 00ff}{\\other}}';
  const group = getGroup(rtf, 'pict');

  assert.equal(group.content, '{\\pict{\\*\\picprop{\\sp{\\sn x}{\\sv \\}y\\{}}}\\pngblip 00ff}');
});

test('getGroups and removeGroups', () => {
  const rtf = '{\\rtf1\\some\\control\\words{\\group content}{\\group content}{\\whatever {\\subgroup content}}}';

  assert.deepEqual(getGroups(rtf, '(group|whatever)'), [
    { start: 25, end: 41, content: '{\\group content}' },
    { start: 41, end: 57, content: '{\\group content}' },
    { start: 57, end: 88, content: '{\\whatever {\\subgroup content}}' },
  ]);

  assert.equal(removeGroups(rtf, '(group|whatever)'), '{\\rtf1\\some\\control\\words}');
});

test('extractGroupContent strips control words and subgroups', () => {
  assert.equal(extractGroupContent('{\\group{\\subgroup subgroupcontent} group content}'), 'group content');
  assert.equal(extractGroupContent('{\\pict\\picw300\\pich300\\pngblip d76df8e7aefc}'), 'd76df8e7aefc');
  assert.equal(extractGroupContent('{\\pict{\\*\\picprop{\\sp{\\sn wzDescription}{\\sv }}}\\picscalex19\\pngblip d76df8e7aefc }').trim(), 'd76df8e7aefc');
});
