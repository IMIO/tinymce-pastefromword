import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { List } from '../src/filter/lists.js';
import { cleanWord } from '../src/filter/word.js';
import { normalizeHtml } from './setup.js';

test('numbering.toNumber converts markers', () => {
  assert.equal(List.numbering.toNumber('XIV', 'upper-roman'), 14);
  assert.equal(List.numbering.toNumber('iv', 'lower-roman'), 4);
  assert.equal(List.numbering.toNumber('d', 'lower-alpha'), 4);
  assert.equal(List.numbering.toNumber('z', 'lower-alpha'), 26);
  assert.equal(List.numbering.toNumber('35', 'decimal'), 35);
  assert.equal(List.numbering.toNumber('404', 'foo'), 1);
});

test('numbering.getStyle guesses the list style of a marker', () => {
  assert.equal(List.numbering.getStyle('4'), 'decimal');
  assert.equal(List.numbering.getStyle('b'), 'lower-alpha');
  assert.equal(List.numbering.getStyle('P'), 'upper-alpha');
  assert.equal(List.numbering.getStyle('i'), 'lower-roman');
  assert.equal(List.numbering.getStyle('X'), 'upper-roman');
  assert.equal(List.numbering.getStyle('c'), 'lower-alpha');
});

test('getSymbolInfo describes markers', () => {
  assert.deepEqual(List.getSymbolInfo('·'), { index: -1, type: 'disc' });
  assert.deepEqual(List.getSymbolInfo('o'), { index: -2, type: 'circle' });
  assert.deepEqual(List.getSymbolInfo('§'), { index: -3, type: 'square' });
  assert.deepEqual(List.getSymbolInfo('3.'), { index: 3, type: 'decimal' });
  assert.deepEqual(List.getSymbolInfo('1.2.5.'), { index: 5, type: 'decimal' });
  assert.deepEqual(List.getSymbolInfo('iv.'), { index: 4, type: 'lower-roman' });
  assert.deepEqual(List.getSymbolInfo('IV)'), { index: 4, type: 'upper-roman' });
  assert.deepEqual(List.getSymbolInfo('b.'), { index: 1, type: 'lower-alpha' });
  assert.deepEqual(List.getSymbolInfo('h.', 'alpha'), { index: 7, type: 'lower-alpha' });
  assert.deepEqual(List.getSymbolInfo('B.'), { index: 1, type: 'upper-alpha' });
  assert.deepEqual(List.getSymbolInfo('-'), { index: -1, type: 'disc' });
});

test('toArabic and getSubsectionSymbol', () => {
  assert.equal(List.toArabic('xiv'), 14);
  assert.equal(List.toArabic('LX'), 60);
  assert.equal(List.toArabic('abc'), 0);
  assert.equal(List.getSubsectionSymbol('1.2.3.'), '3');
  assert.equal(List.getSubsectionSymbol('(a)'), 'a');
  assert.equal(List.getSubsectionSymbol('·'), '1');
});

const item = (level, symbol, text, id = 0) =>
  `<p class="MsoListParagraph" style="margin-left:${level * 0.5}in;text-indent:-.25in;mso-list:l${id} level${level} lfo1">` +
  `<!--[if !supportLists]--><span style="mso-list:Ignore">${symbol}<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp; </span></span><!--[endif]-->${text}<o:p></o:p></p>`;

test('a flat ordered list becomes an ol', () => {
  const html = item(1, '1.', 'One') + item(1, '2.', 'Two') + item(1, '3.', 'Three');
  assert.equal(normalizeHtml(cleanWord(html)), '<ol><li>one</li><li>two</li><li>three</li></ol>');
});

test('a list that does not start at one gets a start attribute', () => {
  const html = item(1, '5.', 'Five') + item(1, '6.', 'Six');
  assert.equal(normalizeHtml(cleanWord(html)), '<ol start="5"><li>five</li><li>six</li></ol>');
});

test('bullet markers give list style types', () => {
  const html = item(1, '·', 'Disc') + item(2, 'o', 'Circle') + item(3, '§', 'Square');
  assert.equal(
    normalizeHtml(cleanWord(html)),
    '<ul><li>disc<ul style="list-style-type:circle"><li>circle' +
      '<ul style="list-style-type:square"><li>square</li></ul></li></ul></li></ul>',
  );
});

test('alpha and roman markers give list style types and values', () => {
  const html = item(1, 'a.', 'A') + item(1, 'b.', 'B') + item(2, 'i.', 'I') + item(2, 'ii.', 'II') + item(1, 'd.', 'D');
  assert.equal(
    normalizeHtml(cleanWord(html)),
    '<ol style="list-style-type:lower-alpha"><li>a</li><li>b' +
      '<ol style="list-style-type:lower-roman"><li>i</li><li>ii</li></ol></li>' +
      '<li value="4">d</li></ol>',
  );
});

test('two lists with different ids are separate lists', () => {
  const html = item(1, '1.', 'One', 0) + item(1, '2.', 'Two', 0) + '<p class="MsoNormal">Between</p>' + item(1, '1.', 'Uno', 1);
  assert.equal(
    normalizeHtml(cleanWord(html)),
    '<ol><li>one</li><li>two</li></ol><p>between</p><ol><li>uno</li></ol>',
  );
});

test('real lists with mso-list items are rebuilt', () => {
  const html =
    '<ul><li class="MsoNormal" style="mso-list:l0 level1 lfo1">One<o:p></o:p></li>' +
    '<li class="MsoNormal" style="mso-list:l0 level1 lfo1">Two<o:p></o:p></li></ul>' +
    '<ol start="3" style="list-style-type:lower-alpha"><li class="MsoNormal" style="mso-list:l1 level1 lfo2">Three</li></ol>';
  assert.equal(
    normalizeHtml(cleanWord(html)),
    '<ul><li>one</li><li>two</li></ul><ol start="3" style="list-style-type:lower-alpha"><li>three</li></ol>',
  );
});

test('list wrappers carry the vertical margins of the items', () => {
  const html =
    '<p class="MsoListParagraph" style="margin-top:24pt;margin-left:.5in;mso-list:l0 level1 lfo1"><span style="mso-list:Ignore">·</span>One</p>' +
    '<p class="MsoListParagraph" style="margin-bottom:12pt;margin-left:.5in;mso-list:l0 level1 lfo1"><span style="mso-list:Ignore">·</span>Two</p>';
  assert.equal(
    normalizeHtml(cleanWord(html)),
    '<ul style="margin-bottom:16px; margin-top:32px"><li>one</li><li>two</li></ul>',
  );
});
