import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  Border,
  convertRgbToHex,
  convertToPx,
  normalizeHex,
  parseBackground,
  parseCssText,
  parseMargin,
  parseShorthandMargins,
  splitMarginShorthand,
  writeCssText,
} from '../src/filter/css.js';

test('parseCssText parses declarations and normalizes colors', () => {
  assert.deepEqual(parseCssText('color: rgb(255, 0, 0); margin-left:36pt; font-family: "Times New Roman", serif'), {
    color: '#ff0000',
    'margin-left': '36pt',
    'font-family': '"Times New Roman", serif',
  });
  assert.deepEqual(parseCssText('COLOR:#ABC', true), { color: '#aabbcc' });
  assert.deepEqual(parseCssText('font-family: Calibri , sans-serif', true), { 'font-family': 'Calibri,sans-serif' });
  assert.deepEqual(parseCssText(''), {});
  assert.deepEqual(parseCssText(undefined), {});
  assert.deepEqual(parseCssText(';'), {});
});

test('writeCssText joins declarations', () => {
  assert.equal(writeCssText({ 'margin-left': '8px', color: 'red' }), 'margin-left:8px; color:red');
  assert.equal(writeCssText({ 'margin-left': '8px', color: 'red' }, true), 'color:red; margin-left:8px');
  assert.equal(writeCssText({}), '');
});

test('color helpers', () => {
  assert.equal(convertRgbToHex('color: rgb(1, 2, 3)'), 'color: #010203');
  assert.equal(normalizeHex('#FFF;'), '#ffffff;');
  assert.equal(normalizeHex('#AbCdEf '), '#abcdef ');
});

test('convertToPx converts absolute units to whole pixels', () => {
  assert.equal(convertToPx('36pt'), 48);
  assert.equal(convertToPx('.5in'), 48);
  assert.equal(convertToPx('1.0in'), 96);
  assert.equal(convertToPx('7.0pt'), 9);
  assert.equal(convertToPx('2.54cm'), 96);
  assert.equal(convertToPx('10mm'), 38);
  assert.equal(convertToPx('12px'), 12);
  assert.equal(convertToPx('-.25in'), -24);
  assert.equal(convertToPx('0'), 0);
  assert.equal(convertToPx('0in'), 0);
  assert.equal(convertToPx('50%'), '50%');
  assert.equal(convertToPx('auto'), 0);
  assert.equal(convertToPx(undefined), 0);
});

test('parseMargin splits the shorthand', () => {
  assert.deepEqual(parseMargin('0in'), { top: '0in', right: '0in', bottom: '0in', left: '0in' });
  assert.deepEqual(parseMargin('1px 2px'), { top: '1px', right: '2px', bottom: '1px', left: '2px' });
  assert.deepEqual(parseMargin('1px 2px 3px'), { top: '1px', right: '2px', bottom: '3px', left: '2px' });
  assert.deepEqual(parseMargin('1px 2px 3px 4px'), { top: '1px', right: '2px', bottom: '3px', left: '4px' });
  assert.deepEqual(parseMargin('0in 0in 8pt auto'), { top: '0in', right: '0in', bottom: '8pt', left: 'auto' });
});

test('parseShorthandMargins and splitMarginShorthand expand margin', () => {
  const styles = { margin: '0in 0in 8pt', color: 'red' };
  parseShorthandMargins(styles);
  assert.deepEqual(styles, { color: 'red', 'margin-top': '0in', 'margin-right': '0in', 'margin-bottom': '8pt', 'margin-left': '0in' });

  const upper = { MARGIN: '1in' };
  parseShorthandMargins(upper);
  assert.deepEqual(upper, { 'margin-top': '1in', 'margin-right': '1in', 'margin-bottom': '1in', 'margin-left': '1in' });

  const split = { margin: '0 auto' };
  splitMarginShorthand(split);
  assert.deepEqual(split, { 'margin-top': '0', 'margin-right': 'auto', 'margin-bottom': '0', 'margin-left': 'auto' });
});

test('parseBackground finds the color', () => {
  assert.deepEqual(parseBackground('yellow'), { color: 'yellow' });
  assert.deepEqual(parseBackground('url(x.png) #FF0000 no-repeat'), { color: '#FF0000', unprocessed: 'url(x.png)  no-repeat' });
  assert.deepEqual(parseBackground('rgba(0, 0, 0, 0.5)'), { color: 'rgba(0, 0, 0, 0.5)' });
});

test('Border parses shorthand values', () => {
  const border = Border.fromCssRule('1.0pt solid windowtext');
  assert.equal(border.width, '1.0pt');
  assert.equal(border.style, 'solid');
  assert.equal(border.color, 'black');
  assert.equal(border.toString(), '1.0pt solid black');

  assert.equal(Border.fromCssRule('none').toString(), 'none');
  assert.equal(Border.fromCssRule('solid #ff0000 2px').toString(), '2px solid #ff0000');
});

test('Border.splitCssValues builds one border per side', () => {
  const borders = Border.splitCssValues(
    {
      'border-width': '1px 2px',
      'border-style': 'solid',
      'border-color': 'red',
      'border-top-color': 'blue',
      'border-left': '3px dashed green',
    },
    {},
  );

  assert.equal(borders['border-top'].toString(), '1px solid blue');
  assert.equal(borders['border-right'].toString(), '2px solid red');
  assert.equal(borders['border-bottom'].toString(), '1px solid red');
  // The border-left shorthand is resolved by the caller. The split only uses the long hand properties.
  assert.equal(borders['border-left'].toString(), '2px solid red');

  const fallback = Border.splitCssValues({}, Border.fromCssRule('1px solid black'));
  assert.equal(fallback['border-top'].toString(), '1px solid black');
  assert.equal(fallback['border-left'].toString(), '1px solid black');
});
