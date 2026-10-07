import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { filterInlinerStyles, inlineStyles, isKnownCssProperty, parseStyleSheet, sortRules } from '../src/filter/inliner.js';

test('parseStyleSheet reads flat rules and skips at-rules', () => {
  const rules = parseStyleSheet(`
    <!--
    @font-face { font-family: Calibri; }
    @page WordSection1 { size: 8.5in 11.0in; margin: 1.0in; }
    p.MsoNormal, li.MsoNormal { margin: 0in; font-size: 11.0pt; font-family: "Calibri", sans-serif; mso-pagination: widow-orphan; }
    @list l0:level1 { mso-level-number-format: bullet; }
    span.Hyperlink { color: blue; text-decoration: underline; text-underline: single; page-break-before: always; }
    -->
  `);

  assert.deepEqual(rules, [
    {
      selector: 'p.MsoNormal, li.MsoNormal',
      styles: { margin: '0in', 'font-size': '11pt', 'font-family': '"Calibri",sans-serif' },
      index: 0,
    },
    { selector: 'span.Hyperlink', styles: { color: 'blue', 'text-decoration': 'underline' }, index: 1 },
  ]);
});

test('isKnownCssProperty drops proprietary properties', () => {
  assert.equal(isKnownCssProperty('margin-left'), true);
  assert.equal(isKnownCssProperty('font-family'), true);
  assert.equal(isKnownCssProperty('line-height'), true);
  assert.equal(isKnownCssProperty('mso-list'), false);
  assert.equal(isKnownCssProperty('tab-stops'), false);
  assert.equal(isKnownCssProperty('-webkit-text-stroke'), false);
  assert.deepEqual(filterInlinerStyles({ color: 'red', 'page-break-after': 'always', 'mso-x': 'y' }), { color: 'red' });
});

test('sortRules puts class selectors first and later rules before earlier ones', () => {
  const sorted = sortRules([
    { selector: 'p', styles: {}, index: 0 },
    { selector: 'p.a', styles: {}, index: 1 },
    { selector: 'h1', styles: {}, index: 2 },
    { selector: 'li.b', styles: {}, index: 3 },
  ]);
  assert.deepEqual(
    sorted.map((rule) => rule.selector),
    ['li.b', 'p.a', 'h1', 'p'],
  );
});

test('inlineStyles applies sheet styles without overwriting inline styles', () => {
  const html = `<html><head><style>
    p.MsoNormal { margin: 0in; font-size: 11.0pt; color: green; }
    p { color: red; text-align: center; }
    v\\:* { behavior: url(#default#VML); }
  </style></head><body><p class="MsoNormal" style="color:blue">Text</p><p>Other</p></body></html>`;

  assert.equal(
    inlineStyles(html),
    '<p class="MsoNormal" style="color:blue; font-size:11pt; margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:0in; text-align:center">Text</p>' +
      '<p style="color:red; text-align:center">Other</p>',
  );
});
