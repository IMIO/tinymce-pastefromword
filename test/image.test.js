import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import {
  bytesToBase64,
  createSrcWithBase64,
  embedRtfImages,
  extractFromRtf,
  extractTagsFromHtml,
  getImageType,
  getImageTypeFromSignature,
  hexToBytes,
} from '../src/filter/image.js';

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rtf');

const RTF = [
  '{\\pict\\picscalex100\\picscaley100\\piccropl0\\piccropr0\\piccropt0\\piccropb0\\picw300\\pich300\\picwgoal3823\\pichgoal3823\\pngblip d76df8e7aefc}',
  '{\\*\\shppict{\\pict{\\*\\picprop{\\sp{\\sn wzDescription}{\\sv }}{\\sp{\\sn wzName}{\\sv }}}\\picscalex19\\picscaley19\\piccropl0\\piccropr0\\piccropt0' +
    '\\piccropb0\\picw300\\pich300\\picwgoal3823\\pichgoal3823\\pngblip d76df8e7aefc }}{\\nonshppict{\\pict{\\*\\picprop{',
  '{\\pict\\picscalex100\\picscaley100\\piccropl0\\piccropr0\\piccropt0\\piccropb0\\picw300\\pich300\\picwgoal3823\\pichgoal3823\\emfblip d76df8e7aefc}',
];

test('hex and base64 helpers', () => {
  assert.deepEqual(Array.from(hexToBytes('08a11d8ada2b')), [8, 161, 29, 138, 218, 43]);
  assert.equal(bytesToBase64(hexToBytes('08a11d8ada2b')), 'CKEditor');
  assert.equal(bytesToBase64(new Uint8Array([77])), 'TQ==');
  assert.equal(bytesToBase64(new Uint8Array([77, 97])), 'TWE=');
  assert.equal(bytesToBase64(new Uint8Array([77, 97, 110])), 'TWFu');
  assert.equal(bytesToBase64(hexToBytes('d76df8e7aefc')), '12345678');
});

test('image type detection', () => {
  assert.equal(getImageType('\\pngblip'), 'image/png');
  assert.equal(getImageType('\\jpegblip'), 'image/jpeg');
  assert.equal(getImageType('\\emfblip'), 'image/emf');
  assert.equal(getImageType('\\wmetafile8'), 'image/wmf');
  assert.equal(getImageType('\\nothing'), 'unknown');
  assert.equal(getImageTypeFromSignature(hexToBytes('89504e470d0a')), 'image/png');
  assert.equal(getImageTypeFromSignature(hexToBytes('ffd8ffe0')), 'image/jpeg');
  assert.equal(getImageTypeFromSignature(hexToBytes('47494638')), 'image/gif');
  assert.equal(getImageTypeFromSignature(hexToBytes('00000000')), null);
});

test('createSrcWithBase64 only supports png, jpeg and gif', () => {
  assert.equal(createSrcWithBase64({ type: 'image/png', hex: 'd76df8e7aefc' }), 'data:image/png;base64,12345678');
  assert.equal(createSrcWithBase64({ type: 'image/emf', hex: 'd76df8e7aefc' }), null);
  assert.equal(createSrcWithBase64({ type: 'image/png', hex: null }), null);
});

test('extractTagsFromHtml reads image sources', () => {
  assert.deepEqual(extractTagsFromHtml('<p><img width="1" src="file:///C:/a.png" alt=""><img src="http://x/y.png" /></p>'), [
    'file:///C:/a.png',
    'http://x/y.png',
  ]);
});

test('embedRtfImages replaces file images with the RTF images', () => {
  assert.equal(embedRtfImages('<img src="file://foo" />', RTF[0]), '<img src="data:image/png;base64,12345678" />');
  assert.equal(embedRtfImages('<img src="file://foo" />', RTF[1]), '<img src="data:image/png;base64,12345678" />');
});

test('embedRtfImages leaves online images alone', () => {
  const html = '<img src="http://example.com/img.png" />';
  assert.equal(embedRtfImages(html, RTF[0]), html);
});

test('embedRtfImages reports unsupported formats', () => {
  const errors = [];
  const html = '<img src="file://foo" />';
  assert.equal(embedRtfImages(html, RTF[2], { onError: (code, data) => errors.push([code, data]) }), html);
  assert.deepEqual(errors, [['pastefromword-unsupported-image', { type: 'image/emf', index: 0 }]]);
});

test('embedRtfImages reports a count mismatch', () => {
  const errors = [];
  const html = '<img src="file://foo" />';
  assert.equal(embedRtfImages(html, RTF[0] + RTF[0], { onError: (code, data) => errors.push([code, data]) }), html);
  assert.deepEqual(errors, [['pastefromword-failed-image-extraction', { rtf: 2, html: 1 }]]);
});

test('embedRtfImages uses the custom src factory', () => {
  const out = embedRtfImages('<img src="file://foo" />', RTF[0], { createSrc: (img) => 'blob:' + img.type });
  assert.equal(out, '<img src="blob:image/png" />');
});

test('embedRtfImages without RTF or images is a no-op', () => {
  assert.equal(embedRtfImages('<p>x</p>', ''), '<p>x</p>');
  assert.equal(embedRtfImages('<p>x</p>', RTF[0]), '<p>x</p>');
});

// The CKEditor fixtures pair an RTF clipboard with the images it must yield.
for (const file of readdirSync(fixturesDir).filter((name) => name.endsWith('.rtf'))) {
  test(`extractFromRtf matches the CKEditor fixture ${file}`, () => {
    const rtf = readFileSync(join(fixturesDir, file), 'utf8');
    const expected = JSON.parse(readFileSync(join(fixturesDir, file.replace(/\.rtf$/, '.json')), 'utf8'));

    assert.deepEqual(extractFromRtf(rtf), expected);
  });
}
