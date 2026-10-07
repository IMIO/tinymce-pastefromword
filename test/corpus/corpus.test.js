// Runs the filter on every Word fixture. The expected outputs of CKEditor depend
// on its content filter, so this only checks invariants of the output.
import '../setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { cleanWord, isWordContent } from '../../src/filter/index.js';

const dir = process.env.PFW_CORPUS_DIR || join(dirname(fileURLToPath(import.meta.url)), '..', 'fixtures', 'word');

const inputs = readdirSync(dir).filter((name) => name.endsWith('.html') && !name.includes('expected'));

for (const file of inputs) {
  test(`corpus: ${file}`, () => {
    const html = readFileSync(join(dir, file), 'utf8');
    const out = cleanWord(html);

    assert.equal(typeof out, 'string');
    assert.doesNotMatch(out, /mso-/i, 'no mso- styles');
    assert.doesNotMatch(out, /\scke-/, 'no cke- helper attributes');
    assert.doesNotMatch(out, /<(o|v|w|m|st1):/i, 'no Office namespaced elements');
    assert.doesNotMatch(out, /<(style|xml|meta|link|script)[\s>]/i, 'no style, xml, meta, link or script elements');
    assert.doesNotMatch(out, /<!--/, 'no comments');
    assert.doesNotMatch(out, /class="[^"]*\bMso/i, 'no Mso classes');
    assert.doesNotMatch(out, /style=""/, 'no empty style attributes');

    if (/mso-list:\s*l\d+\s+level\d+\s+lfo\d+/.test(html) && !/Edge|edge/.test(file)) {
      assert.match(out, /<(ul|ol)[\s>]/, 'lists are rebuilt');
    }

    if (!file.includes('regular_paste')) {
      assert.equal(isWordContent(html), true, 'detected as Word content');
    }
  });
}
