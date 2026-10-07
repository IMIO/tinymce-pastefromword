import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { cleanWord, embedRtfImages, getContentGeneratorName, isWordContent } from '../src/filter/index.js';
import { Element } from '../src/filter/tree.js';
import { normalizeHtml } from './setup.js';

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'word');
const fixture = (name) => readFileSync(join(fixturesDir, name), 'utf8');
const clean = (name, config) => normalizeHtml(cleanWord(fixture(name), config));

test('isWordContent detects Word, Excel and plain HTML', () => {
  assert.equal(isWordContent('<p class="MsoNormal">x</p>'), true);
  assert.equal(isWordContent('<p style="mso-list:l0 level1 lfo1">x</p>'), true);
  assert.equal(isWordContent('<html><head><meta name="Generator" content="Microsoft Word 15"></head><body>x</body></html>'), true);
  assert.equal(isWordContent('<html><head><meta name=Generator content="LibreOffice 7.1"></head><body><p class="MsoNormal">x</p></body></html>'), false);
  assert.equal(isWordContent('<font face="Arial">x</font>'), true);
  assert.equal(isWordContent('<p>plain</p>'), false);
  assert.equal(isWordContent(''), false);
  assert.equal(getContentGeneratorName('<meta name="generator" content="Google Docs">'), 'unknown');
  assert.equal(getContentGeneratorName('<p>x</p>'), undefined);
});

test('paragraphs and headings', () => {
  assert.equal(
    clean('Paragraphs_with_headers__word2013__chrome.html'),
    '<h1>transforming turn-key content</h1>' +
      '<p>our initiative standardizes a wide-ranging partnership, while pursuing this route will enable us to establish the wide-ranging supply-chains.</p>' +
      '<h2>strategic roi</h2>' +
      '<p>controlling should significantly jump-start fine-grained next steps. the powerful champion conservatively broadens bottom lines. as a result, the partners strengthen our decentralized, value-driven, on-message and top-down forward plannings. our one-to-one and intra-organisational controls incentivise the account executive, whilst the cornerstones inspire the community. supply-chain, issue and market practice diligently strengthen the chief strategic planning officer in this space.</p>' +
      '<h2>enterprise schemas</h2>' +
      '<p>a core meeting standardizes large-scale, functional, constructive and in-depth plannings up, down and across the organization. in the same time, our workflows generate rewards. the project manager globally gets from here to here. scale-as-you-grow, carefully thought-out, workflows boost a vision-setting executive talent, whilst the resources take a bite out of our landscapes. the enabler globally establishes multi-source synergies. as a result, the partners foster a cutting-edge decision. the customers optimize a recalibration. the chief legal officer strategically pre-prepares a collaborative wow factor; this is why the chief management office officer learns motivational dashboards.</p>',
  );
});

test('basic formatting', () => {
  assert.equal(clean('Bold__word2013__chrome.html'), '<p><b>this text is only bolded.</b></p>');
  assert.equal(clean('Italic__word2013__chrome.html'), '<p><i>this text is only italicized.</i></p>');
  assert.equal(clean('Underline__word2013__chrome.html'), '<p><u>this text is underlined.</u></p>');
  assert.equal(clean('Link__word2013__chrome.html'), '<p>there is a <a href="http://www.google.com">link here</a>.</p>');
  assert.equal(
    clean('Spacing__word2013__chrome.html'),
    '<p><span style="line-height:300%">there’s some...</span></p><p><span style="line-height:300%">...spacing here.</span></p>',
  );
});

test('fonts and colors are kept as style stacks', () => {
  assert.equal(
    clean('Fonts__word2013__chrome.html'),
    '<p><span style="line-height:normal"><span style="font-family:batang,serif"><span style="color:black">batang!</span></span></span></p>' +
      '<p><span style="line-height:normal"><span style="font-family:comic sans ms"><span style="color:black">comic sans is the devil.</span></span></span></p>' +
      '<p><span style="line-height:normal"><span style="font-family:arial black,sans-serif"><span style="color:black">arial is the new black.</span></span></span></p>',
  );

  const colors = clean('Colors__word2013__chrome.html');
  assert.match(colors, /^<p><span style="line-height:normal"><span style="font-size:48.0pt"><span style="color:red">r<\/span><\/span>/);
  assert.match(colors, /<span style="color:#7030a0">w<\/span>/);
});

test('removeFontStyles drops font family, size, color and line height', () => {
  assert.equal(
    clean('Fonts__word2013__chrome.html', { removeFontStyles: true }),
    '<p>batang!</p><p>comic sans is the devil.</p><p>arial is the new black.</p>',
  );
});

test('text alignment', () => {
  const out = clean('Text_alignment__word2013__chrome.html');
  assert.match(out, /^<p>left<\/p><p align="right" style="text-align:right">right<\/p><p align="center" style="text-align:center">center<\/p><p style="text-align:justify">justified/);
});

test('flat lists', () => {
  assert.equal(
    clean('Ordered_list__word2013__chrome.html'),
    '<ol><li>this</li><li>is</li><li>a</li><li>one</li><li>level</li><li>ordered</li><li>list</li></ol>',
  );
  assert.equal(
    clean('Unordered_list__word2013__chrome.html'),
    '<ul><li>this</li><li>is</li><li>a</li><li>one</li>' +
      '<li>level</li><li>unordered</li><li>list</li></ul>',
  );
});

test('nested lists', () => {
  assert.equal(
    clean('Ordered_list_multiple__word2013__chrome.html'),
    '<ol><li>this</li><li>is</li><li>a<ol style="list-style-type:lower-alpha"><li>multi' +
      '<ol style="list-style-type:lower-roman"><li>level</li></ol></li><li>ordered</li></ol></li><li>list</li></ol>',
  );
  assert.equal(
    clean('Unordered_list_multiple__word2013__chrome.html'),
    '<ul><li>this</li><li>is</li><li>a<ul style="list-style-type:circle"><li>multi' +
      '<ul style="list-style-type:square"><li>level</li></ul></li><li>unordered</li></ul></li><li>list</li></ul>',
  );
});

test('list start values and roman numerals', () => {
  assert.equal(
    clean('Tickets_7584Numbered_list_with_diff_start_value__word2013__chrome.html'),
    '<ol start="5"><li>ertretretre</li><li>tretrete</li></ol><p></p><p>list with roman numerals and start value is 5.</p><p></p>' +
      '<ol start="4" style="list-style-type:upper-roman"><li>first list item.</li><li>second list item.</li></ol><p></p>',
  );
});

test('list items with explicit values', () => {
  assert.equal(
    clean('Multi_dig_list__word2013__chrome.html'),
    '<p><b>multi level list</b></p><ol><li>one<ol><li>two<ol><li>three</li>' +
      '<li value="1">fourth line with extra text where it moves to next line and let’s see how the list gets displayed properly</li>' +
      '<li value="1">five</li></ol></li></ol></li><li>six</li></ol>',
  );
});

test('a list that skips a level gets an item without marker', () => {
  const out = clean('Ordered_list_multiple_edgy__word2013__chrome.html');
  assert.match(out, /<ol><li>one<ol style="list-style-type:lower-alpha"><li style="list-style-type:none"><ol style="list-style-type:lower-roman"><li>level<\/li><li>here<\/li><\/ol><\/li><li>is<\/li><\/ol><\/li><li>skipped<\/li><\/ol>/);
});

test('vertical margins of lists go to the list wrapper', () => {
  assert.equal(
    clean('List_margins__word2016__chrome.html'),
    '<p>before</p><ul style="margin-bottom:80px; margin-top:32px"><li>list item 1</li><li>list item 2</li><li>list item 3</li></ul><p>after</p>',
  );
});

test('tables keep borders, widths and alignment', () => {
  const out = clean('Simple_table__word2013__chrome.html');
  assert.match(out, /^<table style="border-collapse:collapse"><tbody><tr><td style="border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:1px solid black; padding:0in 7px 0in 7px; width:208px" valign="top"><p>x\\y<\/p><\/td>/);
  assert.match(out, /<td style="border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; padding:0in 7px 0in 7px; width:208px" valign="top"><p>2,2<\/p><\/td><\/tr><\/tbody><\/table>/);
  assert.doesNotMatch(out, /class=/);

  const aligned = clean('Table_alignment__word2013__chrome.html');
  assert.match(aligned, /<p>center<\/p><table align="center" style="border-collapse:collapse; border:none">/);
});

test('style sheets are inlined', () => {
  const out = clean('InlineStyles__word2013__chrome.html');
  assert.match(out, /^<p><span style="font-size:10pt"><span style="font-family:verdana,sans-serif"><b><span lang="en-gb" style="font-size:24.0pt"><span style="color:#4472c4">deploy career insite 2.1 salary system test<\/span><\/span><\/b><\/span><\/span><\/p>/);
  assert.match(out, /<h1 style="margin-top:16px"><span style="font-size:16pt"><span style="font-family:verdana,sans-serif"><span style="font-weight:normal">i. test<\/span><\/span><\/span><\/h1>/);
});

test('online images keep their URL', () => {
  assert.equal(
    clean('Image__word2013__chrome.html'),
    '<p><img alt="http://placekitten.com/200/305" src="http://placekitten.com/200/305" style="height:305px; width:200px"></p>',
  );
});

test('Office namespaced elements are unwrapped and objects removed', () => {
  const out = clean('Object__word2013__datatransfer.html');
  assert.match(out, /^<p><span style="font-size:12pt"><span style="font-family:times new roman,serif"><span lang="en-ie" style="font-family:arial,sans-serif">text with arial font<\/span><\/span><\/span><\/p><p><\/p>/);
  assert.doesNotMatch(out, /<o:|<object/);
});

test('page breaks use the factory of the editor', () => {
  const createPageBreak = () => new Element('img', { class: 'mce-pagebreak', src: 'about:blank' });
  const out = clean('Page_break_simple__office365__chrome.html', { createPageBreak });
  assert.match(out, /first page<\/span><\/span><\/span><\/p><span style="font-size:11.0pt"><span style="line-height:115%"><span style="font-family:calibri,sans-serif"><img class="mce-pagebreak" src="about:blank"><\/span><\/span><\/span><p/);

  const without = clean('Page_break_simple__office365__chrome.html');
  assert.match(without, /<br clear="all" style="page-break-before:always">/);
});

test('images come from the RTF clipboard', () => {
  const html = fixture('PFW_image_Simple_offline_image__chrome.html');
  const rtf = fixture('PFW_image_Simple_offline_image__chrome.rtf');
  const out = embedRtfImages(cleanWord(html), rtf);
  const images = out.match(/<img[^>]*>/g);

  assert.equal(images.length, 2);
  for (const image of images) {
    assert.match(image, /src="data:image\/png;base64,iVBORw0KGgo/);
    assert.match(image, /style="width:64px; height:64px"/);
  }
  assert.doesNotMatch(out, /file:\/\//);
});

test('duplicated images in the RTF are all embedded', () => {
  const html = fixture('ImagesExtraction_DuplicatedImage__chrome.html');
  const rtf = fixture('ImagesExtraction_DuplicatedImage__chrome.rtf');
  const out = embedRtfImages(cleanWord(html), rtf);

  assert.equal((out.match(/src="data:image\/jpeg;base64,/g) || []).length, 2);
  assert.doesNotMatch(out, /file:\/\//);
});

test('Excel content is detected through the generator meta tag', () => {
  const html = fixture('Meta_generator_Excel__excel2016__datatransfer.html');
  assert.equal(isWordContent(html), true);
  const out = clean('Meta_generator_Excel__excel2016__datatransfer.html');
  assert.match(out, /<table/);
  assert.doesNotMatch(out, /class="xl\d+"/);
});
