// Exercises the TinyMCE plugin wiring with a fake editor. TinyMCE itself loads in jsdom.
import { test, before } from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
import { normalizeHtml } from './setup.js';

const dom = new JSDOM('<!doctype html><html><body></body></html>', { url: 'http://localhost/' });

// TinyMCE reads these globals when it loads.
dom.window.matchMedia = () => ({
  matches: false,
  addListener() {},
  removeListener() {},
  addEventListener() {},
  removeEventListener() {},
});
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.Node = dom.window.Node;
for (const key of ['navigator', 'HTMLElement', 'Element']) {
  if (!(key in globalThis)) {
    globalThis[key] = dom.window[key];
  }
}

let tinymce;
let pluginFactory;

before(async () => {
  tinymce = (await import('tinymce')).default;
  await import('../src/plugin.js');
  pluginFactory = tinymce.PluginManager.get('pastefromword');
});

const WORD_HTML =
  '<html><head><meta name="Generator" content="Microsoft Word 15"><style>p.MsoNormal { margin: 0in; font-size: 11.0pt; }</style></head>' +
  '<body><p class="MsoNormal" style="mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span style="mso-list:Ignore">1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;</span></span><!--[endif]-->One<o:p></o:p></p>' +
  '<p class="MsoNormal" style="mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span style="mso-list:Ignore">2.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;</span></span><!--[endif]-->Two<o:p></o:p></p></body></html>';

function createFakeEditor(options = {}) {
  const registered = {};
  const handlers = {};
  const commands = {};
  const queryStates = {};
  const ui = { buttons: {}, menuItems: {} };

  const editor = {
    options: {
      register: (name, spec) => {
        registered[name] = spec;
      },
      isRegistered: (name) => name in registered || name in options,
      get: (name) => (name in options ? options[name] : registered[name] ? registered[name].default : undefined),
    },
    on: (name, handler, prepend) => {
      handlers[name] = handlers[name] || [];
      if (prepend) handlers[name].unshift(handler);
      else handlers[name].push(handler);
    },
    off: (name, handler) => {
      handlers[name] = (handlers[name] || []).filter((h) => h !== handler);
    },
    dispatch: (name, args = {}) => {
      let prevented = false;
      const event = { ...args, type: name, preventDefault: () => (prevented = true), isDefaultPrevented: () => prevented };
      for (const handler of handlers[name] || []) handler(event);
      return event;
    },
    addCommand: (name, fn) => {
      commands[name] = fn;
    },
    execCommand: (name) => commands[name](),
    addQueryStateHandler: (name, fn) => {
      queryStates[name] = fn;
    },
    queryCommandState: (name) => queryStates[name](),
    ui: {
      registry: {
        addToggleButton: (name, spec) => (ui.buttons[name] = spec),
        addToggleMenuItem: (name, spec) => (ui.menuItems[name] = spec),
      },
    },
    plugins: {},
    schema: { isValid: () => true },
    translate: (text) => text,
    handlers,
    ui_: ui,
  };

  return editor;
}

const fakeClipboard = (data) => ({ getData: (type) => data[type] || '' });

test('the plugin registers its options, button, menu item and command', () => {
  const editor = createFakeEditor();
  const api = pluginFactory(editor, '');

  assert.equal(editor.options.get('pastefromword_inline_images'), true);
  assert.equal(editor.options.get('pastefromword_remove_font_styles'), false);
  assert.ok(editor.ui_.buttons.pastefromword);
  assert.ok(editor.ui_.menuItems.pastefromword);
  assert.equal(editor.queryCommandState('mceTogglePasteFromWord'), false);
  editor.execCommand('mceTogglePasteFromWord');
  assert.equal(editor.queryCommandState('mceTogglePasteFromWord'), true);
  assert.equal(typeof api.cleanWord, 'function');
  assert.equal(api.isWordContent('<p class="MsoNormal">x</p>'), true);
});

test('Word clipboard content is cleaned in PastePreProcess', () => {
  const editor = createFakeEditor();
  pluginFactory(editor, '');
  editor.dispatch('PreInit');

  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': WORD_HTML }), isDefaultPrevented: () => false });

  const trimmed = WORD_HTML.replace(/^[\s\S]*<body[^>]*>\s*|\s*<\/body[^>]*>[\s\S]*$/gi, '');
  const event = editor.dispatch('PastePreProcess', { content: trimmed, internal: false });

  assert.equal(
    normalizeHtml(event.content),
    '<ol><li><span style="font-size:11pt">one</span></li><li><span style="font-size:11pt">two</span></li></ol>',
  );
});

test('plain HTML is left alone', () => {
  const editor = createFakeEditor();
  pluginFactory(editor, '');
  editor.dispatch('PreInit');

  const html = '<p>Hello</p>';
  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': html }), isDefaultPrevented: () => false });
  const event = editor.dispatch('PastePreProcess', { content: html, internal: false });

  assert.equal(event.content, html);
});

test('plain text paste of Word content is not replaced', () => {
  const editor = createFakeEditor();
  pluginFactory(editor, '');
  editor.dispatch('PreInit');

  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': WORD_HTML }), isDefaultPrevented: () => false });
  const event = editor.dispatch('PastePreProcess', { content: '<p>1. One</p><p>2. Two</p>', internal: false });

  assert.equal(event.content, '<p>1. One</p><p>2. Two</p>');
});

test('the PasteFromWord event can cancel the cleanup and AfterPasteFromWord can change the result', () => {
  const editor = createFakeEditor();
  pluginFactory(editor, '');
  editor.dispatch('PreInit');

  editor.on('PasteFromWord', (e) => e.preventDefault());
  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': WORD_HTML }), isDefaultPrevented: () => false });
  const raw = WORD_HTML.replace(/^[\s\S]*<body[^>]*>\s*|\s*<\/body[^>]*>[\s\S]*$/gi, '');
  assert.equal(editor.dispatch('PastePreProcess', { content: raw, internal: false }).content, raw);

  editor.handlers.PasteFromWord = [];
  editor.on('AfterPasteFromWord', (e) => {
    e.content = e.content.replace(/<ol>/, '<ol class="word">');
  });
  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': WORD_HTML }), isDefaultPrevented: () => false });
  assert.match(editor.dispatch('PastePreProcess', { content: raw, internal: false }).content, /^<ol class="word">/);
});

test('forced cleanup applies to any HTML', () => {
  const editor = createFakeEditor();
  pluginFactory(editor, '');
  editor.dispatch('PreInit');
  editor.execCommand('mceTogglePasteFromWord');

  const html = '<p style="margin:0in; mso-pagination:widow-orphan">Hello<o:p></o:p></p>';
  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': html }), isDefaultPrevented: () => false });
  const event = editor.dispatch('PastePreProcess', { content: html, internal: false });

  assert.equal(event.content, '<p>Hello</p>');
});

test('images from the RTF clipboard are embedded as data URLs without a blob cache', () => {
  const editor = createFakeEditor();
  pluginFactory(editor, '');
  editor.dispatch('PreInit');

  const html = '<html><body><p class="MsoNormal"><img width="1" height="1" src="file:///C:/Users/x/clip_image001.png" v:shapes="Picture_1"></p></body></html>';
  const rtf = '{\\rtf1{\\pict\\picw300\\pich300\\pngblip d76df8e7aefc}}';
  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': html, 'text/rtf': rtf }), isDefaultPrevented: () => false });
  const trimmed = html.replace(/^[\s\S]*<body[^>]*>\s*|\s*<\/body[^>]*>[\s\S]*$/gi, '');
  const event = editor.dispatch('PastePreProcess', { content: trimmed, internal: false });

  assert.equal(event.content, '<p><img src="data:image/png;base64,12345678" style="width:1px; height:1px"></p>');
});

test('images go to the blob cache of the editor when it exists', () => {
  const editor = createFakeEditor();
  const cache = [];
  editor.editorUpload = {
    blobCache: {
      getByData: () => undefined,
      create: (id, blob, base64) => ({ id: () => id, blob: () => blob, base64: () => base64, blobUri: () => 'blob:http://localhost/' + id }),
      add: (info) => cache.push(info),
    },
  };
  pluginFactory(editor, '');
  editor.dispatch('PreInit');

  const html = '<html><body><p class="MsoNormal"><img src="file:///C:/x.png"></p></body></html>';
  const rtf = '{\\rtf1{\\pict\\picw300\\pich300\\pngblip d76df8e7aefc}}';
  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': html, 'text/rtf': rtf }), isDefaultPrevented: () => false });
  const trimmed = html.replace(/^[\s\S]*<body[^>]*>\s*|\s*<\/body[^>]*>[\s\S]*$/gi, '');
  const event = editor.dispatch('PastePreProcess', { content: trimmed, internal: false });

  assert.equal(cache.length, 1);
  assert.equal(cache[0].base64(), '12345678');
  assert.equal(cache[0].blob().type, 'image/png');
  assert.match(event.content, /^<p><img src="blob:http:\/\/localhost\/pastefromword/);
});

test('page breaks become the placeholder of the pagebreak plugin', () => {
  const editor = createFakeEditor();
  editor.plugins.pagebreak = {};
  pluginFactory(editor, '');
  editor.dispatch('PreInit');

  const html = '<p class="MsoNormal">A</p><br clear="all" style="page-break-before:always"><p class="MsoNormal">B</p>';
  editor.dispatch('paste', { clipboardData: fakeClipboard({ 'text/html': html }), isDefaultPrevented: () => false });
  const event = editor.dispatch('PastePreProcess', { content: html, internal: false });

  assert.equal(
    event.content,
    '<p>A</p><img src="' + tinymce.Env.transparentSrc + '" class="mce-pagebreak" data-mce-resize="false" data-mce-placeholder=""><p>B</p>',
  );
});
