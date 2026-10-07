import tinymce from 'tinymce';
import 'tinymce/icons/default/index.js';
import 'tinymce/themes/silver/index.js';
import 'tinymce/models/dom/index.js';
import 'tinymce/plugins/lists/index.js';
import 'tinymce/plugins/table/index.js';
import 'tinymce/plugins/link/index.js';
import 'tinymce/plugins/image/index.js';
import 'tinymce/plugins/pagebreak/index.js';
import 'tinymce/plugins/code/index.js';
// The skin modules register their CSS in tinymce.Resource. The default skin and content_css use them.
import 'tinymce/skins/ui/oxide/skin.js';
import 'tinymce/skins/ui/oxide/content.js';
import 'tinymce/skins/content/default/content.js';

// Importing the plugin registers it in the PluginManager.
import '../src/plugin.js';

const log = document.getElementById('log');

function logEvent(name, detail) {
  const line = `${new Date().toLocaleTimeString()} ${name} ${detail}\n`;
  log.textContent = (line + log.textContent).slice(0, 8000);
}

async function createEditor() {
  await tinymce.remove('#editor');

  await tinymce.init({
    selector: '#editor',
    license_key: 'gpl',
    height: 420,
    plugins: 'lists table link image pagebreak code pastefromword',
    toolbar: 'undo redo | pastefromword | bold italic | bullist numlist | table | pagebreak | code',
    menubar: 'edit insert format',
    menu: { edit: { title: 'Edit', items: 'undo redo | cut copy paste pastefromword | selectall' } },
    pastefromword_remove_font_styles: document.getElementById('opt-remove-font-styles').checked,
    pastefromword_keep_zero_margins: document.getElementById('opt-keep-zero-margins').checked,
    pastefromword_inline_images: document.getElementById('opt-inline-images').checked,
    pastefromword_prompt_cleanup: document.getElementById('opt-prompt').checked,
    setup(editor) {
      editor.on('PasteFromWord', (e) => logEvent('PasteFromWord', `${e.content.length} chars of HTML, ${(e.rtf || '').length} chars of RTF`));
      editor.on('AfterPasteFromWord', (e) => logEvent('AfterPasteFromWord', `${e.content.length} chars`));
      editor.on('PasteFromWordToggle', (e) => logEvent('PasteFromWordToggle', e.state ? 'on' : 'off'));
    },
  });
}

document.getElementById('reinit').addEventListener('click', createEditor);

document.getElementById('clean').addEventListener('click', () => {
  const editor = tinymce.get('editor');
  const raw = document.getElementById('raw').value;
  document.getElementById('cleaned').textContent = editor ? editor.plugins.pastefromword.cleanWord(raw) : 'Editor not ready';
});

createEditor();
