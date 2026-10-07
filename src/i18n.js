import tinymce from 'tinymce';

const TRANSLATIONS = {
  fr_FR: {
    'Paste from Word': 'Coller depuis Word',
    'The text you want to paste seems to be copied from Word. Do you want to clean it before pasting?':
      'Le texte à coller semble provenir de Word. Voulez-vous le nettoyer avant de le coller ?',
  },
  nl: {
    'Paste from Word': 'Plakken vanuit Word',
    'The text you want to paste seems to be copied from Word. Do you want to clean it before pasting?':
      'De tekst die u wilt plakken lijkt uit Word te komen. Wilt u deze opschonen voor het plakken?',
  },
  de: {
    'Paste from Word': 'Aus Word einfügen',
    'The text you want to paste seems to be copied from Word. Do you want to clean it before pasting?':
      'Der Text, den Sie einfügen möchten, scheint aus Word zu stammen. Möchten Sie ihn vor dem Einfügen bereinigen?',
  },
};

let registered = false;

/** Add the plugin strings to the TinyMCE language packs. */
export function registerTranslations() {
  if (registered) {
    return;
  }
  registered = true;

  for (const code of Object.keys(TRANSLATIONS)) {
    tinymce.util.I18n.add(code, TRANSLATIONS[code]);
  }
}
