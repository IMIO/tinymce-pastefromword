import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    lib: {
      entry: 'src/plugin.js',
      name: 'TinyMCEPasteFromWord',
      formats: ['es', 'umd'],
      fileName: 'tinymce-pastefromword',
    },
    rollupOptions: {
      external: ['tinymce'],
      output: {
        globals: {
          tinymce: 'tinymce',
        },
      },
    },
  },
});
