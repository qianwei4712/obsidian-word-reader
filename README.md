# Obsidian Word Reader

[中文文档](README.zh-CN.md)

Obsidian Word Reader is a desktop-only Obsidian plugin for opening `.docx` files directly inside Obsidian as safe, read-only documents.

## Features

- Open `.docx` files in an Obsidian tab.
- Render Word content with headings, paragraphs, lists, tables, images, and page breaks where supported by `docx-preview`.
- Keep the original Word file unchanged.
- Use a white document surface so Word content stays readable in dark themes.
- Zoom between 80%, 100%, 125%, and 150%.
- Fit the rendered document to the pane width.
- Search rendered text in the current document.
- Copy selected rendered text, or extract and copy plain text from the whole document.
- Open the file with the system default Word/WPS-compatible application.
- Create or open a same-name Markdown note linked back to the source `.docx`.

## Install for Development

1. Install dependencies:

   ```bash
   npm install
   ```

2. Build:

   ```bash
   npm run build
   ```

3. Copy the generated `dist` files into an Obsidian vault plugin folder:

   ```text
   dist/main.js
   dist/manifest.json
   dist/styles.css
   ```

4. Enable the plugin in Obsidian.

If you use the local Node installation from this workstation, run npm through:

```text
E:\DevelopHelper\nvm\v18.20.8\node.exe
```

## Known Limits

- This is not a Word editor.
- The plugin never saves changes back to `.docx`.
- Complex Word layouts may not render exactly like Microsoft Word.
- Very large files or files with many images may render slowly.
- Mobile support is not included in this version.

## Recommended Workflow

Keep original Word files in the vault, open them for reading, and create a Markdown summary note for long-term knowledge management. Use Word, WPS, or another external editor when the source document needs to be changed.
