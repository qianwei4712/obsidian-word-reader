# Obsidian Word Reader

[中文文档](README.zh-CN.md)

Obsidian Word Reader is a desktop-only Obsidian plugin for opening `.docx` files directly inside Obsidian as safe, read-only documents.

The plugin is not a Word editor. It is designed to make Word documents easier to read, search, reference, and summarize inside an Obsidian vault while keeping the original file unchanged.

## Features

- Open `.docx` files in an Obsidian tab.
- Render headings, paragraphs, lists, tables, images, and page breaks where supported by `docx-preview`.
- Keep the original Word file unchanged.
- Use a white document surface so Word content remains readable in dark themes.
- Zoom continuously with the toolbar percentage input or `Ctrl` + mouse wheel.
- Fit the rendered document to the pane width.
- Click rendered images to preview them in a larger modal.
- Search rendered text in the current document.
- Copy selected rendered text, or extract and copy plain text from the whole document.
- Open the source file with the system default Word/WPS-compatible application.
- Create or open a same-name Markdown summary note linked back to the source `.docx`.

## Supported Files

| Extension | Status | Notes |
| --- | --- | --- |
| `.docx` | Supported | Rendered inside Obsidian through `docx-preview`. |
| `.doc` | Not rendered directly | Convert to `.docx`, or open with Word, WPS, or another external application. |

## Usage

1. Put a `.docx` file into your Obsidian vault.
2. Click the `.docx` file in the file explorer.
3. Read the document in the Obsidian tab opened by the plugin.
4. Use the toolbar to reload, zoom, fit width, search, copy text, create a summary note, or open the file externally.

### Zoom

- Type a percentage in the zoom input.
- Use `Ctrl` + mouse wheel over the document preview for continuous zoom.
- Use fit width when you want the document to match the current pane.

### Image Preview

Click an image rendered from the Word document to open a larger preview modal. Press `Esc` or close the modal to return to the document.

### Summary Notes

The summary note action creates a same-name Markdown file next to the Word document.

Example:

```text
Report.docx
Report.md
```

The generated note includes frontmatter and starter sections:

```markdown
---
source: "Report.docx"
type: word-note
created: 2026-05-28
---

# Report

Source: [[Report.docx]]

## Summary

## Key Points

## To Do

## Quotes
```

If the same-name Markdown file already exists, the plugin opens it without overwriting content.

## Installation

For local development or manual installation:

1. Install dependencies:

   ```bash
   npm install
   ```

2. Build the plugin:

   ```bash
   npm run build
   ```

3. Copy the generated files from `dist` into your Obsidian vault plugin folder:

   ```text
   dist/main.js
   dist/manifest.json
   dist/styles.css
   ```

   Example target folder:

   ```text
   YourVault/.obsidian/plugins/obsidian-word-reader/
   ```

4. Enable third-party plugins in Obsidian, then enable Obsidian Word Reader.

## Development

Install dependencies:

```bash
npm install
```

Run a production build:

```bash
npm run build
```

Run TypeScript checks:

```bash
npm run typecheck
```

On this workstation, the expected Node.js executable is:

```text
E:\DevelopHelper\nvm\v18.20.8\node.exe
```

## Known Limits

- This is not a Word editor.
- The plugin never saves changes back to `.docx`.
- Legacy `.doc` files are not rendered directly.
- Complex Word layouts may not render exactly like Microsoft Word.
- Very large files or files with many images may render slowly.
- Mobile support is not included in this version.

## Security

The plugin reads `.docx` files through Obsidian's binary vault API for rendering and text extraction only.

It does not modify, overwrite, or write back to the original Word document. Use Word, WPS, or another external editor when the source document needs to be changed.

## Recommended Workflow

Keep original Word files in the vault, open them for reading, and create Markdown summary notes for long-term knowledge management.

This keeps the original document format intact while bringing summaries, decisions, tasks, and quotes into your Obsidian workflow.
