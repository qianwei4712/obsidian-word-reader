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
- Search rendered text in the current document with previous/next navigation.
- Copy selected rendered text, whole-document plain text, selected Markdown, or whole-document Markdown.
- Use the clickable outline panel to jump between rendered headings.
- Switch the plugin interface between Chinese and English.
- Open the source file with the system default Word/WPS-compatible application.
- Show a clear fallback page for legacy `.doc` files with external-open and conversion guidance.
- Show clearer messages for encrypted, damaged, or unsupported Word documents.
- Avoid stale render results when switching or reloading documents.
- Create or open a same-name Markdown summary note linked back to the source `.docx`.

## Supported Files

| Extension | Status | Notes |
| --- | --- | --- |
| `.docx` | Supported | Rendered inside Obsidian through `docx-preview`. |
| `.doc` | Guidance page | Shown inside Obsidian with external-open and `.docx` conversion guidance. |

## Usage

1. Put a `.docx` file into your Obsidian vault.
2. Click the `.docx` file in the file explorer.
3. Read the document in the Obsidian tab opened by the plugin.
4. Use the toolbar to reload, zoom, fit width, show the outline, search, copy text, copy Markdown, create a summary note, or open the file externally.

### Zoom

- Type a percentage in the zoom input.
- Use `Ctrl` + mouse wheel over the document preview for continuous zoom.
- Use fit width when you want the document to match the current pane.

### Image Preview

Click an image rendered from the Word document to open a larger preview modal. Press `Esc` or close the modal to return to the document.

Inside the image preview:

- Use the mouse wheel to zoom.
- Drag the image to pan.
- Double-click to fit the image back into the preview.
- Use the toolbar to fit to window, view actual size, copy the image, or save the image.
- The toolbar shows the original image dimensions and current zoom.

### Search

- Type in the search field to highlight matching rendered text.
- Use the up and down buttons to move between results.
- Press `Enter` for the next result or `Shift` + `Enter` for the previous result.
- The result counter shows the current result and total matches.

### Outline

- Use the outline button to show or hide the heading outline.
- Click an outline item to scroll the rendered document to that heading.
- The outline is built from headings found in the rendered Word document.

### Copy

- Copy selected rendered text as plain text.
- Copy the whole document as plain text when no selection is active.
- Copy selected rendered content as Markdown.
- Copy the whole `.docx` as Markdown when no selection is active.

## Settings

Open Obsidian settings, then go to `Community plugins` -> `Obsidian Word Reader`.

Available settings:

- Plugin interface language.
- Default zoom percentage for newly opened Word previews.
- Whether newly opened Word previews should fit the pane width by default.
- Whether the outline is visible by default.
- Whether rendered images can be clicked for larger preview.
- Large file warning threshold in MB.
- External opening note. The plugin uses the operating system default application for `.docx` files.

## Compatibility and Errors

- Legacy `.doc` files open to an explanation page instead of being rendered directly.
- Encrypted or password-protected documents show a dedicated encrypted document message.
- Damaged or invalid `.docx` files show a damaged document message with next-step guidance.
- Large file warnings include the file size and use the configured threshold from settings.

## Performance and Stability

- Rendering work is guarded by a cancellation token so stale results are discarded.
- Word content is rendered into a temporary buffer before replacing the visible preview.
- The current rendered file state is tracked to avoid unnecessary repeated renders.
- Search highlighting is debounced while typing to reduce work on large documents.
- Closing or unloading a file releases document buffers, search timers, and search result references.

## Stability and Support

The stable reader scope, manual test checklist, support boundaries, and maintenance strategy are documented in [STABILITY.md](STABILITY.md).

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

## Key findings

## Follow-ups

## Quoted excerpts
```

If the same-name Markdown file already exists, the plugin opens it without overwriting content.

## Installation

### Install from a release zip

1. Download `obsidian-word-reader-vX.Y.Z.zip` from the GitHub release page.
2. Create the plugin folder in your vault:

   ```text
   YourVault/.obsidian/plugins/obsidian-word-reader/
   ```

3. Extract the zip into that folder. The files must be directly inside the plugin folder:

   ```text
   YourVault/.obsidian/plugins/obsidian-word-reader/main.js
   YourVault/.obsidian/plugins/obsidian-word-reader/manifest.json
   YourVault/.obsidian/plugins/obsidian-word-reader/styles.css
   ```

4. Restart Obsidian or reload community plugins.
5. In Obsidian, open Settings, enable Community plugins, then enable Obsidian Word Reader.

### Build from source

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

## Local Release

Create and validate a local release package:

```bash
npm run release
```

The release command runs TypeScript checks, builds the plugin, validates version consistency, creates the installable zip, and extracts the current changelog section for release notes.

Expected output:

```text
release/obsidian-word-reader-v1.1.0.zip
release/CHANGELOG-1.1.0.md
```

The zip root contains only the files Obsidian needs:

```text
main.js
manifest.json
styles.css
```

Release artifacts are ignored by Git and should not be committed.

## GitHub Auto Release

GitHub Actions creates a release automatically when a version tag is pushed:

```bash
git tag v1.1.0
git push origin v1.1.0
```

The workflow validates that the tag matches `package.json`, `manifest.json`, and `package-lock.json`, then builds the plugin, creates `release/obsidian-word-reader-v1.1.0.zip`, extracts the matching `CHANGELOG.md` section, and uploads the zip to the GitHub Release.

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

Validate release metadata after building:

```bash
npm run release:check
```

On this workstation, the expected Node.js executable is:

```text
E:\DevelopHelper\nvm\v18.20.8\node.exe
```

## Known Limits

- This is not a Word editor.
- The plugin never saves changes back to `.docx`.
- Legacy `.doc` files are not rendered directly, but the plugin shows external-open and conversion guidance.
- Complex Word layouts may not render exactly like Microsoft Word.
- Very large files or files with many images may render slowly.
- Mobile support is not included in this version.

## Security

The plugin reads `.docx` files through Obsidian's binary vault API for rendering and text extraction only.

It does not modify, overwrite, or write back to the original Word document. Use Word, WPS, or another external editor when the source document needs to be changed.

## Recommended Workflow

Keep original Word files in the vault, open them for reading, and create Markdown summary notes for long-term knowledge management.

This keeps the original document format intact while bringing summaries, decisions, tasks, and quotes into your Obsidian workflow.
