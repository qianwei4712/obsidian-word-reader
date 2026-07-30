# Office Reader

[中文文档](README.zh-CN.md) | [Roadmap](ROADMAP.md)

Office Reader is a desktop-only Obsidian plugin for opening `.docx`, `.pptx`,
and `.xlsx` files directly inside Obsidian as safe, read-only documents.

The plugin is not an Office editor. It provides local reading workflows while
keeping the original files unchanged.

Requires Obsidian Desktop 1.12.7 or newer.

## Features

- Open `.docx` files in an Obsidian tab.
- Open `.pptx` files in an Obsidian tab with local rendering for text, images,
  common shapes, tables, themes, layouts, and masters.
- Open `.xlsx` files in a bounded virtual grid with worksheet tabs, frozen
  panes, merged cells, row/column sizing, basic styles, number/date formats,
  and cached formula results.
- Navigate cells and supported named ranges, search the complete workbook,
  inspect hidden sheets, zoom or fit width, export selected ranges as TSV or
  Markdown, copy formulas separately, and create workbook summary notes.
- Navigate presentations with rendered slide thumbnails, extracted titles,
  previous/next controls, page-number jump, keyboard shortcuts, continuous
  zoom, fit to window, fullscreen reading, and external open.
- Search presentation titles, slide text, tables, and speaker notes; copy the
  current slide text and view speaker notes without leaving Obsidian.
- Copy privacy-safe PPTX render diagnostics for compatibility reports without
  including document content, speaker notes, or absolute vault paths.
- Render headings, paragraphs, lists, tables, images, and page breaks where supported by `docx-preview`.
- Keep the original Word file unchanged.
- Follow Obsidian light and dark themes for the Word preview surface. Complex Word documents with explicit colors may still affect the final rendered appearance.
- Zoom continuously with the toolbar percentage input or `Ctrl` + mouse wheel.
- Fit the rendered document to the pane width.
- Click rendered images to preview them in a larger modal.
- Search rendered text in the current document with previous/next navigation.
- Copy selected rendered text, whole-document plain text, selected Markdown, or whole-document Markdown.
- Use the collapsible outline panel to jump between Word heading levels 1-9,
  including localized and custom styles that inherit a heading outline level,
  and see the current section highlighted while reading.
- Restore zoom, fit mode, navigation state, and reading position for the 100
  most recently used Office documents.
- Switch the plugin interface between Chinese and English.
- Open the source file with the system default Word/WPS-compatible application.
- Show a clear fallback page for legacy `.doc` files with external-open and conversion guidance.
- Show clearer messages for encrypted, damaged, or unsupported Word documents.
- Avoid stale render results when switching or reloading documents.
- Create or open same-name Markdown summary notes linked back to source
  `.docx`, `.pptx`, and `.xlsx` files. Presentation notes include numbered
  slide references; spreadsheet notes include workbook structure, named
  ranges, the current selection, and bounded content previews.

## Supported Files

| Extension | Status | Notes |
| --- | --- | --- |
| `.docx` | Supported | Rendered inside Obsidian through `docx-preview`. |
| `.pptx` | Supported | Rendered locally as a read-only presentation. |
| `.xlsx` | Supported | Rendered locally as a read-only virtualized spreadsheet. |
| `.doc` | Guidance page | Shown inside Obsidian with external-open and `.docx` conversion guidance. |
| `.xls` | Guidance page | Shown inside Obsidian with external-open and `.xlsx` conversion guidance. |
| `.xlsm` | Not registered | Macro-enabled workbooks are outside the supported scope. |

## Usage

1. Put a `.docx`, `.pptx`, or `.xlsx` file into your Obsidian vault.
2. Click the file in the file explorer.
3. Read the document in the Obsidian tab opened by the plugin.
4. Use the toolbar and navigation panels to reload, navigate, zoom, search,
   copy text, inspect notes, create a summary note, or open the file externally.

### Zoom

- Type a percentage in the zoom input.
- Use `Ctrl` + mouse wheel over the document preview for continuous zoom.
- Use fit width when you want the document to match the current pane.

### PowerPoint Reading

- Use the thumbnail and title sidebar to inspect the presentation and jump
  directly to a slide.
- Search all slide titles, body text, tables, and speaker notes from the
  sidebar. Press `Ctrl`/`Cmd` + `F` to focus presentation search.
- Use the arrow buttons, `Page Up`/`Page Down`, `Left`/`Right`, or
  `Space`/`Shift` + `Space` to change slides. Press `Home` or `End` to jump to
  the first or last slide.
- Enter a slide number to jump directly to it.
- Use the percentage input or `Ctrl` + mouse wheel to zoom.
- Use fit window to keep the complete slide visible while resizing a pane.
- Use fullscreen for presentation-focused reading.
- Copy selected text from the rendered slide, or copy all extracted text from
  the current slide when no selection is active.
- Show or hide the current slide's speaker notes.
- Create a same-name presentation summary note with the current slide and
  numbered references for every slide.
- Copy render diagnostics from the toolbar or command palette when reporting a
  presentation compatibility or performance issue.
- The current slide, zoom, fit mode, scroll position, navigation visibility,
  and speaker-note visibility are restored per file.

### Spreadsheet Reading

- Switch among visible worksheets with the tabs below the grid. A separate
  indicator lists hidden and very-hidden worksheets and opens them only when
  requested without changing the source workbook.
- Enter `A1`, a rectangular range, a worksheet-qualified reference, or a
  supported workbook/sheet-scoped named range in the formula-bar name box to
  jump directly to it.
- Scroll large sparse worksheets without mounting their complete row and
  column range. The grid keeps only the viewport, finite overscan, and visible
  frozen rows/columns mounted.
- Use the mouse to drag a rectangular selection, `Shift` + click to extend it,
  or the arrow, page, `Home`, and `End` keys to move the active cell.
- Search values and stored formula text across the complete workbook,
  including hidden worksheets. Press `Ctrl`/`Cmd` + `F` to focus search and
  use `Enter`/`Shift` + `Enter` to move between cross-sheet results.
- Copy displayed values as TSV, export the selected rectangle as a Markdown
  table, or use the separate formula-copy action. Clipboard operations are
  capped at 250,000 cells.
- Inspect the selected cell's complete stored formula text and
  workbook-cached result in the read-only formula bar.
- Read legacy cell-comment authors and complete text from the formula bar.
  Comment markers remain discoverable even when the commented cell is blank.
- Preview package-local PNG, JPEG, GIF, and BMP drawings at their worksheet
  anchors. Common bar, line, area, and pie charts use only workbook-cached
  series data and show an explicit fallback for unsupported chart types.
- Preserve a bounded subset of numeric conditional formatting: literal
  `cellIs` rules, two-/three-color scales, and data bars.
- Create a same-name spreadsheet summary note containing sheet dimensions and
  visibility, named ranges, rich-content counts, the current selection, and up
  to 200 displayed cell previews collected from workbook content.
- Workbook hyperlinks open only after an explicit click and confirmation.
  External workbook references and data connections are never fetched.
- The active worksheet, zoom/fit preference, and scroll position are restored
  per file without persisting cell values or formulas.

### Image Preview

Click an image rendered from the Word document to open a larger preview modal. Press `Esc` or close the modal to return to the document.

Inside the image preview:

- Use the mouse wheel to zoom.
- Drag the image to pan.
- Double-click to fit the image back into the preview.
- Use the toolbar to fit to window, view actual size, copy the image, or save the image.
- Press `Ctrl` + `C` on Windows/Linux or `Cmd` + `C` on macOS while the image preview is open to copy the image.
- Saved images use a default name based on the source document and image dimensions when available.
- The toolbar shows the original image dimensions and current zoom.

### Search

- Type in the search field to highlight matching rendered text.
- Use the up and down buttons to move between results.
- Press `Enter` for the next result or `Shift` + `Enter` for the previous result.
- The result counter shows the current result and total matches.

### Outline

- Use the outline button to show or hide the heading outline.
- Click an outline item to scroll the rendered document to that heading.
- Collapse or expand outline sections with the chevron beside parent headings.
- The outline highlights the current section as the document scrolls.
- The outline is built from headings found in the rendered Word document.

### Copy

- Copy selected rendered text as plain text.
- Copy the whole document as plain text when no selection is active.
- Copy selected rendered content as Markdown.
- Copy the whole `.docx` as Markdown when no selection is active.

## Settings

Open Obsidian settings, then go to `Community plugins` -> `Office Reader`.

Available settings:

- Plugin interface language.
- Default zoom percentage for newly opened Word previews.
- Whether newly opened Word previews should fit the pane width by default.
- Whether newly opened worksheets should fit the grid width by default.
- Whether the outline is visible by default.
- Whether rendered images can be clicked for larger preview.
- Large file warning threshold in MB.
- External opening note. The plugin uses the operating system default
  application for source Office files.

## Compatibility and Errors

- Legacy `.doc` files open to an explanation page instead of being rendered directly.
- Encrypted or password-protected documents show a dedicated encrypted document message.
- Failed `.docx` previews are classified as encrypted, format mismatched, damaged ZIP packages, invalid XML structures, unsupported document structures, or unknown failures where possible.
- Failed `.pptx` previews distinguish format mismatch, encryption, damaged XML
  or ZIP data, unsupported structures, and safe-preview limit violations.
- Failed `.xlsx` previews distinguish format mismatch, encryption, damaged
  workbook XML/ZIP data, unsupported active content, and safe-preview limits.
- Legacy `.xls` files show conversion guidance and an external-open action.
- PPTX archives are checked before decompression for file count, per-entry
  expanded size, total expanded size, ZIP64 usage, encryption, and abnormal
  compression ratios.
- Error pages provide a collapsed diagnostic section and a copy action for issue reports.
- Copied diagnostics use a shared, versioned JSON envelope and include only the
  format, diagnostic kind, file name, size, modification time, a privacy-safe
  summary, and bounded format metrics. Raw renderer errors, document content,
  internal XML, and absolute vault paths are excluded.
- Successful PPTX previews also expose copyable JSON render diagnostics with
  package sizes, slide dimensions, object counts, generated-resource counts,
  explicit font families, and render duration. Document text, speaker notes,
  internal XML, and absolute vault paths are excluded.
- Large file warnings include the file size and use the configured threshold from settings.

## Performance and Stability

- A shared reader shell plus `OfficeReaderAdapter`, `ReaderSession`, and
  `ReaderCapabilities` contracts keep toolbars, status, errors, diagnostics,
  and cleanup behavior consistent. DOCX, PPTX, and XLSX rendering logic lives in
  format-specific adapters and sessions.
- Persisted settings use an explicit schema version with separate `common`,
  `docx`, `pptx`, and `xlsx` sections. Pre-2.4 flat settings migrate
  automatically.
- Rendering work is guarded by a cancellation token so stale results are discarded.
- PPTX file and slide rendering use separate cancellation generations so rapid
  file switching and page navigation cannot commit stale slides.
- The current PPTX slide renders before full metadata indexing finishes.
  Background metadata parsing is cancellable and limited to four workers.
- Presentation search reuses a normalized metadata index. The navigation list
  is virtualized to at most 60 mounted rows, and thumbnail rendering is limited
  to two lower-priority workers.
- Off-screen thumbnails are unmounted and their generated image resources are
  released; stale slide and thumbnail renders are cancelled cooperatively.
- PPTX shape, text, and table rendering yields between bounded work slices,
  while XML, relationship, slide-context, and binary caches have fixed limits.
- Generated presentation resources are also released when a presentation is
  reloaded, switched, or closed.
- XLSX worksheets use a 2,500-cell mount ceiling with finite row/column
  overscan. Frozen panes remain visible without materializing the full sheet.
- Worksheet and shared-string XML is decompressed and validated in chunks.
  Loading reports progress, dense worksheets do not retain the complete
  `sheetData` XML, and parsing can be cancelled between chunks and rows.
- Worksheet parsing, switching, workbook-wide search, summary collection, and
  scheduled grid rendering are cancellable. Package caches, mounted drawings,
  chart sample counts, and image Blob URLs have fixed bounds and are released
  on reload, sheet/file switch, and view close.
- Word content is rendered into a temporary buffer before replacing the visible preview.
- Long documents commit rendered pages and build navigation in cancellable chunks so the interface can update between batches.
- Long previews defer off-screen page painting until pages approach the viewport.
- Embedded images use lazy loading, asynchronous decoding, and recyclable Blob URLs instead of persistent base64 data URLs.
- The current rendered file state is tracked to avoid unnecessary repeated renders.
- DOCX render post-processing collects images, pages, headings, Blob URLs, and
  searchable text in one pass.
- DOCX search builds one reusable text index per render and uses CSS highlights
  without rewriting the document DOM. Reading-state writes are coalesced and
  unchanged states are skipped.
- Fit-width mode is a CSS state change and does not rerender the document.
- Production builds enforce a compressed `dist/main.js` budget of 500 KiB.
- Reading state uses a bounded 100-file LRU store. Persisted entries contain
  only path, modification time, format, position, zoom, and navigation state.
  When a source file changes, stale position/navigation data is cleared while
  the user's zoom and fit preference is retained.
- Development builds log Word file reading, rendering, DOM commit, outline,
  and total preview timings, plus PPTX render duration, object counts,
  resources, and fonts, to the developer console.
- Closing or unloading a file releases document buffers, generated Blob URLs, search timers, and search result references.

## Stability and Support

The stable reader scope, manual test checklist, support boundaries, and maintenance strategy are documented in [STABILITY.md](STABILITY.md).

### Summary Notes

The summary note action creates a same-name Markdown file next to a Word
document, PowerPoint presentation, or Excel workbook.

Example:

```text
Report.docx
Report.md
Quarterly review.pptx
Quarterly review.md
Data register.xlsx
Data register.md
```

The generated note includes frontmatter and starter sections:

```markdown
---
source: "Report.docx"
type: word-note
reader: office-reader
reader_format: docx
created: 2026-05-28
---

# Report

Source: [[Report.docx]]

## Summary

## Key findings

## Follow-ups

## Quoted excerpts
```

Presentation notes use `type: presentation-note` and
`reader_format: pptx`, record the current slide, and include a numbered
reference list for every slide. If the same-name Markdown file already exists,
the plugin opens it without overwriting content.

Spreadsheet notes use `type: spreadsheet-note` and `reader_format: xlsx`,
record the current worksheet and selection, list worksheet dimensions,
visibility, supported named ranges, and comment/image/chart/conditional-rule
counts, and include at most 200 displayed-value preview cells. Existing
same-name Markdown files are opened without being overwritten.

## Installation

### Install from a release zip

1. Download `obsidian-word-reader-X.Y.Z.zip` from the GitHub release page.
2. Create the plugin folder in your vault:

   ```text
   YourVault/.obsidian/plugins/word-reader/
   ```

3. Extract the zip into that folder. The files must be directly inside the plugin folder:

   ```text
   YourVault/.obsidian/plugins/word-reader/main.js
   YourVault/.obsidian/plugins/word-reader/manifest.json
   YourVault/.obsidian/plugins/word-reader/styles.css
   ```

4. Restart Obsidian or reload community plugins.
5. In Obsidian, open Settings, enable Community plugins, then enable Office Reader.

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
release/obsidian-word-reader-X.Y.Z.zip
release/CHANGELOG-X.Y.Z.md
```

The zip root contains only the files Obsidian needs:

```text
main.js
manifest.json
styles.css
```

Release artifacts are ignored by Git and should not be committed.

## GitHub Auto Release

GitHub Actions creates a release automatically when a version tag without a `v` prefix is pushed:

```bash
git tag 2.5.0
git push origin 2.5.0
```

The workflow validates that the tag matches `package.json`, `manifest.json`, and `package-lock.json`, checks the 500 KiB bundle budget, builds the plugin, creates `release/obsidian-word-reader-2.5.0.zip`, extracts the matching `CHANGELOG.md` section, and uploads `main.js`, `manifest.json`, `styles.css`, and the zip to the GitHub Release.

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

Run unit and integration tests:

```bash
npm test
```

Run the production security scan after building:

```bash
npm run build
npm run security:scan
```

Run every automated quality gate:

```bash
npm run check
```

Validate release metadata and zip contents after packaging:

```bash
npm run release:zip
npm run release:check
```

The initial community-directory submission and release review steps are in
[COMMUNITY_PLUGIN_CHECKLIST.md](COMMUNITY_PLUGIN_CHECKLIST.md).

CI and release workflows use Node.js 20.19.0.

## Known Limits

- This is not an Office editor.
- The plugin never saves changes back to `.docx`, `.pptx`, or `.xlsx`.
- Legacy `.xls` files are not rendered directly, and `.xlsm` is not
  registered. Modern threaded comments, floating-object editing, pivot tables,
  slicers, pixel-perfect chart fidelity, and complex conditional formatting
  are outside the 3.2.0 XLSX scope.
- Legacy `.doc` files are not rendered directly, but the plugin shows external-open and conversion guidance.
- Complex Word layouts may not render exactly like Microsoft Word.
- PPTX animation, transitions, audio/video playback, macros, editing, charts,
  SmartArt, SVG/GIF/WebP media, and pixel-perfect PowerPoint fidelity are not
  supported.
- Charts and SmartArt are identified with local placeholders and diagnostics;
  their native PowerPoint appearance is not rendered.
- Presentation thumbnails are generated locally on demand. Very large or
  image-heavy slides can still take longer when they enter the visible range.
- Word previews follow the current Obsidian theme, but explicit colors stored in the Word document may still influence the rendered result.
- Very large files or files with many images may render slowly.
- Mobile support is not included in this version.

## Security

This plugin is designed with security as a top priority:

- **Local-only operations**: The plugin reads `.docx`, `.pptx`, and `.xlsx`
  from your local Obsidian vault. It never automatically fetches workbook
  links, external references, data connections, or remote resources. A safe
  hyperlink may open only after the user clicks and confirms it.
- **No external resources**: The plugin never loads scripts, styles, or assets from the internet. All rendering logic runs locally.
- **Read-only access**: The plugin never modifies, overwrites, or writes back
  to the original Office file. It uses Obsidian's binary vault API for
  rendering and text extraction only.
- **No dynamic script injection**: The plugin creates only structural DOM elements (`div`, `span`, `button`, `input`) for document rendering. No `<script>` elements are created or injected at any point.
- **Safe PPTX archives**: ZIP metadata is validated before decompression, and
  external package relationships are ignored instead of being loaded.
- **Safe XLSX formulas**: Formula text and cached results are displayed without
  running macros, external formula code, or recalculation.
- **Safe XLSX rich content**: Images are restricted to package-local safe
  raster formats; chart previews use cached series only. External drawing,
  chart, and workbook relationships are never resolved. The reviewed
  [XLSM compatibility matrix](XLSM_COMPATIBILITY.md) keeps `.xlsm`
  unregistered.
- **Sandboxed rendering**: Office content is rendered into structural DOM
  containers with no script execution context.
- **Desktop-only**: The plugin requires desktop Obsidian because it uses Electron APIs for image clipboard operations and file dialogs. This is declared in `manifest.json` as `isDesktopOnly: true`.

Use Word, WPS, or another external editor when the source document needs to be changed.

## Recommended Workflow

Keep original Office files in the vault, open them for reading, and create
Markdown summary notes for long-term knowledge management.

This keeps the original document format intact while bringing summaries, decisions, tasks, and quotes into your Obsidian workflow.
