# Stability and Maintenance

## 1.0.0 Stable Scope

The 1.0.0 stable line freezes the core reader scope:

- Open `.docx` files inside Obsidian Desktop as read-only previews.
- Show a guidance page for legacy `.doc` files.
- Keep source Word files unchanged.
- Support zoom, fit width, search, text copy, Markdown summary note creation, external open, image preview, and release packaging.
- Keep rendering, search, and image preview stable for normal vault documents.

New reader features should be planned for minor versions after 1.0.0. Patch releases should focus on regressions, compatibility fixes, documentation corrections, and packaging issues.

## 2.0.0 PPTX Scope

The 2.0.0 line adds local, read-only `.pptx` preview:

- Render text, embedded images, common shapes, tables, theme colors, layouts,
  and masters without loading remote resources.
- Navigate by previous/next slide and page number.
- Support zoom, fit to window, fullscreen reading, external open, and restored
  per-file page state.
- Validate ZIP metadata before decompression with entry-count, per-entry size,
  total expanded size, encryption, ZIP64, and compression-ratio limits.
- Keep animation, transitions, audio/video playback, macros, editing, charts,
  SmartArt, SVG/GIF/WebP media, and exact PowerPoint layout fidelity outside
  the supported scope.

## 2.1.0 PPTX Workflow Scope

The 2.1.0 line adds presentation reading and knowledge-capture workflows:

- Show rendered thumbnails and extracted titles in a slide navigation panel.
- Search slide titles, body text, tables, and speaker notes locally.
- Copy selected rendered text or all extracted text from the current slide.
- Show speaker notes for the current slide without changing the source file.
- Create or open a same-name Markdown presentation note that records the
  current slide and includes numbered references for every slide.
- Restore navigation-panel and speaker-note visibility per presentation.
- Support arrow keys, page keys, `Space`/`Shift` + `Space`, `Home`, `End`, and
  `Ctrl`/`Cmd` + `F` presentation shortcuts.

## 2.2.0 PPTX Quality and Performance Scope

The 2.2.0 line bounds presentation rendering work and adds compatibility
diagnostics:

- Mount thumbnails only for visible and nearby navigation entries, with a
  fixed upper bound on concurrently rendered previews.
- Cooperatively cancel stale slide and thumbnail work before it can commit.
- Release generated image resources after cancellation, thumbnail unmount,
  reload, file switch, and view close.
- Copy privacy-safe render diagnostics with package, slide, object, resource,
  font, and timing metrics, excluding content and absolute vault paths.
- Maintain generated compatibility fixtures and a structural visual baseline
  for tables, charts, SmartArt, themes, layouts, and masters.
- Keep charts and SmartArt outside native-fidelity support while identifying
  them with explicit placeholders and separate diagnostics.

## 2.3.0 Global Performance Scope

The 2.3.0 line governs startup responsiveness, interaction latency, memory,
and bundle size across DOCX and PPTX:

- Render the current PPTX slide before full metadata indexing completes.
- Limit metadata parsing to four workers and thumbnail rendering to two
  lower-priority workers; cancel stale work during navigation and scrolling.
- Virtualize presentation navigation to at most 60 mounted rows for a
  1,000-slide deck.
- Yield during PPTX shape, text, and table rendering, targeting work slices of
  approximately 8 ms.
- Bound XML, relationship, slide-context, and binary caches.
- Build one DOCX text-search index per render, update the current result in
  constant time, and avoid rewriting rendered DOM for highlights.
- Coalesce DOCX reading-state writes, skip unchanged states, merge render
  post-processing scans, and switch fit-width through CSS.
- Keep `dist/main.js` at or below 500 KiB and maintain generated 200/1,000
  slide PPTX plus 100-page DOCX regression fixtures.

## 2.4.0 Shared Reader Architecture Scope

The 2.4.0 line closes the reader architecture before XLSX work begins:

- Keep the existing `word-reader-view` and `pptx-reader-view` identifiers while
  routing format behavior through `OfficeReaderAdapter`, `ReaderSession`, and
  `ReaderCapabilities`.
- Build DOCX and PPTX views on the same `OfficeReaderShell` toolbar, status,
  error, and diagnostic conventions while preserving legacy CSS classes.
- Persist settings with an explicit schema version and separate `common`,
  `docx`, `pptx`, and reserved `xlsx` sections; migrate pre-2.4 data on load.
- Keep at most 100 reading states containing only file identity, modification
  time, format, position, zoom, and navigation state.
- Invalidate stale position/navigation data when file modification time
  changes while retaining the user's zoom and fit preference.
- Use one privacy-safe diagnostic envelope and one summary-note frontmatter
  convention across formats.

## 2.5.0 OOXML and XLSX Validation Scope

The 2.5.0 line completes pre-release XLSX validation without exposing an XLSX
view:

- Apply shared ZIP entry, expansion, compression-ratio, encryption, ZIP64,
  path, macro, ActiveX, OLE, and script-media checks to public DOCX/PPTX
  loading and the unregistered XLSX research parser.
- Parse workbook metadata and sparse worksheet cells, styles, cached formulas,
  merges, frozen panes, hidden sheets, hyperlinks, and raster-image metadata
  without recalculation or remote-resource access.
- Keep only a visible grid window and finite overscan, with a 2,500-cell
  default mount ceiling.
- Maintain generated normal, damaged, encrypted, malicious, and extreme-scale
  fixtures plus the 100,000-row benchmark budget.
- Pin direct runtime dependencies, inventory compatible licenses, audit
  production bundle inputs, and retain the 500 KiB `2.x` bundle ceiling and
  8 MiB `3.0.0` release target.
- Do not register `.xlsx`, an XLSX view, a public feature flag, or a technical
  preview entry before `3.0.0`.

## 3.0.0 XLSX Read-only MVP Scope

The 3.0.0 line exposes the validated XLSX core through a public, local,
read-only spreadsheet session:

- Register `.xlsx` and legacy `.xls` with `xlsx-reader-view`; render `.xlsx`
  locally and show external-open/conversion guidance for `.xls`. Keep `.xlsm`
  unregistered.
- Switch visible worksheets and render only the viewport, finite overscan, and
  visible frozen rows/columns under the 2,500-cell mount ceiling.
- Preserve merged cells, row/column sizing, basic font/fill/border/alignment
  styles, number/date formats, formula text, and workbook-cached results.
- Search the current worksheet with cancellable chunked work, navigate results,
  zoom or fit width, and restore active sheet/scroll/zoom without persisting
  cell content.
- Select rectangular ranges and copy displayed values by default; expose
  formula copy as a separate operation with a 250,000-cell materialization
  ceiling.
- Never recalculate formulas or dereference external workbook references,
  connections, or remote resources. Open only safe-protocol hyperlinks after a
  direct user action and confirmation.
- Cancel stale workbook, worksheet, search, and scheduled grid work during
  reload, sheet/file switching, and view close; clear bounded package caches.
- Keep `dist/main.js` within 500 KiB and the three-file release archive below
  the 8 MiB 3.0.0 target.

## 3.1.0 Workbook Navigation and Knowledge Extraction Scope

The 3.1.0 line extends the local read-only XLSX session without broadening its
network or execution boundary:

- Use the formula-bar name box to navigate static A1 cells/ranges,
  worksheet-qualified references, and supported workbook- or sheet-scoped
  named ranges. Ignore dynamic names, unions, formulas, and external workbook
  targets.
- Search displayed values and stored formula text across visible, hidden, and
  very-hidden worksheets with cancellable sequential parsing and cross-sheet
  result navigation.
- Show an explicit hidden-sheet count and user-initiated navigation menu
  without modifying workbook visibility.
- Present complete stored formula text, displayed or cached results, and the
  cached-only safety notice in a read-only formula bar.
- Copy displayed values as TSV, export the complete selected rectangle as a
  Markdown table, and keep formula copy separate. All clipboard paths retain
  the 250,000-cell materialization ceiling.
- Create or open a same-name XLSX summary note containing worksheet dimensions
  and visibility, named ranges, the current selection, and no more than 200
  displayed-value preview cells collected across the workbook.
- Cancel workbook search and summary collection during reload, worksheet/file
  switching, and view close; continue clearing bounded package caches.

## 3.2.0 XLSX Rich Content and Large Workbook Scope

The 3.2.0 line adds bounded rich-content previews and removes full-worksheet
XML retention from the large-workbook path:

- Parse legacy cell-comment authors and complete rich text, expose comment
  markers for populated and blank cells, and show the selected comment in the
  read-only formula bar.
- Render package-local PNG/JPEG/GIF/BMP images with one-cell or two-cell
  anchors. Image bytes, mounted drawings, and recyclable Blob URLs all have
  fixed limits; external image relationships remain ignored.
- Render simplified bar, line, area, and pie previews from workbook-cached
  series only. Unsupported chart types show an explicit fallback, formulas are
  not recalculated, and external chart data is never fetched.
- Apply literal numeric `cellIs` rules, two-/three-color scales, and data bars.
  Parse no more than 256 conditional-formatting rules per worksheet, and let
  differential styles override only properties declared by the rule.
- Decompress and validate worksheet and shared-string XML in chunks, parse
  worksheet rows incrementally, report loading progress, and cancel between
  chunks and bounded row batches without retaining complete `sheetData` XML.
- Keep no more than 12 worksheet drawings mounted, sample no more than 48
  rendered chart points, and release rich-content resources on reload,
  worksheet/file switch, and view close.
- Benchmark at least 20,000 dense rows and 80,000 cells, including elapsed
  parse time, peak heap, input-chunk count, maximum worksheet-data buffer,
  cancellation, and virtual-grid DOM budgets.
- Keep `.xlsm` unregistered after reviewing macro-enabled content types, VBA,
  macro/dialog sheets, ActiveX, OLE, script media, external relationships, and
  cached values in `XLSM_COMPATIBILITY.md`.

## 3.3.0 Performance Baseline and Observability Scope

- `npm run performance:check` measures DOCX, PPTX, and XLSX in five isolated
  processes and reports every raw sample plus median and p95 aggregates.
- Every format records package load, parse, first readable content, search,
  navigation, scroll, peak heap, DOM, cache bounds, cancellation, and cleanup.
- XLSX retains the 3.2.0 first-window, scroll, heap, DOM, dense-parse, and
  streamed-buffer limits. Production `main.js` remains at most `500 KiB`; the
  three-file release zip remains at most `8 MiB`.
- CI uploads the full JSON result, a compact trend point, and a Markdown job
  summary. Budget failures identify the format, aggregate, stage/metric,
  observed value, and threshold.
- DOCX, PPTX, and XLSX expose the same Obsidian Desktop performance-diagnostic
  schema. It contains numeric timing/resource observations only and excludes
  document content, notes, cell values, internal XML, and absolute vault paths.
- Node.js gates are calibrated with the manual Obsidian Desktop worksheet in
  `benchmarks/OBSIDIAN_DESKTOP_BASELINE.md`, including activation, long tasks,
  actual DOM, and resources remaining after close.

## Manual Test Checklist

Run this checklist before publishing a stable release:

- Build and package:
  - `npm run typecheck`
  - `npm test`
  - `npm run build`
  - `npm run performance:check`
  - `npm run dependencies:audit`
  - `npm run security:scan`
  - `npm run release`
  - Confirm the release zip contains only `main.js`, `manifest.json`, and `styles.css`.
- Installation:
  - Install the release zip into a test vault at `.obsidian/plugins/word-reader/`.
  - Enable the plugin from Obsidian Community plugins.
  - Confirm Obsidian loads the plugin without console errors.
- `.docx` preview:
  - Open a small `.docx` file.
  - Open a document with headings, lists, tables, images, and page breaks.
  - Confirm the preview is read-only and does not modify the source file.
- `.doc` compatibility:
  - Open a legacy `.doc` file.
  - Confirm the guidance page appears.
  - Confirm external open is available on Desktop.
- `.pptx` preview:
  - Open presentations containing text, embedded images, common shapes, tables,
    theme colors, layouts, and masters.
  - Test previous/next navigation, page-number jump, zoom, fit to window,
    fullscreen reading, keyboard navigation, and external open.
  - Confirm the sidebar shows rendered thumbnails, extracted titles, and the
    active slide, and that clicking an entry navigates to the correct slide.
  - Scroll a presentation with at least 50 slides and confirm only visible and
    nearby thumbnails are mounted; off-screen thumbnails return to placeholders.
  - Open a 1,000-slide generated fixture and confirm no more than 60
    navigation rows are mounted and the current slide renders before all
    titles finish indexing.
  - Rapidly scroll, search, and navigate while thumbnails are rendering and
    confirm stale previews do not replace current entries.
  - Confirm off-screen thumbnail images release their Blob URLs and render
    again when their entries return to the visible range.
  - Search for text in slide titles, body text, tables, and speaker notes;
    confirm matching snippets and note-only matches are identified correctly.
  - Copy selected rendered text, then clear the selection and copy all
    extracted text from the current slide.
  - Show and hide speaker notes, navigate between slides, and confirm the
    displayed notes follow the current slide.
  - Create a same-name presentation summary note and confirm it records the
    current slide and contains numbered references for every slide.
  - Confirm an existing same-name Markdown note opens without being overwritten.
  - Test `Space`, `Shift` + `Space`, `Home`, `End`, and `Ctrl`/`Cmd` + `F`.
  - Reopen a presentation and confirm page, zoom, fit mode, and scroll position
    are restored together with navigation and speaker-note visibility.
  - Rapidly switch files, pages, and zoom levels and confirm no stale slide is
    displayed.
  - Open damaged, encrypted, renamed non-PPTX, over-limit, and unsupported
    presentations and confirm each shows understandable recovery guidance.
  - Confirm presentations with external image relationships do not make
    network requests.
  - Open fixtures containing tables, chart frames, SmartArt frames, theme
    colors, layouts, and masters; confirm supported content renders and chart
    or SmartArt frames show explicit placeholders.
  - Copy PPTX render diagnostics from the toolbar and command palette. Confirm
    the JSON contains package, slide, object, resource, font, and timing
    metrics but no slide text, speaker notes, XML, or absolute vault paths.
  - In a development build, confirm the console logs PPTX render duration,
    object counts, generated resources, and explicit font families.
- `.xlsx` preview:
  - Open a workbook with multiple visible sheets and switch rapidly between
    tabs; confirm stale worksheet parses never replace the active sheet.
  - Confirm frozen rows/columns remain visible while scrolling and merged
    cells, custom row/column sizes, basic styles, numbers, dates, and cached
    formula results render correctly.
  - Open the generated 100,000-row sparse workbook, scroll to the tail, and
    confirm no more than 2,500 cells are mounted.
  - Search displayed values and formula text across visible and hidden
    worksheets; navigate between cross-sheet results, change queries rapidly,
    and confirm cancelled searches leave no stale highlights.
  - Enter direct cells, rectangular ranges, qualified worksheet references,
    workbook names, and sheet-local names in the name box. Confirm invalid,
    dynamic, and external names are rejected.
  - Confirm the hidden-sheet indicator lists hidden and very-hidden sheets and
    opens them only after direct user selection.
  - Drag a rectangular selection and copy displayed values as TSV. Export the
    same selection as Markdown and confirm all data rows remain present. Use
    the separate formula-copy action and confirm formula cells begin with `=`.
  - Confirm the formula bar shows complete stored formula text and cached
    values without recalculating the workbook.
  - Open legacy comments by different authors, including a comment on a blank
    cell. Confirm the grid marker and formula-bar detail show the complete
    author and text without editing the workbook.
  - Open package-local PNG/JPEG/GIF/BMP drawings with one-cell and two-cell
    anchors. Confirm external image relationships never load, no more than 12
    drawings mount together, and Blob URLs are released after sheet/file
    switching and view close.
  - Open cached bar, line, area, and pie charts. Confirm cached labels/values
    render locally, unsupported chart types show a fallback, and formula or
    external chart sources are never evaluated or fetched.
  - Confirm literal numeric `cellIs`, two-/three-color-scale, and data-bar
    rules affect only matching cells, while unsupported or malformed rules
    remain inert.
  - Open the generated 20,000-row/80,000-cell dense workbook. Confirm loading
    progress advances, parsing can be cancelled between chunks, scrolling
    remains virtualized, and the performance check reports a bounded
    worksheet-data buffer and heap.
  - Create an XLSX summary note and confirm it contains worksheet dimensions,
    visibility, named ranges, rich-content counts, the current selection, and
    bounded workbook previews. Confirm an existing same-name note opens
    without being replaced.
  - Click an internal or safe external hyperlink, cancel the confirmation, and
    confirm nothing opens. Confirm `javascript:`, `file:`, external-workbook,
    and data-connection targets never load.
  - Reopen the workbook and confirm active worksheet, scroll position, zoom,
    and fit preference restore without saved cell values or formulas.
  - Open damaged, encrypted, ZIP64, over-limit, macro, OLE, and script-media
    fixtures and confirm they fail with understandable recovery guidance.
  - Open a legacy `.xls` and confirm conversion guidance plus external open;
    confirm `.xlsm` is not registered.
- Error diagnostics:
  - Open encrypted, invalid, and damaged test documents where available.
  - Confirm error pages show an appropriate category and recovery guidance.
  - Confirm diagnostic details are collapsed by default.
  - Confirm copied diagnostics are valid JSON.
  - Confirm copied diagnostics contain file name, size, modification time, category, a privacy-safe summary, and an error fingerprint.
  - Confirm copied diagnostics do not contain raw renderer errors, document text, internal XML content, or absolute vault paths.
  - Confirm the external-open action remains available.
- Navigation and reading:
  - Test toolbar zoom input.
  - Test `Ctrl` + mouse wheel zoom.
  - Test fit width.
  - Confirm scrolling updates the highlighted current outline section.
  - Collapse parent outline headings and confirm their descendants hide and restore correctly.
  - Reopen a document and confirm zoom, fit width, outline visibility, collapsed sections, and scroll position are restored.
  - Open more than 100 distinct documents and confirm persisted reading state remains capped at 100 entries.
  - Test search, previous/next navigation, and current result highlighting.
  - Search repeatedly in a 100-page document and confirm navigation does not
    rewrite or disturb the rendered content.
  - Test large document loading status with a file above the configured warning size.
  - Confirm a long document shows reading, rendering, preview preparation, and navigation-building status updates before the final preview state.
  - Confirm a document with at least 12 rendered pages remains responsive while pages are committed and while searching for a frequent term.
  - Confirm rapidly changing the search query or switching files does not leave stale highlights or outline entries.
  - Confirm image-heavy documents still support image preview, copy, and save after switching from base64 data URLs to Blob URLs.
  - Confirm reloading, switching, and closing image-heavy documents does not leave broken previews in an already-open image modal.
  - In a development build, confirm the developer console logs read, render, DOM commit, outline, total duration, page count, and image count metrics.
- Copy and notes:
  - Copy selected rendered text.
  - Copy whole-document plain text.
  - Copy selected content as Markdown.
  - Copy whole-document Markdown.
  - Create or open a same-name Markdown summary note.
- Image preview:
  - Click a rendered image.
  - Test image zoom, pan, fit, actual size, copy, and save.
- 1.2.x theme and image behavior:
  - Test the Word preview surface in both Obsidian light and dark themes.
  - Confirm fallback message pages remain readable in both themes.
  - Confirm the image preview modal remains readable in both themes.
  - Confirm `Ctrl` + `C` on Windows/Linux or `Cmd` + `C` on macOS copies the current image while the image preview modal is open.
  - Confirm saved image default names are based on the source document and include image dimensions when available.
  - Do not include PPTX preview, reading-position restore, or outline folding in the 1.2.x manual test scope.
- Settings and language:
  - Switch between Chinese and English.
  - Confirm settings text, toolbar labels, notices, status messages, errors, and image modal text follow the selected language.
  - Confirm default zoom, default fit width, outline visibility, image preview, and large file warning settings persist after reload.
  - Confirm changing interface language preserves the current document reading position and outline state.
- Release automation:
  - Confirm unit and integration tests pass in CI.
  - Confirm the security scan passes against source and `dist/main.js`.
  - Confirm `node scripts/release-check.mjs --tag X.Y.Z` passes for the release tag.
  - Push an `X.Y.Z` tag in a test release flow.
  - Confirm GitHub Actions creates a release and uploads `main.js`, `manifest.json`, `styles.css`, and the zip.

## Support Boundaries

Supported:

- Obsidian Desktop 1.12.7 or newer.
- Local desktop vaults.
- `.docx` read-only preview through `docx-preview`.
- `.pptx` local, read-only preview for text, embedded images, common shapes,
  tables, themes, layouts, and masters.
- `.pptx` thumbnail/title navigation, local full-presentation search, current
  slide text copy, speaker-note viewing, and numbered presentation summary notes.
- `.pptx` on-demand thumbnail mounting, cooperative render cancellation,
  generated-resource cleanup, and privacy-safe render diagnostics.
- `.xlsx` local, read-only virtual grid for worksheet switching, frozen panes,
  merged cells, row/column sizes, basic styles and formats, cached formulas,
  name-box/named-range navigation, workbook-wide search, hidden-sheet
  indicators, zoom, TSV/Markdown range copy, summary notes, and confirmed
  hyperlinks.
- Plain text and Markdown copy from the already rendered DOCX content.
- `.doc` detection with external-open and conversion guidance.
- `.xls` detection with external-open and `.xlsx` conversion guidance.

Not supported:

- Editing or saving Word files.
- Mobile-specific support.
- Direct rendering of legacy `.doc` binary files.
- Password-protected or encrypted Word documents.
- Password-protected or encrypted PowerPoint presentations.
- Password-protected or encrypted Excel workbooks.
- Direct rendering of legacy `.xls`, and registration of `.xlsm`.
- XLSX modern threaded comments, floating-object editing, pivot tables,
  slicers, pixel-perfect chart fidelity, and complex conditional formatting.
- PPTX animation, transitions, audio/video playback, macros, editing, charts,
  SmartArt, SVG/GIF/WebP media, and pixel-perfect PowerPoint rendering.
- Perfect Microsoft Word layout fidelity for complex documents.
- Remote vault adapters that do not expose local file paths for external opening.

## Maintenance Strategy

- Use semantic versioning:
  - Patch versions fix regressions and documentation errors.
  - Minor versions add compatible reader, navigation, copy, or workflow features.
  - Major versions may change support boundaries or core architecture.
- Keep release artifacts reproducible through `npm run release`.
- Keep `README.md`, `README.zh-CN.md`, and `CHANGELOG.md` updated in every release.
- Keep the manual checklist current when user-facing behavior changes.
- Prefer small, focused releases over large mixed changes.

## 稳定性与维护策略

### 1.0.0 稳定版范围

1.0.0 稳定线冻结核心阅读器能力：

- 在 Obsidian 桌面端以只读方式打开 `.docx` 文件。
- 为旧版 `.doc` 文件显示说明页面。
- 不修改原始 Word 文件。
- 支持缩放、适配宽度、搜索、文本复制、Markdown 摘要笔记、外部打开、图片预览和发布打包。
- 保持常见 vault 文档的渲染、搜索和图片预览稳定。

新的阅读器功能放到 1.0.0 之后的小版本中规划。补丁版本优先修复回归、兼容性、文档和打包问题。

### 2.0.0 PPTX 范围

2.0.0 稳定线新增本地只读 `.pptx` 预览：

- 在不加载远程资源的前提下渲染文本、内嵌图片、常见形状、表格、主题颜色、版式和母版。
- 支持上一张/下一张、页码跳转、缩放、适应窗口、全屏阅读、外部打开和按文件恢复阅读页码。
- 解压前校验条目数量、单条目大小、总解压大小、加密、ZIP64 和异常压缩率。
- 动画、切换效果、音视频播放、宏、编辑、图表、SmartArt、SVG/GIF/WebP 媒体和
  PowerPoint 像素级还原不在支持范围内。

### 2.1.0 PPTX 阅读工作流范围

2.1.0 稳定线新增演示文稿阅读和知识整理能力：

- 在幻灯片导航面板显示真实缩略图和提取标题。
- 在本地搜索幻灯片标题、正文、表格和演讲者备注。
- 复制选中的渲染文本，或复制当前幻灯片的全部提取文本。
- 查看当前幻灯片的演讲者备注，不修改源文件。
- 创建或打开同名 Markdown 演示文稿笔记，记录当前页并为全部幻灯片生成页码引用。
- 按演示文稿恢复导航面板和演讲者备注的显示状态。
- 支持方向键、翻页键、`Space`/`Shift` + `Space`、`Home`、`End` 和
  `Ctrl`/`Cmd` + `F` 快捷键。

### 2.2.0 PPTX 质量与性能范围

2.2.0 稳定线限制演示文稿渲染开销，并增加兼容性诊断：

- 只为可见及邻近的导航条目挂载缩略图，并限制同时渲染的预览数量。
- 过期的幻灯片和缩略图任务会在提交前被协作取消。
- 取消任务、卸载缩略图、重新加载、切换文件或关闭视图后释放生成的图片资源。
- 可复制包含文档包、幻灯片、对象、资源、字体和耗时指标的隐私安全诊断，不包含
  文档内容和 vault 绝对路径。
- 维护运行时生成的兼容性样本和结构化视觉基线，覆盖表格、图表、SmartArt、
  主题、版式和母版。
- 图表和 SmartArt 仍不承诺原生外观还原，但会显示明确占位并在诊断中单独统计。

### 2.3.0 全局性能治理范围

2.3.0 稳定线统一治理 DOCX/PPTX 的首屏响应、交互延迟、内存和包体积：

- PPTX 当前页在全量元数据索引完成前开始渲染。
- 元数据解析最多四个 worker，缩略图渲染最多两个低优先级 worker，并可在
  切页或滚动时取消过期任务。
- 1,000 页演示文稿最多同时挂载 60 个导航行。
- PPTX 形状、文本和表格渲染按约 8 ms 的目标时间片主动让出主线程。
- XML、关系、幻灯片上下文和二进制缓存均有容量上限。
- DOCX 每次渲染只建立一次搜索索引，当前结果切换为常量时间，且不改写文档 DOM。
- DOCX 阅读状态按帧合并并跳过未变化值，渲染后扫描合并，适配宽度使用 CSS。
- `dist/main.js` 不超过 500 KiB，并维护 200/1,000 页 PPTX 与 100 页 DOCX
  的生成式回归样本。

### 2.4.0 共享阅读器架构范围

2.4.0 在开始 XLSX 工作前完成阅读器架构收口：

- 保留 `word-reader-view` 和 `pptx-reader-view` 标识，通过
  `OfficeReaderAdapter`、`ReaderSession` 和 `ReaderCapabilities` 分派格式行为。
- DOCX/PPTX 使用相同的 `OfficeReaderShell` 工具栏、状态、错误和诊断约定，
  同时保留旧 CSS 类。
- 设置带显式 schema 版本，按 `common`、`docx`、`pptx` 和预留的 `xlsx`
  分区；加载时迁移 2.4 之前的数据。
- 最多保存 100 条阅读状态，只包含文件标识、修改时间、格式、位置、缩放和导航。
- 文件修改时间变化时清除过期位置与导航，同时保留缩放和适配偏好。
- 各格式共用隐私安全诊断信封和摘要笔记 frontmatter 约定。

### 2.5.0 OOXML 与 XLSX 验证范围

2.5.0 完成 XLSX 发布前验证，但不开放 XLSX 视图：

- DOCX/PPTX 公开加载路径和未注册 XLSX 研究解析器统一执行 ZIP 条目、解压量、
  压缩比、加密、ZIP64、路径、宏、ActiveX、OLE 和脚本型媒体检查。
- 本地解析工作簿元数据和稀疏工作表单元格、样式、公式缓存、合并、冻结、
  隐藏表、超链接与光栅图片元数据，不重算公式，也不访问远程资源。
- 虚拟网格只保留可视窗口和有限 overscan，默认最多挂载 2,500 个单元格。
- 维护正常、损坏、加密、恶意和极端规模生成式样本，以及 100,000 行基准预算。
- 顶层运行时依赖使用精确版本，生成兼容许可证清单并审计生产 bundle 输入；
  继续保持 `2.x` 500 KiB bundle 上限和 `3.0.0` 8 MiB 发布目标。
- `3.0.0` 前不注册 `.xlsx`、XLSX 视图、公开功能开关或技术预览入口。

### 3.0.0 XLSX 只读 MVP 范围

3.0.0 将已验证的 XLSX 内核通过公开、本地、只读电子表格会话交付：

- 使用 `xlsx-reader-view` 注册 `.xlsx` 和旧版 `.xls`；`.xlsx` 本地渲染，
  `.xls` 显示外部打开和转换说明，`.xlsm` 保持不注册。
- 支持切换可见工作表；在 2,500 单元格挂载上限内，只渲染可视区域、有限
  overscan 和当前可见的冻结行列。
- 保留合并单元格、行列尺寸、基础字体/填充/边框/对齐、数字/日期格式、公式
  文本和工作簿已有缓存结果。
- 以可取消分片搜索当前工作表，支持结果跳转、缩放/适配宽度，并恢复当前表、
  滚动和缩放；不持久化单元格内容。
- 支持矩形选区，默认复制显示值；复制公式为独立操作，单次最多物化
  250,000 个单元格。
- 绝不重算公式或解析外部工作簿引用、连接和远程资源；仅在用户直接操作并确认
  后打开安全协议的超链接。
- 重新加载、切表、切文件和关闭视图时取消过期工作簿、工作表、搜索和计划网格
  任务，并清空有界包缓存。
- `dist/main.js` 继续不超过 500 KiB，三文件发布包不超过 3.0.0 的 8 MiB 目标。

### 3.1.0 工作簿导航与知识提取范围

3.1.0 在不扩大网络和代码执行边界的前提下完善本地 XLSX 只读会话：

- 通过公式栏名称框跳转静态 A1 单元格/矩形区域、带工作表限定的引用，以及受
  支持的工作簿级或工作表级命名区域；忽略动态名称、多区域并集、公式和外部
  工作簿目标。
- 以可取消的顺序解析在可见、隐藏和深度隐藏工作表中搜索显示值与已保存公式
  文本，并支持跨表结果跳转。
- 明确显示隐藏工作表数量和用户主动导航菜单，不修改工作簿可见性。
- 在只读公式栏显示完整的已保存公式文本、显示值或缓存结果，以及只使用缓存的
  安全提示。
- 将显示值复制为 TSV、把完整矩形选区导出为 Markdown 表格，并保持公式复制
  独立；全部剪贴板路径继续遵守 250,000 单元格物化上限。
- 创建或打开同名 XLSX 摘要笔记，包含工作表尺寸/可见性、命名区域、当前
  选区，以及整个工作簿最多 200 个显示值预览。
- 重新加载、切表、切文件和关闭视图时取消工作簿搜索与摘要收集，并继续清空
  有界包缓存。

### 3.2.0 XLSX 丰富内容与大工作簿范围

3.2.0 增加有界的丰富内容预览，并从大工作簿路径中移除完整工作表 XML 保留：

- 解析旧式单元格批注的作者和完整富文本，为有值及空白批注单元格显示标记，并
  在只读公式栏中显示所选批注。
- 使用单单元格或双单元格锚点渲染包内 PNG/JPEG/GIF/BMP 图片；图片字节、
  挂载绘图和可回收 Blob URL 都有固定上限，外部图片关系继续忽略。
- 仅使用工作簿缓存序列简化渲染柱状图、折线图、面积图和饼图；不支持的图表
  显示明确占位，不重算公式，也不抓取外部图表数据。
- 应用数值常量 `cellIs` 规则、双色/三色色阶和数据条；每张工作表最多解析
  256 条条件格式规则，差异样式只覆盖规则明确声明的属性。
- 分片解压和校验工作表及共享字符串 XML，逐行增量解析工作表、报告加载进度，
  并可在分片与有界行批次之间取消，不保留完整 `sheetData` XML。
- 同屏最多挂载 12 个工作表绘图，单图最多抽样渲染 48 个点；重新加载、切表、
  切文件和关闭视图时释放丰富内容资源。
- 稠密基准至少覆盖 20,000 行和 80,000 个单元格，记录解析耗时、峰值堆、
  输入分片数、最大工作表数据缓冲、取消结果和虚拟网格 DOM 预算。
- 在 `XLSM_COMPATIBILITY.md` 中评估启用宏的内容类型、VBA、宏/对话框工作表、
  ActiveX、OLE、脚本媒体、外部关系和缓存值后，继续保持 `.xlsm` 不注册。

### 支持边界

支持：

- Obsidian 桌面端 1.12.7 或更新版本。
- 本地桌面 vault。
- 基于 `docx-preview` 的 `.docx` 只读预览。
- `.pptx` 文本、内嵌图片、常见形状、表格、主题、版式和母版的本地只读预览。
- `.pptx` 缩略图/标题导航、本地全文搜索、当前页文本复制、演讲者备注查看和带页码摘要笔记。
- `.pptx` 按需缩略图挂载、协作取消渲染、生成资源清理和隐私安全渲染诊断。
- `.xlsx` 工作表切换、冻结窗格、合并单元格、行列尺寸、基础样式与格式、
  公式缓存、名称框/命名区域、工作簿级搜索、隐藏表提示、缩放、TSV/Markdown
  选区复制、旧式批注、包内光栅图片、常见图表缓存预览、数值条件格式子集、
  摘要笔记和确认后超链接的本地只读虚拟网格。
- 从已渲染 DOCX 内容复制纯文本和 Markdown。
- `.doc` 检测、外部打开和转换说明。
- `.xls` 检测、外部打开和 `.xlsx` 转换说明。

不支持：

- 编辑或保存 Word 文件。
- 移动端专项支持。
- 直接渲染旧版 `.doc` 二进制文件。
- 密码保护或加密 Word 文档。
- 密码保护或加密 PowerPoint 演示文稿。
- 密码保护或加密 Excel 工作簿。
- 直接渲染旧版 `.xls`，以及注册 `.xlsm`。
- XLSX 现代线程批注、浮动对象编辑、数据透视表、切片器、像素级图表还原和
  复杂条件格式。
- PPTX 动画、切换效果、音视频播放、宏、编辑、图表、SmartArt、
  SVG/GIF/WebP 媒体和 PowerPoint 像素级还原。
- 复杂 Word 排版与 Microsoft Word 完全一致。
- 不提供本地路径的远程 vault 适配器外部打开。
