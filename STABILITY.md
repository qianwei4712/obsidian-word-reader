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

## Manual Test Checklist

Run this checklist before publishing a stable release:

- Build and package:
  - `npm run typecheck`
  - `npm test`
  - `npm run build`
  - `npm run performance:check`
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
  - Open more than 50 distinct documents and confirm persisted reading state remains capped at 50 entries.
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
- Plain text and Markdown copy from the already rendered DOCX content.
- `.doc` detection with external-open and conversion guidance.

Not supported:

- Editing or saving Word files.
- Mobile-specific support.
- Direct rendering of legacy `.doc` binary files.
- Password-protected or encrypted Word documents.
- Password-protected or encrypted PowerPoint presentations.
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

### 支持边界

支持：

- Obsidian 桌面端 1.12.7 或更新版本。
- 本地桌面 vault。
- 基于 `docx-preview` 的 `.docx` 只读预览。
- `.pptx` 文本、内嵌图片、常见形状、表格、主题、版式和母版的本地只读预览。
- `.pptx` 缩略图/标题导航、本地全文搜索、当前页文本复制、演讲者备注查看和带页码摘要笔记。
- `.pptx` 按需缩略图挂载、协作取消渲染、生成资源清理和隐私安全渲染诊断。
- 从已渲染 DOCX 内容复制纯文本和 Markdown。
- `.doc` 检测、外部打开和转换说明。

不支持：

- 编辑或保存 Word 文件。
- 移动端专项支持。
- 直接渲染旧版 `.doc` 二进制文件。
- 密码保护或加密 Word 文档。
- 密码保护或加密 PowerPoint 演示文稿。
- PPTX 动画、切换效果、音视频播放、宏、编辑、图表、SmartArt、
  SVG/GIF/WebP 媒体和 PowerPoint 像素级还原。
- 复杂 Word 排版与 Microsoft Word 完全一致。
- 不提供本地路径的远程 vault 适配器外部打开。
