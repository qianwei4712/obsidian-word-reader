# Changelog

## 3.1.0 - 2026-07-29

### English

#### Added

- Added an editable spreadsheet name box for direct `A1`/range navigation,
  qualified worksheet references, and workbook- or sheet-scoped named ranges.
- Added workbook-wide search across visible, hidden, and very-hidden
  worksheets, with cancellable sequential parsing and cross-sheet result
  navigation.
- Added an explicit hidden-worksheet indicator and local navigation menu while
  retaining the active hidden sheet in the worksheet tab strip.
- Expanded the read-only formula bar with a standard name box, full stored
  formula text, displayed or workbook-cached results, and the existing
  cached-only safety notice.
- Added explicit displayed-value TSV copy and Markdown-table export for the
  selected rectangular range. Markdown export preserves every selected row
  under generated spreadsheet column headings.
- Added same-name XLSX summary notes containing workbook and worksheet
  dimensions, visibility, named ranges, the current selection, and bounded
  displayed-value previews from workbook content.

#### Security and performance

- Named ranges accept only local static A1 cell or rectangular range targets;
  external workbook references, formulas, unions, and dynamic names are not
  resolved.
- Workbook search and summary collection are cancellable during reload,
  worksheet/file switching, and view close. Worksheet/package caches remain
  bounded and are released with the workbook.
- Summary previews are capped at 200 populated cells across the workbook, and
  clipboard materialization retains the existing 250,000-cell ceiling.

#### Changed

- Updated XLSX search from current-sheet scope to the complete workbook and
  made result counters identify the target worksheet and cell.
- Updated version metadata, documentation, release checks, and the manual
  checklist for the 3.1.0 workbook navigation and knowledge-extraction scope.

### 中文

#### 新增

- 新增可输入的电子表格名称框，支持直接跳转 `A1`/矩形区域、带工作表限定的
  引用，以及工作簿级或工作表级命名区域。
- 新增覆盖可见、隐藏和深度隐藏工作表的工作簿级搜索；按工作表顺序解析、
  支持协作取消，并可跨表跳转结果。
- 新增明确的隐藏工作表数量提示和本地导航菜单；打开隐藏表后仍会在工作表
  标签栏保留当前表。
- 完善只读公式栏，包含标准名称框、完整的已保存公式文本、显示值或工作簿缓存
  结果，以及原有的“只使用缓存、不重算”安全提示。
- 新增明确的显示值 TSV 复制和选中矩形区域 Markdown 表格导出；Markdown
  使用生成的列标题并保留全部选中数据行。
- 新增同名 XLSX 摘要笔记，包含工作簿/工作表尺寸、可见性、命名区域、当前
  选区，以及从工作簿内容生成的有界显示值预览。

#### 安全与性能

- 命名区域只解析包内静态 A1 单元格或矩形区域；不解析外部工作簿引用、公式、
  多区域并集和动态名称。
- 工作簿搜索和摘要收集会在重新加载、切表、切文件和关闭视图时取消；工作表
  与包缓存继续保持有界，并随工作簿释放。
- 摘要预览在整个工作簿内最多收集 200 个非空单元格；剪贴板物化继续遵守
  250,000 单元格上限。

#### 变更

- XLSX 搜索从当前工作表升级为整个工作簿，结果计数会显示目标工作表与单元格。
- 更新 3.1.0 工作簿导航与知识提取范围对应的版本元数据、文档、发布检查和
  手动验收清单。

## 3.0.0 - 2026-07-29

### English

#### Added

- Added the public, local-only `xlsx-reader-view` for `.xlsx` files, with a
  conversion guidance page and external-open action for legacy `.xls` files.
  Macro-enabled `.xlsm` remains unregistered.
- Added visible worksheet switching and a read-only virtual grid that mounts
  only the viewport, finite overscan, and visible frozen rows/columns under a
  2,500-cell ceiling.
- Added rendering for merged cells, custom row/column dimensions, basic font,
  fill, border and alignment styles, number/date formats, stored formula text,
  and workbook-cached formula results.
- Added cancellable current-worksheet search, result navigation, continuous
  zoom, fit width, rectangular mouse/keyboard selection, displayed-value TSV
  copy, and a separate formula-copy action.
- Added a selected-cell details bar and confirmed workbook hyperlinks.
  Internal cell locations can navigate within the workbook; only `http`,
  `https`, and `mailto` external targets are permitted after confirmation.
- Added per-file restoration of the active worksheet, scroll position, zoom,
  and fit preference without persisting cell values or formulas.
- Added localized XLSX status, error, legacy-format, copy, search, and safety
  messages plus an XLSX default-fit-width setting.

#### Security

- XLSX formulas remain cached-only and are never recalculated. External
  workbook references, data connections, remote resources, macros, ActiveX,
  OLE objects, and script-capable media are never executed or fetched.
- Workbook, worksheet, search, animation-frame, and state-write work is
  cancelled during reload, sheet/file switching, and view close. Bounded
  package caches are cleared when the workbook is released.
- Clipboard materialization is capped at 250,000 selected cells, while the
  sparse 100,000-row benchmark continues to enforce bounded DOM and memory.

#### Changed

- Updated the supported product boundary, manifest description, release
  metadata, documentation, manual checklist, and production-bundle audit for
  `.xlsx`.

### 中文

#### 新增

- 新增公开且完全本地的 `.xlsx` 只读视图 `xlsx-reader-view`；旧版 `.xls`
  显示转换说明和外部打开操作，含宏 `.xlsm` 仍不注册。
- 新增可见工作表切换和只读虚拟网格；在 2,500 单元格上限内只挂载可视区域、
  有限 overscan 和当前可见的冻结行列。
- 新增合并单元格、自定义行列尺寸、基础字体/填充/边框/对齐、数字/日期格式、
  已保存公式文本和工作簿已有公式缓存结果的显示。
- 新增可取消的当前工作表搜索、结果跳转、连续缩放、适配宽度、鼠标/键盘矩形
  选区、显示值 TSV 复制和独立的公式复制操作。
- 新增当前单元格详情栏和需确认的工作簿超链接；内部位置可以在工作簿中跳转，
  外部目标仅允许用户确认后的 `http`、`https` 和 `mailto`。
- 新增按文件恢复当前工作表、滚动位置、缩放和适配偏好；不会持久化单元格值或
  公式。
- 新增本地化 XLSX 状态、错误、旧格式、复制、搜索和安全提示，以及电子表格
  默认适配宽度设置。

#### 安全

- XLSX 公式继续只使用缓存，绝不重新计算；外部工作簿引用、数据连接、远程
  资源、宏、ActiveX、OLE 对象和脚本型媒体绝不执行或自动抓取。
- 重新加载、切表、切文件和关闭视图时取消工作簿、工作表、搜索、动画帧和状态
  写入任务；释放工作簿时清空有界包缓存。
- 剪贴板单次最多物化 250,000 个选中单元格，100,000 行稀疏表基准继续约束
  DOM 和内存。

#### 变更

- 更新产品支持边界、manifest 描述、发布元数据、文档、手动清单和生产 bundle
  审计以正式包含 `.xlsx`。

## 2.5.0 - 2026-07-28

### English

#### Added

- Added a shared OOXML ZIP safety layer with immutable global ceilings of
  10,000 entries, 128 MiB per entry, 512 MiB total expansion, and a 200:1
  compression ratio. PPTX retains its stricter 2,000-entry, 64 MiB per-entry,
  and 256 MiB total defaults.
- Added an unregistered XLSX validation core for workbook and sparse worksheet
  parsing, shared/inline strings, merged cells, frozen panes, hidden sheets,
  row and column sizes, basic cell styles, date and number formatting, cached
  formulas, hyperlink metadata, and safe raster image parts.
- Added a bounded virtual-grid model and a generated 100,000-row sparse
  workbook benchmark that records first-window time, p95 scroll calculation,
  sampled peak heap growth, estimated DOM nodes, and cancellation behavior.
- Added a generated XLSX fixture catalog covering valid, damaged, encrypted,
  malicious-compression, macro, OLE, script-media, external-relationship, data
  connection, and extreme-scale cases.
- Added a generated third-party runtime license inventory plus exact direct
  runtime version, esbuild input, 500 KiB bundle, and 8 MiB release-target
  audits.

#### Security

- OOXML packages now reject encrypted ZIP entries, unsafe paths, duplicate or
  malformed directory entries, ZIP64 envelopes, macros, ActiveX, embedded OLE
  objects, and script-capable media before content rendering.
- XLSX formulas are never recalculated: research parsing exposes only formula
  text and the cached value already stored in the workbook. External
  relationships and data connections are recorded as ignored metadata and are
  never dereferenced.

#### Changed

- Pinned all top-level runtime registry dependencies to exact versions.
- Kept XLSX research modules out of the production bundle and did not register
  an `.xlsx` extension, public view, feature flag, or technical-preview entry.

### 中文

#### 新增

- 新增统一 OOXML ZIP 安全层，全局硬上限固定为 10,000 个条目、单项
  128 MiB、总解压 512 MiB 和 200:1 压缩比；PPTX 继续使用更严格的
  2,000 个条目、单项 64 MiB 和总计 256 MiB 默认值。
- 新增未注册的 XLSX 验证内核，覆盖工作簿与稀疏工作表解析、共享/内联字符串、
  合并单元格、冻结窗格、隐藏工作表、行列尺寸、基础单元格样式、日期与数字
  格式、公式缓存、超链接元数据和安全光栅图片部件。
- 新增有界虚拟网格模型和生成式 100,000 行稀疏工作簿基准，记录首屏耗时、
  滚动计算 p95、采样峰值堆增量、估算 DOM 节点数和取消加载结果。
- 新增生成式 XLSX 样本目录，覆盖正常、损坏、加密、恶意压缩、宏、OLE、
  脚本型媒体、外部关系、数据连接和极端规模场景。
- 新增第三方运行时许可证清单，以及顶层运行时精确版本、esbuild 输入、
  500 KiB bundle 和 8 MiB 发布目标审计。

#### 安全

- OOXML 包在渲染前拒绝加密 ZIP 条目、不安全路径、重复或损坏的目录项、
  ZIP64 包络、宏、ActiveX、嵌入式 OLE 对象和脚本型媒体。
- XLSX 公式绝不重算；研究解析只暴露公式文本和工作簿已有缓存值。外部关系与
  数据连接仅记录为忽略项，绝不自动访问。

#### 变更

- 所有顶层运行时 registry 依赖改为精确版本。
- XLSX 研究模块不进入生产 bundle，且未注册 `.xlsx` 扩展名、公开视图、
  功能开关或技术预览入口。

## 2.4.0 - 2026-07-27

### English

#### Added

- Added the shared `OfficeReaderShell`, `OfficeReaderAdapter`,
  `ReaderSession`, and `ReaderCapabilities` contracts. Toolbar commands and
  controls now use declared format capabilities.
- Added versioned settings with `common`, `docx`, `pptx`, and reserved `xlsx`
  sections, plus automatic migration from the pre-2.4 flat data shape.
- Added `office-reader-*` shared CSS classes while retaining the existing
  `word-reader-*` and `pptx-reader-*` compatibility classes.

#### Changed

- Moved DOCX and PPTX rendering, navigation, search, copy, state, and cleanup
  behavior into format-specific sessions and adapters. The legacy view modules
  now remain as small Obsidian registration entry points.
- Expanded the bounded reading-state LRU from 50 to 100 Office files. The new
  persisted shape records only file identity, modification time, format,
  position, zoom, and navigation state; legacy entries migrate without losing
  their first restored position.
- File modification now invalidates stale position and navigation state while
  retaining the user's zoom and fit preference.
- Unified DOCX and PPTX diagnostic JSON under a versioned, privacy-safe
  envelope and removed raw PPTX error messages from the error details panel.
- Unified DOCX and PPTX summary-note creation and frontmatter conventions with
  `reader: office-reader` and `reader_format` metadata.

### 中文

#### 新增

- 新增共享 `OfficeReaderShell`、`OfficeReaderAdapter`、`ReaderSession` 和
  `ReaderCapabilities` 契约；工具栏命令和控件改为按格式声明的能力启用。
- 新增带版本号的设置结构，按 `common`、`docx`、`pptx` 和预留的 `xlsx`
  分区，并自动迁移 2.4 之前的扁平数据。
- 新增 `office-reader-*` 通用 CSS 类，同时保留现有 `word-reader-*` 和
  `pptx-reader-*` 兼容类。

#### 变更

- DOCX/PPTX 的渲染、导航、搜索、复制、状态和资源清理逻辑迁入各自的会话与
  适配器；旧视图模块只保留精简的 Obsidian 注册入口。
- 阅读状态 LRU 从 50 条扩展为 100 条 Office 文件记录。新结构只保存文件
  标识、修改时间、格式、位置、缩放和导航状态；旧记录首次恢复时不会丢失位置。
- 文件修改后会清除已失效的位置和导航状态，同时保留用户的缩放与适配偏好。
- DOCX/PPTX 诊断统一为带版本号的隐私安全 JSON 信封，并从 PPTX 错误详情中
  移除原始错误消息。
- DOCX/PPTX 摘要笔记统一创建流程和 frontmatter 约定，新增
  `reader: office-reader` 与 `reader_format` 元数据。

## 2.3.2 - 2026-07-26

### English

#### Fixed

- DOCX outlines now use Word paragraph outline levels and inherited paragraph
  styles, including localized or custom heading styles and heading levels 1
  through 9, instead of relying only on rendered HTML class names.

### 中文

#### 修复

- DOCX 大纲改为识别 Word 段落大纲级别和段落样式继承，支持本地化或自定义
  标题样式以及 1 至 9 级标题，不再只依赖渲染后的 HTML 类名。

## 2.3.1 - 2026-07-26

### English

#### Fixed

- Removed the local lint-disable comment from the Obsidian 1.12.7 settings
  compatibility refresh path so official plugin checks can pass.

### 中文

#### 修复

- 移除 Obsidian 1.12.7 设置页兼容刷新路径中的本地 lint 禁用注释，确保官方插件
  检查可以通过。

## 2.3.0 - 2026-06-14

### English

#### Added

- Added priority-aware PPTX metadata and thumbnail queues, a virtualized slide
  navigation window, time-sliced rendering, and bounded XML, context, binary,
  and media caches.
- Added generated performance regression fixtures for 200- and 1,000-slide
  presentations and a 100-page Word document.
- Added a production bundle budget that fails builds above 500 KiB.
- Restored compatibility with the current Obsidian Desktop stable line by
  setting the minimum app version to 1.12.7 and keeping a classic settings tab
  fallback.

#### Changed

- PPTX now renders the current slide before full-presentation metadata
  indexing finishes, while background metadata parsing is cancellable and
  limited to four workers.
- DOCX search now builds one reusable text index and uses CSS highlights
  without rewriting the rendered document DOM.
- DOCX reading-state writes are coalesced and unchanged states are skipped;
  fit-width changes now use CSS without rerendering the document.
- Removed Mammoth from the runtime and dependency tree. Whole-document text
  and Markdown copy reuse the already rendered document.

### 中文

#### 新增

- 新增带优先级的 PPTX 元数据与缩略图任务队列、虚拟化幻灯片导航窗口、分时
  渲染，以及容量有界的 XML、上下文、二进制和媒体缓存。
- 新增 200/1,000 页 PPTX 与 100 页 DOCX 的运行时生成性能回归样本。
- 新增生产 bundle 体积预算，超过 500 KiB 时构建失败。

#### 变更

- PPTX 会在全量元数据索引完成前优先渲染当前幻灯片；后台元数据解析可取消，
  且最多使用四个并发 worker。
- DOCX 搜索改为每次渲染只建立一次可复用文本索引，并使用 CSS Highlight，
  不再重写已渲染文档 DOM。
- DOCX 阅读状态写入按动画帧合并且跳过未变化状态；适配宽度改为 CSS 切换，
  不再重新渲染文档。
- 从运行时和依赖树中移除 Mammoth，整篇纯文本和 Markdown 复制复用已渲染文档。
- 将最低 Obsidian 版本调整为 1.12.7，并保留经典设置页回退，以兼容当前
  Obsidian Desktop 稳定版。

## 2.2.0 - 2026-06-14

### English

#### Added

- Added viewport-driven, bounded thumbnail mounting so large presentations
  render only visible and nearby slide previews.
- Added a privacy-safe render diagnostics action and command with slide,
  package, object, resource, font, and timing metrics.
- Added generated compatibility fixtures and structural visual regression
  coverage for tables, charts, SmartArt, themes, layouts, and masters.

#### Changed

- PPTX slide and thumbnail rendering now cooperatively cancels stale work and
  releases generated image resources after cancellation, unmount, reload,
  file switch, or view close.
- Chart and SmartArt frames now render explicit local placeholders and appear
  as separate diagnostic counters in addition to the overall placeholder
  count.
- Development builds now log PPTX render performance and compatibility
  counters to the developer console.

### 中文

#### 新增

- 新增由可视区域驱动且数量有界的缩略图挂载机制，大型演示文稿只渲染当前可见
  及邻近的幻灯片预览。
- 新增注重隐私的渲染诊断操作和命令，包含幻灯片、文档包、对象、资源、字体与
  耗时指标。
- 新增运行时生成的兼容性样本和结构化视觉回归测试，覆盖表格、图表、SmartArt、
  主题、版式和母版。

#### 变更

- PPTX 幻灯片和缩略图渲染现在会协作取消过期任务，并在取消、卸载、重新加载、
  切换文件或关闭视图后释放生成的图片资源。
- 图表和 SmartArt 框架现在显示明确的本地占位，并在总占位数量之外提供独立的
  诊断计数。
- 开发构建会在开发者控制台记录 PPTX 渲染性能和兼容性计数。

## 2.1.0 - 2026-06-13

### English

#### Added

- Added a presentation sidebar with rendered slide thumbnails, extracted slide
  titles, active-slide tracking, and direct slide navigation.
- Added presentation-wide search across slide text, tables, titles, and speaker
  notes, with matching snippets and note-only result indicators.
- Added current-slide text copy, speaker-note viewing, and presentation summary
  notes that record the current slide and include numbered references for every
  slide.
- Added `Space`/`Shift` + `Space`, `Home`, `End`, and `Ctrl`/`Cmd` + `F`
  presentation shortcuts alongside the existing arrow and page-key navigation.
- Added PPTX metadata, speaker-note, search, summary-note, and reading-state
  test coverage.

#### Changed

- Restored slide navigation and speaker-note visibility per presentation.
- Released thumbnail image resources when presentations are reloaded, switched,
  or closed.

### 中文

#### 新增

- 新增演示文稿侧边栏，提供真实幻灯片缩略图、标题提取、当前页高亮和直接跳转。
- 新增覆盖幻灯片标题、正文、表格和演讲者备注的全文搜索，并显示匹配摘要及仅备注命中提示。
- 新增当前幻灯片文本复制、演讲者备注查看，以及记录当前页并列出全部幻灯片页码引用的演示文稿摘要笔记。
- 在原有方向键和翻页键之外，新增 `Space`/`Shift` + `Space`、`Home`、
  `End` 和 `Ctrl`/`Cmd` + `F` 演示文稿快捷键。
- 新增 PPTX 元数据、演讲者备注、搜索、摘要笔记和阅读状态测试。

#### 变更

- 按演示文稿恢复幻灯片导航和演讲者备注的显示状态。
- 重新加载、切换或关闭演示文稿时释放缩略图图片资源。

## 2.0.2 - 2026-06-13

### English

#### Changed

- Renamed the community-directory display name from Word Reader to Office
  Reader while retaining the stable `word-reader` plugin ID for update
  compatibility.
- Updated English and Chinese documentation, manual installation paths,
  diagnostics, logs, and build metadata to use the new display name.

### 中文

#### 变更

- 社区插件目录显示名称由 Word Reader 更新为 Office Reader，同时保留稳定的
  `word-reader` 插件 ID，确保现有安装可以继续升级。
- 更新中英文说明文档、手动安装目录、诊断信息、日志和构建元数据，使其统一使用
  新显示名称。

## 2.0.1 - 2026-06-13

### English

#### Fixed

- Fixed external opening in Obsidian Desktop by loading Electron through the
  guarded CommonJS plugin runtime instead of a browser dynamic import.
- Applied the same runtime fix to image copy, image save, and Node.js file
  access, with a production-bundle regression check.

### 中文

#### 修复

- 修复 Obsidian 桌面端外部打开失败的问题，Electron 改为通过带平台守卫的
  CommonJS 插件运行时加载，不再使用浏览器动态导入。
- 同步修复图片复制、图片另存和 Node.js 文件访问，并新增生产包回归检查。

## 2.0.0 - 2026-06-13

### English

#### Added

- Added local, read-only `.pptx` preview with text, images, common shapes,
  tables, theme colors, slide layouts, and slide masters.
- Added previous/next navigation, page-number jump, continuous zoom, fit to
  window, fullscreen reading, external open, and per-file page restoration.
- Added pre-decompression ZIP safety limits for entry count, per-entry expanded
  size, total expanded size, ZIP64, encryption, and abnormal compression ratios.
- Added generated PPTX package and DOM rendering integration tests.

#### Security

- PPTX rendering ignores external relationships, rejects script-capable SVG
  media, and never loads remote resources.
- Desktop Electron and Node.js integrations use guarded CommonJS loading so
  external open, image copy, and image save work in Obsidian's plugin runtime.
- Animation, audio/video playback, macros, editing, and source-file write-back
  are outside the supported preview scope.
- Rapid file, page, and zoom changes invalidate stale rendering work and
  release generated Blob URLs.

### 中文

#### 新增

- 新增本地只读 `.pptx` 预览，支持文本、图片、常见形状、表格、主题颜色、
  幻灯片版式和母版。
- 新增上一张/下一张、页码跳转、连续缩放、适应窗口、全屏阅读、外部打开，
  并按文件恢复阅读页码。
- 新增解压前 ZIP 安全限制，覆盖条目数量、单条目解压大小、总解压大小、
  ZIP64、加密和异常压缩率。
- 新增运行时生成 PPTX 包及 DOM 渲染的集成测试。

#### 安全

- PPTX 渲染会忽略外部关系、拒绝可包含脚本或远程引用的 SVG 媒体，并且
  不加载任何远程资源。
- 桌面端 Electron 与 Node.js 集成改用带平台守卫的 CommonJS 加载，确保
  外部打开、图片复制和图片另存可在 Obsidian 插件运行时正常工作。
- 不支持动画、音视频播放、宏、编辑和源文件回写。
- 快速切换文件、页码或缩放时会作废过期渲染任务，并释放生成的 Blob URL。

## 1.6.0 - 2026-06-12

### English

#### Added

- Added automatic current-section highlighting in the document outline.
- Added collapsible outline sections with restored collapse state per file.
- Added reading-state restoration for zoom, fit-width mode, outline visibility,
  collapsed sections, and scroll position across the 50 most recently used
  Word documents.
- Added unit coverage for shared reader lifecycle, status, zoom, outline,
  resource, and bounded LRU state modules.

#### Changed

- Extracted reusable reader primitives for file lifecycle cancellation, loading
  status, privacy-safe diagnostics, zoom behavior, external opening, and
  retained resource cleanup as the foundation for later Office formats.
- Word error diagnostics and Blob URL cleanup now use the shared reader
  infrastructure while preserving existing DOCX behavior.

### 中文

#### 新增

- 大纲会根据当前滚动位置自动高亮正在阅读的章节。
- 大纲章节支持折叠，并按文件恢复折叠状态。
- 为最近使用的 50 个 Word 文档恢复缩放、适配宽度、大纲显示、折叠章节和
  滚动位置等阅读状态。
- 增加公共阅读生命周期、状态、缩放、大纲、资源和有界 LRU 状态模块的单元测试。

#### 变更

- 抽取文件生命周期取消、加载状态、隐私安全诊断、缩放、外部打开和资源释放等
  公共阅读能力，为后续 Office 格式提供基础。
- Word 错误诊断和 Blob URL 清理改用公共阅读基础设施，同时保持现有 DOCX 行为兼容。

## 1.5.0 - 2026-06-07

### English

#### Added

- Added Node.js-based unit tests for settings normalization, localization,
  summary note generation, Word error classification, and privacy-safe
  diagnostics.
- Added a generated minimal DOCX integration test covering plain-text and
  Markdown extraction.
- Added an AST-based security scan for dynamic code, script/HTML injection,
  remote execution, and undisclosed network request primitives, plus
  production bundle checks.
- Added `versions.json` and a community plugin submission checklist based on
  the current Obsidian release requirements.

#### Changed

- CI and release workflows now require tests and the security scan to pass.
- Release validation now checks community submission metadata, manifest ID and
  description rules, and `versions.json` compatibility.
- Extracted settings, summary-note, and Word-error logic into independently
  testable production modules.

### 中文

#### 新增

- 增加基于 Node.js 的单元测试，覆盖设置归一化、界面语言、摘要笔记生成、
  Word 错误分类和隐私安全诊断信息。
- 增加运行时生成最小 DOCX 的集成测试，验证纯文本和 Markdown 提取。
- 增加基于 AST 的安全扫描，检查动态代码、脚本或 HTML 注入、远程执行和
  未披露网络请求风险，并检查生产构建产物。
- 增加 `versions.json` 和依据当前 Obsidian 发布要求整理的社区插件提交清单。

#### 变更

- CI 和发布流程现在必须通过自动化测试与安全扫描。
- 发布校验新增社区提交元数据、manifest ID、描述规则和 `versions.json`
  兼容性检查。
- 将设置、摘要笔记和 Word 错误处理逻辑拆分为可独立测试的生产模块。

## 1.4.1 - 2026-06-07

### English

#### Added

- Added the official `eslint-plugin-obsidianmd` recommended rules, zero-warning local checks, and CI/release lint gates.
- Added repository-level development rules documenting required Obsidian compatibility and release checks.

#### Changed

- Migrated plugin settings to Obsidian's declarative `getSettingDefinitions()` API and raised the minimum supported Obsidian version to 1.13.0.
- Simplified command IDs so Obsidian can add the plugin namespace without duplicated prefixes.

#### Fixed

- Made document creation and DOM type checks safe across Obsidian popout windows.
- Replaced direct `fetch` and CommonJS `require()` usage with local resource loading and guarded desktop-only module imports.
- Removed an unsafe settings argument and other automated review warnings.

### 中文

#### 新增

- 接入官方 `eslint-plugin-obsidianmd` 推荐规则、零告警本地检查，以及 CI 和发布流程中的 lint 门禁。
- 增加仓库级开发规则，明确 Obsidian 兼容性要求和发布前必需检查。

#### 变更

- 设置页迁移到 Obsidian 声明式 `getSettingDefinitions()` API，并将最低支持版本提升到 Obsidian 1.13.0。
- 简化命令 ID，由 Obsidian 自动添加插件命名空间，避免重复前缀。

#### 修复

- 修复文档创建与 DOM 类型检查在 Obsidian 弹出窗口中的跨窗口兼容性。
- 使用本地资源读取和受桌面端保护的模块导入，替代直接使用 `fetch` 和 CommonJS `require()`。
- 移除设置读取中的不安全参数以及其他自动审查告警。

## 1.4.0 - 2026-06-07

### English

#### Added

- Added clearer reading, rendering, preview preparation, and navigation-building status feedback for large and long Word documents.
- Added development-build render timing diagnostics for file reading, document rendering, DOM commit, outline creation, and total preview time.

#### Changed

- Long documents now commit rendered pages, build outlines, and apply search highlights in cancellable chunks to reduce main-thread stalls.
- Image-heavy documents now use lazy, asynchronously decoded Blob URL images, with generated image and font resources released when previews change or close.
- Long previews defer off-screen page painting to reduce rendering work during normal reading.

### 中文

#### 新增

- 为大型和长篇 Word 文档增加更清晰的读取、渲染、预览准备和导航生成状态反馈。
- 在开发构建中增加渲染耗时诊断，覆盖文件读取、文档渲染、DOM 提交、大纲生成和预览总耗时。

#### 变更

- 长文档现在会分片提交渲染页面、生成大纲和处理搜索高亮，并支持取消过期任务，以减少主线程长时间阻塞。
- 图片密集文档改用懒加载、异步解码的 Blob URL 图片，并在预览切换或关闭时释放生成的图片和字体资源。
- 长文档预览会延迟绘制屏幕外页面，减少正常阅读过程中的渲染开销。

## 1.3.3 - 2026-06-07

### English

#### Fixed

- Removed the redundant plugin-name heading from the settings page to follow Obsidian's settings UI guidelines and pass automated review.

### 中文

#### 修复

- 移除设置页面中重复的插件名称标题，以遵循 Obsidian 设置界面规范并通过自动审核。

## 1.3.2 - 2026-06-07

### English

#### Fixed

- Replaced the settings page HTML heading with Obsidian's native `Setting.setHeading()` API for consistent UI and automated review compliance.

### 中文

#### 修复

- 将设置页面的 HTML 标题替换为 Obsidian 原生 `Setting.setHeading()` API，以保持界面一致并通过自动审核。

## 1.3.1 - 2026-06-06

### English

#### Fixed

- Removed `!important` declarations from the Word preview theme overrides and used more specific selectors to satisfy Obsidian's automated CSS review.

### 中文

#### 修复

- 移除 Word 预览主题覆盖样式中的 `!important`，改用更具体的选择器，以通过 Obsidian CSS 自动审核。

## 1.3.0 - 2026-06-06

### English

#### Added

- Added structured error categories for encrypted, format-mismatched, damaged ZIP, invalid XML, unsupported structure, and unknown Word document failures.
- Added collapsible, privacy-conscious diagnostic details with file metadata, a safe category summary, and an error fingerprint.
- Added one-click diagnostic copying for issue reports without including document content or absolute vault paths.

### 中文

#### 新增

- 为加密、格式不匹配、ZIP 文档包损坏、XML 结构无效、不支持的结构和未知 Word 文档错误增加结构化分类。
- 增加默认收起且注重隐私的诊断详情，包含文件元数据、隐私安全的分类摘要和错误指纹。
- 增加一键复制诊断信息，便于提交问题，同时不包含文档正文或 vault 绝对路径。

## 1.2.1 - 2026-06-06

### English

#### Docs

- Updated README theme and image preview documentation to match the current Obsidian theme-following behavior.
- Added manual test coverage for light/dark theme rendering, image preview shortcuts, and image save naming.

### 中文

#### 文档

- 更新 README 中的主题和图片预览说明，使其与当前跟随 Obsidian 主题的行为一致。
- 增加深浅色主题渲染、图片预览快捷键和图片另存命名的手动测试覆盖。

## 1.2.0 - 2026-06-05

### English

#### Added

- Added theme-following document and image preview surfaces so Word previews, fallback messages, and image modals blend better with Obsidian dark and light themes.
- Added `Ctrl`/`Cmd` + `C` image copying while the image preview modal is open.
- Improved image save defaults with source-document-based names that include image dimensions when available.

### 中文

#### 新增

- 增加跟随 Obsidian 主题的文档和图片预览界面，让 Word 预览、兜底提示和图片弹窗更自然地适配深浅色主题。
- 增加图片预览弹窗中的 `Ctrl`/`Cmd` + `C` 快捷复制图片。
- 优化图片另存默认文件名，基于源文档名命名，并在可用时包含图片尺寸。

## 1.1.6 - 2026-06-05

### English

#### Added

- Added actionable recovery tips for encrypted, damaged, and renderer-incompatible Word documents so users can fix common `.docx` failures more easily.

### 中文

#### 新增

- 为加密、损坏以及渲染器不兼容的 Word 文档增加可操作的修复建议，帮助用户更容易处理常见 `.docx` 打开失败问题。

## 1.1.5 - 2026-05-31

### English

#### Fixed

- Fixed dynamic script injection by forcing esbuild to compile `jszip` from source instead of using its pre-compiled dist file, completely removing `createElement("script")` from the plugin build.

### 中文

#### 修复

- 通过强制 esbuild 从源码编译 `jszip` 而非使用其预编译产物，彻底清除了构建产物中的 `createElement("script")`，修复了动态脚本注入的问题。

## 1.1.4 - 2026-05-30

### English

#### Added

- Added explicit security declarations in README and code comments to clarify that no dynamic script injection is performed, addressing Obsidian marketplace security scan requirements.

### 中文

#### 新增

- 在 README 和代码注释中添加了明确的安全声明，澄清插件未执行任何动态脚本注入，以回应 Obsidian 上架安全扫描的要求。

## 1.1.3 - 2026-05-30

### English

#### Fixed

- Replaced transitive `immediate` and `setimmediate` packages with local safe shims so release builds no longer inherit dynamic `<script>` creation patterns flagged by marketplace security scans.

### 中文

#### 修复

- 用本地安全 shim 替换传递依赖中的 `immediate` 和 `setimmediate`，使发布产物不再继承会被上架安全扫描标记的动态 `<script>` 创建模式。

## 1.1.2 - 2026-05-30

### English

#### Fixed

- Fixed rendered `.docx` content selection by explicitly enabling text selection in the preview area while keeping images non-selectable for image preview interactions.

### 中文

#### 修复

- 修复 `.docx` 渲染预览区文本无法稳定选中的问题，显式启用正文文本选择，同时保持图片不可选中以兼容图片预览交互。

## 1.1.1 - 2026-05-30

### English

#### Fixed

- Fixed the GitHub release workflow version export step so tag-triggered releases run correctly on GitHub Actions.
- Fixed GitHub release automation to publish Obsidian-compatible tags without a `v` prefix and upload `main.js`, `manifest.json`, and `styles.css` with the release.

### 中文

#### 修复

- 修复 GitHub Release workflow 中的版本导出步骤，确保 tag 触发的发布流程能在 GitHub Actions 上正常运行。
- 修复 GitHub Release 自动发布流程，改为发布不带 `v` 前缀且兼容 Obsidian 的 tag，并随 release 上传 `main.js`、`manifest.json` 和 `styles.css`。

## 1.1.0 - 2026-05-29

### English

#### Added

- Added Markdown copy for selected rendered content and whole `.docx` documents.
- Added a clickable outline panel based on rendered document headings.
- Added full Chinese and English interface text for toolbar labels, notices, statuses, errors, image preview, settings, commands, and summary note templates.
- Added a setting for showing the outline by default.

#### Changed

- Summary note templates now follow the selected plugin language.

### 中文

#### 新增

- 增加选中渲染内容和整篇 `.docx` 文档的 Markdown 复制能力。
- 增加基于渲染标题的可点击大纲面板。
- 为工具栏、通知、状态、错误、图片预览、设置、命令和摘要笔记模板补齐中文与英文界面文本。
- 增加默认显示大纲的设置项。

#### 变更

- 摘要笔记模板现在会跟随插件选择的界面语言。

## 1.0.0 - 2026-05-29

### English

#### Added

- Added a stable release checklist, support boundaries, and long-term maintenance strategy.

#### Changed

- Froze the core reader scope for the stable line.

### 中文

#### 新增

- 增加稳定版发布检查清单、支持边界和长期维护策略。

#### 变更

- 冻结稳定线的核心阅读器范围。

## 0.9.0 - 2026-05-29

### English

#### Added

- Added release validation scripts for version consistency, changelog coverage, and build output checks.
- Added dependency-free release zip packaging for Obsidian manual installation.
- Added GitHub Actions release automation for version tags.

#### Changed

- Standardized release artifacts as `release/obsidian-word-reader-X.Y.Z.zip`.
- Updated package, manifest, and lockfile versions to `0.9.0`.

#### Docs

- Expanded manual installation, source build, local release, and GitHub auto-release instructions.
- Normalized the bilingual changelog structure for release notes.

### 中文

#### 新增

- 增加发布校验脚本，用于检查版本一致性、changelog 覆盖和构建产物。
- 增加无额外依赖的发布 zip 打包能力，便于 Obsidian 手动安装。
- 增加基于 GitHub tag 的自动发布 workflow。

#### 变更

- 将发布产物统一为 `release/obsidian-word-reader-X.Y.Z.zip`。
- 将 package、manifest 和 lockfile 版本更新为 `0.9.0`。

#### 文档

- 补充手动安装、源码构建、本地发布和 GitHub 自动发布说明。
- 规范中英文双语 changelog 结构，便于生成 release notes。

## 0.8.0 - 2026-05-29

### English

#### Added

- Added clearer reading and rendering status messages for large documents.

#### Changed

- Improved render cancellation by rendering into a temporary buffer and only replacing the visible preview when the latest render is still current.
- Avoided duplicate rendering when the same file state and layout mode are already displayed.
- Debounced search highlighting to reduce repeated work while typing in large documents.
- Improved unload cleanup for render state, search timers, search matches, and buffered document data.

### 中文

#### 新增

- 为大文档增加更清晰的读取和渲染状态提示。

#### 变更

- 优化渲染取消机制：先渲染到临时缓冲区，只有最新渲染仍然有效时才替换可见预览。
- 避免同一文件状态和同一布局模式已经显示时重复渲染。
- 对搜索高亮进行防抖处理，减少在大文档中输入搜索词时的重复计算。
- 改进卸载清理，释放渲染状态、搜索定时器、搜索结果和文档缓冲数据。

## 0.7.0 - 2026-05-29

### English

#### Added

- Added legacy `.doc` file registration with an in-view guidance page.
- Added external-open and `.docx` conversion guidance for `.doc` files.
- Added dedicated messages for encrypted or password-protected Word documents.
- Added dedicated messages for damaged or invalid Word documents.

#### Changed

- Large file warnings now include the file size and use the configured threshold.
- Generic render failures now include clearer next-step guidance and an external-open action.

### 中文

#### 新增

- 增加旧版 `.doc` 文件注册，并在视图中显示说明页面。
- 为 `.doc` 文件增加外部打开和转换为 `.docx` 的说明。
- 为加密或受密码保护的 Word 文档增加专门提示。
- 为损坏或无效的 Word 文档增加专门提示。

#### 变更

- 大文件提醒现在会显示文件大小，并使用设置中的提醒阈值。
- 通用渲染失败现在会显示更清晰的下一步建议，并提供外部打开操作。

## 0.6.0 - 2026-05-29

### English

#### Added

- Added mouse wheel zoom inside the image preview modal.
- Added drag-to-pan support for previewed images.
- Added double-click reset to fit the image back into the preview modal.
- Added actual-size viewing for previewed images.
- Added copy image and save image actions.
- Added original image dimensions and current zoom display.
- Added Chinese and English language switching for the settings page.

### 中文

#### 新增

- 支持在图片预览弹窗中使用鼠标滚轮缩放。
- 支持拖拽平移查看预览图片。
- 支持双击将图片重新适配到预览窗口。
- 支持查看图片原始尺寸。
- 支持复制图片和另存图片。
- 支持显示图片原始尺寸和当前缩放比例。
- 支持设置页面在中文和英文之间切换。

## 0.5.0 - 2026-05-28

### English

#### Added

- Added previous and next controls for rendered text search results.
- Added `Enter` and `Shift` + `Enter` keyboard navigation in the search input.
- Added current search result highlighting.

#### Changed

- Search result counts now show the current result and total result count.
- Search navigation now scrolls the active result into view.

### 中文

#### 新增

- 增加渲染文本搜索结果的上一个和下一个控制按钮。
- 支持在搜索框中使用 `Enter` 和 `Shift` + `Enter` 键盘跳转。
- 增加当前搜索结果高亮。

#### 变更

- 搜索结果计数现在显示当前结果序号和总结果数。
- 搜索跳转现在会自动滚动到当前结果。

## 0.4.0 - 2026-05-28

### English

#### Added

- Added an Obsidian settings tab.
- Added persisted settings for default zoom, default fit width, image preview, and large file warning size.
- Added an external opening note explaining that `.docx` files use the operating system default application.

#### Changed

- Word previews now use the configured default zoom and fit-width behavior when opened.
- Large file warnings now use the configured size threshold.
- Image click preview can now be disabled from settings.

### 中文

#### 新增

- 增加 Obsidian 插件设置页。
- 增加默认缩放、默认适配宽度、图片预览和大文件提醒阈值的持久化设置。
- 增加外部打开说明，解释 `.docx` 文件会使用操作系统默认应用。

#### 变更

- Word 预览打开时会使用配置的默认缩放和适配宽度行为。
- 大文件提醒现在使用配置的大小阈值。
- 现在可以在设置中关闭图片点击预览。

## 0.3.0 - 2026-05-28

### English

#### Added

- Added continuous document zoom with the toolbar percentage input.
- Added `Ctrl` + mouse wheel zoom in the Word preview.
- Added click-to-preview image modal for rendered Word images.

#### Changed

- Replaced fixed zoom levels with percentage-based zoom.
- Improved zoom behavior to keep the cursor position stable while zooming.

#### Docs

- Updated README files for zoom, image preview, and `.doc` support limits.

### 中文

#### 新增

- 增加工具栏百分比输入框的连续文档缩放。
- 增加 Word 预览区域内的 `Ctrl` + 鼠标滚轮缩放。
- 增加渲染图片的点击放大预览弹窗。

#### 变更

- 将固定缩放档位改为基于百分比的缩放。
- 改进缩放行为，使缩放时尽量保持鼠标指向位置稳定。

#### 文档

- 更新 README，说明缩放、图片预览和 `.doc` 支持限制。

## 0.2.0 - 2026-05-28

### English

#### Added

- Added read-only `.docx` view registration.
- Added rendering through `docx-preview`.
- Added toolbar actions for reload, zoom, fit width, search, copy text, external open, and Markdown note creation.
- Added safe same-name Markdown summary note generation.

### 中文

#### 新增

- 增加只读 `.docx` 视图注册。
- 增加基于 `docx-preview` 的渲染能力。
- 增加刷新、缩放、适配宽度、搜索、复制文本、外部打开和创建 Markdown 笔记的工具栏操作。
- 增加安全的同名 Markdown 摘要笔记生成能力。

## 0.1.0 - 2026-05-28

### English

#### Added

- Initial project scaffold.

### 中文

#### 新增

- 初始化项目脚手架。
