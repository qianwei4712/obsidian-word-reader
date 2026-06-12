# Changelog

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
