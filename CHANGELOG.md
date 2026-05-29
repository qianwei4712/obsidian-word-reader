# Changelog

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

- Standardized release artifacts as `release/obsidian-word-reader-vX.Y.Z.zip`.
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

- 将发布产物统一为 `release/obsidian-word-reader-vX.Y.Z.zip`。
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
