# Stability and Maintenance

## 1.0.0 Stable Scope

The 1.0.0 stable line freezes the core reader scope:

- Open `.docx` files inside Obsidian Desktop as read-only previews.
- Show a guidance page for legacy `.doc` files.
- Keep source Word files unchanged.
- Support zoom, fit width, search, text copy, Markdown summary note creation, external open, image preview, and release packaging.
- Keep rendering, search, and image preview stable for normal vault documents.

New reader features should be planned for minor versions after 1.0.0. Patch releases should focus on regressions, compatibility fixes, documentation corrections, and packaging issues.

## Manual Test Checklist

Run this checklist before publishing a stable release:

- Build and package:
  - `npm run typecheck`
  - `npm run build`
  - `npm run release`
  - Confirm the release zip contains only `main.js`, `manifest.json`, and `styles.css`.
- Installation:
  - Install the release zip into a test vault at `.obsidian/plugins/obsidian-word-reader/`.
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
  - Test search, previous/next navigation, and current result highlighting.
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
- Release automation:
  - Confirm `node scripts/release-check.mjs --tag X.Y.Z` passes for the release tag.
  - Push an `X.Y.Z` tag in a test release flow.
  - Confirm GitHub Actions creates a release and uploads `main.js`, `manifest.json`, `styles.css`, and the zip.

## Support Boundaries

Supported:

- Obsidian Desktop.
- Local desktop vaults.
- `.docx` read-only preview through `docx-preview`.
- Plain text and Markdown extraction through `mammoth`.
- `.doc` detection with external-open and conversion guidance.

Not supported:

- Editing or saving Word files.
- Mobile-specific support.
- Direct rendering of legacy `.doc` binary files.
- Password-protected or encrypted Word documents.
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

### 支持边界

支持：

- Obsidian 桌面端。
- 本地桌面 vault。
- 基于 `docx-preview` 的 `.docx` 只读预览。
- 基于 `mammoth` 的纯文本和 Markdown 提取。
- `.doc` 检测、外部打开和转换说明。

不支持：

- 编辑或保存 Word 文件。
- 移动端专项支持。
- 直接渲染旧版 `.doc` 二进制文件。
- 密码保护或加密 Word 文档。
- 复杂 Word 排版与 Microsoft Word 完全一致。
- 不提供本地路径的远程 vault 适配器外部打开。
