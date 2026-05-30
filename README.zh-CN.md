# Obsidian Word Reader

[English README](README.md)

Obsidian Word Reader 是一个 Obsidian 桌面端插件，用于在 Obsidian 内直接打开 `.docx` 文件，并以只读方式阅读 Word 文档。

插件的目标不是把 Obsidian 变成 Word 编辑器，而是提供一个稳定、安全、适合知识整理的 Word 阅读入口，让 Word 文档可以在 vault 中被阅读、搜索、引用和摘要化，同时保持原始文件不变。

## 功能

- 在 Obsidian 标签页中打开 `.docx` 文件。
- 尽量保留标题、段落、列表、表格、图片和分页效果。
- 只读预览，不修改原始 Word 文件。
- 使用白底文档区域，避免暗色主题影响 Word 内容阅读。
- 支持通过工具栏百分比输入框或 `Ctrl` + 鼠标滚轮连续缩放。
- 支持适配当前窗格宽度。
- 支持点击渲染内容中的图片，在弹窗中放大预览。
- 支持在当前渲染内容中搜索，并在搜索结果间前后跳转。
- 支持复制选中文本、整篇纯文本、选中内容 Markdown 或整篇 Markdown。
- 支持通过可点击大纲在渲染标题之间跳转。
- 支持插件界面在中文和英文之间切换。
- 支持用系统默认程序打开原始 Word 文件。
- 旧版 `.doc` 文件会显示说明页面，提供外部打开和转换建议。
- 加密、损坏或不受支持的 Word 文档会显示更明确的错误提示。
- 切换或重新加载文档时会避免旧渲染结果污染当前预览。
- 支持创建同名 Markdown 摘要笔记，并自动链接回原始 `.docx`。

## 支持的文件

| 扩展名 | 状态 | 说明 |
| --- | --- | --- |
| `.docx` | 支持 | 通过 `docx-preview` 在 Obsidian 内渲染。 |
| `.doc` | 说明页面 | 在 Obsidian 中显示外部打开和 `.docx` 转换说明。 |

## 使用方式

1. 将 `.docx` 文件放入 Obsidian vault。
2. 在文件列表中点击该 `.docx`。
3. 插件会在 Obsidian 标签页中打开只读预览。
4. 使用顶部工具栏进行刷新、缩放、适配宽度、显示大纲、搜索、复制文本、复制 Markdown、创建摘要笔记或外部打开。

### 缩放

- 可以在工具栏的缩放输入框中输入百分比。
- 可以在文档预览区域使用 `Ctrl` + 鼠标滚轮进行连续缩放。
- 需要让文档适配当前窗格时，可以使用适配宽度按钮。

### 图片预览

点击 Word 文档中渲染出的图片，可以在弹窗中查看更大的预览。按 `Esc` 或关闭弹窗即可返回文档阅读。

在图片预览中：

- 使用鼠标滚轮缩放。
- 拖拽图片进行平移查看。
- 双击将图片重新适配到预览窗口。
- 使用工具栏适配窗口、查看原始尺寸、复制图片或另存图片。
- 工具栏会显示图片原始尺寸和当前缩放比例。

### 搜索

- 在搜索框中输入关键词后，会高亮匹配的渲染文本。
- 使用向上和向下按钮在结果之间跳转。
- 按 `Enter` 跳转到下一个结果，按 `Shift` + `Enter` 跳转到上一个结果。
- 结果计数会显示当前结果序号和总匹配数量。

### 大纲

- 使用大纲按钮显示或隐藏标题大纲。
- 点击大纲条目会滚动到对应的渲染标题。
- 大纲来自 Word 文档渲染结果中的标题结构。

### 复制

- 可以将选中的渲染文本复制为纯文本。
- 没有选区时，可以复制整篇文档纯文本。
- 可以将选中的渲染内容复制为 Markdown。
- 没有选区时，可以将整篇 `.docx` 转换并复制为 Markdown。

## 设置

打开 Obsidian 设置，然后进入 `第三方插件` -> `Obsidian Word Reader`。

当前支持的设置：

- 插件界面语言。
- 新打开 Word 预览时使用的默认缩放比例。
- 新打开 Word 预览时是否默认适配当前窗格宽度。
- 是否默认显示大纲。
- 是否允许点击渲染出的图片进行放大预览。
- 大文件提醒阈值，单位为 MB。
- 外部打开说明。插件会使用操作系统为 `.docx` 配置的默认应用。

## 兼容性与错误处理

- 旧版 `.doc` 文件会打开说明页面，不会直接内嵌渲染。
- 加密或受密码保护的文档会显示专门的加密文档提示。
- 损坏或无效的 `.docx` 文件会显示损坏文档提示，并给出下一步处理建议。
- 大文件提醒会显示文件大小，并使用设置中的提醒阈值。

## 性能与稳定性

- 渲染流程使用取消 token，过期渲染结果会被丢弃。
- Word 内容会先渲染到临时缓冲区，再替换可见预览。
- 插件会跟踪当前已渲染的文件状态，避免不必要的重复渲染。
- 搜索高亮在输入时会进行防抖处理，减少大文档中的重复计算。
- 关闭或卸载文件时会释放文档缓冲、搜索定时器和搜索结果引用。

## 稳定性与支持边界

稳定版核心范围、手动测试清单、支持边界和长期维护策略见 [STABILITY.md](STABILITY.md)。

### 摘要笔记

点击“创建摘要笔记”后，插件会在原 Word 文件同目录创建一个同名 `.md` 文件。

例如：

```text
报告.docx
报告.md
```

生成的 Markdown 会包含 frontmatter 和基础章节：

```markdown
---
source: "报告.docx"
type: word-note
created: 2026-05-28
---

# 报告

原文：[[报告.docx]]

## 摘要

## 关键结论

## 待处理

## 引用摘录
```

如果同名 Markdown 已经存在，插件会直接打开已有笔记，不会覆盖内容。

## 安装

### 通过发布 zip 安装

1. 从 GitHub Release 页面下载 `obsidian-word-reader-X.Y.Z.zip`。
2. 在 vault 中创建插件目录：

   ```text
   YourVault/.obsidian/plugins/obsidian-word-reader/
   ```

3. 将 zip 解压到该目录。文件必须直接位于插件目录根部：

   ```text
   YourVault/.obsidian/plugins/obsidian-word-reader/main.js
   YourVault/.obsidian/plugins/obsidian-word-reader/manifest.json
   YourVault/.obsidian/plugins/obsidian-word-reader/styles.css
   ```

4. 重启 Obsidian，或重新加载第三方插件。
5. 在 Obsidian 设置中启用第三方插件，并启用 Obsidian Word Reader。

### 从源码构建

1. 安装依赖：

   ```bash
   npm install
   ```

2. 构建插件：

   ```bash
   npm run build
   ```

3. 将 `dist` 中生成的文件复制到 Obsidian vault 的插件目录：

   ```text
   dist/main.js
   dist/manifest.json
   dist/styles.css
   ```

## 本地发布

创建并校验本地发布包：

```bash
npm run release
```

该命令会执行 TypeScript 检查、构建插件、校验版本一致性、生成可安装 zip，并提取当前版本的 changelog 作为 release notes。

预期输出：

```text
release/obsidian-word-reader-1.1.1.zip
release/CHANGELOG-1.1.1.md
```

zip 根目录只包含 Obsidian 需要的文件：

```text
main.js
manifest.json
styles.css
```

发布产物会被 Git 忽略，不需要提交。

## GitHub 自动发布

推送不带 `v` 前缀的版本 tag 后，GitHub Actions 会自动创建 release：

```bash
git tag 1.1.1
git push origin 1.1.1
```

workflow 会校验 tag 是否与 `package.json`、`manifest.json` 和 `package-lock.json` 一致，然后构建插件、生成 `release/obsidian-word-reader-1.1.1.zip`、提取对应版本的 `CHANGELOG.md` 内容，并把 `main.js`、`manifest.json`、`styles.css` 和 zip 一起上传到 GitHub Release。

## 开发

安装依赖：

```bash
npm install
```

执行生产构建：

```bash
npm run build
```

执行 TypeScript 检查：

```bash
npm run typecheck
```

构建后校验发布元数据：

```bash
npm run release:check
```

当前机器指定的 Node.js 路径：

```text
E:\DevelopHelper\nvm\v18.20.8\node.exe
```

## 已知限制

- 这不是 Word 编辑器。
- 插件不会把修改保存回 `.docx`。
- 旧版 `.doc` 文件不会直接内嵌渲染，但会显示外部打开和转换说明。
- 复杂 Word 排版无法保证与 Microsoft Word 100% 一致。
- 大文件或包含大量图片的文档可能渲染较慢。
- 当前版本优先支持 Obsidian Desktop，不包含移动端专项适配。

## 安全说明

插件读取 `.docx` 时使用 Obsidian vault 的二进制读取接口，只用于渲染和文本提取。

插件不会修改、覆盖或回写原始 Word 文件。需要编辑 Word 内容时，请使用 Word、WPS 或其他外部编辑器。

## 推荐工作流

建议将原始 Word 文件保留在 vault 中，用插件进行阅读，并为重要文档创建同名 Markdown 摘要笔记。

这样可以同时保留原始文档格式，并把摘要、结论、待办和摘录纳入 Obsidian 的知识库工作流。
