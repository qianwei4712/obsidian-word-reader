# Obsidian Word Reader

Obsidian Word Reader 是一个 Obsidian 桌面端插件，用于在 Obsidian 内直接打开 `.docx` 文件，并以只读方式阅读 Word 文档。

它的目标不是把 Obsidian 变成 Word 编辑器，而是提供一个稳定、安全、适合知识整理的 Word 阅读入口。

## 功能

- 在 Obsidian 标签页中打开 `.docx` 文件。
- 尽量保留标题、段落、列表、表格、图片和分页效果。
- 只读预览，不修改原始 Word 文件。
- 使用白底文档区域，避免暗色主题影响 Word 内容阅读。
- 支持缩放：80%、100%、125%、150%。
- 支持适配窗口宽度。
- 支持在当前渲染内容中搜索。
- 支持复制选中文本，或提取并复制整篇纯文本。
- 支持用系统默认程序打开原始 Word 文件。
- 支持创建同名 Markdown 摘要笔记，并自动链接回原始 `.docx`。

## 开发安装

1. 安装依赖：

   ```bash
   npm install
   ```

2. 构建插件：

   ```bash
   npm run build
   ```

3. 构建完成后，发布文件会生成到 `dist` 目录：

   ```text
   dist/main.js
   dist/manifest.json
   dist/styles.css
   ```

4. 将 `dist` 目录中的三个文件复制到 Obsidian vault 的插件目录中，例如：

   ```text
   YourVault/.obsidian/plugins/obsidian-word-reader/
   ```

5. 在 Obsidian 设置中启用第三方插件，并启用 Obsidian Word Reader。

如果使用当前机器上的指定 Node 版本，可以使用：

```text
E:\DevelopHelper\nvm\v18.20.8\node.exe
```

## 使用方式

1. 将 `.docx` 文件放入 Obsidian vault。
2. 在文件列表中点击该 `.docx`。
3. 插件会在 Obsidian 标签页中打开只读预览。
4. 使用顶部工具栏进行缩放、适配宽度、搜索、复制文本、创建摘要笔记或外部打开。

## 摘要笔记

点击“创建摘要笔记”后，插件会在原 Word 文件同目录创建一个同名 `.md` 文件。

例如：

```text
报告.docx
报告.md
```

生成的 Markdown 会包含：

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

## 已知限制

- 这不是 Word 编辑器。
- 插件不会把修改保存回 `.docx`。
- 复杂 Word 排版无法保证与 Microsoft Word 100% 一致。
- 大文件或包含大量图片的文档可能渲染较慢。
- 当前版本优先支持 Obsidian Desktop，不包含移动端专项适配。

## 安全说明

插件读取 `.docx` 时使用 Obsidian vault 的二进制读取接口，只用于渲染和文本提取。

插件不会修改、覆盖或回写原始 Word 文件。需要编辑 Word 内容时，请使用 Word、WPS 或其他外部编辑器。

## 推荐工作流

建议将原始 Word 文件保留在 vault 中，用插件进行阅读，并为重要文档创建同名 Markdown 摘要笔记。

这样可以同时保留原始文档格式，并把摘要、结论、待办和摘录纳入 Obsidian 的知识库工作流。
