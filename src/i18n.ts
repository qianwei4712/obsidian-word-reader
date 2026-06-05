export type WordReaderLanguage = "en" | "zh-CN";

export interface WordReaderText {
  settings: {
    languageName: string;
    languageDesc: string;
    languageZh: string;
    languageEn: string;
    defaultZoomName: string;
    defaultZoomDesc: string;
    defaultFitWidthName: string;
    defaultFitWidthDesc: string;
    showOutlineName: string;
    showOutlineDesc: string;
    imagePreviewName: string;
    imagePreviewDesc: string;
    largeFileWarningName: string;
    largeFileWarningDesc: string;
    externalOpeningName: string;
    externalOpeningDesc: string;
  };
  commands: {
    reload: string;
    copyText: string;
    copyMarkdown: string;
    createNote: string;
    openExternal: string;
  };
  toolbar: {
    reload: string;
    zoomPercentage: string;
    fitWidth: string;
    showOutline: string;
    hideOutline: string;
    searchPlaceholder: string;
    searchDocument: string;
    previousSearchResult: string;
    nextSearchResult: string;
    copyText: string;
    copyMarkdown: string;
    createSummaryNote: string;
    openExternally: string;
  };
  notices: {
    legacyDoc: string;
    copiedSelectedText: string;
    copiedPlainText: string;
    copiedSelectedMarkdown: string;
    copiedMarkdown: string;
    copiedImage: string;
    couldNotCopyImage: (message: string) => string;
    startedImageDownload: string;
    savedImage: string;
    couldNotSaveImage: (message: string) => string;
    openedExistingSummaryNote: string;
    createdSummaryNote: string;
    externalDesktopOnly: string;
    externalLocalVaultOnly: string;
    largeDocument: (sizeMb: string) => string;
  };
  status: {
    extractingPlainText: string;
    copiedPlainTextFrom: (fileName: string) => string;
    extractingMarkdown: string;
    copiedMarkdownFrom: (fileName: string) => string;
    preview: (fileName: string) => string;
    reading: (fileName: string) => string;
    rendering: (fileName: string) => string;
    largeDocument: (sizeMb: string) => string;
    legacyDoc: (fileName: string) => string;
  };
  outline: {
    title: string;
    empty: string;
  };
  legacyDoc: {
    title: string;
    body: string;
    openExternally: string;
    convertToDocx: string;
  };
  errors: {
    encryptedTitle: string;
    encryptedBody: string;
    encryptedStatus: string;
    encryptedTips: string[];
    damagedTitle: string;
    damagedBody: string;
    damagedStatus: string;
    damagedTips: string[];
    genericTitle: string;
    genericBody: string;
    genericStatus: string;
    genericTips: string[];
  };
  imagePreview: {
    fitToWindow: string;
    actualSize: string;
    copyImage: string;
    saveImageAs: string;
    loading: string;
    loadFailed: string;
    unknownSize: string;
    saveDialogTitle: string;
    imageFilterName: (extension: string) => string;
    allFiles: string;
  };
  summaryNote: {
    sourceLabel: string;
    summaryHeading: string;
    keyFindingsHeading: string;
    followUpsHeading: string;
    quotesHeading: string;
  };
}

export function getWordReaderText(language: WordReaderLanguage): WordReaderText {
  return WORD_READER_TEXT[language] ?? WORD_READER_TEXT["zh-CN"];
}

export function normalizeLanguage(value: unknown): WordReaderLanguage {
  return value === "en" || value === "zh-CN" ? value : "zh-CN";
}

const WORD_READER_TEXT: Record<WordReaderLanguage, WordReaderText> = {
  en: {
    settings: {
      languageName: "Interface language",
      languageDesc: "Choose the display language for this plugin.",
      languageZh: "中文",
      languageEn: "English",
      defaultZoomName: "Default zoom",
      defaultZoomDesc: "Initial zoom percentage for newly opened Word previews.",
      defaultFitWidthName: "Fit width by default",
      defaultFitWidthDesc:
        "Open newly rendered Word previews fitted to the pane width.",
      showOutlineName: "Show outline by default",
      showOutlineDesc:
        "Open Word previews with the clickable heading outline visible.",
      imagePreviewName: "Image preview",
      imagePreviewDesc:
        "Allow clicking rendered Word images to open a larger preview.",
      largeFileWarningName: "Large file warning",
      largeFileWarningDesc:
        "Show a rendering warning when a Word file is larger than this size in MB.",
      externalOpeningName: "External opening",
      externalOpeningDesc:
        "External opening uses your operating system's default application for .docx files. Change the default app in Windows if you want Word, WPS, or another editor to open documents.",
    },
    commands: {
      reload: "Reload current Word document",
      copyText: "Copy text from current Word document",
      copyMarkdown: "Copy current Word document as Markdown",
      createNote: "Create summary note for current Word document",
      openExternal: "Open current Word document externally",
    },
    toolbar: {
      reload: "Reload",
      zoomPercentage: "Zoom percentage",
      fitWidth: "Fit width",
      showOutline: "Show outline",
      hideOutline: "Hide outline",
      searchPlaceholder: "Search",
      searchDocument: "Search document",
      previousSearchResult: "Previous search result",
      nextSearchResult: "Next search result",
      copyText: "Copy text",
      copyMarkdown: "Copy as Markdown",
      createSummaryNote: "Create summary note",
      openExternally: "Open externally",
    },
    notices: {
      legacyDoc:
        "Legacy .doc files are not readable inside Obsidian. Open the file externally or convert it to .docx.",
      copiedSelectedText: "Copied selected text",
      copiedPlainText: "Copied plain text",
      copiedSelectedMarkdown: "Copied selected Markdown",
      copiedMarkdown: "Copied Markdown",
      copiedImage: "Copied image",
      couldNotCopyImage: (message) => `Could not copy image: ${message}`,
      startedImageDownload: "Started image download",
      savedImage: "Saved image",
      couldNotSaveImage: (message) => `Could not save image: ${message}`,
      openedExistingSummaryNote: "Opened existing summary note",
      createdSummaryNote: "Created summary note",
      externalDesktopOnly:
        "External opening is only available in Obsidian Desktop",
      externalLocalVaultOnly: "External opening requires a local desktop vault",
      largeDocument: (sizeMb) =>
        `Large Word document (${sizeMb} MB). Rendering may take a while.`,
    },
    status: {
      extractingPlainText: "Extracting plain text...",
      copiedPlainTextFrom: (fileName) => `Copied plain text from ${fileName}`,
      extractingMarkdown: "Extracting Markdown...",
      copiedMarkdownFrom: (fileName) => `Copied Markdown from ${fileName}`,
      preview: (fileName) => `Read-only preview: ${fileName}`,
      reading: (fileName) => `Reading ${fileName}...`,
      rendering: (fileName) => `Rendering ${fileName}...`,
      largeDocument: (sizeMb) =>
        `Large document detected (${sizeMb} MB). Rendering may take a while...`,
      legacyDoc: (fileName) => `Legacy .doc file: ${fileName}`,
    },
    outline: {
      title: "Outline",
      empty: "No headings found",
    },
    legacyDoc: {
      title: "Legacy .doc file",
      body: "This plugin renders .docx files directly in Obsidian. The older .doc format is a different binary Word format, so it cannot be rendered safely here.",
      openExternally:
        "Open the file with Word, WPS, or your system default editor.",
      convertToDocx:
        "Convert or save the file as .docx, then open the converted file in Obsidian.",
    },
    errors: {
      encryptedTitle: "Encrypted Word document",
      encryptedBody:
        "This document appears to be encrypted or password protected. Obsidian Word Reader cannot unlock it. Open it in Word or WPS, remove the password if appropriate, then save a readable .docx copy.",
      encryptedStatus: "Encrypted Word document",
      encryptedTips: [
        "Open the document in Word or WPS with the correct password.",
        'Go to "File \u2192 Info \u2192 Protect Document \u2192 Encrypt with Password" and clear the password field.',
        "Save the document as a new .docx file, then open it in Obsidian.",
      ],
      damagedTitle: "Damaged or unsupported Word document",
      damagedBody:
        "This document could not be parsed as a valid .docx file. It may be damaged, incomplete, or saved in a format that looks like .docx but is not readable by the renderer. Try opening it in Word or WPS and saving a fresh .docx copy.",
      damagedStatus: "Damaged Word document",
      damagedTips: [
        "Open the document in Word or WPS \u2014 Word has built-in repair capability for corrupted files.",
        'Use "File \u2192 Save As" to re-save it as a new .docx file, then open the copy in Obsidian.',
        "If the file was generated by a third-party tool (WPS, LibreOffice, etc.), try re-saving with Microsoft Word.",
      ],
      genericTitle: "Word document could not be rendered",
      genericBody:
        "This Word document could not be rendered inside Obsidian. This often happens when the document contains minor formatting quirks that Word tolerates but the renderer does not. The suggestions below usually resolve the issue.",
      genericStatus: "Failed to open Word document",
      genericTips: [
        "Open the document in Word or WPS \u2014 these apps can automatically fix minor formatting issues.",
        'Use "File \u2192 Save As" to save a fresh .docx copy, then open the new file in Obsidian.',
        "This error can also be caused by complex nested tables, SmartArt graphics, or custom XML structures. Re-saving usually fixes these.",
        "If the document was generated by a third-party tool (WPS, LibreOffice, online converter), try re-saving it with Microsoft Word.",
      ],
    },
    imagePreview: {
      fitToWindow: "Fit to window",
      actualSize: "Actual size",
      copyImage: "Copy image",
      saveImageAs: "Save image as",
      loading: "Loading image...",
      loadFailed: "Image could not be loaded",
      unknownSize: "Unknown size",
      saveDialogTitle: "Save image as",
      imageFilterName: (extension) => `${extension.toUpperCase()} image`,
      allFiles: "All files",
    },
    summaryNote: {
      sourceLabel: "Source",
      summaryHeading: "Summary",
      keyFindingsHeading: "Key findings",
      followUpsHeading: "Follow-ups",
      quotesHeading: "Quoted excerpts",
    },
  },
  "zh-CN": {
    settings: {
      languageName: "界面语言",
      languageDesc: "选择本插件的显示语言。",
      languageZh: "中文",
      languageEn: "English",
      defaultZoomName: "默认缩放比例",
      defaultZoomDesc: "新打开 Word 预览时使用的初始缩放百分比。",
      defaultFitWidthName: "默认适配宽度",
      defaultFitWidthDesc: "新渲染的 Word 预览默认适配当前窗格宽度。",
      showOutlineName: "默认显示大纲",
      showOutlineDesc: "打开 Word 预览时默认显示可点击的标题大纲。",
      imagePreviewName: "图片预览",
      imagePreviewDesc: "允许点击 Word 中渲染出的图片并打开更大的预览。",
      largeFileWarningName: "大文件提醒",
      largeFileWarningDesc:
        "当 Word 文件大于此大小时显示渲染提醒，单位为 MB。",
      externalOpeningName: "外部打开",
      externalOpeningDesc:
        "外部打开会使用操作系统为 .docx 文件配置的默认应用。如果想用 Word、WPS 或其他编辑器打开，请在 Windows 中修改默认应用。",
    },
    commands: {
      reload: "重新加载当前 Word 文档",
      copyText: "复制当前 Word 文档文本",
      copyMarkdown: "复制当前 Word 文档为 Markdown",
      createNote: "为当前 Word 文档创建摘要笔记",
      openExternal: "外部打开当前 Word 文档",
    },
    toolbar: {
      reload: "重新加载",
      zoomPercentage: "缩放百分比",
      fitWidth: "适配宽度",
      showOutline: "显示大纲",
      hideOutline: "隐藏大纲",
      searchPlaceholder: "搜索",
      searchDocument: "搜索文档",
      previousSearchResult: "上一个搜索结果",
      nextSearchResult: "下一个搜索结果",
      copyText: "复制文本",
      copyMarkdown: "复制为 Markdown",
      createSummaryNote: "创建摘要笔记",
      openExternally: "外部打开",
    },
    notices: {
      legacyDoc:
        "旧版 .doc 文件不能在 Obsidian 内直接读取。请外部打开，或转换为 .docx。",
      copiedSelectedText: "已复制选中文本",
      copiedPlainText: "已复制纯文本",
      copiedSelectedMarkdown: "已复制选中内容的 Markdown",
      copiedMarkdown: "已复制 Markdown",
      copiedImage: "已复制图片",
      couldNotCopyImage: (message) => `无法复制图片：${message}`,
      startedImageDownload: "已开始下载图片",
      savedImage: "已保存图片",
      couldNotSaveImage: (message) => `无法保存图片：${message}`,
      openedExistingSummaryNote: "已打开已有摘要笔记",
      createdSummaryNote: "已创建摘要笔记",
      externalDesktopOnly: "外部打开仅支持 Obsidian 桌面端",
      externalLocalVaultOnly: "外部打开需要使用本地桌面 vault",
      largeDocument: (sizeMb) =>
        `较大的 Word 文档（${sizeMb} MB），渲染可能需要一些时间。`,
    },
    status: {
      extractingPlainText: "正在提取纯文本...",
      copiedPlainTextFrom: (fileName) => `已从 ${fileName} 复制纯文本`,
      extractingMarkdown: "正在提取 Markdown...",
      copiedMarkdownFrom: (fileName) => `已从 ${fileName} 复制 Markdown`,
      preview: (fileName) => `只读预览：${fileName}`,
      reading: (fileName) => `正在读取 ${fileName}...`,
      rendering: (fileName) => `正在渲染 ${fileName}...`,
      largeDocument: (sizeMb) =>
        `检测到大文档（${sizeMb} MB），渲染可能需要一些时间...`,
      legacyDoc: (fileName) => `旧版 .doc 文件：${fileName}`,
    },
    outline: {
      title: "大纲",
      empty: "未找到标题",
    },
    legacyDoc: {
      title: "旧版 .doc 文件",
      body: "本插件可以在 Obsidian 内直接渲染 .docx 文件。旧版 .doc 是另一种二进制 Word 格式，不能在这里安全渲染。",
      openExternally: "使用 Word、WPS 或系统默认编辑器打开该文件。",
      convertToDocx: "将文件转换或另存为 .docx 后，再在 Obsidian 中打开转换后的文件。",
    },
    errors: {
      encryptedTitle: "加密的 Word 文档",
      encryptedBody:
        "该文档可能已加密或受密码保护。Obsidian Word Reader 不能解锁文档。请用 Word 或 WPS 打开，按需移除密码后另存为可读取的 .docx 副本。",
      encryptedStatus: "加密的 Word 文档",
      encryptedTips: [
        "用 Word 或 WPS 输入正确密码打开该文档。",
        "点击「文件 → 信息 → 保护文档 → 用密码进行加密」，清空密码框后确认。",
        "将文档另存为新的 .docx 文件，再在 Obsidian 中打开。",
      ],
      damagedTitle: "损坏或不支持的 Word 文档",
      damagedBody:
        "该文档无法作为有效 .docx 解析。它可能已损坏、不完整，或扩展名像 .docx 但实际格式无法被渲染器读取。请尝试用 Word 或 WPS 打开并另存为新的 .docx 副本。",
      damagedStatus: "损坏的 Word 文档",
      damagedTips: [
        "用 Word 或 WPS 打开该文档——Word 对损坏文件有内置修复能力。",
        "点击「文件 → 另存为」，重新保存为新的 .docx 文件，然后在 Obsidian 中打开副本。",
        "如果文件由第三方工具生成（WPS、LibreOffice 等），尝试用 Microsoft Word 重新保存。",
      ],
      genericTitle: "无法渲染 Word 文档",
      genericBody:
        "该 Word 文档无法在 Obsidian 内渲染。这通常是因为文档内部存在轻微的格式瑕疵——Word 能自动容错，但渲染器要求更严格的格式。以下方法通常可以解决问题：",
      genericStatus: "打开 Word 文档失败",
      genericTips: [
        "用 Word 或 WPS 打开该文档——这些软件能自动修复轻微的格式问题。",
        "打开后直接「文件 → 另存为」一份新的 .docx 文件，再在 Obsidian 中打开新文件。",
        "此错误也可能由复杂的表格嵌套、SmartArt 图形或自定义 XML 结构引起，另存为通常可以修复。",
        "如果文档由第三方工具（WPS、LibreOffice、在线转换器等）生成，尝试用 Microsoft Word 重新保存。",
      ],
    },
    imagePreview: {
      fitToWindow: "适配窗口",
      actualSize: "原始尺寸",
      copyImage: "复制图片",
      saveImageAs: "另存图片",
      loading: "正在加载图片...",
      loadFailed: "图片无法加载",
      unknownSize: "未知尺寸",
      saveDialogTitle: "另存图片",
      imageFilterName: (extension) => `${extension.toUpperCase()} 图片`,
      allFiles: "所有文件",
    },
    summaryNote: {
      sourceLabel: "原文",
      summaryHeading: "摘要",
      keyFindingsHeading: "关键结论",
      followUpsHeading: "待处理",
      quotesHeading: "引用摘录",
    },
  },
};
