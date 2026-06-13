import type { WordReaderLanguage } from "../i18n";

export interface PptxReaderText {
  displayName: string;
  commands: {
    previousSlide: string;
    nextSlide: string;
    toggleFullscreen: string;
    copySlideText: string;
    createSummaryNote: string;
    toggleNotes: string;
    focusSearch: string;
  };
  toolbar: {
    reload: string;
    previousSlide: string;
    nextSlide: string;
    pageNumber: string;
    zoomPercentage: string;
    fitWindow: string;
    enterFullscreen: string;
    exitFullscreen: string;
    showNavigation: string;
    hideNavigation: string;
    searchPlaceholder: string;
    searchPresentation: string;
    copyText: string;
    createSummaryNote: string;
    showNotes: string;
    hideNotes: string;
    openExternally: string;
  };
  status: {
    reading: (fileName: string) => string;
    indexing: (fileName: string) => string;
    rendering: (page: number, total: number) => string;
    ready: (fileName: string, page: number, total: number) => string;
  };
  navigation: {
    title: string;
    slideLabel: (page: number) => string;
    slideCount: (slides: number) => string;
    searchCount: (matches: number, slides: number) => string;
    notesMatch: string;
    noMatches: string;
  };
  notes: {
    title: string;
    empty: string;
  };
  errors: {
    formatMismatchTitle: string;
    formatMismatchBody: string;
    encryptedTitle: string;
    encryptedBody: string;
    damagedTitle: string;
    damagedBody: string;
    limitTitle: string;
    limitBody: string;
    unsupportedTitle: string;
    unsupportedBody: string;
    unknownTitle: string;
    unknownBody: string;
    tryAgain: string;
    openExternally: string;
    details: string;
  };
  notices: {
    externalDesktopOnly: string;
    externalLocalVaultOnly: string;
    externalFailed: (message: string) => string;
    fullscreenFailed: (message: string) => string;
    copiedSelectedText: string;
    copiedSlideText: string;
    copyFailed: (message: string) => string;
    openedExistingSummaryNote: string;
    createdSummaryNote: string;
  };
  summaryNote: {
    sourceLabel: string;
    currentSlideLabel: string;
    summaryHeading: string;
    keySlidesHeading: string;
    followUpsHeading: string;
    slideReferencesHeading: string;
    slideReference: (page: number, title: string) => string;
  };
}

const ENGLISH: PptxReaderText = {
  displayName: "PowerPoint Reader",
  commands: {
    previousSlide: "Previous slide",
    nextSlide: "Next slide",
    toggleFullscreen: "Toggle presentation fullscreen",
    copySlideText: "Copy text from current slide",
    createSummaryNote: "Create presentation summary note",
    toggleNotes: "Toggle speaker notes",
    focusSearch: "Search current presentation",
  },
  toolbar: {
    reload: "Reload",
    previousSlide: "Previous slide",
    nextSlide: "Next slide",
    pageNumber: "Slide number",
    zoomPercentage: "Zoom percentage",
    fitWindow: "Fit slide to window",
    enterFullscreen: "Enter fullscreen",
    exitFullscreen: "Exit fullscreen",
    showNavigation: "Show slide navigation",
    hideNavigation: "Hide slide navigation",
    searchPlaceholder: "Search presentation",
    searchPresentation: "Search all slides and speaker notes",
    copyText: "Copy slide text",
    createSummaryNote: "Create summary note",
    showNotes: "Show speaker notes",
    hideNotes: "Hide speaker notes",
    openExternally: "Open externally",
  },
  status: {
    reading: (fileName) => `Reading ${fileName}...`,
    indexing: (fileName) => `Indexing slides in ${fileName}...`,
    rendering: (page, total) => `Rendering slide ${page} of ${total}...`,
    ready: (fileName, page, total) =>
      `${fileName} - slide ${page} of ${total}`,
  },
  navigation: {
    title: "Slides",
    slideLabel: (page) => `Slide ${page}`,
    slideCount: (slides) => `${slides} slides`,
    searchCount: (matches, slides) =>
      `${matches} matches in ${slides} slides`,
    notesMatch: "Speaker notes",
    noMatches: "No matching slides",
  },
  notes: {
    title: "Speaker notes",
    empty: "No speaker notes for this slide.",
  },
  errors: {
    formatMismatchTitle: "Not a valid PowerPoint presentation",
    formatMismatchBody:
      "This file is not a readable ZIP-based .pptx presentation.",
    encryptedTitle: "Encrypted presentation",
    encryptedBody:
      "Encrypted or password-protected presentations cannot be previewed.",
    damagedTitle: "Damaged presentation",
    damagedBody:
      "The presentation package is incomplete or contains invalid XML.",
    limitTitle: "Presentation exceeds safe preview limits",
    limitBody:
      "The archive contains too many files, expands too far, or has an unsafe compression ratio.",
    unsupportedTitle: "Unsupported presentation",
    unsupportedBody:
      "This presentation does not contain readable slides or uses an unsupported package structure.",
    unknownTitle: "Could not open presentation",
    unknownBody: "The presentation could not be rendered locally.",
    tryAgain: "Reload the file after confirming it is a valid .pptx presentation.",
    openExternally: "Open it in PowerPoint or another presentation application.",
    details: "Technical details",
  },
  notices: {
    externalDesktopOnly:
      "External opening is only available in Obsidian Desktop",
    externalLocalVaultOnly:
      "External opening requires a local desktop vault",
    externalFailed: (message) => `Could not open the presentation: ${message}`,
    fullscreenFailed: (message) => `Could not enter fullscreen: ${message}`,
    copiedSelectedText: "Copied selected slide text",
    copiedSlideText: "Copied current slide text",
    copyFailed: (message) => `Could not copy slide text: ${message}`,
    openedExistingSummaryNote: "Opened existing presentation summary note",
    createdSummaryNote: "Created presentation summary note",
  },
  summaryNote: {
    sourceLabel: "Source",
    currentSlideLabel: "Current slide",
    summaryHeading: "Summary",
    keySlidesHeading: "Key slides",
    followUpsHeading: "Follow-ups",
    slideReferencesHeading: "Slide references",
    slideReference: (page, title) =>
      title ? `Slide ${page} - ${title}` : `Slide ${page}`,
  },
};

const CHINESE: PptxReaderText = {
  displayName: "PowerPoint 阅读器",
  commands: {
    previousSlide: "上一张幻灯片",
    nextSlide: "下一张幻灯片",
    toggleFullscreen: "切换演示全屏",
    copySlideText: "复制当前幻灯片文本",
    createSummaryNote: "创建演示文稿摘要笔记",
    toggleNotes: "切换演讲者备注",
    focusSearch: "搜索当前演示文稿",
  },
  toolbar: {
    reload: "重新加载",
    previousSlide: "上一张幻灯片",
    nextSlide: "下一张幻灯片",
    pageNumber: "幻灯片页码",
    zoomPercentage: "缩放百分比",
    fitWindow: "适应窗口",
    enterFullscreen: "进入全屏",
    exitFullscreen: "退出全屏",
    showNavigation: "显示幻灯片导航",
    hideNavigation: "隐藏幻灯片导航",
    searchPlaceholder: "搜索演示文稿",
    searchPresentation: "搜索全部幻灯片和演讲者备注",
    copyText: "复制幻灯片文本",
    createSummaryNote: "创建摘要笔记",
    showNotes: "显示演讲者备注",
    hideNotes: "隐藏演讲者备注",
    openExternally: "外部打开",
  },
  status: {
    reading: (fileName) => `正在读取 ${fileName}...`,
    indexing: (fileName) => `正在为 ${fileName} 建立幻灯片索引...`,
    rendering: (page, total) => `正在渲染第 ${page} / ${total} 张...`,
    ready: (fileName, page, total) =>
      `${fileName} - 第 ${page} / ${total} 张`,
  },
  navigation: {
    title: "幻灯片",
    slideLabel: (page) => `第 ${page} 张`,
    slideCount: (slides) => `共 ${slides} 张`,
    searchCount: (matches, slides) =>
      `${slides} 张幻灯片中共 ${matches} 处匹配`,
    notesMatch: "演讲者备注",
    noMatches: "没有匹配的幻灯片",
  },
  notes: {
    title: "演讲者备注",
    empty: "当前幻灯片没有演讲者备注。",
  },
  errors: {
    formatMismatchTitle: "不是有效的 PowerPoint 演示文稿",
    formatMismatchBody: "该文件不是可读取的 ZIP 格式 .pptx 演示文稿。",
    encryptedTitle: "演示文稿已加密",
    encryptedBody: "暂不支持预览加密或受密码保护的演示文稿。",
    damagedTitle: "演示文稿已损坏",
    damagedBody: "演示文稿包不完整，或其中包含无效的 XML。",
    limitTitle: "演示文稿超出安全预览限制",
    limitBody: "压缩包文件数量、解压后大小或压缩率超出安全限制。",
    unsupportedTitle: "不支持的演示文稿",
    unsupportedBody: "演示文稿没有可读取的幻灯片，或使用了不支持的包结构。",
    unknownTitle: "无法打开演示文稿",
    unknownBody: "无法在本地渲染该演示文稿。",
    tryAgain: "确认文件是有效的 .pptx 后重新加载。",
    openExternally: "请使用 PowerPoint 或其他演示文稿应用打开。",
    details: "技术详情",
  },
  notices: {
    externalDesktopOnly: "仅 Obsidian 桌面版支持外部打开",
    externalLocalVaultOnly: "外部打开要求仓库位于本地文件系统",
    externalFailed: (message) => `无法打开演示文稿：${message}`,
    fullscreenFailed: (message) => `无法进入全屏：${message}`,
    copiedSelectedText: "已复制选中的幻灯片文本",
    copiedSlideText: "已复制当前幻灯片文本",
    copyFailed: (message) => `无法复制幻灯片文本：${message}`,
    openedExistingSummaryNote: "已打开现有演示文稿摘要笔记",
    createdSummaryNote: "已创建演示文稿摘要笔记",
  },
  summaryNote: {
    sourceLabel: "原文",
    currentSlideLabel: "当前幻灯片",
    summaryHeading: "摘要",
    keySlidesHeading: "关键幻灯片",
    followUpsHeading: "待处理",
    slideReferencesHeading: "幻灯片引用",
    slideReference: (page, title) =>
      title ? `第 ${page} 张 - ${title}` : `第 ${page} 张`,
  },
};

export function getPptxReaderText(
  language: WordReaderLanguage,
): PptxReaderText {
  return language === "en" ? ENGLISH : CHINESE;
}
