import type { WordReaderLanguage } from "../i18n";

export interface PptxReaderText {
  displayName: string;
  commands: {
    previousSlide: string;
    nextSlide: string;
    toggleFullscreen: string;
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
    openExternally: string;
  };
  status: {
    reading: (fileName: string) => string;
    rendering: (page: number, total: number) => string;
    ready: (fileName: string, page: number, total: number) => string;
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
  };
}

const ENGLISH: PptxReaderText = {
  displayName: "PowerPoint Reader",
  commands: {
    previousSlide: "Previous slide",
    nextSlide: "Next slide",
    toggleFullscreen: "Toggle presentation fullscreen",
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
    openExternally: "Open externally",
  },
  status: {
    reading: (fileName) => `Reading ${fileName}...`,
    rendering: (page, total) => `Rendering slide ${page} of ${total}...`,
    ready: (fileName, page, total) =>
      `${fileName} - slide ${page} of ${total}`,
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
  },
};

const CHINESE: PptxReaderText = {
  displayName: "PowerPoint 阅读器",
  commands: {
    previousSlide: "上一张幻灯片",
    nextSlide: "下一张幻灯片",
    toggleFullscreen: "切换演示全屏",
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
    openExternally: "外部打开",
  },
  status: {
    reading: (fileName) => `正在读取 ${fileName}...`,
    rendering: (page, total) => `正在渲染第 ${page} / ${total} 张...`,
    ready: (fileName, page, total) =>
      `${fileName} - 第 ${page} / ${total} 张`,
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
  },
};

export function getPptxReaderText(
  language: WordReaderLanguage,
): PptxReaderText {
  return language === "en" ? ENGLISH : CHINESE;
}
