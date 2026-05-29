import { App, PluginSettingTab, Setting } from "obsidian";

import type WordReaderPlugin from "./main";

export type WordReaderLanguage = "en" | "zh-CN";

export interface WordReaderSettings {
  language: WordReaderLanguage;
  defaultZoomPercent: number;
  defaultFitWidth: boolean;
  enableImagePreview: boolean;
  largeFileWarningMb: number;
}

export const DEFAULT_SETTINGS: WordReaderSettings = {
  language: "zh-CN",
  defaultZoomPercent: 100,
  defaultFitWidth: false,
  enableImagePreview: true,
  largeFileWarningMb: 25,
};

const MIN_ZOOM_PERCENT = 25;
const MAX_ZOOM_PERCENT = 400;
const MIN_LARGE_FILE_WARNING_MB = 1;
const MAX_LARGE_FILE_WARNING_MB = 500;

export function normalizeSettings(
  settings: Partial<WordReaderSettings> | null | undefined,
): WordReaderSettings {
  return {
    language: normalizeLanguage(settings?.language),
    defaultZoomPercent: clampInteger(
      settings?.defaultZoomPercent ?? DEFAULT_SETTINGS.defaultZoomPercent,
      MIN_ZOOM_PERCENT,
      MAX_ZOOM_PERCENT,
      DEFAULT_SETTINGS.defaultZoomPercent,
    ),
    defaultFitWidth:
      settings?.defaultFitWidth ?? DEFAULT_SETTINGS.defaultFitWidth,
    enableImagePreview:
      settings?.enableImagePreview ?? DEFAULT_SETTINGS.enableImagePreview,
    largeFileWarningMb: clampInteger(
      settings?.largeFileWarningMb ?? DEFAULT_SETTINGS.largeFileWarningMb,
      MIN_LARGE_FILE_WARNING_MB,
      MAX_LARGE_FILE_WARNING_MB,
      DEFAULT_SETTINGS.largeFileWarningMb,
    ),
  };
}

export class WordReaderSettingTab extends PluginSettingTab {
  constructor(
    app: App,
    private readonly plugin: WordReaderPlugin,
  ) {
    super(app, plugin);
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    const text = SETTINGS_TEXT[this.plugin.settings.language];

    containerEl.createEl("h2", { text: "Obsidian Word Reader" });

    new Setting(containerEl)
      .setName(text.languageName)
      .setDesc(text.languageDesc)
      .addDropdown((dropdown) => {
        dropdown
          .addOption("zh-CN", "中文")
          .addOption("en", "English")
          .setValue(this.plugin.settings.language)
          .onChange(async (value) => {
            this.plugin.settings.language = normalizeLanguage(value);
            await this.plugin.saveSettings();
            this.display();
          });
      });

    new Setting(containerEl)
      .setName(text.defaultZoomName)
      .setDesc(text.defaultZoomDesc)
      .addText((input) => {
        input
          .setPlaceholder("100")
          .setValue(String(this.plugin.settings.defaultZoomPercent))
          .onChange(async (value) => {
            this.plugin.settings.defaultZoomPercent = clampInteger(
              Number(value),
              MIN_ZOOM_PERCENT,
              MAX_ZOOM_PERCENT,
              DEFAULT_SETTINGS.defaultZoomPercent,
            );
            await this.plugin.saveSettings();
          });

        input.inputEl.type = "number";
        input.inputEl.min = String(MIN_ZOOM_PERCENT);
        input.inputEl.max = String(MAX_ZOOM_PERCENT);
        input.inputEl.step = "5";
      });

    new Setting(containerEl)
      .setName(text.defaultFitWidthName)
      .setDesc(text.defaultFitWidthDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.defaultFitWidth)
          .onChange(async (value) => {
            this.plugin.settings.defaultFitWidth = value;
            await this.plugin.saveSettings();
          });
      });

    new Setting(containerEl)
      .setName(text.imagePreviewName)
      .setDesc(text.imagePreviewDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.enableImagePreview)
          .onChange(async (value) => {
            this.plugin.settings.enableImagePreview = value;
            await this.plugin.saveSettings();
          });
      });

    new Setting(containerEl)
      .setName(text.largeFileWarningName)
      .setDesc(text.largeFileWarningDesc)
      .addText((input) => {
        input
          .setPlaceholder("25")
          .setValue(String(this.plugin.settings.largeFileWarningMb))
          .onChange(async (value) => {
            this.plugin.settings.largeFileWarningMb = clampInteger(
              Number(value),
              MIN_LARGE_FILE_WARNING_MB,
              MAX_LARGE_FILE_WARNING_MB,
              DEFAULT_SETTINGS.largeFileWarningMb,
            );
            await this.plugin.saveSettings();
          });

        input.inputEl.type = "number";
        input.inputEl.min = String(MIN_LARGE_FILE_WARNING_MB);
        input.inputEl.max = String(MAX_LARGE_FILE_WARNING_MB);
        input.inputEl.step = "1";
      });

    new Setting(containerEl)
      .setName(text.externalOpeningName)
      .setDesc(text.externalOpeningDesc);
  }
}

const SETTINGS_TEXT: Record<
  WordReaderLanguage,
  {
    languageName: string;
    languageDesc: string;
    defaultZoomName: string;
    defaultZoomDesc: string;
    defaultFitWidthName: string;
    defaultFitWidthDesc: string;
    imagePreviewName: string;
    imagePreviewDesc: string;
    largeFileWarningName: string;
    largeFileWarningDesc: string;
    externalOpeningName: string;
    externalOpeningDesc: string;
  }
> = {
  en: {
    languageName: "Settings language",
    languageDesc: "Choose the display language for this plugin settings page.",
    defaultZoomName: "Default zoom",
    defaultZoomDesc: "Initial zoom percentage for newly opened Word previews.",
    defaultFitWidthName: "Fit width by default",
    defaultFitWidthDesc:
      "Open newly rendered Word previews fitted to the pane width.",
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
  "zh-CN": {
    languageName: "设置界面语言",
    languageDesc: "选择本插件设置页面的显示语言。",
    defaultZoomName: "默认缩放比例",
    defaultZoomDesc: "新打开 Word 预览时使用的初始缩放百分比。",
    defaultFitWidthName: "默认适配宽度",
    defaultFitWidthDesc: "新渲染的 Word 预览默认适配当前窗格宽度。",
    imagePreviewName: "图片预览",
    imagePreviewDesc: "允许点击 Word 中渲染出的图片并打开更大的预览。",
    largeFileWarningName: "大文件提醒",
    largeFileWarningDesc:
      "当 Word 文件大于此大小时显示渲染提醒，单位为 MB。",
    externalOpeningName: "外部打开",
    externalOpeningDesc:
      "外部打开会使用操作系统为 .docx 文件配置的默认应用。如果想用 Word、WPS 或其他编辑器打开，请在 Windows 中修改默认应用。",
  },
};

function normalizeLanguage(value: unknown): WordReaderLanguage {
  return value === "en" || value === "zh-CN"
    ? value
    : DEFAULT_SETTINGS.language;
}

function clampInteger(
  value: number,
  min: number,
  max: number,
  fallback: number,
): number {
  if (!Number.isFinite(value)) {
    return fallback;
  }

  return Math.min(Math.max(Math.round(value), min), max);
}
