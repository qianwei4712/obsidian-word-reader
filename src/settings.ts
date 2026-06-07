import { App, PluginSettingTab, Setting } from "obsidian";

import type WordReaderPlugin from "./main";
import {
  getWordReaderText,
  normalizeLanguage,
  type WordReaderLanguage,
} from "./i18n";

export type { WordReaderLanguage };

export interface WordReaderSettings {
  language: WordReaderLanguage;
  defaultZoomPercent: number;
  defaultFitWidth: boolean;
  showOutlineByDefault: boolean;
  enableImagePreview: boolean;
  largeFileWarningMb: number;
}

export const DEFAULT_SETTINGS: WordReaderSettings = {
  language: "zh-CN",
  defaultZoomPercent: 100,
  defaultFitWidth: false,
  showOutlineByDefault: true,
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
    showOutlineByDefault:
      settings?.showOutlineByDefault ??
      DEFAULT_SETTINGS.showOutlineByDefault,
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
    const text = getWordReaderText(this.plugin.settings.language).settings;

    new Setting(containerEl).setName("Obsidian Word Reader").setHeading();

    new Setting(containerEl)
      .setName(text.languageName)
      .setDesc(text.languageDesc)
      .addDropdown((dropdown) => {
        dropdown
          .addOption("zh-CN", text.languageZh)
          .addOption("en", text.languageEn)
          .setValue(this.plugin.settings.language)
          .onChange(async (value) => {
            this.plugin.settings.language = normalizeLanguage(value);
            await this.plugin.saveSettings();
            this.plugin.refreshWordReaderViews();
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
      .setName(text.showOutlineName)
      .setDesc(text.showOutlineDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.showOutlineByDefault)
          .onChange(async (value) => {
            this.plugin.settings.showOutlineByDefault = value;
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
