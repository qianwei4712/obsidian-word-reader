import { App, PluginSettingTab, Setting } from "obsidian";

import type WordReaderPlugin from "./main";

export interface WordReaderSettings {
  defaultZoomPercent: number;
  defaultFitWidth: boolean;
  enableImagePreview: boolean;
  largeFileWarningMb: number;
}

export const DEFAULT_SETTINGS: WordReaderSettings = {
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

    containerEl.createEl("h2", { text: "Obsidian Word Reader" });

    new Setting(containerEl)
      .setName("Default zoom")
      .setDesc("Initial zoom percentage for newly opened Word previews.")
      .addText((text) => {
        text
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

        text.inputEl.type = "number";
        text.inputEl.min = String(MIN_ZOOM_PERCENT);
        text.inputEl.max = String(MAX_ZOOM_PERCENT);
        text.inputEl.step = "5";
      });

    new Setting(containerEl)
      .setName("Fit width by default")
      .setDesc("Open newly rendered Word previews fitted to the pane width.")
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.defaultFitWidth)
          .onChange(async (value) => {
            this.plugin.settings.defaultFitWidth = value;
            await this.plugin.saveSettings();
          });
      });

    new Setting(containerEl)
      .setName("Image preview")
      .setDesc("Allow clicking rendered Word images to open a larger preview.")
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.enableImagePreview)
          .onChange(async (value) => {
            this.plugin.settings.enableImagePreview = value;
            await this.plugin.saveSettings();
          });
      });

    new Setting(containerEl)
      .setName("Large file warning")
      .setDesc("Show a rendering warning when a Word file is larger than this size in MB.")
      .addText((text) => {
        text
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

        text.inputEl.type = "number";
        text.inputEl.min = String(MIN_LARGE_FILE_WARNING_MB);
        text.inputEl.max = String(MAX_LARGE_FILE_WARNING_MB);
        text.inputEl.step = "1";
      });

    new Setting(containerEl)
      .setName("External opening")
      .setDesc(
        "External opening uses your operating system's default application for .docx files. Change the default app in Windows if you want Word, WPS, or another editor to open documents.",
      );
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
