import {
  App,
  PluginSettingTab,
  type SettingDefinitionItem,
} from "obsidian";

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
  settings: unknown,
): WordReaderSettings {
  const source = isRecord(settings) ? settings : {};
  return {
    language: normalizeLanguage(source.language),
    defaultZoomPercent: clampInteger(
      readNumber(source.defaultZoomPercent, DEFAULT_SETTINGS.defaultZoomPercent),
      MIN_ZOOM_PERCENT,
      MAX_ZOOM_PERCENT,
      DEFAULT_SETTINGS.defaultZoomPercent,
    ),
    defaultFitWidth: readBoolean(
      source.defaultFitWidth,
      DEFAULT_SETTINGS.defaultFitWidth,
    ),
    showOutlineByDefault: readBoolean(
      source.showOutlineByDefault,
      DEFAULT_SETTINGS.showOutlineByDefault,
    ),
    enableImagePreview: readBoolean(
      source.enableImagePreview,
      DEFAULT_SETTINGS.enableImagePreview,
    ),
    largeFileWarningMb: clampInteger(
      readNumber(
        source.largeFileWarningMb,
        DEFAULT_SETTINGS.largeFileWarningMb,
      ),
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

  getSettingDefinitions(): SettingDefinitionItem[] {
    const text = getWordReaderText(this.plugin.settings.language).settings;

    return [
      {
        name: text.languageName,
        desc: text.languageDesc,
        control: {
          type: "dropdown",
          key: "language",
          defaultValue: DEFAULT_SETTINGS.language,
          options: {
            "zh-CN": text.languageZh,
            en: text.languageEn,
          },
        },
      },
      {
        name: text.defaultZoomName,
        desc: text.defaultZoomDesc,
        control: {
          type: "number",
          key: "defaultZoomPercent",
          defaultValue: DEFAULT_SETTINGS.defaultZoomPercent,
          placeholder: "100",
          min: MIN_ZOOM_PERCENT,
          max: MAX_ZOOM_PERCENT,
          step: 5,
        },
      },
      {
        name: text.defaultFitWidthName,
        desc: text.defaultFitWidthDesc,
        control: {
          type: "toggle",
          key: "defaultFitWidth",
          defaultValue: DEFAULT_SETTINGS.defaultFitWidth,
        },
      },
      {
        name: text.showOutlineName,
        desc: text.showOutlineDesc,
        control: {
          type: "toggle",
          key: "showOutlineByDefault",
          defaultValue: DEFAULT_SETTINGS.showOutlineByDefault,
        },
      },
      {
        name: text.imagePreviewName,
        desc: text.imagePreviewDesc,
        control: {
          type: "toggle",
          key: "enableImagePreview",
          defaultValue: DEFAULT_SETTINGS.enableImagePreview,
        },
      },
      {
        name: text.largeFileWarningName,
        desc: text.largeFileWarningDesc,
        control: {
          type: "number",
          key: "largeFileWarningMb",
          defaultValue: DEFAULT_SETTINGS.largeFileWarningMb,
          placeholder: "25",
          min: MIN_LARGE_FILE_WARNING_MB,
          max: MAX_LARGE_FILE_WARNING_MB,
          step: 1,
        },
      },
      {
        name: text.externalOpeningName,
        desc: text.externalOpeningDesc,
      },
    ];
  }

  getControlValue(key: string): unknown {
    return getSettingValue(this.plugin.settings, key);
  }

  async setControlValue(key: string, value: unknown): Promise<void> {
    const previousLanguage = this.plugin.settings.language;
    this.plugin.settings = normalizeSettings({
      ...this.plugin.settings,
      [key]: value,
    });
    await this.plugin.saveSettings();

    if (this.plugin.settings.language !== previousLanguage) {
      this.plugin.refreshWordReaderViews();
      this.update();
    }
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function readNumber(value: unknown, fallback: number): number {
  return typeof value === "number" ? value : fallback;
}

function readBoolean(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function getSettingValue(
  settings: WordReaderSettings,
  key: string,
): unknown {
  switch (key) {
    case "language":
      return settings.language;
    case "defaultZoomPercent":
      return settings.defaultZoomPercent;
    case "defaultFitWidth":
      return settings.defaultFitWidth;
    case "showOutlineByDefault":
      return settings.showOutlineByDefault;
    case "enableImagePreview":
      return settings.enableImagePreview;
    case "largeFileWarningMb":
      return settings.largeFileWarningMb;
    default:
      return undefined;
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
