import {
  App,
  PluginSettingTab,
  Setting,
  type SettingDefinitionItem,
} from "obsidian";

import type WordReaderPlugin from "./main";
import {
  getWordReaderText,
} from "./i18n";
import {
  DEFAULT_OFFICE_READER_SETTINGS,
  DEFAULT_SETTINGS,
  migrateSettings,
  normalizeOfficeReaderSettings,
  normalizeSettings,
  type OfficeReaderSettings,
  type WordReaderLanguage,
  type WordReaderSettings,
} from "./settingsModel";

export {
  DEFAULT_OFFICE_READER_SETTINGS,
  DEFAULT_SETTINGS,
  migrateSettings,
  normalizeOfficeReaderSettings,
  normalizeSettings,
  type OfficeReaderSettings,
  type WordReaderLanguage,
  type WordReaderSettings,
};

const MIN_ZOOM_PERCENT = 25;
const MAX_ZOOM_PERCENT = 400;
const MIN_LARGE_FILE_WARNING_MB = 1;
const MAX_LARGE_FILE_WARNING_MB = 500;

export class WordReaderSettingTab extends PluginSettingTab {
  constructor(
    app: App,
    private readonly plugin: WordReaderPlugin,
  ) {
    super(app, plugin);
  }

  getSettingDefinitions(): SettingDefinitionItem[] {
    const text =
      getWordReaderText(this.plugin.settings.common.language).settings;

    return [
      {
        name: text.languageName,
        desc: text.languageDesc,
        control: {
          type: "dropdown",
          key: "common.language",
          defaultValue: DEFAULT_OFFICE_READER_SETTINGS.common.language,
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
          key: "common.defaultZoomPercent",
          defaultValue:
            DEFAULT_OFFICE_READER_SETTINGS.common.defaultZoomPercent,
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
          key: "docx.defaultFitWidth",
          defaultValue: DEFAULT_OFFICE_READER_SETTINGS.docx.defaultFitWidth,
        },
      },
      {
        name: text.showOutlineName,
        desc: text.showOutlineDesc,
        control: {
          type: "toggle",
          key: "docx.showOutlineByDefault",
          defaultValue:
            DEFAULT_OFFICE_READER_SETTINGS.docx.showOutlineByDefault,
        },
      },
      {
        name: text.xlsxFitWidthName,
        desc: text.xlsxFitWidthDesc,
        control: {
          type: "toggle",
          key: "xlsx.defaultFitWidth",
          defaultValue: DEFAULT_OFFICE_READER_SETTINGS.xlsx.defaultFitWidth,
        },
      },
      {
        name: text.imagePreviewName,
        desc: text.imagePreviewDesc,
        control: {
          type: "toggle",
          key: "docx.enableImagePreview",
          defaultValue:
            DEFAULT_OFFICE_READER_SETTINGS.docx.enableImagePreview,
        },
      },
      {
        name: text.largeFileWarningName,
        desc: text.largeFileWarningDesc,
        control: {
          type: "number",
          key: "common.largeFileWarningMb",
          defaultValue:
            DEFAULT_OFFICE_READER_SETTINGS.common.largeFileWarningMb,
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

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    const text =
      getWordReaderText(this.plugin.settings.common.language).settings;

    new Setting(containerEl)
      .setName(text.languageName)
      .setDesc(text.languageDesc)
      .addDropdown((dropdown) => {
        dropdown
          .addOption("zh-CN", text.languageZh)
          .addOption("en", text.languageEn)
          .setValue(this.plugin.settings.common.language)
          .onChange(async (value) => {
            await this.setControlValue("common.language", value);
          });
      });

    new Setting(containerEl)
      .setName(text.defaultZoomName)
      .setDesc(text.defaultZoomDesc)
      .addText((input) => {
        input
          .setPlaceholder("100")
          .setValue(String(this.plugin.settings.common.defaultZoomPercent))
          .onChange(async (value) => {
            await this.setControlValue(
              "common.defaultZoomPercent",
              Number(value),
            );
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
          .setValue(this.plugin.settings.docx.defaultFitWidth)
          .onChange(async (value) => {
            await this.setControlValue("docx.defaultFitWidth", value);
          });
      });

    new Setting(containerEl)
      .setName(text.showOutlineName)
      .setDesc(text.showOutlineDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.docx.showOutlineByDefault)
          .onChange(async (value) => {
            await this.setControlValue("docx.showOutlineByDefault", value);
          });
      });

    new Setting(containerEl)
      .setName(text.xlsxFitWidthName)
      .setDesc(text.xlsxFitWidthDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.xlsx.defaultFitWidth)
          .onChange(async (value) => {
            await this.setControlValue("xlsx.defaultFitWidth", value);
          });
      });

    new Setting(containerEl)
      .setName(text.imagePreviewName)
      .setDesc(text.imagePreviewDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.docx.enableImagePreview)
          .onChange(async (value) => {
            await this.setControlValue("docx.enableImagePreview", value);
          });
      });

    new Setting(containerEl)
      .setName(text.largeFileWarningName)
      .setDesc(text.largeFileWarningDesc)
      .addText((input) => {
        input
          .setPlaceholder("25")
          .setValue(String(this.plugin.settings.common.largeFileWarningMb))
          .onChange(async (value) => {
            await this.setControlValue(
              "common.largeFileWarningMb",
              Number(value),
            );
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

  getControlValue(key: string): unknown {
    return getSettingValue(this.plugin.settings, key);
  }

  async setControlValue(key: string, value: unknown): Promise<void> {
    const previousLanguage = this.plugin.settings.common.language;
    this.plugin.settings = updateSetting(this.plugin.settings, key, value);
    await this.plugin.saveSettings();

    if (this.plugin.settings.common.language !== previousLanguage) {
      this.plugin.refreshWordReaderViews();
      (this as { update?: () => void }).update?.();
      (this as { display: () => void }).display();
    }
  }
}

function getSettingValue(
  settings: OfficeReaderSettings,
  key: string,
): unknown {
  switch (key) {
    case "common.language":
      return settings.common.language;
    case "common.defaultZoomPercent":
      return settings.common.defaultZoomPercent;
    case "docx.defaultFitWidth":
      return settings.docx.defaultFitWidth;
    case "docx.showOutlineByDefault":
      return settings.docx.showOutlineByDefault;
    case "xlsx.defaultFitWidth":
      return settings.xlsx.defaultFitWidth;
    case "docx.enableImagePreview":
      return settings.docx.enableImagePreview;
    case "common.largeFileWarningMb":
      return settings.common.largeFileWarningMb;
    default:
      return undefined;
  }
}

function updateSetting(
  settings: OfficeReaderSettings,
  key: string,
  value: unknown,
): OfficeReaderSettings {
  const next = {
    ...settings,
    common: { ...settings.common },
    docx: { ...settings.docx },
    pptx: { ...settings.pptx },
    xlsx: { ...settings.xlsx },
  };
  switch (key) {
    case "common.language":
      next.common.language =
        typeof value === "string" ? value as WordReaderLanguage : "zh-CN";
      break;
    case "common.defaultZoomPercent":
      next.common.defaultZoomPercent = Number(value);
      break;
    case "docx.defaultFitWidth":
      next.docx.defaultFitWidth = Boolean(value);
      break;
    case "docx.showOutlineByDefault":
      next.docx.showOutlineByDefault = Boolean(value);
      break;
    case "xlsx.defaultFitWidth":
      next.xlsx.defaultFitWidth = Boolean(value);
      break;
    case "docx.enableImagePreview":
      next.docx.enableImagePreview = Boolean(value);
      break;
    case "common.largeFileWarningMb":
      next.common.largeFileWarningMb = Number(value);
      break;
  }
  return normalizeOfficeReaderSettings(next);
}
