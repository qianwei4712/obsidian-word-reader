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
  DEFAULT_SETTINGS,
  normalizeSettings,
  type WordReaderLanguage,
  type WordReaderSettings,
} from "./settingsModel";

export {
  DEFAULT_SETTINGS,
  normalizeSettings,
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

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    const text = getWordReaderText(this.plugin.settings.language).settings;

    new Setting(containerEl)
      .setName(text.languageName)
      .setDesc(text.languageDesc)
      .addDropdown((dropdown) => {
        dropdown
          .addOption("zh-CN", text.languageZh)
          .addOption("en", text.languageEn)
          .setValue(this.plugin.settings.language)
          .onChange(async (value) => {
            await this.setControlValue("language", value);
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
            await this.setControlValue("defaultZoomPercent", Number(value));
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
            await this.setControlValue("defaultFitWidth", value);
          });
      });

    new Setting(containerEl)
      .setName(text.showOutlineName)
      .setDesc(text.showOutlineDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.showOutlineByDefault)
          .onChange(async (value) => {
            await this.setControlValue("showOutlineByDefault", value);
          });
      });

    new Setting(containerEl)
      .setName(text.imagePreviewName)
      .setDesc(text.imagePreviewDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.enableImagePreview)
          .onChange(async (value) => {
            await this.setControlValue("enableImagePreview", value);
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
            await this.setControlValue("largeFileWarningMb", Number(value));
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
    const previousLanguage = this.plugin.settings.language;
    this.plugin.settings = normalizeSettings({
      ...this.plugin.settings,
      [key]: value,
    });
    await this.plugin.saveSettings();

    if (this.plugin.settings.language !== previousLanguage) {
      this.plugin.refreshWordReaderViews();
      (this as { update?: () => void }).update?.();
      // eslint-disable-next-line @typescript-eslint/no-deprecated -- Obsidian 1.12.7 still uses display() for setting tabs.
      this.display();
    }
  }
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
