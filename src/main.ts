import { Plugin } from "obsidian";

import { openExternalDocx } from "./commands/openExternal";
import { createNoteFromDocx } from "./commands/createNoteFromDocx";
import { WordView, VIEW_TYPE_WORD_READER } from "./WordView";
import {
  DEFAULT_SETTINGS,
  WordReaderSettingTab,
  type WordReaderSettings,
  normalizeSettings,
} from "./settings";
import { getWordReaderText, type WordReaderText } from "./i18n";

export default class WordReaderPlugin extends Plugin {
  settings: WordReaderSettings = DEFAULT_SETTINGS;

  async onload(): Promise<void> {
    await this.loadSettings();
    this.addSettingTab(new WordReaderSettingTab(this.app, this));
    const text = this.text;

    this.registerView(
      VIEW_TYPE_WORD_READER,
      (leaf) => new WordView(leaf, this),
    );

    this.registerExtensions(["docx", "doc"], VIEW_TYPE_WORD_READER);

    this.addCommand({
      id: "reload",
      name: text.commands.reload,
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void view.reload();
        }

        return true;
      },
    });

    this.addCommand({
      id: "copy-text",
      name: text.commands.copyText,
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void view.copyText();
        }

        return true;
      },
    });

    this.addCommand({
      id: "copy-markdown",
      name: text.commands.copyMarkdown,
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void view.copyMarkdown();
        }

        return true;
      },
    });

    this.addCommand({
      id: "create-note",
      name: text.commands.createNote,
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void createNoteFromDocx(this.app, view.file, this.text);
        }

        return true;
      },
    });

    this.addCommand({
      id: "open-external",
      name: text.commands.openExternal,
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void openExternalDocx(this.app, view.file, this.text);
        }

        return true;
      },
    });
  }

  private getActiveWordView(): WordView | null {
    return this.app.workspace.getActiveViewOfType(WordView);
  }

  get text(): WordReaderText {
    return getWordReaderText(this.settings.language);
  }

  refreshWordReaderViews(): void {
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_WORD_READER)) {
      if (leaf.view instanceof WordView) {
        leaf.view.refreshInterfaceLanguage();
      }
    }
  }

  async loadSettings(): Promise<void> {
    const loadedSettings: unknown = await this.loadData();
    this.settings = normalizeSettings(loadedSettings);
  }

  async saveSettings(): Promise<void> {
    this.settings = normalizeSettings(this.settings);
    await this.saveData(this.settings);
  }
}
