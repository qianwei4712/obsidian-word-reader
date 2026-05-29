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

export default class WordReaderPlugin extends Plugin {
  settings: WordReaderSettings = DEFAULT_SETTINGS;

  async onload(): Promise<void> {
    await this.loadSettings();
    this.addSettingTab(new WordReaderSettingTab(this.app, this));

    this.registerView(
      VIEW_TYPE_WORD_READER,
      (leaf) => new WordView(leaf, this),
    );

    this.registerExtensions(["docx", "doc"], VIEW_TYPE_WORD_READER);

    this.addCommand({
      id: "word-reader-reload",
      name: "Reload current Word document",
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
      id: "word-reader-copy-text",
      name: "Copy text from current Word document",
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
      id: "word-reader-create-note",
      name: "Create summary note for current Word document",
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void createNoteFromDocx(this.app, view.file);
        }

        return true;
      },
    });

    this.addCommand({
      id: "word-reader-open-external",
      name: "Open current Word document externally",
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void openExternalDocx(this.app, view.file);
        }

        return true;
      },
    });
  }

  private getActiveWordView(): WordView | null {
    return this.app.workspace.getActiveViewOfType(WordView);
  }

  async loadSettings(): Promise<void> {
    this.settings = normalizeSettings(await this.loadData());
  }

  async saveSettings(): Promise<void> {
    this.settings = normalizeSettings(this.settings);
    await this.saveData(this.settings);
  }
}
