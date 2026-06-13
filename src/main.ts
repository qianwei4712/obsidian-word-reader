import { Plugin } from "obsidian";

import { openExternalDocx } from "./commands/openExternal";
import { createNoteFromDocx } from "./commands/createNoteFromDocx";
import {
  PptxView,
  VIEW_TYPE_PPTX_READER,
} from "./PptxView";
import { getPptxReaderText } from "./pptx/pptxI18n";
import { WordView, VIEW_TYPE_WORD_READER } from "./WordView";
import {
  DEFAULT_SETTINGS,
  WordReaderSettingTab,
  type WordReaderSettings,
  normalizeSettings,
} from "./settings";
import { getWordReaderText, type WordReaderText } from "./i18n";
import {
  ReadingStateStore,
  type ReaderViewState,
} from "./reader/readingState";

const DATA_SAVE_DEBOUNCE_MS = 500;

export default class WordReaderPlugin extends Plugin {
  settings: WordReaderSettings = DEFAULT_SETTINGS;
  private readingStates = new ReadingStateStore();
  private dataSaveTimer: number | null = null;
  private dataSavePromise: Promise<void> = Promise.resolve();

  async onload(): Promise<void> {
    await this.loadSettings();
    this.addSettingTab(new WordReaderSettingTab(this.app, this));
    const text = this.text;

    this.registerView(
      VIEW_TYPE_WORD_READER,
      (leaf) => new WordView(leaf, this),
    );
    this.registerView(
      VIEW_TYPE_PPTX_READER,
      (leaf) => new PptxView(leaf, this),
    );

    this.registerExtensions(["docx", "doc"], VIEW_TYPE_WORD_READER);
    this.registerExtensions(["pptx"], VIEW_TYPE_PPTX_READER);

    this.addCommand({
      id: "reload",
      name: text.commands.reload,
      checkCallback: (checking) => {
        const view = this.getActiveReaderView();
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
        const view = this.getActiveReaderView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          if (view instanceof PptxView) {
            void view.openExternal();
          } else {
            void openExternalDocx(this.app, view.file, this.text);
          }
        }

        return true;
      },
    });

    const pptxText = getPptxReaderText(this.settings.language);
    this.addCommand({
      id: "previous-slide",
      name: pptxText.commands.previousSlide,
      checkCallback: (checking) => {
        const view = this.getActivePptxView();
        if (!view?.file) {
          return false;
        }
        if (!checking) {
          void view.previousSlide();
        }
        return true;
      },
    });
    this.addCommand({
      id: "next-slide",
      name: pptxText.commands.nextSlide,
      checkCallback: (checking) => {
        const view = this.getActivePptxView();
        if (!view?.file) {
          return false;
        }
        if (!checking) {
          void view.nextSlide();
        }
        return true;
      },
    });
    this.addCommand({
      id: "toggle-presentation-fullscreen",
      name: pptxText.commands.toggleFullscreen,
      checkCallback: (checking) => {
        const view = this.getActivePptxView();
        if (!view?.file) {
          return false;
        }
        if (!checking) {
          void view.toggleFullscreen();
        }
        return true;
      },
    });
    this.addCommand({
      id: "copy-presentation-text",
      name: pptxText.commands.copySlideText,
      checkCallback: (checking) => {
        const view = this.getActivePptxView();
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
      id: "create-presentation-note",
      name: pptxText.commands.createSummaryNote,
      checkCallback: (checking) => {
        const view = this.getActivePptxView();
        if (!view?.file) {
          return false;
        }
        if (!checking) {
          void view.createSummaryNote();
        }
        return true;
      },
    });
    this.addCommand({
      id: "toggle-presentation-notes",
      name: pptxText.commands.toggleNotes,
      checkCallback: (checking) => {
        const view = this.getActivePptxView();
        if (!view?.file) {
          return false;
        }
        if (!checking) {
          view.toggleNotes();
        }
        return true;
      },
    });
    this.addCommand({
      id: "search-presentation",
      name: pptxText.commands.focusSearch,
      checkCallback: (checking) => {
        const view = this.getActivePptxView();
        if (!view?.file) {
          return false;
        }
        if (!checking) {
          view.focusSearch();
        }
        return true;
      },
    });
  }

  onunload(): void {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
      this.dataSaveTimer = null;
      void this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not save plugin data", error);
      });
    }
  }

  private getActiveWordView(): WordView | null {
    return this.app.workspace.getActiveViewOfType(WordView);
  }

  private getActivePptxView(): PptxView | null {
    return this.app.workspace.getActiveViewOfType(PptxView);
  }

  private getActiveReaderView(): WordView | PptxView | null {
    return this.getActiveWordView() ?? this.getActivePptxView();
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
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_PPTX_READER)) {
      if (leaf.view instanceof PptxView) {
        leaf.view.refreshInterfaceLanguage();
      }
    }
  }

  async loadSettings(): Promise<void> {
    const loadedData: unknown = await this.loadData();
    this.settings = normalizeSettings(loadedData);
    this.readingStates = new ReadingStateStore(
      undefined,
      getReadingStateData(loadedData),
    );
  }

  async saveSettings(): Promise<void> {
    this.settings = normalizeSettings(this.settings);
    await this.flushData();
  }

  getReadingState(path: string): ReaderViewState | undefined {
    return this.readingStates.get(path);
  }

  updateReadingState(path: string, state: ReaderViewState): void {
    this.readingStates.set(path, state);
    this.scheduleDataSave();
  }

  async flushData(): Promise<void> {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
      this.dataSaveTimer = null;
    }
    await this.persistData();
  }

  private scheduleDataSave(): void {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
    }
    this.dataSaveTimer = window.setTimeout(() => {
      this.dataSaveTimer = null;
      void this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not save plugin data", error);
      });
    }, DATA_SAVE_DEBOUNCE_MS);
  }

  private persistData(): Promise<void> {
    const data = {
      ...this.settings,
      readingStates: this.readingStates.serialize(),
    };
    this.dataSavePromise = this.dataSavePromise
      .catch(() => undefined)
      .then(async () => {
        await this.saveData(data);
      });
    return this.dataSavePromise;
  }
}

function getReadingStateData(data: unknown): unknown {
  if (typeof data !== "object" || data === null) {
    return undefined;
  }
  return (data as Record<string, unknown>).readingStates;
}
