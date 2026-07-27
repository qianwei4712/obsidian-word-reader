import { Plugin, TFile } from "obsidian";

import {
  PptxView,
  VIEW_TYPE_PPTX_READER,
} from "./PptxView";
import {
  WordView,
  VIEW_TYPE_WORD_READER,
} from "./WordView";
import { DOCX_ADAPTER } from "./docx/DocxAdapter";
import { getWordReaderText, type WordReaderText } from "./i18n";
import { PPTX_ADAPTER } from "./pptx/PptxAdapter";
import { getPptxReaderText } from "./pptx/pptxI18n";
import {
  ReadingStateStore,
  type ReaderFileIdentity,
  type ReaderFormat,
  type ReaderViewState,
} from "./reader/readingState";
import {
  isReaderSession,
  type ReaderSession,
} from "./reader/session";
import {
  DEFAULT_OFFICE_READER_SETTINGS,
  WordReaderSettingTab,
  type OfficeReaderSettings,
  migrateSettings,
  normalizeOfficeReaderSettings,
} from "./settings";

const DATA_SAVE_DEBOUNCE_MS = 500;

export default class WordReaderPlugin extends Plugin {
  settings: OfficeReaderSettings = normalizeOfficeReaderSettings(
    DEFAULT_OFFICE_READER_SETTINGS,
  );
  private readingStates = new ReadingStateStore();
  private dataSaveTimer: number | null = null;
  private dataSavePromise: Promise<void> = Promise.resolve();

  async onload(): Promise<void> {
    await this.loadSettings();
    this.addSettingTab(new WordReaderSettingTab(this.app, this));
    const text = this.text;
    const pptxText = getPptxReaderText(this.settings.common.language);

    this.registerView(
      VIEW_TYPE_WORD_READER,
      (leaf) => new WordView(leaf, this),
    );
    this.registerView(
      VIEW_TYPE_PPTX_READER,
      (leaf) => new PptxView(leaf, this),
    );
    this.registerExtensions(
      [...DOCX_ADAPTER.extensions],
      DOCX_ADAPTER.viewType,
    );
    this.registerExtensions(
      [...PPTX_ADAPTER.extensions],
      PPTX_ADAPTER.viewType,
    );

    this.registerReaderCommand(
      "reload",
      text.commands.reload,
      (session) => session.capabilities.reload,
      (session) => session.reload(),
    );
    this.registerReaderCommand(
      "copy-text",
      text.commands.copyText,
      (session) => session.capabilities.copyText && Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-markdown",
      text.commands.copyMarkdown,
      (session) =>
        session.capabilities.copyMarkdown && Boolean(session.copyMarkdown),
      (session) => session.copyMarkdown?.(),
    );
    this.registerReaderCommand(
      "create-note",
      text.commands.createNote,
      (session) =>
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
    this.registerReaderCommand(
      "open-external",
      text.commands.openExternal,
      (session) =>
        session.capabilities.openExternal && Boolean(session.openExternal),
      (session) => session.openExternal?.(),
    );

    this.registerReaderCommand(
      "previous-slide",
      pptxText.commands.previousSlide,
      (session) =>
        session.capabilities.paged && Boolean(session.previousPage),
      (session) => session.previousPage?.(),
    );
    this.registerReaderCommand(
      "next-slide",
      pptxText.commands.nextSlide,
      (session) =>
        session.capabilities.paged && Boolean(session.nextPage),
      (session) => session.nextPage?.(),
    );
    this.registerReaderCommand(
      "toggle-presentation-fullscreen",
      pptxText.commands.toggleFullscreen,
      (session) =>
        session.capabilities.fullscreen && Boolean(session.toggleFullscreen),
      (session) => session.toggleFullscreen?.(),
    );
    this.registerReaderCommand(
      "copy-presentation-text",
      pptxText.commands.copySlideText,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.copyText &&
        Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-presentation-render-diagnostics",
      pptxText.commands.copyRenderDiagnostics,
      (session) =>
        session.capabilities.diagnostics && Boolean(session.copyDiagnostics),
      (session) => session.copyDiagnostics?.(),
    );
    this.registerReaderCommand(
      "create-presentation-note",
      pptxText.commands.createSummaryNote,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
    this.registerReaderCommand(
      "toggle-presentation-notes",
      pptxText.commands.toggleNotes,
      (session) =>
        session.capabilities.notes && Boolean(session.toggleNotes),
      (session) => session.toggleNotes?.(),
    );
    this.registerReaderCommand(
      "search-presentation",
      pptxText.commands.focusSearch,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.search &&
        Boolean(session.focusSearch),
      (session) => session.focusSearch?.(),
    );
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

  get text(): WordReaderText {
    return getWordReaderText(this.settings.common.language);
  }

  refreshWordReaderViews(): void {
    for (const viewType of [
      VIEW_TYPE_WORD_READER,
      VIEW_TYPE_PPTX_READER,
    ]) {
      for (const leaf of this.app.workspace.getLeavesOfType(viewType)) {
        if (isReaderSession(leaf.view)) {
          leaf.view.refreshInterfaceLanguage();
        }
      }
    }
  }

  async loadSettings(): Promise<void> {
    const loadedData: unknown = await this.loadData();
    this.settings = migrateSettings(loadedData);
    this.readingStates = new ReadingStateStore(
      undefined,
      getReadingStateData(loadedData),
    );
    if (!hasCurrentSettingsSchema(loadedData)) {
      await this.persistData();
    }
  }

  async saveSettings(): Promise<void> {
    this.settings = normalizeOfficeReaderSettings(this.settings);
    await this.flushData();
  }

  getReadingState(
    file: TFile | ReaderFileIdentity | string,
  ): ReaderViewState | undefined {
    return typeof file === "string"
      ? this.readingStates.get(file)
      : this.readingStates.get(toReaderIdentity(file));
  }

  updateReadingState(
    file: TFile | ReaderFileIdentity | string,
    state: ReaderViewState,
  ): void {
    if (typeof file === "string") {
      this.readingStates.set(file, state);
    } else {
      this.readingStates.set(toReaderIdentity(file), state);
    }
    this.scheduleDataSave();
  }

  async flushData(): Promise<void> {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
      this.dataSaveTimer = null;
    }
    await this.persistData();
  }

  private registerReaderCommand(
    id: string,
    name: string,
    isAvailable: (session: ReaderSession) => boolean,
    run: (session: ReaderSession) => void | Promise<void>,
  ): void {
    this.addCommand({
      id,
      name,
      checkCallback: (checking) => {
        const session = this.getActiveReaderSession();
        if (!session?.file || !isAvailable(session)) {
          return false;
        }
        if (!checking) {
          void run(session);
        }
        return true;
      },
    });
  }

  private getActiveReaderSession(): ReaderSession | null {
    const view = this.app.workspace.getMostRecentLeaf()?.view;
    return isReaderSession(view) ? view : null;
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

function hasCurrentSettingsSchema(data: unknown): boolean {
  if (typeof data !== "object" || data === null) {
    return false;
  }
  const source = data as Record<string, unknown>;
  const settingsAreCurrent =
    source.schemaVersion === DEFAULT_OFFICE_READER_SETTINGS.schemaVersion &&
    typeof source.common === "object" &&
    source.common !== null &&
    typeof source.docx === "object" &&
    source.docx !== null &&
    typeof source.pptx === "object" &&
    source.pptx !== null &&
    typeof source.xlsx === "object" &&
    source.xlsx !== null;
  if (!settingsAreCurrent) {
    return false;
  }
  const readingStates = source.readingStates;
  return (
    readingStates === undefined ||
    (
      Array.isArray(readingStates) &&
      readingStates.every(isCurrentReadingStateEntry)
    )
  );
}

function isCurrentReadingStateEntry(value: unknown): boolean {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const entry = value as Record<string, unknown>;
  return (
    typeof entry.path === "string" &&
    typeof entry.mtime === "number" &&
    typeof entry.format === "string" &&
    typeof entry.position === "object" &&
    entry.position !== null &&
    typeof entry.zoom === "object" &&
    entry.zoom !== null &&
    typeof entry.navigation === "object" &&
    entry.navigation !== null
  );
}

function toReaderIdentity(
  file: TFile | ReaderFileIdentity,
): ReaderFileIdentity {
  if ("format" in file) {
    return file;
  }
  return {
    path: file.path,
    mtime: file.stat.mtime,
    format: getReaderFormat(file.extension),
  };
}

function getReaderFormat(extension: string): ReaderFormat {
  switch (extension.toLowerCase()) {
    case "pptx":
      return "pptx";
    case "xlsx":
      return "xlsx";
    default:
      return "docx";
  }
}
