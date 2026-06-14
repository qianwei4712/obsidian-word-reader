import {
  FileView,
  Notice,
  TFile,
  WorkspaceLeaf,
  setIcon,
} from "obsidian";

import type WordReaderPlugin from "./main";
import { createNoteFromPptx } from "./commands/createNoteFromPptx";
import { openExternalFile } from "./commands/openExternal";
import {
  createPptxDiagnosticReport,
  formatPptxDiagnosticReport,
} from "./pptx/pptxDiagnostics";
import { PptxPackage } from "./pptx/pptxPackage";
import { classifyPptxError } from "./pptx/pptxErrors";
import {
  getPptxReaderText,
  type PptxReaderText,
} from "./pptx/pptxI18n";
import {
  formatSlideText,
  PptxSearchIndex,
  type PptxSlideMetadata,
} from "./pptx/pptxMetadata";
import { PptxNavigationWindow } from "./pptx/pptxNavigationWindow";
import {
  PptxRenderCancelledError,
  renderPptxSlide,
  type PptxRenderDiagnostics,
} from "./pptx/pptxRenderer";
import { PptxTaskQueue } from "./pptx/pptxTaskQueue";
import { ReaderLifecycle } from "./reader/lifecycle";
import type { ReaderViewState } from "./reader/readingState";
import {
  releaseResources,
  RetainedResourceRegistry,
} from "./reader/resources";
import {
  ReaderStatusController,
  type ReaderStatus,
} from "./reader/status";
import { normalizeZoom, preserveZoomAnchor } from "./reader/zoom";

export const VIEW_TYPE_PPTX_READER = "pptx-reader-view";

declare const __DEV__: boolean;

const MIN_ZOOM = 0.1;
const MAX_ZOOM = 4;
const ZOOM_STEP = 0.05;
const FIT_PADDING = 48;
const THUMBNAIL_WIDTH = 140;
const PPTX_SEARCH_DEBOUNCE_MS = 150;

interface RenderOptions {
  restoreScroll?: boolean;
}

interface PptxNavigationEntry {
  element: HTMLElement;
  titleEl: HTMLElement;
  snippetEl: HTMLElement;
  matchEl: HTMLElement;
  thumbnailEl: HTMLElement;
  resources: Set<string>;
  renderVersion: number;
  rendering: boolean;
  rendered: boolean;
  mounted: boolean;
}

export class PptxView extends FileView {
  private readonly plugin: WordReaderPlugin;
  private rootEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private bodyEl: HTMLElement | null = null;
  private navigationEl: HTMLElement | null = null;
  private slideListEl: HTMLElement | null = null;
  private navigationEmptyEl: HTMLElement | null = null;
  private navigationTopSpacerEl: HTMLElement | null = null;
  private navigationRowsEl: HTMLElement | null = null;
  private navigationBottomSpacerEl: HTMLElement | null = null;
  private searchInputEl: HTMLInputElement | null = null;
  private searchCountEl: HTMLElement | null = null;
  private viewportEl: HTMLElement | null = null;
  private canvasEl: HTMLElement | null = null;
  private notesEl: HTMLElement | null = null;
  private notesContentEl: HTMLElement | null = null;
  private stageEl: HTMLElement | null = null;
  private pageInputEl: HTMLInputElement | null = null;
  private pageCountEl: HTMLElement | null = null;
  private zoomInputEl: HTMLInputElement | null = null;
  private previousButtonEl: HTMLButtonElement | null = null;
  private nextButtonEl: HTMLButtonElement | null = null;
  private fitButtonEl: HTMLButtonElement | null = null;
  private fullscreenButtonEl: HTMLButtonElement | null = null;
  private navigationButtonEl: HTMLButtonElement | null = null;
  private notesButtonEl: HTMLButtonElement | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private presentation: PptxPackage | null = null;
  private slideMetadata: PptxSlideMetadata[] = [];
  private readonly loadedMetadataIndices = new Set<number>();
  private metadataPromise: Promise<PptxSlideMetadata[]> | null = null;
  private readonly searchIndex = new PptxSearchIndex();
  private filteredSlideIndices: number[] = [];
  private navigationResults = new Map<
    number,
    ReturnType<PptxSearchIndex["search"]>[number]
  >();
  private readonly navigationEntries = new Map<number, PptxNavigationEntry>();
  private readonly navigationWindow = new PptxNavigationWindow();
  private readonly thumbnailQueue = new PptxTaskQueue(2);
  private navigationRenderFrameId: number | null = null;
  private searchTimer: number | null = null;
  private metadataUiFrameId: number | null = null;
  private currentSlideIndex = 0;
  private activeFilePath: string | null = null;
  private pendingScrollPosition: { left: number; top: number } | null = null;
  private zoom = 1;
  private fitWindow = true;
  private navigationVisible = true;
  private notesVisible = false;
  private searchQuery = "";
  private renderedWidth = 0;
  private renderedHeight = 0;
  private currentRenderDiagnostics: PptxRenderDiagnostics | null = null;
  private readonly loadLifecycle = new ReaderLifecycle();
  private readonly slideLifecycle = new ReaderLifecycle();
  private readonly thumbnailLifecycle = new ReaderLifecycle();
  private readonly slideResources = new RetainedResourceRegistry((resource) => {
    URL.revokeObjectURL(resource);
  });
  private readonly readerStatus = new ReaderStatusController((status) => {
    this.renderStatus(status);
  });
  private readonly handleFullscreenChange = (): void => {
    this.updateFullscreenControl();
    this.applyScale();
  };

  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf);
    this.plugin = plugin;
    this.zoom = plugin.settings.defaultZoomPercent / 100;
    this.fitWindow = true;
  }

  getViewType(): string {
    return VIEW_TYPE_PPTX_READER;
  }

  getDisplayText(): string {
    return this.file?.name ?? this.text.displayName;
  }

  getIcon(): string {
    return "presentation";
  }

  private get text(): PptxReaderText {
    return getPptxReaderText(this.plugin.settings.language);
  }

  async onOpen(): Promise<void> {
    this.buildLayout();
    this.contentEl.doc.addEventListener(
      "fullscreenchange",
      this.handleFullscreenChange,
    );
    const resizeObserver = new ResizeObserver(() => {
      if (this.fitWindow) {
        this.applyScale();
      }
    });
    resizeObserver.observe(this.contentEl);
    this.resizeObserver = resizeObserver;
  }

  async onClose(): Promise<void> {
    this.contentEl.doc.removeEventListener(
      "fullscreenchange",
      this.handleFullscreenChange,
    );
    this.resizeObserver?.disconnect();
    this.resizeObserver = null;
    this.loadLifecycle.cancel();
    this.slideLifecycle.cancel();
    this.thumbnailLifecycle.cancel();
    this.releasePresentation();
    this.clearCanvas();
  }

  async onLoadFile(file: TFile): Promise<void> {
    this.searchQuery = "";
    if (this.searchInputEl) {
      this.searchInputEl.value = "";
    }
    this.restoreReadingState(file.path);
    await this.loadPresentation(file);
  }

  async onUnloadFile(): Promise<void> {
    this.saveReadingState();
    await this.plugin.flushData();
    this.activeFilePath = null;
    this.loadLifecycle.cancel();
    this.slideLifecycle.cancel();
    this.thumbnailLifecycle.cancel();
    this.releasePresentation();
    this.clearCanvas();
    this.setStatus("");
  }

  async reload(): Promise<void> {
    if (this.file) {
      await this.loadPresentation(this.file);
    }
  }

  async openExternal(): Promise<void> {
    if (!this.file) {
      return;
    }
    try {
      await openExternalFile(this.app, this.file, {
        desktopOnly: this.text.notices.externalDesktopOnly,
        localVaultOnly: this.text.notices.externalLocalVaultOnly,
      });
    } catch (error) {
      new Notice(this.text.notices.externalFailed(getErrorMessage(error)));
    }
  }

  async previousSlide(): Promise<void> {
    await this.goToSlide(this.currentSlideIndex - 1);
  }

  async nextSlide(): Promise<void> {
    await this.goToSlide(this.currentSlideIndex + 1);
  }

  async copyText(): Promise<void> {
    const metadata = await this.getCurrentSlideMetadata();
    if (!metadata) {
      return;
    }
    try {
      const selectedText = this.getSelectedSlideText();
      await navigator.clipboard.writeText(
        selectedText || formatSlideText(metadata),
      );
      new Notice(
        selectedText
          ? this.text.notices.copiedSelectedText
          : this.text.notices.copiedSlideText,
      );
    } catch (error) {
      new Notice(this.text.notices.copyFailed(getErrorMessage(error)));
    }
  }

  async copyRenderDiagnostics(): Promise<void> {
    const file = this.file;
    const presentation = this.presentation;
    const render = this.currentRenderDiagnostics;
    if (!file || !presentation || !render) {
      new Notice(this.text.notices.noRenderDiagnostics);
      return;
    }
    try {
      const entries = [...this.navigationEntries.values()];
      const report = createPptxDiagnosticReport(
        file,
        presentation,
        this.currentSlideIndex,
        render,
        {
          mounted: entries.filter((entry) => entry.mounted).length,
          rendered: entries.filter((entry) => entry.rendered).length,
          rendering: entries.filter((entry) => entry.rendering).length,
          resourceCount: entries.reduce(
            (count, entry) => count + entry.resources.size,
            0,
          ),
        },
      );
      await navigator.clipboard.writeText(
        formatPptxDiagnosticReport(report),
      );
      new Notice(this.text.notices.copiedRenderDiagnostics);
    } catch (error) {
      new Notice(this.text.notices.copyFailed(getErrorMessage(error)));
    }
  }

  async createSummaryNote(): Promise<void> {
    const presentation = this.presentation;
    if (!this.file || !presentation) {
      return;
    }
    const metadata = await this.getCompleteMetadata(presentation);
    if (metadata.length === 0 || this.presentation !== presentation) {
      return;
    }
    await createNoteFromPptx(
      this.app,
      this.file,
      metadata,
      this.currentSlideIndex,
      this.text,
    );
  }

  toggleNotes(): void {
    this.notesVisible = !this.notesVisible;
    this.applyNotesVisibility();
    this.saveReadingState();
  }

  focusSearch(): void {
    if (!this.navigationVisible) {
      this.navigationVisible = true;
      this.applyNavigationVisibility();
      this.saveReadingState();
    }
    this.searchInputEl?.focus();
    this.searchInputEl?.select();
  }

  async toggleFullscreen(): Promise<void> {
    if (!this.rootEl) {
      return;
    }
    try {
      if (this.contentEl.doc.fullscreenElement === this.rootEl) {
        await this.contentEl.doc.exitFullscreen();
      } else {
        await this.rootEl.requestFullscreen();
      }
    } catch (error) {
      new Notice(this.text.notices.fullscreenFailed(getErrorMessage(error)));
    }
  }

  refreshInterfaceLanguage(): void {
    const file = this.file;
    this.saveReadingState();
    this.slideLifecycle.cancel();
    this.thumbnailLifecycle.cancel();
    this.thumbnailQueue.clear();
    this.releaseThumbnailResources();
    this.buildLayout();
    if (file && this.presentation) {
      this.thumbnailLifecycle.begin();
      this.buildNavigationList();
      this.updateNotes();
      void this.renderCurrentSlide({ restoreScroll: true });
      this.refreshThumbnailQueue();
    }
  }

  private buildLayout(): void {
    this.contentEl.empty();
    const text = this.text;
    this.rootEl = this.contentEl.createDiv({
      cls: "word-reader-root pptx-reader-root",
    });
    this.rootEl.tabIndex = 0;
    this.rootEl.addEventListener("keydown", (event) => {
      this.handleKeyDown(event);
    });

    const toolbarEl = this.rootEl.createDiv({
      cls: "word-reader-toolbar pptx-reader-toolbar",
    });
    this.createIconButton(toolbarEl, "refresh-cw", text.toolbar.reload, () => {
      void this.reload();
    });
    this.navigationButtonEl = this.createIconButton(
      toolbarEl,
      "panel-left",
      text.toolbar.hideNavigation,
      () => {
        this.navigationVisible = !this.navigationVisible;
        this.applyNavigationVisibility();
        this.saveReadingState();
      },
    );
    this.previousButtonEl = this.createIconButton(
      toolbarEl,
      "chevron-left",
      text.toolbar.previousSlide,
      () => {
        void this.previousSlide();
      },
    );
    this.nextButtonEl = this.createIconButton(
      toolbarEl,
      "chevron-right",
      text.toolbar.nextSlide,
      () => {
        void this.nextSlide();
      },
    );

    const pageControlEl = toolbarEl.createDiv({
      cls: "pptx-reader-page-control",
    });
    this.pageInputEl = pageControlEl.createEl("input", {
      cls: "pptx-reader-page-input",
      attr: {
        type: "number",
        min: "1",
        step: "1",
        "aria-label": text.toolbar.pageNumber,
        title: text.toolbar.pageNumber,
      },
    });
    this.pageInputEl.addEventListener("change", () => {
      const page = Number(this.pageInputEl?.value);
      if (Number.isInteger(page)) {
        void this.goToSlide(page - 1);
      } else {
        this.updateNavigationControls();
      }
    });
    this.pageCountEl = pageControlEl.createSpan({
      cls: "pptx-reader-page-count",
      text: "/ 0",
    });

    this.zoomInputEl = toolbarEl.createEl("input", {
      cls: "word-reader-zoom",
      attr: {
        type: "number",
        min: String(MIN_ZOOM * 100),
        max: String(MAX_ZOOM * 100),
        step: String(ZOOM_STEP * 100),
        "aria-label": text.toolbar.zoomPercentage,
        title: text.toolbar.zoomPercentage,
      },
    });
    this.zoomInputEl.addEventListener("change", () => {
      this.setZoom(Number(this.zoomInputEl?.value) / 100);
    });

    this.fitButtonEl = this.createIconButton(
      toolbarEl,
      "maximize-2",
      text.toolbar.fitWindow,
      () => {
        this.fitWindow = !this.fitWindow;
        this.applyScale();
        this.saveReadingState();
      },
    );
    this.createIconButton(toolbarEl, "copy", text.toolbar.copyText, () => {
      void this.copyText();
    });
    this.createIconButton(
      toolbarEl,
      "activity",
      text.toolbar.copyRenderDiagnostics,
      () => {
        void this.copyRenderDiagnostics();
      },
    );
    this.createIconButton(
      toolbarEl,
      "notebook-pen",
      text.toolbar.createSummaryNote,
      () => {
        void this.createSummaryNote();
      },
    );
    this.notesButtonEl = this.createIconButton(
      toolbarEl,
      "sticky-note",
      text.toolbar.showNotes,
      () => {
        this.toggleNotes();
      },
    );
    this.fullscreenButtonEl = this.createIconButton(
      toolbarEl,
      "fullscreen",
      text.toolbar.enterFullscreen,
      () => {
        void this.toggleFullscreen();
      },
    );
    this.createIconButton(
      toolbarEl,
      "external-link",
      text.toolbar.openExternally,
      () => {
        void this.openExternal();
      },
    );

    this.statusEl = this.rootEl.createDiv({
      cls: "word-reader-status pptx-reader-status",
    });

    this.bodyEl = this.rootEl.createDiv({
      cls: "pptx-reader-body",
    });
    this.navigationEl = this.bodyEl.createDiv({
      cls: "pptx-reader-navigation",
    });
    const navigationHeaderEl = this.navigationEl.createDiv({
      cls: "pptx-reader-navigation-header",
    });
    navigationHeaderEl.createDiv({
      cls: "pptx-reader-navigation-title",
      text: text.navigation.title,
    });
    this.searchInputEl = navigationHeaderEl.createEl("input", {
      cls: "pptx-reader-search",
      attr: {
        type: "search",
        placeholder: text.toolbar.searchPlaceholder,
        "aria-label": text.toolbar.searchPresentation,
      },
    });
    this.searchInputEl.value = this.searchQuery;
    this.searchInputEl.addEventListener("input", () => {
      this.searchQuery = this.searchInputEl?.value ?? "";
      this.scheduleNavigationSearch();
    });
    this.searchInputEl.addEventListener("keydown", (event) => {
      if (event.key === "Enter" && this.searchQuery.trim()) {
        const firstResult = this.searchIndex.search(
          this.searchQuery,
          this.allSlideIndices,
        )[0];
        if (firstResult) {
          event.preventDefault();
          void this.goToSlide(firstResult.slideIndex);
        }
      }
    });
    this.searchCountEl = navigationHeaderEl.createDiv({
      cls: "pptx-reader-search-count",
    });
    this.slideListEl = this.navigationEl.createDiv({
      cls: "pptx-reader-slide-list",
    });
    this.slideListEl.addEventListener("scroll", () => {
      this.scheduleNavigationWindowRender();
    });
    this.navigationTopSpacerEl = this.slideListEl.createDiv({
      cls: "pptx-reader-navigation-spacer",
    });
    this.navigationRowsEl = this.slideListEl.createDiv({
      cls: "pptx-reader-navigation-rows",
    });
    this.navigationBottomSpacerEl = this.slideListEl.createDiv({
      cls: "pptx-reader-navigation-spacer",
    });
    this.navigationEmptyEl = this.navigationEl.createDiv({
      cls: "pptx-reader-navigation-empty",
      text: text.navigation.noMatches,
    });

    const mainEl = this.bodyEl.createDiv({
      cls: "pptx-reader-main",
    });
    this.viewportEl = mainEl.createDiv({
      cls: "pptx-reader-viewport",
    });
    this.viewportEl.addEventListener("wheel", (event) => {
      this.handleWheelZoom(event);
    });
    this.canvasEl = this.viewportEl.createDiv({
      cls: "pptx-reader-canvas",
    });
    this.notesEl = mainEl.createDiv({
      cls: "pptx-reader-notes",
    });
    this.notesEl.createDiv({
      cls: "pptx-reader-notes-title",
      text: text.notes.title,
    });
    this.notesContentEl = this.notesEl.createDiv({
      cls: "pptx-reader-notes-content",
    });

    this.updateNavigationControls();
    this.applyNavigationVisibility();
    this.applyNotesVisibility();
    this.applyNavigationSearch();
    this.updateNotes();
    this.updateFullscreenControl();
    this.applyScale();
    this.readerStatus.refresh();
  }

  private handleKeyDown(event: KeyboardEvent): void {
    const target = event.target as Node | null;
    if (
      target?.instanceOf(HTMLElement) &&
      target.closest("input, textarea, select, button, a, [contenteditable='true']")
    ) {
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "f") {
      event.preventDefault();
      this.focusSearch();
      return;
    }
    if (event.altKey || event.ctrlKey || event.metaKey) {
      return;
    }

    switch (event.key) {
      case "ArrowLeft":
      case "PageUp":
        event.preventDefault();
        void this.previousSlide();
        break;
      case "ArrowRight":
      case "PageDown":
        event.preventDefault();
        void this.nextSlide();
        break;
      case " ":
        event.preventDefault();
        void (event.shiftKey ? this.previousSlide() : this.nextSlide());
        break;
      case "Home":
        event.preventDefault();
        void this.goToSlide(0);
        break;
      case "End":
        event.preventDefault();
        void this.goToSlide((this.presentation?.slideCount ?? 1) - 1);
        break;
    }
  }

  private async loadPresentation(file: TFile): Promise<void> {
    const token = this.loadLifecycle.begin();
    this.slideLifecycle.cancel();
    this.thumbnailLifecycle.cancel();
    this.releasePresentation();
    this.clearCanvas();
    this.setStatus(this.text.status.reading(file.name), false, true);

    try {
      const buffer = await this.app.vault.readBinary(file);
      if (!this.loadLifecycle.isCurrent(token)) {
        return;
      }
      const presentation = await PptxPackage.load(buffer);
      if (!this.loadLifecycle.isCurrent(token)) {
        return;
      }
      this.presentation = presentation;
      this.currentSlideIndex = clamp(
        this.currentSlideIndex,
        0,
        presentation.slideCount - 1,
      );
      this.slideMetadata = this.allSlideIndices.map((index) =>
        createEmptySlideMetadata(index),
      );
      this.loadedMetadataIndices.clear();
      this.searchIndex.clear();
      this.thumbnailLifecycle.begin();
      this.buildNavigationList();
      this.updateNotes();
      this.updateNavigationControls();
      this.metadataPromise = presentation.indexSlideMetadata({
        concurrency: 4,
        priorityIndex: this.currentSlideIndex,
        isCancelled: () =>
          !this.loadLifecycle.isCurrent(token) ||
          this.presentation !== presentation,
        onMetadata: (metadata) => {
          this.slideMetadata[metadata.index] = metadata;
          this.loadedMetadataIndices.add(metadata.index);
          this.searchIndex.set(metadata);
          this.scheduleMetadataUiRefresh();
        },
      });
      void this.metadataPromise.catch((error: unknown) => {
        if (
          this.loadLifecycle.isCurrent(token) &&
          this.presentation === presentation &&
          error instanceof Error &&
          error.name !== "PptxMetadataCancelledError" &&
          typeof __DEV__ !== "undefined" &&
          __DEV__
        ) {
          console.debug("[Office Reader] PPTX metadata indexing stopped", error);
        }
      });
      await this.renderCurrentSlide({ restoreScroll: true });
      this.refreshThumbnailQueue();
    } catch (error) {
      if (!this.loadLifecycle.isCurrent(token)) {
        return;
      }
      this.releasePresentation();
      this.showError(error);
    }
  }

  private async goToSlide(index: number): Promise<void> {
    if (!this.presentation) {
      return;
    }
    const nextIndex = clamp(index, 0, this.presentation.slideCount - 1);
    if (
      nextIndex === this.currentSlideIndex &&
      this.stageEl?.isConnected
    ) {
      this.updateNavigationControls();
      return;
    }
    this.currentSlideIndex = nextIndex;
    this.pendingScrollPosition = { left: 0, top: 0 };
    this.updateNavigationControls();
    this.saveReadingState();
    void this.loadMetadataForSlide(nextIndex);
    await this.renderCurrentSlide({ restoreScroll: true });
  }

  private async renderCurrentSlide(
    options: RenderOptions = {},
  ): Promise<void> {
    const presentation = this.presentation;
    const file = this.file;
    if (!presentation || !file || !this.canvasEl) {
      return;
    }

    const token = this.slideLifecycle.begin();
    const index = this.currentSlideIndex;
    this.thumbnailQueue.setPaused(true, true);
    this.clearCanvas();
    this.setStatus(
      this.text.status.rendering(index + 1, presentation.slideCount),
      false,
      true,
    );

    let resources = new Set<string>();
    let adoptedResources = false;
    try {
      const slideContext = await presentation.getSlideContext(index);
      if (
        !this.slideLifecycle.isCurrent(token) ||
        this.presentation !== presentation
      ) {
        return;
      }
      const rendered = await renderPptxSlide(
        presentation,
        slideContext,
        this.canvasEl.ownerDocument,
        {
          isCancelled: () =>
            !this.slideLifecycle.isCurrent(token) ||
            this.presentation !== presentation,
        },
      );
      resources = rendered.resources;
      if (
        !this.slideLifecycle.isCurrent(token) ||
        this.presentation !== presentation ||
        !this.canvasEl
      ) {
        return;
      }

      this.slideResources.replace(resources);
      adoptedResources = true;
      this.stageEl = rendered.element;
      this.renderedWidth = rendered.width;
      this.renderedHeight = rendered.height;
      this.currentRenderDiagnostics = rendered.diagnostics;
      this.canvasEl.appendChild(rendered.element);
      logPptxRenderPerformance(
        file.name,
        index,
        rendered.diagnostics,
      );
      this.applyScale();
      if (options.restoreScroll) {
        this.restoreScrollPosition();
      }
      this.setStatus(
        this.text.status.ready(
          file.name,
          index + 1,
          presentation.slideCount,
        ),
      );
    } catch (error) {
      if (error instanceof PptxRenderCancelledError) {
        return;
      }
      if (
        this.slideLifecycle.isCurrent(token) &&
        this.presentation === presentation
      ) {
        this.showError(error);
      }
    } finally {
      if (!adoptedResources) {
        releaseResources(resources, (resource) => {
          URL.revokeObjectURL(resource);
        });
      }
      if (
        this.slideLifecycle.isCurrent(token) &&
        this.presentation === presentation
      ) {
        this.thumbnailQueue.setPaused(false);
        this.refreshThumbnailQueue();
      }
    }
  }

  private buildNavigationList(): void {
    this.releaseThumbnailResources();
    this.navigationEntries.clear();
    this.navigationRowsEl?.empty();
    this.filteredSlideIndices = this.allSlideIndices;
    if (!this.navigationRowsEl) {
      return;
    }
    this.applyNavigationSearch();
  }

  private scheduleNavigationSearch(): void {
    if (this.searchTimer !== null) {
      window.clearTimeout(this.searchTimer);
    }
    this.searchTimer = window.setTimeout(() => {
      this.searchTimer = null;
      this.applyNavigationSearch();
    }, PPTX_SEARCH_DEBOUNCE_MS);
  }

  private applyNavigationSearch(resetScroll = true): void {
    const results = this.searchIndex.search(
      this.searchQuery,
      this.allSlideIndices,
    );
    this.navigationResults = new Map(
      results.map((result) => [result.slideIndex, result]),
    );
    const searching = this.searchQuery.trim().length > 0;
    const totalMatches = results.reduce(
      (total, result) => total + result.matchCount,
      0,
    );
    this.filteredSlideIndices = results.map((result) => result.slideIndex);

    this.navigationEmptyEl?.toggleClass("is-visible", results.length === 0);
    this.searchCountEl?.setText(
      searching
        ? this.text.navigation.searchCount(totalMatches, results.length)
        : this.text.navigation.slideCount(this.slideMetadata.length),
    );
    if (resetScroll && this.slideListEl) {
      this.slideListEl.scrollTop = 0;
    }
    this.renderNavigationWindow();
  }

  private scheduleNavigationWindowRender(): void {
    if (this.navigationRenderFrameId !== null || !this.slideListEl) {
      return;
    }
    const renderWindow = this.slideListEl.win;
    this.navigationRenderFrameId = renderWindow.requestAnimationFrame(() => {
      this.navigationRenderFrameId = null;
      this.renderNavigationWindow();
    });
  }

  private renderNavigationWindow(): void {
    const slideListEl = this.slideListEl;
    const rowsEl = this.navigationRowsEl;
    if (!slideListEl || !rowsEl) {
      return;
    }
    const navigationWindow = this.navigationWindow.calculate(
      this.filteredSlideIndices,
      slideListEl.scrollTop,
      slideListEl.clientHeight,
    );
    const desired = new Set(navigationWindow.indices);

    for (const [index, entry] of [...this.navigationEntries]) {
      if (!desired.has(index)) {
        this.unmountThumbnail(index);
        entry.element.remove();
        this.navigationEntries.delete(index);
      }
    }

    const fragment = rowsEl.ownerDocument.createDocumentFragment();
    for (const index of navigationWindow.indices) {
      const entry =
        this.navigationEntries.get(index) ??
        this.createNavigationEntry(index);
      this.updateNavigationEntry(entry, index);
      entry.mounted = true;
      fragment.appendChild(entry.element);
    }
    rowsEl.appendChild(fragment);
    if (this.navigationTopSpacerEl) {
      this.navigationTopSpacerEl.style.height =
        `${navigationWindow.topSpacer}px`;
    }
    if (this.navigationBottomSpacerEl) {
      this.navigationBottomSpacerEl.style.height =
        `${navigationWindow.bottomSpacer}px`;
    }
    this.updateActiveNavigationEntry(false);
    this.refreshThumbnailQueue();
  }

  private createNavigationEntry(index: number): PptxNavigationEntry {
    if (!this.navigationRowsEl) {
      throw new Error("The PPTX navigation list is not available.");
    }
    const entryEl = this.navigationRowsEl.ownerDocument.createElement("button");
    entryEl.className = "pptx-reader-slide-entry";
    entryEl.type = "button";
    entryEl.addEventListener("click", () => {
      void this.goToSlide(index);
    });
    const thumbnailEl = entryEl.createDiv({
      cls: "pptx-reader-thumbnail",
    });
    thumbnailEl.createDiv({
      cls: "pptx-reader-thumbnail-placeholder",
      text: String(index + 1),
    });
    const detailsEl = entryEl.createDiv({
      cls: "pptx-reader-slide-details",
    });
    detailsEl.createDiv({
      cls: "pptx-reader-slide-page",
      text: this.text.navigation.slideLabel(index + 1),
    });
    const titleEl = detailsEl.createDiv({
      cls: "pptx-reader-slide-title",
    });
    const snippetEl = detailsEl.createDiv({
      cls: "pptx-reader-slide-snippet",
    });
    const matchEl = detailsEl.createDiv({
      cls: "pptx-reader-slide-match",
    });
    const entry: PptxNavigationEntry = {
      element: entryEl,
      titleEl,
      snippetEl,
      matchEl,
      thumbnailEl,
      resources: new Set<string>(),
      renderVersion: 0,
      rendering: false,
      rendered: false,
      mounted: true,
    };
    this.navigationEntries.set(index, entry);
    return entry;
  }

  private updateNavigationEntry(
    entry: PptxNavigationEntry,
    index: number,
  ): void {
    const metadata = this.slideMetadata[index] ?? createEmptySlideMetadata(index);
    const title =
      metadata.title || this.text.navigation.slideLabel(index + 1);
    const result = this.navigationResults.get(index);
    entry.element.setAttribute(
      "aria-label",
      `${this.text.navigation.slideLabel(index + 1)}: ${title}`,
    );
    entry.titleEl.setText(title);
    entry.snippetEl.setText(result?.snippet ?? "");
    entry.matchEl.setText(
      this.searchQuery.trim() && result
        ? [
            `${result.matchCount}`,
            result.matchedNotes ? this.text.navigation.notesMatch : "",
          ]
            .filter(Boolean)
            .join(" - ")
        : "",
    );
  }

  private updateActiveNavigationEntry(ensureVisible = true): void {
    for (const [index, entry] of this.navigationEntries) {
      const active = index === this.currentSlideIndex;
      entry.element.toggleClass("is-active", active);
      entry.element.setAttribute("aria-current", active ? "true" : "false");
    }
    if (!ensureVisible || !this.navigationVisible || !this.slideListEl) {
      return;
    }
    const position = this.filteredSlideIndices.indexOf(this.currentSlideIndex);
    if (position < 0) {
      return;
    }
    const top = position * this.navigationWindow.rowHeight;
    const bottom = top + this.navigationWindow.rowHeight;
    const visibleTop = this.slideListEl.scrollTop;
    const visibleBottom = visibleTop + this.slideListEl.clientHeight;
    if (top < visibleTop || bottom > visibleBottom) {
      this.slideListEl.scrollTop = top;
      this.renderNavigationWindow();
    }
  }

  private refreshThumbnailQueue(): void {
    const presentation = this.presentation;
    if (!presentation) {
      return;
    }
    const lifecycleToken = this.thumbnailLifecycle.currentToken;
    for (const [index, entry] of this.navigationEntries) {
      if (!entry.mounted || entry.rendered || entry.rendering) {
        continue;
      }
      this.thumbnailQueue.schedule(
        index,
        Math.abs(index - this.currentSlideIndex),
        async (isQueueCancelled) => {
          await this.renderThumbnail(
            index,
            entry,
            presentation,
            lifecycleToken,
            isQueueCancelled,
          );
        },
      );
    }
  }

  private async renderThumbnail(
    index: number,
    entry: PptxNavigationEntry,
    presentation: PptxPackage,
    lifecycleToken: number,
    isQueueCancelled: () => boolean,
  ): Promise<void> {
    if (entry.rendered || entry.rendering || !entry.mounted) {
      return;
    }
    entry.rendering = true;
    entry.renderVersion += 1;
    const renderVersion = entry.renderVersion;
    let resources = new Set<string>();
    let adopted = false;
    const isCancelled = (): boolean =>
      !this.thumbnailLifecycle.isCurrent(lifecycleToken) ||
      isQueueCancelled() ||
      this.presentation !== presentation ||
      !entry.mounted ||
      entry.renderVersion !== renderVersion ||
      !entry.thumbnailEl.isConnected;

    try {
      const context = await presentation.getSlideContext(index);
      if (isCancelled()) {
        return;
      }
      const rendered = await renderPptxSlide(
        presentation,
        context,
        entry.thumbnailEl.ownerDocument,
        { isCancelled },
      );
      resources = rendered.resources;
      if (isCancelled()) {
        return;
      }
      entry.resources = resources;
      adopted = true;
      entry.rendered = true;
      const scale = THUMBNAIL_WIDTH / rendered.width;
      rendered.element.addClass("pptx-reader-thumbnail-stage");
      rendered.element.setAttribute("aria-hidden", "true");
      rendered.element.style.transform = `scale(${scale})`;
      entry.thumbnailEl.empty();
      entry.thumbnailEl.style.height = `${rendered.height * scale}px`;
      entry.thumbnailEl.appendChild(rendered.element);
    } catch (error) {
      if (!(error instanceof PptxRenderCancelledError) && entry.mounted) {
        this.showThumbnailPlaceholder(entry, index);
      }
    } finally {
      if (entry.renderVersion === renderVersion) {
        entry.rendering = false;
      }
      if (!adopted) {
        releaseResources(resources, (resource) => {
          URL.revokeObjectURL(resource);
        });
      }
      if (
        entry.mounted &&
        !entry.rendered &&
        this.thumbnailLifecycle.isCurrent(lifecycleToken) &&
        this.presentation === presentation
      ) {
        entry.thumbnailEl.win.setTimeout(() => {
          this.refreshThumbnailQueue();
        }, 0);
      }
    }
  }

  private unmountThumbnail(index: number): void {
    const entry = this.navigationEntries.get(index);
    if (!entry) {
      return;
    }
    this.thumbnailQueue.cancel(index);
    entry.mounted = false;
    entry.renderVersion += 1;
    entry.rendering = false;
    entry.rendered = false;
    releaseResources(entry.resources, (resource) => {
      URL.revokeObjectURL(resource);
    });
    entry.resources.clear();
    this.showThumbnailPlaceholder(entry, index);
  }

  private showThumbnailPlaceholder(
    entry: PptxNavigationEntry,
    index: number,
  ): void {
    entry.thumbnailEl.empty();
    entry.thumbnailEl.style.removeProperty("height");
    entry.thumbnailEl.createDiv({
      cls: "pptx-reader-thumbnail-placeholder",
      text: String(index + 1),
    });
  }

  private updateNotes(): void {
    if (!this.notesContentEl) {
      return;
    }
    const notes = this.slideMetadata[this.currentSlideIndex]?.notes.trim();
    this.notesContentEl.setText(notes || this.text.notes.empty);
    this.notesContentEl.toggleClass("is-empty", !notes);
  }

  private get allSlideIndices(): number[] {
    const count = this.presentation?.slideCount ?? this.slideMetadata.length;
    return Array.from({ length: count }, (_, index) => index);
  }

  private scheduleMetadataUiRefresh(): void {
    if (this.metadataUiFrameId !== null) {
      return;
    }
    const renderWindow = this.contentEl.win;
    this.metadataUiFrameId = renderWindow.requestAnimationFrame(() => {
      this.metadataUiFrameId = null;
      this.applyNavigationSearch(false);
      this.updateNotes();
    });
  }

  private async loadMetadataForSlide(index: number): Promise<void> {
    const presentation = this.presentation;
    if (!presentation || this.loadedMetadataIndices.has(index)) {
      return;
    }
    try {
      const metadata = await presentation.getSlideMetadata(index);
      if (this.presentation !== presentation) {
        return;
      }
      this.slideMetadata[index] = metadata;
      this.loadedMetadataIndices.add(index);
      this.searchIndex.set(metadata);
      this.scheduleMetadataUiRefresh();
    } catch {
      // Background indexing reports package failures without blocking rendering.
    }
  }

  private async getCurrentSlideMetadata(): Promise<PptxSlideMetadata | null> {
    const presentation = this.presentation;
    if (!presentation) {
      return null;
    }
    const metadata = await presentation.getSlideMetadata(
      this.currentSlideIndex,
    );
    if (this.presentation !== presentation) {
      return null;
    }
    this.slideMetadata[metadata.index] = metadata;
    this.loadedMetadataIndices.add(metadata.index);
    this.searchIndex.set(metadata);
    return metadata;
  }

  private async getCompleteMetadata(
    presentation: PptxPackage,
  ): Promise<PptxSlideMetadata[]> {
    const pending =
      this.metadataPromise ??
      presentation.indexSlideMetadata({
        concurrency: 4,
        priorityIndex: this.currentSlideIndex,
        isCancelled: () => this.presentation !== presentation,
      });
    const metadata = await pending;
    if (this.presentation !== presentation) {
      return [];
    }
    this.slideMetadata = metadata;
    this.loadedMetadataIndices.clear();
    for (const slide of metadata) {
      this.loadedMetadataIndices.add(slide.index);
      this.searchIndex.set(slide);
    }
    return metadata;
  }

  private applyNavigationVisibility(): void {
    this.navigationEl?.toggleClass("is-hidden", !this.navigationVisible);
    if (!this.navigationButtonEl) {
      return;
    }
    const label = this.navigationVisible
      ? this.text.toolbar.hideNavigation
      : this.text.toolbar.showNavigation;
    this.navigationButtonEl.setAttribute("aria-label", label);
    this.navigationButtonEl.setAttribute("title", label);
    this.navigationButtonEl.setAttribute(
      "aria-pressed",
      String(this.navigationVisible),
    );
    this.navigationButtonEl.toggleClass("is-active", this.navigationVisible);
    if (this.fitWindow) {
      this.applyScale();
    }
  }

  private applyNotesVisibility(): void {
    this.notesEl?.toggleClass("is-hidden", !this.notesVisible);
    if (!this.notesButtonEl) {
      return;
    }
    const label = this.notesVisible
      ? this.text.toolbar.hideNotes
      : this.text.toolbar.showNotes;
    this.notesButtonEl.setAttribute("aria-label", label);
    this.notesButtonEl.setAttribute("title", label);
    this.notesButtonEl.setAttribute(
      "aria-pressed",
      String(this.notesVisible),
    );
    this.notesButtonEl.toggleClass("is-active", this.notesVisible);
    if (this.fitWindow) {
      this.applyScale();
    }
  }

  private getSelectedSlideText(): string {
    const selection = this.canvasEl?.ownerDocument.defaultView?.getSelection();
    if (!selection || selection.isCollapsed || !this.stageEl) {
      return "";
    }
    const anchorNode = selection.anchorNode;
    const focusNode = selection.focusNode;
    return anchorNode &&
      focusNode &&
      this.stageEl.contains(anchorNode) &&
      this.stageEl.contains(focusNode)
      ? selection.toString().trim()
      : "";
  }

  private setZoom(value: number, anchor?: { left: number; top: number }): void {
    const nextZoom = normalizeZoom(value, {
      min: MIN_ZOOM,
      max: MAX_ZOOM,
      step: ZOOM_STEP,
    });
    if (nextZoom === null || !this.viewportEl) {
      this.updateZoomControl(this.getEffectiveScale());
      return;
    }

    const previousZoom = this.getEffectiveScale();
    const pointer = anchor ?? {
      left: this.viewportEl.clientWidth / 2,
      top: this.viewportEl.clientHeight / 2,
    };
    const nextScroll = preserveZoomAnchor(
      {
        left: this.viewportEl.scrollLeft,
        top: this.viewportEl.scrollTop,
      },
      pointer,
      previousZoom,
      nextZoom,
    );
    this.fitWindow = false;
    this.zoom = nextZoom;
    this.applyScale();
    this.viewportEl.scrollLeft = Math.max(0, nextScroll.left);
    this.viewportEl.scrollTop = Math.max(0, nextScroll.top);
    this.saveReadingState();
  }

  private handleWheelZoom(event: WheelEvent): void {
    if (!event.ctrlKey || !this.viewportEl) {
      return;
    }
    event.preventDefault();
    const rect = this.viewportEl.getBoundingClientRect();
    const direction = event.deltaY < 0 ? 1 : -1;
    this.setZoom(this.getEffectiveScale() + direction * ZOOM_STEP, {
      left: event.clientX - rect.left,
      top: event.clientY - rect.top,
    });
  }

  private applyScale(): void {
    const scale = this.getEffectiveScale();
    if (this.stageEl && this.canvasEl) {
      this.stageEl.style.transform = `scale(${scale})`;
      this.canvasEl.style.width = `${this.renderedWidth * scale}px`;
      this.canvasEl.style.height = `${this.renderedHeight * scale}px`;
    }
    this.fitButtonEl?.toggleClass("is-active", this.fitWindow);
    this.fitButtonEl?.setAttribute(
      "aria-pressed",
      String(this.fitWindow),
    );
    this.updateZoomControl(scale);
  }

  private getEffectiveScale(): number {
    if (
      !this.fitWindow ||
      !this.viewportEl ||
      !this.renderedWidth ||
      !this.renderedHeight
    ) {
      return this.zoom;
    }
    const availableWidth = Math.max(
      1,
      this.viewportEl.clientWidth - FIT_PADDING,
    );
    const availableHeight = Math.max(
      1,
      this.viewportEl.clientHeight - FIT_PADDING,
    );
    return clamp(
      Math.min(
        availableWidth / this.renderedWidth,
        availableHeight / this.renderedHeight,
      ),
      MIN_ZOOM,
      MAX_ZOOM,
    );
  }

  private updateZoomControl(scale: number): void {
    if (this.zoomInputEl) {
      this.zoomInputEl.value = String(Math.round(scale * 100));
    }
  }

  private updateNavigationControls(): void {
    const count = this.presentation?.slideCount ?? 0;
    const page = count > 0 ? this.currentSlideIndex + 1 : 1;
    if (this.pageInputEl) {
      this.pageInputEl.value = String(page);
      this.pageInputEl.max = String(Math.max(count, 1));
      this.pageInputEl.disabled = count === 0;
    }
    this.pageCountEl?.setText(`/ ${count}`);
    this.previousButtonEl?.toggleAttribute(
      "disabled",
      count === 0 || this.currentSlideIndex <= 0,
    );
    this.nextButtonEl?.toggleAttribute(
      "disabled",
      count === 0 || this.currentSlideIndex >= count - 1,
    );
    this.updateActiveNavigationEntry();
    this.updateNotes();
  }

  private updateFullscreenControl(): void {
    if (!this.fullscreenButtonEl) {
      return;
    }
    const fullscreen = this.contentEl.doc.fullscreenElement === this.rootEl;
    const label = fullscreen
      ? this.text.toolbar.exitFullscreen
      : this.text.toolbar.enterFullscreen;
    this.fullscreenButtonEl.setAttribute("aria-label", label);
    this.fullscreenButtonEl.setAttribute("title", label);
    this.fullscreenButtonEl.toggleClass("is-active", fullscreen);
  }

  private createIconButton(
    parentEl: HTMLElement,
    icon: string,
    label: string,
    onClick: () => void,
  ): HTMLButtonElement {
    const buttonEl = parentEl.createEl("button", {
      attr: {
        type: "button",
        "aria-label": label,
        title: label,
      },
    });
    setIcon(buttonEl, icon);
    buttonEl.addEventListener("click", onClick);
    return buttonEl;
  }

  private showError(error: unknown): void {
    const info = classifyPptxError(error, this.text);
    this.clearCanvas();
    if (!this.canvasEl) {
      return;
    }
    const messageEl = this.canvasEl.createDiv({
      cls: "pptx-reader-error",
    });
    messageEl.createDiv({
      cls: "word-reader-message-title",
      text: info.title,
    });
    messageEl.createDiv({
      cls: "word-reader-message-body",
      text: info.body,
    });
    const listEl = messageEl.createEl("ul", {
      cls: "word-reader-message-list",
    });
    for (const tip of info.tips) {
      listEl.createEl("li", { text: tip });
    }
    const detailsEl = messageEl.createEl("details", {
      cls: "word-reader-diagnostics",
    });
    detailsEl.createEl("summary", { text: this.text.errors.details });
    detailsEl.createEl("pre", {
      cls: "pptx-reader-error-details",
      text: info.details,
    });
    const actionsEl = messageEl.createDiv({
      cls: "word-reader-message-actions",
    });
    this.createIconButton(
      actionsEl,
      "refresh-cw",
      this.text.toolbar.reload,
      () => {
        void this.reload();
      },
    );
    this.createIconButton(
      actionsEl,
      "external-link",
      this.text.toolbar.openExternally,
      () => {
        void this.openExternal();
      },
    );
    this.setStatus(info.title, true);
  }

  private setStatus(
    message: string,
    isError = false,
    isLoading = false,
  ): void {
    this.readerStatus.set(
      message,
      isError ? "error" : isLoading ? "loading" : "idle",
    );
  }

  private renderStatus(status: ReaderStatus): void {
    if (!this.statusEl) {
      return;
    }
    this.statusEl.setText(status.message);
    this.statusEl.toggleClass("is-error", status.kind === "error");
    this.statusEl.toggleClass("is-loading", status.kind === "loading");
  }

  private restoreReadingState(path: string): void {
    this.activeFilePath = path;
    const state = this.plugin.getReadingState(path);
    const defaultZoom = this.plugin.settings.defaultZoomPercent / 100;
    this.zoom =
      normalizeZoom(state?.zoom ?? defaultZoom, {
        min: MIN_ZOOM,
        max: MAX_ZOOM,
        step: ZOOM_STEP,
      }) ?? defaultZoom;
    this.fitWindow = state?.fitWidth ?? true;
    this.navigationVisible = state?.outlineVisible ?? true;
    this.notesVisible = state?.notesVisible ?? false;
    this.currentSlideIndex = Math.max(0, (state?.page ?? 1) - 1);
    this.pendingScrollPosition = {
      left: state?.scrollLeft ?? 0,
      top: state?.scrollTop ?? 0,
    };
  }

  private restoreScrollPosition(): void {
    if (!this.viewportEl || !this.pendingScrollPosition) {
      return;
    }
    this.viewportEl.scrollLeft = this.pendingScrollPosition.left;
    this.viewportEl.scrollTop = this.pendingScrollPosition.top;
    this.pendingScrollPosition = null;
  }

  private saveReadingState(): void {
    if (!this.activeFilePath) {
      return;
    }
    const state: ReaderViewState = {
      zoom: this.zoom,
      fitWidth: this.fitWindow,
      outlineVisible: this.navigationVisible,
      scrollLeft: this.viewportEl?.scrollLeft ?? 0,
      scrollTop: this.viewportEl?.scrollTop ?? 0,
      collapsedOutlineIds: [],
      page: this.currentSlideIndex + 1,
      notesVisible: this.notesVisible,
    };
    this.plugin.updateReadingState(this.activeFilePath, state);
  }

  private clearCanvas(): void {
    this.slideResources.releaseActive();
    this.stageEl = null;
    this.renderedWidth = 0;
    this.renderedHeight = 0;
    this.currentRenderDiagnostics = null;
    this.canvasEl?.empty();
  }

  private releasePresentation(): void {
    this.thumbnailLifecycle.cancel();
    this.thumbnailQueue.clear();
    this.cancelNavigationWork();
    this.releaseThumbnailResources();
    this.presentation = null;
    this.metadataPromise = null;
    this.slideMetadata = [];
    this.loadedMetadataIndices.clear();
    this.searchIndex.clear();
    this.filteredSlideIndices = [];
    this.navigationResults.clear();
    this.navigationEntries.clear();
    this.navigationRowsEl?.empty();
    this.navigationTopSpacerEl?.style.removeProperty("height");
    this.navigationBottomSpacerEl?.style.removeProperty("height");
    this.slideResources.releaseActive();
    this.applyNavigationSearch();
    this.updateNotes();
    this.updateNavigationControls();
  }

  private releaseThumbnailResources(): void {
    for (const index of this.navigationEntries.keys()) {
      this.unmountThumbnail(index);
    }
  }

  private cancelNavigationWork(): void {
    if (this.searchTimer !== null) {
      window.clearTimeout(this.searchTimer);
      this.searchTimer = null;
    }
    if (this.navigationRenderFrameId !== null) {
      this.contentEl.win.cancelAnimationFrame(this.navigationRenderFrameId);
      this.navigationRenderFrameId = null;
    }
    if (this.metadataUiFrameId !== null) {
      this.contentEl.win.cancelAnimationFrame(this.metadataUiFrameId);
      this.metadataUiFrameId = null;
    }
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function createEmptySlideMetadata(index: number): PptxSlideMetadata {
  return {
    index,
    title: "",
    text: "",
    notes: "",
  };
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function logPptxRenderPerformance(
  fileName: string,
  slideIndex: number,
  diagnostics: PptxRenderDiagnostics,
): void {
  if (typeof __DEV__ === "undefined" || !__DEV__) {
    return;
  }
  console.debug("[Office Reader] PPTX render performance", {
    fileName,
    slide: slideIndex + 1,
    ...diagnostics,
    durationMs: Math.round(diagnostics.durationMs),
  });
}
