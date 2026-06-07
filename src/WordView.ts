import {
  App,
  FileView,
  Modal,
  Notice,
  Platform,
  TFile,
  WorkspaceLeaf,
  setIcon,
} from "obsidian";

import type WordReaderPlugin from "./main";
import { createNoteFromDocx } from "./commands/createNoteFromDocx";
import { openExternalDocx } from "./commands/openExternal";
import { renderDocx } from "./renderer/docxRenderer";
import { extractMarkdown, extractPlainText } from "./renderer/textExtractor";
import type { WordReaderText } from "./i18n";

export const VIEW_TYPE_WORD_READER = "word-reader-view";

const MIN_ZOOM = 0.25;
const MAX_ZOOM = 4;
const ZOOM_STEP = 0.05;
const MIN_IMAGE_ZOOM = 0.1;
const MAX_IMAGE_ZOOM = 8;
const IMAGE_ZOOM_STEP = 1.1;
const SEARCH_DEBOUNCE_MS = 150;
const DOM_COMMIT_CHUNK_SIZE = 4;
const OUTLINE_CHUNK_SIZE = 50;
const SEARCH_SCAN_CHUNK_SIZE = 500;
const SEARCH_MUTATION_CHUNK_SIZE = 100;
const LONG_DOCUMENT_PAGE_THRESHOLD = 12;

declare const __DEV__: boolean;

export class WordView extends FileView {
  private readonly plugin: WordReaderPlugin;
  private rootEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private bodyEl: HTMLElement | null = null;
  private outlineEl: HTMLElement | null = null;
  private outlineListEl: HTMLElement | null = null;
  private scrollEl: HTMLElement | null = null;
  private documentEl: HTMLElement | null = null;
  private zoomInputEl: HTMLInputElement | null = null;
  private searchInputEl: HTMLInputElement | null = null;
  private outlineToggleButtonEl: HTMLButtonElement | null = null;
  private searchPreviousButtonEl: HTMLButtonElement | null = null;
  private searchNextButtonEl: HTMLButtonElement | null = null;
  private searchCountEl: HTMLElement | null = null;
  private searchMatchEls: HTMLElement[] = [];
  private currentSearchIndex = -1;
  private searchHighlightTimer: number | null = null;
  private searchTaskToken = 0;
  private outlineTaskToken = 0;
  private buffer: ArrayBuffer | null = null;
  private renderToken = 0;
  private renderedDocumentKey: string | null = null;
  private pendingRenderKey: string | null = null;
  private documentBlobUrls = new Set<string>();
  private retainedBlobUrls = new Map<string, number>();
  private pendingBlobUrlRevocations = new Set<string>();
  private zoom = 1;
  private fitWidth = false;
  private outlineVisible = false;
  private searchQuery = "";

  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf);
    this.plugin = plugin;
    this.zoom = plugin.settings.defaultZoomPercent / 100;
    this.fitWidth = plugin.settings.defaultFitWidth;
    this.outlineVisible = plugin.settings.showOutlineByDefault;
  }

  getViewType(): string {
    return VIEW_TYPE_WORD_READER;
  }

  getDisplayText(): string {
    return this.file?.name ?? "Word Reader";
  }

  getIcon(): string {
    return "file-text";
  }

  private get text(): WordReaderText {
    return this.plugin.text;
  }

  async onOpen(): Promise<void> {
    this.buildLayout();
  }

  async onLoadFile(file: TFile): Promise<void> {
    await this.loadDocx(file);
  }

  async onUnloadFile(): Promise<void> {
    this.renderToken += 1;
    this.cancelSearchWork();
    this.releaseDocumentState();
    this.clearRenderedDocument();
    this.updateSearchState();
    this.setStatus("");
  }

  async reload(): Promise<void> {
    if (!this.file) {
      return;
    }

    await this.loadDocx(this.file, { force: true });
  }

  async copyText(): Promise<void> {
    const text = this.text;
    if (!this.file) {
      return;
    }

    if (this.isLegacyDocFile(this.file)) {
      new Notice(text.notices.legacyDoc);
      return;
    }

    const selectedText = this.getSelectedRenderedText();
    if (selectedText.length > 0) {
      await navigator.clipboard.writeText(selectedText);
      new Notice(text.notices.copiedSelectedText);
      return;
    }

    if (!this.buffer) {
      this.buffer = await this.app.vault.readBinary(this.file);
    }

    this.setStatus(text.status.extractingPlainText);
    const plainText = await extractPlainText(this.buffer);
    await navigator.clipboard.writeText(plainText);
    this.setStatus(this.text.status.copiedPlainTextFrom(this.file.name));
    new Notice(this.text.notices.copiedPlainText);
  }

  async copyMarkdown(): Promise<void> {
    const text = this.text;
    if (!this.file) {
      return;
    }

    if (this.isLegacyDocFile(this.file)) {
      new Notice(text.notices.legacyDoc);
      return;
    }

    const selectedMarkdown = this.getSelectedRenderedMarkdown();
    if (selectedMarkdown.length > 0) {
      await navigator.clipboard.writeText(selectedMarkdown);
      new Notice(text.notices.copiedSelectedMarkdown);
      return;
    }

    if (!this.buffer) {
      this.buffer = await this.app.vault.readBinary(this.file);
    }

    this.setStatus(text.status.extractingMarkdown);
    const markdown = await extractMarkdown(this.buffer);
    await navigator.clipboard.writeText(markdown);
    this.setStatus(text.status.copiedMarkdownFrom(this.file.name));
    new Notice(text.notices.copiedMarkdown);
  }

  refreshInterfaceLanguage(): void {
    const file = this.file;
    this.renderToken += 1;
    this.cancelSearchWork();
    this.buildLayout();
    if (file) {
      void this.loadDocx(file, { force: true });
    }
  }

  private buildLayout(): void {
    this.contentEl.empty();
    const text = this.text;

    this.rootEl = this.contentEl.createDiv({ cls: "word-reader-root" });
    const toolbarEl = this.rootEl.createDiv({ cls: "word-reader-toolbar" });

    this.createIconButton(toolbarEl, "refresh-cw", text.toolbar.reload, () => {
      void this.reload();
    });

    this.zoomInputEl = toolbarEl.createEl("input", {
      cls: "word-reader-zoom",
      attr: {
        type: "number",
        min: String(Math.round(MIN_ZOOM * 100)),
        max: String(Math.round(MAX_ZOOM * 100)),
        step: String(Math.round(ZOOM_STEP * 100)),
        "aria-label": text.toolbar.zoomPercentage,
        title: text.toolbar.zoomPercentage,
      },
    });
    this.updateZoomControl();
    this.zoomInputEl.addEventListener("change", () => {
      this.setZoom(Number(this.zoomInputEl?.value) / 100);
    });

    this.createIconButton(toolbarEl, "panel-top-open", text.toolbar.fitWidth, () => {
      this.fitWidth = !this.fitWidth;
      if (this.buffer && this.file) {
        void this.renderCurrentBuffer(
          ++this.renderToken,
          this.getRenderKey(this.file),
        );
      } else {
        this.applyDocumentOptions();
      }
    });

    this.outlineToggleButtonEl = this.createIconButton(
      toolbarEl,
      "list-tree",
      this.outlineVisible ? text.toolbar.hideOutline : text.toolbar.showOutline,
      () => {
        this.outlineVisible = !this.outlineVisible;
        this.applyOutlineVisibility();
        if (this.outlineVisible && this.renderedDocumentKey) {
          void this.updateOutline(this.renderToken);
        }
      },
    );

    this.searchInputEl = toolbarEl.createEl("input", {
      cls: "word-reader-search",
      attr: {
        type: "search",
        placeholder: text.toolbar.searchPlaceholder,
        "aria-label": text.toolbar.searchDocument,
      },
    });
    this.searchInputEl.value = this.searchQuery;
    this.searchInputEl.addEventListener("input", () => {
      this.searchQuery = this.searchInputEl?.value ?? "";
      this.scheduleSearchHighlights();
    });
    this.searchInputEl.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") {
        return;
      }

      event.preventDefault();
      this.navigateSearch(event.shiftKey ? -1 : 1);
    });

    this.searchPreviousButtonEl = this.createIconButton(
      toolbarEl,
      "chevron-up",
      text.toolbar.previousSearchResult,
      () => {
        this.navigateSearch(-1);
      },
    );

    this.searchNextButtonEl = this.createIconButton(
      toolbarEl,
      "chevron-down",
      text.toolbar.nextSearchResult,
      () => {
        this.navigateSearch(1);
      },
    );

    this.searchCountEl = toolbarEl.createSpan({
      cls: "word-reader-search-count",
      text: "0",
    });
    this.updateSearchState();

    this.createIconButton(toolbarEl, "copy", text.toolbar.copyText, () => {
      void this.copyText();
    });

    this.createIconButton(toolbarEl, "clipboard-copy", text.toolbar.copyMarkdown, () => {
      void this.copyMarkdown();
    });

    this.createIconButton(toolbarEl, "file-plus", text.toolbar.createSummaryNote, () => {
      if (this.file) {
        void createNoteFromDocx(this.app, this.file, this.text);
      }
    });

    this.createIconButton(toolbarEl, "external-link", text.toolbar.openExternally, () => {
      if (this.file) {
        void openExternalDocx(this.app, this.file, this.text);
      }
    });

    this.statusEl = this.rootEl.createDiv({ cls: "word-reader-status" });
    this.bodyEl = this.rootEl.createDiv({ cls: "word-reader-body" });
    this.outlineEl = this.bodyEl.createDiv({ cls: "word-reader-outline" });
    this.outlineEl.createDiv({
      cls: "word-reader-outline-title",
      text: text.outline.title,
    });
    this.outlineListEl = this.outlineEl.createDiv({
      cls: "word-reader-outline-list",
    });
    this.scrollEl = this.bodyEl.createDiv({ cls: "word-reader-scroll" });
    this.scrollEl.addEventListener("wheel", (event) => {
      this.handleCtrlWheelZoom(event);
    });
    this.documentEl = this.scrollEl.createDiv({ cls: "word-reader-document" });
    this.documentEl.addEventListener("click", (event) => {
      this.handleDocumentClick(event);
    });
    this.applyDocumentOptions();
    this.applyOutlineVisibility();
  }

  private async loadDocx(
    file: TFile,
    options: { force?: boolean } = {},
  ): Promise<void> {
    this.ensureLayout();

    const token = ++this.renderToken;
    const loadStartedAt = performance.now();
    this.cancelSearchWork();
    this.clearSearchHighlightsState();
    this.updateSearchState();

    if (this.isLegacyDocFile(file)) {
      this.releaseDocumentState();
      this.clearRenderedDocument();
      this.showLegacyDocMessage(file);
      return;
    }

    const renderKey = this.getRenderKey(file);
    if (
      !options.force &&
      this.buffer &&
      this.renderedDocumentKey === renderKey &&
      this.documentEl?.hasChildNodes()
    ) {
      this.applyDocumentOptions();
      void this.updateOutline(token);
      this.scheduleSearchHighlights();
      this.setStatus(this.text.status.preview(file.name));
      return;
    }

    this.clearRenderedDocument();
    this.pendingRenderKey = renderKey;
    this.renderedDocumentKey = null;
    this.setStatus(this.text.status.reading(file.name), false, true);

    const largeFileWarningBytes =
      this.plugin.settings.largeFileWarningMb * 1024 * 1024;
    if (file.stat.size > largeFileWarningBytes) {
      const fileSizeMb = formatFileSizeMb(file.stat.size);
      new Notice(this.text.notices.largeDocument(fileSizeMb));
      this.setStatus(
        this.text.status.largeDocument(fileSizeMb),
        false,
        true,
      );
    }

    try {
      await yieldToBrowser();
      if (token !== this.renderToken) {
        return;
      }

      const readStartedAt = performance.now();
      const buffer = await this.app.vault.readBinary(file);
      if (token !== this.renderToken) {
        return;
      }

      this.buffer = buffer;
      await this.renderCurrentBuffer(token, renderKey, {
        loadStartedAt,
        readDurationMs: performance.now() - readStartedAt,
      });
    } catch (error) {
      if (token !== this.renderToken) {
        return;
      }

      this.showError(error, file);
      this.releaseDocumentState();
    }
  }

  private async renderCurrentBuffer(
    token: number,
    renderKey: string,
    timings?: {
      loadStartedAt: number;
      readDurationMs: number;
    },
  ): Promise<void> {
    if (!this.buffer || !this.documentEl || !this.file) {
      return;
    }

    if (
      this.renderedDocumentKey === renderKey &&
      this.documentEl.hasChildNodes()
    ) {
      this.applyDocumentOptions();
      void this.updateOutline(token);
      this.scheduleSearchHighlights();
      this.setStatus(this.text.status.preview(this.file.name));
      return;
    }

    this.cancelSearchWork();
    this.clearSearchHighlightsState();
    this.updateSearchState();
    this.setStatus(this.text.status.rendering(this.file.name), false, true);
    await yieldToBrowser();
    if (token !== this.renderToken) {
      return;
    }

    // SECURITY: Only structural DOM elements are created for safe document rendering.
    // No <script> elements are created or injected. All content comes from trusted local .docx files.
    const renderTargetEl = activeDocument.createElement("div");
    renderTargetEl.className = "word-reader-render-buffer";
    let renderedBlobUrls = new Set<string>();
    let adoptedBlobUrls = false;

    try {
      const rendererStartedAt = performance.now();
      await renderDocx(this.buffer, renderTargetEl, {
        fitWidth: this.fitWidth,
      });
      const rendererDurationMs = performance.now() - rendererStartedAt;
      renderedBlobUrls = collectBlobUrls(renderTargetEl);

      if (token !== this.renderToken) {
        return;
      }

      const imageCount = prepareRenderedImages(renderTargetEl);
      const pageCount = renderTargetEl.querySelectorAll(
        ".docx-wrapper > section.docx",
      ).length;
      const commitStartedAt = performance.now();

      this.clearRenderedDocument();
      this.documentBlobUrls = renderedBlobUrls;
      adoptedBlobUrls = true;
      await this.commitRenderedDocument(renderTargetEl, token, this.file.name);

      if (token !== this.renderToken) {
        return;
      }

      const commitDurationMs = performance.now() - commitStartedAt;
      this.renderedDocumentKey = renderKey;
      this.pendingRenderKey = null;
      this.applyDocumentOptions();
      this.documentEl.toggleClass(
        "is-long-document",
        pageCount >= LONG_DOCUMENT_PAGE_THRESHOLD,
      );

      this.setStatus(
        this.text.status.buildingNavigation(this.file.name),
        false,
        true,
      );
      const outlineStartedAt = performance.now();
      await this.updateOutline(token);
      const outlineDurationMs = performance.now() - outlineStartedAt;

      if (token !== this.renderToken) {
        return;
      }

      this.scheduleSearchHighlights();
      this.setStatus(this.text.status.preview(this.file.name));
      logRenderPerformance({
        fileName: this.file.name,
        fileSizeBytes: this.file.stat.size,
        readDurationMs: timings?.readDurationMs ?? 0,
        rendererDurationMs,
        commitDurationMs,
        outlineDurationMs,
        totalDurationMs:
          performance.now() - (timings?.loadStartedAt ?? rendererStartedAt),
        pageCount,
        imageCount,
        blobUrlCount: renderedBlobUrls.size,
      });
    } finally {
      if (!adoptedBlobUrls) {
        for (const url of collectBlobUrls(renderTargetEl)) {
          renderedBlobUrls.add(url);
        }
        revokeBlobUrls(renderedBlobUrls);
      }
    }
  }

  private async commitRenderedDocument(
    renderTargetEl: HTMLElement,
    token: number,
    fileName: string,
  ): Promise<void> {
    if (!this.documentEl) {
      return;
    }

    const targetEl = this.documentEl;
    const topLevelNodes = Array.from(renderTargetEl.childNodes);
    const totalUnits = topLevelNodes.reduce((total, node) => {
      if (node.instanceOf(HTMLElement) && node.matches(".docx-wrapper")) {
        return total + Math.max(node.childNodes.length, 1);
      }
      return total + 1;
    }, 0);
    let completedUnits = 0;

    for (const node of topLevelNodes) {
      if (token !== this.renderToken || this.documentEl !== targetEl) {
        return;
      }

      if (node.instanceOf(HTMLElement) && node.matches(".docx-wrapper")) {
        const wrapperEl = node.cloneNode(false) as HTMLElement;
        targetEl.appendChild(wrapperEl);
        const wrapperChildren = Array.from(node.childNodes);

        if (wrapperChildren.length === 0) {
          completedUnits += 1;
          continue;
        }

        for (
          let index = 0;
          index < wrapperChildren.length;
          index += DOM_COMMIT_CHUNK_SIZE
        ) {
          if (token !== this.renderToken || this.documentEl !== targetEl) {
            return;
          }

          const fragment = activeDocument.createDocumentFragment();
          for (const child of wrapperChildren.slice(
            index,
            index + DOM_COMMIT_CHUNK_SIZE,
          )) {
            fragment.appendChild(child);
          }
          wrapperEl.appendChild(fragment);
          completedUnits += Math.min(
            DOM_COMMIT_CHUNK_SIZE,
            wrapperChildren.length - index,
          );
          this.setStatus(
            this.text.status.preparingPreview(
              fileName,
              Math.round((completedUnits / totalUnits) * 100),
            ),
            false,
            true,
          );

          if (index + DOM_COMMIT_CHUNK_SIZE < wrapperChildren.length) {
            await yieldToBrowser();
          }
        }
        continue;
      }

      targetEl.appendChild(node);
      completedUnits += 1;
    }
  }

  private ensureLayout(): void {
    if (!this.rootEl || !this.documentEl || !this.statusEl) {
      this.buildLayout();
    }
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

  private applyDocumentOptions(): void {
    if (!this.documentEl) {
      return;
    }

    this.documentEl.style.setProperty("--word-reader-zoom", String(this.zoom));
    this.documentEl.toggleClass("is-fit-width", this.fitWidth);
    this.documentEl.toggleClass(
      "can-preview-images",
      this.plugin.settings.enableImagePreview,
    );
    this.updateZoomControl();
  }

  private applyOutlineVisibility(): void {
    this.bodyEl?.toggleClass("is-outline-visible", this.outlineVisible);
    const label = this.outlineVisible
      ? this.text.toolbar.hideOutline
      : this.text.toolbar.showOutline;
    if (this.outlineToggleButtonEl) {
      this.outlineToggleButtonEl.setAttribute("aria-label", label);
      this.outlineToggleButtonEl.setAttribute("title", label);
    }
  }

  private async updateOutline(token: number): Promise<void> {
    if (!this.documentEl || !this.outlineListEl) {
      return;
    }

    const taskToken = ++this.outlineTaskToken;
    const documentEl = this.documentEl;
    const outlineListEl = this.outlineListEl;
    outlineListEl.empty();
    const headings = getDocumentHeadings(documentEl);

    if (headings.length === 0) {
      outlineListEl.createDiv({
        cls: "word-reader-outline-empty",
        text: this.text.outline.empty,
      });
      return;
    }

    for (let index = 0; index < headings.length; index += OUTLINE_CHUNK_SIZE) {
      if (
        taskToken !== this.outlineTaskToken ||
        token !== this.renderToken ||
        this.documentEl !== documentEl ||
        this.outlineListEl !== outlineListEl
      ) {
        return;
      }

      const fragment = activeDocument.createDocumentFragment();
      for (const heading of headings.slice(index, index + OUTLINE_CHUNK_SIZE)) {
        const buttonEl = activeDocument.createElement("button");
        buttonEl.className = `word-reader-outline-item level-${heading.level}`;
        buttonEl.type = "button";
        buttonEl.title = heading.text;
        buttonEl.textContent = heading.text;
        buttonEl.addEventListener("click", () => {
          heading.element.scrollIntoView({
            block: "start",
            inline: "nearest",
          });
        });
        fragment.appendChild(buttonEl);
      }
      outlineListEl.appendChild(fragment);

      if (index + OUTLINE_CHUNK_SIZE < headings.length) {
        await yieldToBrowser();
      }
    }
  }

  private handleCtrlWheelZoom(event: WheelEvent): void {
    if (!event.ctrlKey) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const scrollEl = this.scrollEl;
    const scrollRect = scrollEl?.getBoundingClientRect();
    const offsetX = scrollRect ? event.clientX - scrollRect.left : 0;
    const offsetY = scrollRect ? event.clientY - scrollRect.top : 0;
    const contentX = scrollEl ? (scrollEl.scrollLeft + offsetX) / this.zoom : 0;
    const contentY = scrollEl ? (scrollEl.scrollTop + offsetY) / this.zoom : 0;
    const direction = event.deltaY < 0 ? 1 : -1;

    this.setZoom(this.zoom + direction * ZOOM_STEP);

    if (scrollEl) {
      scrollEl.scrollLeft = contentX * this.zoom - offsetX;
      scrollEl.scrollTop = contentY * this.zoom - offsetY;
    }
  }

  private setZoom(value: number): void {
    if (!Number.isFinite(value)) {
      this.updateZoomControl();
      return;
    }

    this.zoom = clamp(roundToStep(value, ZOOM_STEP), MIN_ZOOM, MAX_ZOOM);
    this.applyDocumentOptions();
  }

  private updateZoomControl(): void {
    if (!this.zoomInputEl) {
      return;
    }

    this.zoomInputEl.value = String(Math.round(this.zoom * 100));
  }

  private handleDocumentClick(event: MouseEvent): void {
    if (!this.plugin.settings.enableImagePreview) {
      return;
    }

    const target = event.targetNode;
    if (!target?.instanceOf(Element)) {
      return;
    }

    const imageEl = target.closest("img");
    if (!imageEl?.instanceOf(HTMLImageElement)) {
      return;
    }

    const src = imageEl.currentSrc || imageEl.src;
    if (!src) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const releaseBlobUrl = src.startsWith("blob:")
      ? this.retainBlobUrl(src)
      : undefined;
    new ImagePreviewModal(
      this.app,
      src,
      imageEl.alt || this.file?.name || "Word document image",
      this.file?.basename || "word-image",
      this.text,
      releaseBlobUrl,
    ).open();
  }

  private setStatus(
    message: string,
    isError = false,
    isLoading = false,
  ): void {
    if (!this.statusEl) {
      return;
    }

    this.statusEl.setText(message);
    this.statusEl.toggleClass("is-error", isError);
    this.statusEl.toggleClass("is-loading", isLoading);
  }

  private showLegacyDocMessage(file: TFile): void {
    if (!this.documentEl) {
      return;
    }

    this.clearRenderedDocument();
    this.documentEl.createDiv({
      cls: "word-reader-message-title",
      text: this.text.legacyDoc.title,
    });
    this.documentEl.createDiv({
      cls: "word-reader-message-body",
      text: this.text.legacyDoc.body,
    });
    this.documentEl.createEl("ul", {
      cls: "word-reader-message-list",
    }).append(
      createMessageListItem(this.text.legacyDoc.openExternally),
      createMessageListItem(this.text.legacyDoc.convertToDocx),
    );
    const actionsEl = this.documentEl.createDiv({
      cls: "word-reader-message-actions",
    });
    this.createIconButton(actionsEl, "external-link", this.text.toolbar.openExternally, () => {
      void openExternalDocx(this.app, file, this.text);
    });
    this.setStatus(this.text.status.legacyDoc(file.name));
  }

  private showError(error: unknown, file: TFile): void {
    const message = error instanceof Error ? error.message : String(error);
    const errorInfo = classifyWordError(message, this.text, this.buffer);
    const diagnostics = createWordDiagnostics(errorInfo.kind, message, file);
    this.clearRenderedDocument();
    this.documentEl?.createDiv({
      cls: "word-reader-message-title",
      text: errorInfo.title,
    });
    this.documentEl?.createDiv({
      cls: "word-reader-message-body",
      text: errorInfo.body,
    });

    if (errorInfo.tips.length > 0) {
      const listEl = this.documentEl?.createEl("ul", {
        cls: "word-reader-message-list",
      });
      if (listEl) {
        for (const tip of errorInfo.tips) {
          listEl.appendChild(createMessageListItem(tip));
        }
      }
    }

    const detailsEl = this.documentEl?.createEl("details", {
      cls: "word-reader-diagnostics",
    });
    if (detailsEl) {
      detailsEl.createEl("summary", {
        text: this.text.errors.diagnosticDetails,
      });
      const listEl = detailsEl.createEl("dl", {
        cls: "word-reader-diagnostics-list",
      });
      appendDiagnosticEntry(
        listEl,
        this.text.errors.diagnosticCategory,
        diagnostics.category,
      );
      appendDiagnosticEntry(
        listEl,
        this.text.errors.diagnosticFileName,
        diagnostics.fileName,
      );
      appendDiagnosticEntry(
        listEl,
        this.text.errors.diagnosticFileSize,
        `${diagnostics.fileSizeBytes} bytes`,
      );
      appendDiagnosticEntry(
        listEl,
        this.text.errors.diagnosticModified,
        diagnostics.modifiedAt,
      );
      appendDiagnosticEntry(
        listEl,
        this.text.errors.diagnosticErrorSummary,
        diagnostics.errorSummary,
      );
    }

    const actionsEl = this.documentEl?.createDiv({
      cls: "word-reader-message-actions",
    });
    if (actionsEl) {
      this.createIconButton(
        actionsEl,
        "clipboard-copy",
        this.text.errors.copyDiagnostics,
        () => {
          void this.copyDiagnostics(formatWordDiagnostics(diagnostics));
        },
      );
      this.createIconButton(actionsEl, "external-link", this.text.toolbar.openExternally, () => {
        void openExternalDocx(this.app, file, this.text);
      });
    }

    this.setStatus(errorInfo.status, true);
  }

  private async copyDiagnostics(diagnostics: string): Promise<void> {
    try {
      await navigator.clipboard.writeText(diagnostics);
      new Notice(this.text.notices.copiedDiagnostics);
    } catch (error) {
      new Notice(
        this.text.notices.couldNotCopyDiagnostics(getErrorMessage(error)),
      );
    }
  }

  private getSelectedRenderedText(): string {
    const selection = window.getSelection();
    if (!selection || selection.isCollapsed || !this.documentEl) {
      return "";
    }

    const anchorNode = selection.anchorNode;
    const focusNode = selection.focusNode;
    if (
      anchorNode &&
      focusNode &&
      this.documentEl.contains(anchorNode) &&
      this.documentEl.contains(focusNode)
    ) {
      return selection.toString().trim();
    }

    return "";
  }

  private getSelectedRenderedMarkdown(): string {
    const selection = window.getSelection();
    if (!selection || selection.isCollapsed || !this.documentEl) {
      return "";
    }

    const anchorNode = selection.anchorNode;
    const focusNode = selection.focusNode;
    if (
      !anchorNode ||
      !focusNode ||
      !this.documentEl.contains(anchorNode) ||
      !this.documentEl.contains(focusNode)
    ) {
      return "";
    }

    const fragment = activeDocument.createDocumentFragment();
    for (let index = 0; index < selection.rangeCount; index += 1) {
      fragment.appendChild(selection.getRangeAt(index).cloneContents());
    }

    return htmlFragmentToMarkdown(fragment).trim();
  }

  private scheduleSearchHighlights(): void {
    this.clearSearchTimer();
    const taskToken = ++this.searchTaskToken;
    this.searchHighlightTimer = window.setTimeout(() => {
      this.searchHighlightTimer = null;
      void this.updateSearchHighlights(taskToken);
    }, SEARCH_DEBOUNCE_MS);
  }

  private cancelSearchWork(): void {
    this.clearSearchTimer();
    this.searchTaskToken += 1;
  }

  private clearSearchTimer(): void {
    if (this.searchHighlightTimer === null) {
      return;
    }

    window.clearTimeout(this.searchHighlightTimer);
    this.searchHighlightTimer = null;
  }

  private async updateSearchHighlights(taskToken: number): Promise<void> {
    if (!this.documentEl) {
      return;
    }

    const documentEl = this.documentEl;
    const shouldContinue = (): boolean =>
      taskToken === this.searchTaskToken && this.documentEl === documentEl;

    await clearSearchHighlights(documentEl, shouldContinue);
    if (!shouldContinue()) {
      return;
    }

    this.clearSearchHighlightsState();

    const query = this.searchQuery.trim();
    if (!query) {
      this.updateSearchState();
      return;
    }

    const matches = await highlightText(documentEl, query, shouldContinue);
    if (!shouldContinue()) {
      return;
    }

    this.searchMatchEls = matches;
    if (this.searchMatchEls.length > 0) {
      this.currentSearchIndex = 0;
    }

    this.updateSearchState({ scrollToCurrent: this.currentSearchIndex >= 0 });
  }

  private navigateSearch(direction: number): void {
    const matchCount = this.searchMatchEls.length;
    if (matchCount === 0) {
      this.scheduleSearchHighlights();
      return;
    }

    this.currentSearchIndex =
      this.currentSearchIndex < 0
        ? 0
        : (this.currentSearchIndex + direction + matchCount) % matchCount;
    this.updateSearchState({ scrollToCurrent: true });
  }

  private updateSearchState(options?: { scrollToCurrent?: boolean }): void {
    const matchCount = this.searchMatchEls.length;

    for (const [index, matchEl] of this.searchMatchEls.entries()) {
      matchEl.toggleClass(
        "is-current",
        index === this.currentSearchIndex && matchCount > 0,
      );
    }

    if (matchCount === 0) {
      this.searchCountEl?.setText("0");
    } else {
      this.searchCountEl?.setText(
        `${this.currentSearchIndex + 1} / ${matchCount}`,
      );
    }

    const hasMatches = matchCount > 0;
    this.searchPreviousButtonEl?.toggleAttribute("disabled", !hasMatches);
    this.searchNextButtonEl?.toggleAttribute("disabled", !hasMatches);

    if (options?.scrollToCurrent && this.currentSearchIndex >= 0) {
      this.searchMatchEls[this.currentSearchIndex]?.scrollIntoView({
        block: "center",
        inline: "nearest",
      });
    }
  }

  private isLegacyDocFile(file: TFile): boolean {
    return file.extension.toLowerCase() === "doc";
  }

  private clearSearchHighlightsState(): void {
    this.searchMatchEls = [];
    this.currentSearchIndex = -1;
  }

  private releaseDocumentState(): void {
    this.cancelSearchWork();
    this.buffer = null;
    this.renderedDocumentKey = null;
    this.pendingRenderKey = null;
    this.clearSearchHighlightsState();
    this.revokeDocumentBlobUrls();
  }

  private clearRenderedDocument(): void {
    this.cancelSearchWork();
    this.outlineTaskToken += 1;
    this.revokeDocumentBlobUrls();
    this.documentEl?.empty();
    this.documentEl?.removeClass("is-long-document");
    this.outlineListEl?.empty();
  }

  private retainBlobUrl(url: string): () => void {
    this.retainedBlobUrls.set(url, (this.retainedBlobUrls.get(url) ?? 0) + 1);
    let released = false;

    return () => {
      if (released) {
        return;
      }
      released = true;

      const remaining = (this.retainedBlobUrls.get(url) ?? 1) - 1;
      if (remaining > 0) {
        this.retainedBlobUrls.set(url, remaining);
        return;
      }

      this.retainedBlobUrls.delete(url);
      if (this.pendingBlobUrlRevocations.delete(url)) {
        URL.revokeObjectURL(url);
      }
    };
  }

  private revokeDocumentBlobUrls(): void {
    for (const url of this.documentBlobUrls) {
      if ((this.retainedBlobUrls.get(url) ?? 0) > 0) {
        this.pendingBlobUrlRevocations.add(url);
      } else {
        URL.revokeObjectURL(url);
      }
    }
    this.documentBlobUrls.clear();
  }

  private getRenderKey(file: TFile): string {
    return [
      file.path,
      file.stat.mtime,
      file.stat.size,
      this.fitWidth ? "fit" : "page",
    ].join(":");
  }
}

async function clearSearchHighlights(
  rootEl: HTMLElement,
  shouldContinue: () => boolean,
): Promise<void> {
  const highlights = Array.from(
    rootEl.querySelectorAll("span.word-reader-highlight"),
  );

  for (
    let index = 0;
    index < highlights.length;
    index += SEARCH_MUTATION_CHUNK_SIZE
  ) {
    if (!shouldContinue()) {
      return;
    }

    const parents = new Set<Node>();
    for (const highlightEl of highlights.slice(
      index,
      index + SEARCH_MUTATION_CHUNK_SIZE,
    )) {
      const parentEl = highlightEl.parentNode;
      if (!parentEl) {
        continue;
      }

      while (highlightEl.firstChild) {
        parentEl.insertBefore(highlightEl.firstChild, highlightEl);
      }
      parentEl.removeChild(highlightEl);
      parents.add(parentEl);
    }
    for (const parentEl of parents) {
      parentEl.normalize();
    }

    if (index + SEARCH_MUTATION_CHUNK_SIZE < highlights.length) {
      await yieldToBrowser();
    }
  }
}

type WordErrorKind =
  | "encrypted"
  | "format-mismatch"
  | "zip-container"
  | "xml-structure"
  | "unsupported-structure"
  | "unknown";

interface WordErrorInfo {
  kind: WordErrorKind;
  title: string;
  body: string;
  status: string;
  tips: string[];
}

interface WordDiagnostics {
  category: WordErrorKind;
  fileName: string;
  fileSizeBytes: number;
  modifiedAt: string;
  errorSummary: string;
}

function classifyWordError(
  message: string,
  text: WordReaderText,
  buffer: ArrayBuffer | null,
): WordErrorInfo {
  const normalizedMessage = message.toLowerCase();

  if (
    normalizedMessage.includes("password") ||
    normalizedMessage.includes("encrypt") ||
    normalizedMessage.includes("protected") ||
    (buffer && isEncryptedOfficeBuffer(buffer))
  ) {
    return {
      kind: "encrypted",
      title: text.errors.encryptedTitle,
      body: text.errors.encryptedBody,
      status: text.errors.encryptedStatus,
      tips: text.errors.encryptedTips,
    };
  }

  if (buffer && !hasZipSignature(buffer)) {
    return {
      kind: "format-mismatch",
      title: text.errors.formatMismatchTitle,
      body: text.errors.formatMismatchBody,
      status: text.errors.formatMismatchStatus,
      tips: text.errors.formatMismatchTips,
    };
  }

  if (
    normalizedMessage.includes("corrupt") ||
    normalizedMessage.includes("zip") ||
    normalizedMessage.includes("central directory") ||
    normalizedMessage.includes("archive") ||
    normalizedMessage.includes("crc") ||
    normalizedMessage.includes("unexpected end")
  ) {
    return {
      kind: "zip-container",
      title: text.errors.zipContainerTitle,
      body: text.errors.zipContainerBody,
      status: text.errors.zipContainerStatus,
      tips: text.errors.zipContainerTips,
    };
  }

  if (
    normalizedMessage.includes("xml") ||
    normalizedMessage.includes("malformed") ||
    normalizedMessage.includes("parse error") ||
    normalizedMessage.includes("invalid character") ||
    normalizedMessage.includes("unclosed") ||
    normalizedMessage.includes("namespace")
  ) {
    return {
      kind: "xml-structure",
      title: text.errors.xmlStructureTitle,
      body: text.errors.xmlStructureBody,
      status: text.errors.xmlStructureStatus,
      tips: text.errors.xmlStructureTips,
    };
  }

  if (
    normalizedMessage.includes("unsupported") ||
    normalizedMessage.includes("not supported") ||
    normalizedMessage.includes("not implemented") ||
    normalizedMessage.includes("smartart") ||
    normalizedMessage.includes("custom xml") ||
    normalizedMessage.includes("drawing")
  ) {
    return {
      kind: "unsupported-structure",
      title: text.errors.unsupportedStructureTitle,
      body: text.errors.unsupportedStructureBody,
      status: text.errors.unsupportedStructureStatus,
      tips: text.errors.unsupportedStructureTips,
    };
  }

  return {
    kind: "unknown",
    title: text.errors.genericTitle,
    body: text.errors.genericBody,
    status: text.errors.genericStatus,
    tips: text.errors.genericTips,
  };
}

function hasZipSignature(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength < 4) {
    return false;
  }

  const bytes = new Uint8Array(buffer, 0, 4);
  return (
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    ((bytes[2] === 0x03 && bytes[3] === 0x04) ||
      (bytes[2] === 0x05 && bytes[3] === 0x06) ||
      (bytes[2] === 0x07 && bytes[3] === 0x08))
  );
}

function isEncryptedOfficeBuffer(buffer: ArrayBuffer): boolean {
  const bytes = new Uint8Array(buffer);
  const oleSignature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (
    bytes.length < oleSignature.length ||
    !oleSignature.every((value, index) => bytes[index] === value)
  ) {
    return false;
  }

  return (
    containsUtf16LeText(bytes, "EncryptedPackage") ||
    containsUtf16LeText(bytes, "EncryptionInfo")
  );
}

function containsUtf16LeText(bytes: Uint8Array, text: string): boolean {
  const pattern = Array.from(text).flatMap((character) => [
    character.charCodeAt(0),
    0,
  ]);

  for (let start = 0; start <= bytes.length - pattern.length; start += 1) {
    let matches = true;
    for (let offset = 0; offset < pattern.length; offset += 1) {
      if (bytes[start + offset] !== pattern[offset]) {
        matches = false;
        break;
      }
    }
    if (matches) {
      return true;
    }
  }

  return false;
}

function createWordDiagnostics(
  category: WordErrorKind,
  message: string,
  file: TFile,
): WordDiagnostics {
  return {
    category,
    fileName: file.name,
    fileSizeBytes: file.stat.size,
    modifiedAt: new Date(file.stat.mtime).toISOString(),
    errorSummary: createSafeDiagnosticSummary(category, message),
  };
}

function createSafeDiagnosticSummary(
  category: WordErrorKind,
  message: string,
): string {
  const summaries: Record<WordErrorKind, string> = {
    encrypted: "Encrypted or protected Office container detected",
    "format-mismatch": "File content does not match the .docx container format",
    "zip-container": "The ZIP-based Office package could not be read",
    "xml-structure": "An internal document structure could not be parsed",
    "unsupported-structure": "The renderer reported an unsupported structure",
    unknown: "The renderer returned an unclassified error",
  };
  return `${summaries[category]} (fingerprint: ${fingerprintMessage(message)})`;
}

function fingerprintMessage(message: string): string {
  let hash = 0x811c9dc5;
  for (const character of message) {
    hash ^= character.charCodeAt(0);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function formatWordDiagnostics(diagnostics: WordDiagnostics): string {
  return JSON.stringify(
    {
      product: "Obsidian Word Reader",
      category: diagnostics.category,
      fileName: diagnostics.fileName,
      fileSizeBytes: diagnostics.fileSizeBytes,
      modifiedAt: diagnostics.modifiedAt,
      errorSummary: diagnostics.errorSummary,
    },
    null,
    2,
  );
}

function appendDiagnosticEntry(
  listEl: HTMLDListElement,
  label: string,
  value: string,
): void {
  listEl.createEl("dt", { text: label });
  listEl.createEl("dd", { text: value });
}

function createMessageListItem(text: string): HTMLLIElement {
  const itemEl = activeDocument.createElement("li");
  itemEl.textContent = text;
  return itemEl;
}

function formatFileSizeMb(sizeBytes: number): string {
  return (sizeBytes / 1024 / 1024).toFixed(1);
}

function prepareRenderedImages(rootEl: HTMLElement): number {
  const imageEls = Array.from(rootEl.querySelectorAll("img"));
  for (const imageEl of imageEls) {
    imageEl.loading = "lazy";
    imageEl.decoding = "async";
  }
  return imageEls.length;
}

function collectBlobUrls(rootEl: HTMLElement): Set<string> {
  const urls = new Set<string>();
  const blobUrlPattern = /blob:[^)'"\s]+/g;
  const elements = [rootEl, ...Array.from(rootEl.querySelectorAll("*"))];

  for (const element of elements) {
    for (const attribute of Array.from(element.attributes)) {
      for (const match of attribute.value.matchAll(blobUrlPattern)) {
        urls.add(match[0]);
      }
    }

    if (element.instanceOf(HTMLStyleElement)) {
      for (const match of (element.textContent ?? "").matchAll(blobUrlPattern)) {
        urls.add(match[0]);
      }
    }
  }

  return urls;
}

function revokeBlobUrls(urls: Iterable<string>): void {
  for (const url of urls) {
    URL.revokeObjectURL(url);
  }
}

function yieldToBrowser(): Promise<void> {
  return new Promise((resolve) => {
    let completed = false;
    const finish = (): void => {
      if (completed) {
        return;
      }

      completed = true;
      window.clearTimeout(timeoutId);
      window.cancelAnimationFrame(frameId);
      resolve();
    };
    const timeoutId = window.setTimeout(finish, 50);
    const frameId = window.requestAnimationFrame(finish);
  });
}

interface RenderPerformanceMetrics {
  fileName: string;
  fileSizeBytes: number;
  readDurationMs: number;
  rendererDurationMs: number;
  commitDurationMs: number;
  outlineDurationMs: number;
  totalDurationMs: number;
  pageCount: number;
  imageCount: number;
  blobUrlCount: number;
}

function logRenderPerformance(metrics: RenderPerformanceMetrics): void {
  if (typeof __DEV__ === "undefined" || !__DEV__) {
    return;
  }

  console.debug("[Word Reader] Render performance", {
    ...metrics,
    readDurationMs: Math.round(metrics.readDurationMs),
    rendererDurationMs: Math.round(metrics.rendererDurationMs),
    commitDurationMs: Math.round(metrics.commitDurationMs),
    outlineDurationMs: Math.round(metrics.outlineDurationMs),
    totalDurationMs: Math.round(metrics.totalDurationMs),
  });
}

async function highlightText(
  rootEl: HTMLElement,
  query: string,
  shouldContinue: () => boolean,
): Promise<HTMLElement[]> {
  const normalizedQuery = query.toLowerCase();
  const walker = activeDocument.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parentEl = node.parentElement;
      if (
        !parentEl ||
        parentEl.closest("style, script, .word-reader-toolbar")
      ) {
        return NodeFilter.FILTER_REJECT;
      }
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const textNodes: Text[] = [];
  let scannedNodeCount = 0;
  while (walker.nextNode()) {
    const textNode = walker.currentNode as Text;
    if (textNode.nodeValue?.toLowerCase().includes(normalizedQuery)) {
      textNodes.push(textNode);
    }
    scannedNodeCount += 1;
    if (scannedNodeCount % SEARCH_SCAN_CHUNK_SIZE === 0) {
      await yieldToBrowser();
      if (!shouldContinue()) {
        return [];
      }
    }
  }

  const matches: HTMLElement[] = [];
  const escapedQuery = escapeRegExp(query);
  const regex = new RegExp(escapedQuery, "gi");

  for (
    let index = 0;
    index < textNodes.length;
    index += SEARCH_MUTATION_CHUNK_SIZE
  ) {
    if (!shouldContinue()) {
      return [];
    }

    for (const textNode of textNodes.slice(
      index,
      index + SEARCH_MUTATION_CHUNK_SIZE,
    )) {
      if (!textNode.isConnected || !rootEl.contains(textNode)) {
        continue;
      }

      const text = textNode.nodeValue ?? "";
      const fragment = activeDocument.createDocumentFragment();
      let lastIndex = 0;
      let match: RegExpExecArray | null;
      regex.lastIndex = 0;

      while ((match = regex.exec(text)) !== null) {
        if (match.index > lastIndex) {
          fragment.appendChild(
            activeDocument.createTextNode(text.slice(lastIndex, match.index)),
          );
        }

        // SECURITY: Only <span> elements are created for text highlighting.
        // No script execution or external resource loading occurs.
        const markEl = activeDocument.createElement("span");
        markEl.className = "word-reader-highlight";
        markEl.textContent = match[0];
        fragment.appendChild(markEl);
        matches.push(markEl);

        lastIndex = match.index + match[0].length;
      }

      if (lastIndex < text.length) {
        fragment.appendChild(
          activeDocument.createTextNode(text.slice(lastIndex)),
        );
      }

      textNode.replaceWith(fragment);
    }

    if (index + SEARCH_MUTATION_CHUNK_SIZE < textNodes.length) {
      await yieldToBrowser();
    }
  }

  return matches;
}

interface OutlineHeading {
  element: HTMLElement;
  level: number;
  text: string;
}

function getDocumentHeadings(rootEl: HTMLElement): OutlineHeading[] {
  const headingEls = Array.from(
    rootEl.querySelectorAll<HTMLElement>(
      "h1, h2, h3, h4, h5, h6, [class*='Heading'], [class*='heading']",
    ),
  );
  const seen = new Set<HTMLElement>();
  const headings: OutlineHeading[] = [];

  for (const headingEl of headingEls) {
    if (seen.has(headingEl)) {
      continue;
    }
    seen.add(headingEl);

    const text = normalizeWhitespace(headingEl.textContent ?? "");
    if (!text) {
      continue;
    }

    headings.push({
      element: headingEl,
      level: getHeadingLevel(headingEl),
      text,
    });
  }

  return headings;
}

function getHeadingLevel(element: HTMLElement): number {
  const tagMatch = /^H([1-6])$/.exec(element.tagName);
  if (tagMatch) {
    return Number(tagMatch[1]);
  }

  const classText = Array.from(element.classList).join(" ");
  const classMatch = /heading[-_ ]?([1-6])/i.exec(classText);
  return classMatch ? Number(classMatch[1]) : 2;
}

function htmlFragmentToMarkdown(fragment: DocumentFragment): string {
  return markdownForChildNodes(fragment).replace(/\n{3,}/g, "\n\n").trim();
}

function markdownForChildNodes(parent: Node): string {
  return Array.from(parent.childNodes).map(markdownForNode).join("");
}

function markdownForNode(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent ?? "";
  }

  if (!node.instanceOf(HTMLElement)) {
    return markdownForChildNodes(node);
  }

  const tagName = node.tagName.toLowerCase();
  const content = markdownForChildNodes(node).trim();

  switch (tagName) {
    case "h1":
    case "h2":
    case "h3":
    case "h4":
    case "h5":
    case "h6":
      return `${"#".repeat(Number(tagName.slice(1)))} ${content}\n\n`;
    case "p":
    case "div":
      return content ? `${content}\n\n` : "";
    case "strong":
    case "b":
      return content ? `**${content}**` : "";
    case "em":
    case "i":
      return content ? `*${content}*` : "";
    case "code":
      return content ? `\`${content.replace(/`/g, "\\`")}\`` : "";
    case "br":
      return "\n";
    case "a": {
      const href = node.getAttribute("href");
      return href && content ? `[${content}](${href})` : content;
    }
    case "ul":
      return `${markdownForList(node, "-")}\n`;
    case "ol":
      return `${markdownForList(node, "1.")}\n`;
    case "li":
      return content;
    case "table":
      return `${markdownForTable(node)}\n\n`;
    case "tr":
    case "tbody":
    case "thead":
    case "span":
    default:
      return markdownForChildNodes(node);
  }
}

function markdownForList(listEl: HTMLElement, marker: string): string {
  return Array.from(listEl.children)
    .filter((child): child is HTMLElement => child.instanceOf(HTMLElement))
    .filter((child) => child.tagName.toLowerCase() === "li")
    .map((itemEl) => `${marker} ${markdownForChildNodes(itemEl).trim()}`)
    .join("\n");
}

function markdownForTable(tableEl: HTMLElement): string {
  const rows = Array.from(tableEl.querySelectorAll("tr"))
    .map((rowEl) =>
      Array.from(rowEl.children).map((cellEl) =>
        normalizeWhitespace(cellEl.textContent ?? "").replace(/\|/g, "\\|"),
      ),
    )
    .filter((row) => row.length > 0);

  if (rows.length === 0) {
    return "";
  }

  const header = rows[0];
  const divider = header.map(() => "---");
  const body = rows.slice(1);
  return [header, divider, ...body]
    .map((row) => `| ${row.join(" | ")} |`)
    .join("\n");
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function roundToStep(value: number, step: number): number {
  return Math.round(value / step) * step;
}

class ImagePreviewModal extends Modal {
  private viewportEl: HTMLElement | null = null;
  private imageEl: HTMLImageElement | null = null;
  private statusEl: HTMLElement | null = null;
  private scale = 1;
  private translateX = 0;
  private translateY = 0;
  private naturalWidth = 0;
  private naturalHeight = 0;
  private isDragging = false;
  private dragStartX = 0;
  private dragStartY = 0;
  private dragStartTranslateX = 0;
  private dragStartTranslateY = 0;
  private readonly handleKeyDown = (event: KeyboardEvent): void => {
    if (
      event.key.toLowerCase() !== "c" ||
      (!event.ctrlKey && !event.metaKey) ||
      event.altKey ||
      event.shiftKey
    ) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    void this.copyImage();
  };

  constructor(
    app: App,
    private readonly src: string,
    private readonly alt: string,
    private readonly suggestedName: string,
    private readonly text: WordReaderText,
    private readonly releaseResource?: () => void,
  ) {
    super(app);
  }

  onOpen(): void {
    this.modalEl.addClass("word-reader-image-modal-container");
    this.contentEl.empty();
    this.titleEl.setText(this.alt);
    this.contentEl.addClass("word-reader-image-modal");
    this.contentEl.doc.addEventListener("keydown", this.handleKeyDown);

    const toolbarEl = this.contentEl.createDiv({
      cls: "word-reader-image-toolbar",
    });
    this.createToolbarButton(toolbarEl, "maximize", this.text.imagePreview.fitToWindow, () => {
      this.resetView();
    });
    this.createToolbarButton(toolbarEl, "scan", this.text.imagePreview.actualSize, () => {
      this.showActualSize();
    });
    this.createToolbarButton(toolbarEl, "copy", this.text.imagePreview.copyImage, () => {
      void this.copyImage();
    });
    this.createToolbarButton(toolbarEl, "download", this.text.imagePreview.saveImageAs, () => {
      void this.saveImageAs();
    });
    this.statusEl = toolbarEl.createSpan({
      cls: "word-reader-image-status",
      text: this.text.imagePreview.loading,
    });

    this.viewportEl = this.contentEl.createDiv({
      cls: "word-reader-image-viewport",
    });
    this.viewportEl.addEventListener("wheel", (event) => {
      this.handleWheel(event);
    });
    this.viewportEl.addEventListener("pointerdown", (event) => {
      this.handlePointerDown(event);
    });
    this.viewportEl.addEventListener("pointermove", (event) => {
      this.handlePointerMove(event);
    });
    this.viewportEl.addEventListener("pointerup", (event) => {
      this.handlePointerUp(event);
    });
    this.viewportEl.addEventListener("pointercancel", (event) => {
      this.handlePointerUp(event);
    });
    this.viewportEl.addEventListener("dblclick", () => {
      this.resetView();
    });

    this.imageEl = this.viewportEl.createEl("img", {
      cls: "word-reader-image-preview",
      attr: {
        src: this.src,
        alt: this.alt,
      },
    });
    this.imageEl.draggable = false;
    this.imageEl.addEventListener("load", () => {
      this.naturalWidth = this.imageEl?.naturalWidth ?? 0;
      this.naturalHeight = this.imageEl?.naturalHeight ?? 0;
      this.applyNaturalSize();
      this.resetView();
    });
    this.imageEl.addEventListener("error", () => {
      this.statusEl?.setText(this.text.imagePreview.loadFailed);
    });
  }

  onClose(): void {
    this.contentEl.doc.removeEventListener("keydown", this.handleKeyDown);
    this.releaseResource?.();
    this.contentEl.empty();
  }

  private createToolbarButton(
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

  private handleWheel(event: WheelEvent): void {
    if (!this.viewportEl || !this.imageEl) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const rect = this.viewportEl.getBoundingClientRect();
    const pointerX = event.clientX - rect.left - rect.width / 2;
    const pointerY = event.clientY - rect.top - rect.height / 2;
    const nextScale = clamp(
      this.scale * (event.deltaY < 0 ? IMAGE_ZOOM_STEP : 1 / IMAGE_ZOOM_STEP),
      MIN_IMAGE_ZOOM,
      MAX_IMAGE_ZOOM,
    );

    this.zoomAroundPoint(nextScale, pointerX, pointerY);
  }

  private handlePointerDown(event: PointerEvent): void {
    if (!this.viewportEl || event.button !== 0) {
      return;
    }

    event.preventDefault();
    this.isDragging = true;
    this.dragStartX = event.clientX;
    this.dragStartY = event.clientY;
    this.dragStartTranslateX = this.translateX;
    this.dragStartTranslateY = this.translateY;
    this.viewportEl.addClass("is-dragging");
    this.viewportEl.setPointerCapture(event.pointerId);
  }

  private handlePointerMove(event: PointerEvent): void {
    if (!this.isDragging) {
      return;
    }

    this.translateX =
      this.dragStartTranslateX + event.clientX - this.dragStartX;
    this.translateY =
      this.dragStartTranslateY + event.clientY - this.dragStartY;
    this.applyTransform();
  }

  private handlePointerUp(event: PointerEvent): void {
    if (!this.viewportEl || !this.isDragging) {
      return;
    }

    this.isDragging = false;
    this.viewportEl.removeClass("is-dragging");
    if (this.viewportEl.hasPointerCapture(event.pointerId)) {
      this.viewportEl.releasePointerCapture(event.pointerId);
    }
  }

  private resetView(): void {
    if (!this.viewportEl || !this.naturalWidth || !this.naturalHeight) {
      return;
    }

    const rect = this.viewportEl.getBoundingClientRect();
    const fitScale =
      rect.width > 0 && rect.height > 0
        ? Math.min(
            1,
            rect.width / this.naturalWidth,
            rect.height / this.naturalHeight,
          )
        : 1;
    this.scale = clamp(fitScale, MIN_IMAGE_ZOOM, 1);
    this.translateX = 0;
    this.translateY = 0;
    this.applyTransform();
  }

  private showActualSize(): void {
    this.scale = 1;
    this.translateX = 0;
    this.translateY = 0;
    this.applyTransform();
  }

  private zoomAroundPoint(
    nextScale: number,
    pointerX: number,
    pointerY: number,
  ): void {
    const imagePointX = (pointerX - this.translateX) / this.scale;
    const imagePointY = (pointerY - this.translateY) / this.scale;
    this.scale = nextScale;
    this.translateX = pointerX - imagePointX * this.scale;
    this.translateY = pointerY - imagePointY * this.scale;
    this.applyTransform();
  }

  private applyNaturalSize(): void {
    if (!this.imageEl || !this.naturalWidth || !this.naturalHeight) {
      return;
    }

    this.imageEl.style.width = `${this.naturalWidth}px`;
    this.imageEl.style.height = `${this.naturalHeight}px`;
  }

  private applyTransform(): void {
    if (!this.imageEl) {
      return;
    }

    this.imageEl.style.transform = `translate(${this.translateX}px, ${this.translateY}px) scale(${this.scale})`;
    this.updateStatus();
  }

  private updateStatus(): void {
    if (!this.statusEl) {
      return;
    }

    const dimensions =
      this.naturalWidth && this.naturalHeight
        ? `${this.naturalWidth} x ${this.naturalHeight}px`
        : this.text.imagePreview.unknownSize;
    this.statusEl.setText(`${dimensions} - ${Math.round(this.scale * 100)}%`);
  }

  private async copyImage(): Promise<void> {
    try {
      const clipboard = await getElectronClipboard();
      if (clipboard) {
        const image = await this.loadElectronImage();
        clipboard.writeImage(image);
      } else {
        await copyImageWithWebClipboard(await loadImageData(this.src));
      }

      new Notice(this.text.notices.copiedImage);
    } catch (error) {
      new Notice(this.text.notices.couldNotCopyImage(getErrorMessage(error)));
    }
  }

  private async saveImageAs(): Promise<void> {
    try {
      const imageData = await loadImageData(this.src);
      const fileName = this.getImageFileName(imageData.extension);
      const dialog = await getElectronDialog();
      if (!dialog) {
        downloadImageData(imageData, fileName);
        new Notice(this.text.notices.startedImageDownload);
        return;
      }

      try {
        const result = await dialog.showSaveDialog({
          title: this.text.imagePreview.saveDialogTitle,
          defaultPath: fileName,
          filters: [
            {
              name: this.text.imagePreview.imageFilterName(
                imageData.extension,
              ),
              extensions: [imageData.extension],
            },
            { name: this.text.imagePreview.allFiles, extensions: ["*"] },
          ],
        });

        if (result.canceled || !result.filePath) {
          return;
        }

        await writeImageFile(result.filePath, imageData.bytes);
        new Notice(this.text.notices.savedImage);
      } catch {
        downloadImageData(imageData, fileName);
        new Notice(this.text.notices.startedImageDownload);
      }
    } catch (error) {
      new Notice(this.text.notices.couldNotSaveImage(getErrorMessage(error)));
    }
  }

  private getImageFileName(extension: string): string {
    const baseName = sanitizeFileName(this.suggestedName);
    const dimensions =
      this.naturalWidth && this.naturalHeight
        ? `-${this.naturalWidth}x${this.naturalHeight}`
        : "";
    return `${baseName}-image${dimensions}.${extension}`;
  }

  private async loadElectronImage(): Promise<ElectronNativeImage> {
    const nativeImage = await getElectronNativeImage();
    if (this.src.startsWith("data:")) {
      const image = nativeImage.createFromDataURL(this.src);
      if (!image.isEmpty?.()) {
        return image;
      }
    }

    const imageData = await loadImageData(this.src);
    return nativeImage.createFromBuffer(Buffer.from(imageData.bytes));
  }
}

interface ElectronNativeImage {
  isEmpty?(): boolean;
}

interface ElectronNativeImageModule {
  createFromBuffer(buffer: Uint8Array): ElectronNativeImage;
  createFromDataURL(dataUrl: string): ElectronNativeImage;
}

interface ElectronClipboard {
  writeImage(image: ElectronNativeImage): void;
}

interface ElectronDialog {
  showSaveDialog(options: {
    title: string;
    defaultPath: string;
    filters: Array<{
      name: string;
      extensions: string[];
    }>;
  }): Promise<{
    canceled: boolean;
    filePath?: string;
  }>;
}

interface ImageData {
  bytes: Uint8Array;
  extension: string;
}

async function loadImageData(src: string): Promise<ImageData> {
  const { bytes, mimeType } = await loadLocalImageBytes(src);
  return {
    bytes,
    extension: getImageExtension(mimeType),
  };
}

function loadLocalImageBytes(
  src: string,
): Promise<{ bytes: Uint8Array; mimeType: string }> {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open("GET", src, true);
    request.responseType = "arraybuffer";
    request.addEventListener("load", () => {
      if (request.status !== 0 && (request.status < 200 || request.status >= 300)) {
        reject(new Error(`Image request failed: ${request.status}`));
        return;
      }

      const response: unknown = request.response;
      if (!(response instanceof ArrayBuffer)) {
        reject(new Error("Image request returned an invalid response"));
        return;
      }

      resolve({
        bytes: new Uint8Array(response),
        mimeType: request.getResponseHeader("content-type") ?? "",
      });
    });
    request.addEventListener("error", () => {
      reject(new Error("Image request failed"));
    });
    request.send();
  });
}

function getImageExtension(mimeType: string): string {
  switch (mimeType.toLowerCase()) {
    case "image/jpeg":
      return "jpg";
    case "image/gif":
      return "gif";
    case "image/webp":
      return "webp";
    case "image/bmp":
      return "bmp";
    case "image/svg+xml":
      return "svg";
    case "image/png":
    default:
      return "png";
  }
}

function getImageMimeType(extension: string): string {
  switch (extension) {
    case "jpg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "webp":
      return "image/webp";
    case "bmp":
      return "image/bmp";
    case "svg":
      return "image/svg+xml";
    case "png":
    default:
      return "image/png";
  }
}

function createImageBlob(imageData: ImageData): Blob {
  const buffer = new ArrayBuffer(imageData.bytes.byteLength);
  new Uint8Array(buffer).set(imageData.bytes);
  return new Blob([buffer], {
    type: getImageMimeType(imageData.extension),
  });
}

async function copyImageWithWebClipboard(imageData: ImageData): Promise<void> {
  if (!window.ClipboardItem || !navigator.clipboard?.write) {
    throw new Error("Clipboard image writing is unavailable");
  }

  const blob = createImageBlob(imageData);
  await navigator.clipboard.write([
    new window.ClipboardItem({
      [blob.type]: blob,
    }),
  ]);
}

function downloadImageData(imageData: ImageData, fileName: string): void {
  const blob = createImageBlob(imageData);
  const url = URL.createObjectURL(blob);
  const linkEl = activeDocument.createElement("a");
  linkEl.href = url;
  linkEl.download = fileName;
  linkEl.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1000);
}

function sanitizeFileName(value: string): string {
  const sanitized = Array.from(value, (character) => {
    const codePoint = character.codePointAt(0) ?? 0;
    return codePoint < 32 || '<>:"/\\|?*'.includes(character)
      ? "-"
      : character;
  })
    .join("")
    .trim();
  return sanitized || "word-image";
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function getElectronClipboard(): Promise<ElectronClipboard | null> {
  if (!Platform.isDesktopApp) {
    return null;
  }

  const electron = await import("electron");
  return electron.clipboard;
}

async function getElectronDialog(): Promise<ElectronDialog | null> {
  if (!Platform.isDesktopApp) {
    return null;
  }

  const electron = await import("electron");
  return electron.dialog;
}

async function getElectronNativeImage(): Promise<ElectronNativeImageModule> {
  if (!Platform.isDesktopApp) {
    throw new Error("Native image support is only available on desktop");
  }

  const electron = await import("electron");
  return electron.nativeImage;
}

async function writeImageFile(
  path: string,
  data: Uint8Array,
): Promise<void> {
  if (!Platform.isDesktopApp) {
    throw new Error("Saving images outside the vault is only available on desktop");
  }

  const { writeFile } = await import("node:fs/promises");
  await writeFile(path, data);
}
