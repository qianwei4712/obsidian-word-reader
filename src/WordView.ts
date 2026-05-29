import {
  App,
  FileView,
  Modal,
  Notice,
  TFile,
  WorkspaceLeaf,
  setIcon,
} from "obsidian";

import type WordReaderPlugin from "./main";
import { createNoteFromDocx } from "./commands/createNoteFromDocx";
import { openExternalDocx } from "./commands/openExternal";
import { renderDocx } from "./renderer/docxRenderer";
import { extractPlainText } from "./renderer/textExtractor";

export const VIEW_TYPE_WORD_READER = "word-reader-view";

const MIN_ZOOM = 0.25;
const MAX_ZOOM = 4;
const ZOOM_STEP = 0.05;
const MIN_IMAGE_ZOOM = 0.1;
const MAX_IMAGE_ZOOM = 8;
const IMAGE_ZOOM_STEP = 1.1;

export class WordView extends FileView {
  private readonly plugin: WordReaderPlugin;
  private rootEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private scrollEl: HTMLElement | null = null;
  private documentEl: HTMLElement | null = null;
  private zoomInputEl: HTMLInputElement | null = null;
  private searchInputEl: HTMLInputElement | null = null;
  private searchPreviousButtonEl: HTMLButtonElement | null = null;
  private searchNextButtonEl: HTMLButtonElement | null = null;
  private searchCountEl: HTMLElement | null = null;
  private searchMatchEls: HTMLElement[] = [];
  private currentSearchIndex = -1;
  private buffer: ArrayBuffer | null = null;
  private renderToken = 0;
  private zoom = 1;
  private fitWidth = false;
  private searchQuery = "";

  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf);
    this.plugin = plugin;
    this.zoom = plugin.settings.defaultZoomPercent / 100;
    this.fitWidth = plugin.settings.defaultFitWidth;
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

  async onOpen(): Promise<void> {
    this.buildLayout();
  }

  async onLoadFile(file: TFile): Promise<void> {
    await this.loadDocx(file);
  }

  async onUnloadFile(): Promise<void> {
    this.buffer = null;
    this.searchMatchEls = [];
    this.currentSearchIndex = -1;
    this.documentEl?.empty();
    this.updateSearchState();
    this.setStatus("");
  }

  async reload(): Promise<void> {
    if (!this.file) {
      return;
    }

    await this.loadDocx(this.file);
  }

  async copyText(): Promise<void> {
    if (!this.file) {
      return;
    }

    if (this.isLegacyDocFile(this.file)) {
      new Notice("Legacy .doc files are not readable inside Obsidian. Open the file externally or convert it to .docx.");
      return;
    }

    const selectedText = this.getSelectedRenderedText();
    if (selectedText.length > 0) {
      await navigator.clipboard.writeText(selectedText);
      new Notice("Copied selected text");
      return;
    }

    if (!this.buffer) {
      this.buffer = await this.app.vault.readBinary(this.file);
    }

    this.setStatus("Extracting plain text...");
    const text = await extractPlainText(this.buffer);
    await navigator.clipboard.writeText(text);
    this.setStatus(`Copied plain text from ${this.file.name}`);
    new Notice("Copied plain text");
  }

  private buildLayout(): void {
    this.contentEl.empty();

    this.rootEl = this.contentEl.createDiv({ cls: "word-reader-root" });
    const toolbarEl = this.rootEl.createDiv({ cls: "word-reader-toolbar" });

    this.createIconButton(toolbarEl, "refresh-cw", "Reload", () => {
      void this.reload();
    });

    this.zoomInputEl = toolbarEl.createEl("input", {
      cls: "word-reader-zoom",
      attr: {
        type: "number",
        min: String(Math.round(MIN_ZOOM * 100)),
        max: String(Math.round(MAX_ZOOM * 100)),
        step: String(Math.round(ZOOM_STEP * 100)),
        "aria-label": "Zoom percentage",
        title: "Zoom percentage",
      },
    });
    this.updateZoomControl();
    this.zoomInputEl.addEventListener("change", () => {
      this.setZoom(Number(this.zoomInputEl?.value) / 100);
    });

    this.createIconButton(toolbarEl, "panel-top-open", "Fit width", () => {
      this.fitWidth = !this.fitWidth;
      if (this.buffer) {
        void this.renderCurrentBuffer(++this.renderToken);
      } else {
        this.applyDocumentOptions();
      }
    });

    this.searchInputEl = toolbarEl.createEl("input", {
      cls: "word-reader-search",
      attr: {
        type: "search",
        placeholder: "Search",
        "aria-label": "Search document",
      },
    });
    this.searchInputEl.addEventListener("input", () => {
      this.searchQuery = this.searchInputEl?.value ?? "";
      this.updateSearchHighlights();
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
      "Previous search result",
      () => {
        this.navigateSearch(-1);
      },
    );

    this.searchNextButtonEl = this.createIconButton(
      toolbarEl,
      "chevron-down",
      "Next search result",
      () => {
        this.navigateSearch(1);
      },
    );

    this.searchCountEl = toolbarEl.createSpan({
      cls: "word-reader-search-count",
      text: "0",
    });
    this.updateSearchState();

    this.createIconButton(toolbarEl, "copy", "Copy text", () => {
      void this.copyText();
    });

    this.createIconButton(toolbarEl, "file-plus", "Create summary note", () => {
      if (this.file) {
        void createNoteFromDocx(this.app, this.file);
      }
    });

    this.createIconButton(toolbarEl, "external-link", "Open externally", () => {
      if (this.file) {
        void openExternalDocx(this.app, this.file);
      }
    });

    this.statusEl = this.rootEl.createDiv({ cls: "word-reader-status" });
    this.scrollEl = this.rootEl.createDiv({ cls: "word-reader-scroll" });
    this.scrollEl.addEventListener("wheel", (event) => {
      this.handleCtrlWheelZoom(event);
    });
    this.documentEl = this.scrollEl.createDiv({ cls: "word-reader-document" });
    this.documentEl.addEventListener("click", (event) => {
      this.handleDocumentClick(event);
    });
    this.applyDocumentOptions();
  }

  private async loadDocx(file: TFile): Promise<void> {
    this.ensureLayout();

    const token = ++this.renderToken;
    this.buffer = null;
    this.documentEl?.empty();
    this.searchMatchEls = [];
    this.currentSearchIndex = -1;
    this.updateSearchState();

    if (this.isLegacyDocFile(file)) {
      this.showLegacyDocMessage(file);
      return;
    }

    this.setStatus(`Loading ${file.name}...`);

    const largeFileWarningBytes =
      this.plugin.settings.largeFileWarningMb * 1024 * 1024;
    if (file.stat.size > largeFileWarningBytes) {
      const fileSizeMb = formatFileSizeMb(file.stat.size);
      new Notice(
        `Large Word document (${fileSizeMb} MB). Rendering may take a while.`,
      );
      this.setStatus(
        `Large document detected (${fileSizeMb} MB). Rendering may take a while...`,
      );
    }

    try {
      const buffer = await this.app.vault.readBinary(file);
      if (token !== this.renderToken) {
        return;
      }

      this.buffer = buffer;
      await this.renderCurrentBuffer(token);
    } catch (error) {
      if (token !== this.renderToken) {
        return;
      }

      this.showError(error, file);
    }
  }

  private async renderCurrentBuffer(token: number): Promise<void> {
    if (!this.buffer || !this.documentEl || !this.file) {
      return;
    }

    this.documentEl.empty();
    this.setStatus(`Rendering ${this.file.name}...`);

    await renderDocx(this.buffer, this.documentEl, {
      fitWidth: this.fitWidth,
    });

    if (token !== this.renderToken) {
      return;
    }

    this.applyDocumentOptions();
    this.updateSearchHighlights();
    this.setStatus(`Read-only preview: ${this.file.name}`);
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

    if (!(event.target instanceof Element)) {
      return;
    }

    const imageEl = event.target.closest("img");
    if (!(imageEl instanceof HTMLImageElement)) {
      return;
    }

    const src = imageEl.currentSrc || imageEl.src;
    if (!src) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    new ImagePreviewModal(
      this.app,
      src,
      imageEl.alt || this.file?.name || "Word document image",
      this.file?.basename || "word-image",
    ).open();
  }

  private setStatus(message: string, isError = false): void {
    if (!this.statusEl) {
      return;
    }

    this.statusEl.setText(message);
    this.statusEl.toggleClass("is-error", isError);
  }

  private showLegacyDocMessage(file: TFile): void {
    if (!this.documentEl) {
      return;
    }

    this.documentEl.empty();
    this.documentEl.createDiv({
      cls: "word-reader-message-title",
      text: "Legacy .doc file",
    });
    this.documentEl.createDiv({
      cls: "word-reader-message-body",
      text: "This plugin renders .docx files directly in Obsidian. The older .doc format is a different binary Word format, so it cannot be rendered safely here.",
    });
    this.documentEl.createEl("ul", {
      cls: "word-reader-message-list",
    }).append(
      createMessageListItem("Open the file with Word, WPS, or your system default editor."),
      createMessageListItem("Convert or save the file as .docx, then open the converted file in Obsidian."),
    );
    const actionsEl = this.documentEl.createDiv({
      cls: "word-reader-message-actions",
    });
    this.createIconButton(actionsEl, "external-link", "Open externally", () => {
      void openExternalDocx(this.app, file);
    });
    this.setStatus(`Legacy .doc file: ${file.name}`);
  }

  private showError(error: unknown, file: TFile): void {
    const message = error instanceof Error ? error.message : String(error);
    const errorInfo = classifyWordError(message);
    this.documentEl?.empty();
    this.documentEl?.createDiv({
      cls: "word-reader-message-title",
      text: errorInfo.title,
    });
    this.documentEl?.createDiv({
      cls: "word-reader-message-body",
      text: errorInfo.body,
    });

    const actionsEl = this.documentEl?.createDiv({
      cls: "word-reader-message-actions",
    });
    if (actionsEl) {
      this.createIconButton(actionsEl, "external-link", "Open externally", () => {
        void openExternalDocx(this.app, file);
      });
    }

    this.setStatus(`${errorInfo.status}: ${message}`, true);
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

  private updateSearchHighlights(): void {
    if (!this.documentEl) {
      return;
    }

    clearSearchHighlights(this.documentEl);
    this.searchMatchEls = [];
    this.currentSearchIndex = -1;

    const query = this.searchQuery.trim();
    if (!query) {
      this.updateSearchState();
      return;
    }

    this.searchMatchEls = highlightText(this.documentEl, query);
    if (this.searchMatchEls.length > 0) {
      this.currentSearchIndex = 0;
    }

    this.updateSearchState({ scrollToCurrent: this.currentSearchIndex >= 0 });
  }

  private navigateSearch(direction: number): void {
    const matchCount = this.searchMatchEls.length;
    if (matchCount === 0) {
      this.updateSearchHighlights();
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
}

function clearSearchHighlights(rootEl: HTMLElement): void {
  const highlights = Array.from(
    rootEl.querySelectorAll("span.word-reader-highlight"),
  );

  for (const highlightEl of highlights) {
    const parentEl = highlightEl.parentNode;
    if (!parentEl) {
      continue;
    }

    while (highlightEl.firstChild) {
      parentEl.insertBefore(highlightEl.firstChild, highlightEl);
    }
    parentEl.removeChild(highlightEl);
    parentEl.normalize();
  }
}

interface WordErrorInfo {
  title: string;
  body: string;
  status: string;
}

function classifyWordError(message: string): WordErrorInfo {
  const normalizedMessage = message.toLowerCase();

  if (
    normalizedMessage.includes("password") ||
    normalizedMessage.includes("encrypt") ||
    normalizedMessage.includes("protected")
  ) {
    return {
      title: "Encrypted Word document",
      body: "This document appears to be encrypted or password protected. Obsidian Word Reader cannot unlock it. Open it in Word or WPS, remove the password if appropriate, then save a readable .docx copy.",
      status: "Encrypted Word document",
    };
  }

  if (
    normalizedMessage.includes("corrupt") ||
    normalizedMessage.includes("invalid") ||
    normalizedMessage.includes("malformed") ||
    normalizedMessage.includes("zip") ||
    normalizedMessage.includes("central directory") ||
    normalizedMessage.includes("xml")
  ) {
    return {
      title: "Damaged or unsupported Word document",
      body: "This document could not be parsed as a valid .docx file. It may be damaged, incomplete, or saved in a format that looks like .docx but is not readable by the renderer. Try opening it in Word or WPS and saving a fresh .docx copy.",
      status: "Damaged Word document",
    };
  }

  return {
    title: "Word document could not be rendered",
    body: "This Word document could not be rendered inside Obsidian. Open it externally to inspect the file, then try saving a fresh .docx copy if the problem continues.",
    status: "Failed to open Word document",
  };
}

function createMessageListItem(text: string): HTMLLIElement {
  const itemEl = document.createElement("li");
  itemEl.textContent = text;
  return itemEl;
}

function formatFileSizeMb(sizeBytes: number): string {
  return (sizeBytes / 1024 / 1024).toFixed(1);
}

function highlightText(rootEl: HTMLElement, query: string): HTMLElement[] {
  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parentEl = node.parentElement;
      if (!parentEl || parentEl.closest(".word-reader-toolbar")) {
        return NodeFilter.FILTER_REJECT;
      }

      if (!node.nodeValue?.toLowerCase().includes(query.toLowerCase())) {
        return NodeFilter.FILTER_REJECT;
      }

      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const textNodes: Text[] = [];
  while (walker.nextNode()) {
    textNodes.push(walker.currentNode as Text);
  }

  const matches: HTMLElement[] = [];
  const escapedQuery = escapeRegExp(query);
  const regex = new RegExp(escapedQuery, "gi");

  for (const textNode of textNodes) {
    const text = textNode.nodeValue ?? "";
    const fragment = document.createDocumentFragment();
    let lastIndex = 0;
    let match: RegExpExecArray | null;

    while ((match = regex.exec(text)) !== null) {
      if (match.index > lastIndex) {
        fragment.appendText(text.slice(lastIndex, match.index));
      }

      const markEl = document.createElement("span");
      markEl.className = "word-reader-highlight";
      markEl.textContent = match[0];
      fragment.appendChild(markEl);
      matches.push(markEl);

      lastIndex = match.index + match[0].length;
    }

    if (lastIndex < text.length) {
      fragment.appendText(text.slice(lastIndex));
    }

    textNode.replaceWith(fragment);
  }

  return matches;
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

  constructor(
    app: App,
    private readonly src: string,
    private readonly alt: string,
    private readonly suggestedName: string,
  ) {
    super(app);
  }

  onOpen(): void {
    this.modalEl.addClass("word-reader-image-modal-container");
    this.contentEl.empty();
    this.titleEl.setText(this.alt);
    this.contentEl.addClass("word-reader-image-modal");

    const toolbarEl = this.contentEl.createDiv({
      cls: "word-reader-image-toolbar",
    });
    this.createToolbarButton(toolbarEl, "maximize", "Fit to window", () => {
      this.resetView();
    });
    this.createToolbarButton(toolbarEl, "scan", "Actual size", () => {
      this.showActualSize();
    });
    this.createToolbarButton(toolbarEl, "copy", "Copy image", () => {
      void this.copyImage();
    });
    this.createToolbarButton(toolbarEl, "download", "Save image as", () => {
      void this.saveImageAs();
    });
    this.statusEl = toolbarEl.createSpan({
      cls: "word-reader-image-status",
      text: "Loading image...",
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
      this.statusEl?.setText("Image could not be loaded");
    });
  }

  onClose(): void {
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
        : "Unknown size";
    this.statusEl.setText(`${dimensions} - ${Math.round(this.scale * 100)}%`);
  }

  private async copyImage(): Promise<void> {
    try {
      const clipboard = getElectronClipboard();
      if (clipboard) {
        const image = await this.loadElectronImage();
        clipboard.writeImage(image);
      } else {
        await copyImageWithWebClipboard(await loadImageData(this.src));
      }

      new Notice("Copied image");
    } catch (error) {
      new Notice(`Could not copy image: ${getErrorMessage(error)}`);
    }
  }

  private async saveImageAs(): Promise<void> {
    try {
      const imageData = await loadImageData(this.src);
      const fileName = `${sanitizeFileName(this.suggestedName)}.${imageData.extension}`;
      const dialog = getElectronDialog();
      if (!dialog) {
        downloadImageData(imageData, fileName);
        new Notice("Started image download");
        return;
      }

      try {
        const result = await dialog.showSaveDialog({
          title: "Save image as",
          defaultPath: fileName,
          filters: [
            {
              name: `${imageData.extension.toUpperCase()} image`,
              extensions: [imageData.extension],
            },
            { name: "All files", extensions: ["*"] },
          ],
        });

        if (result.canceled || !result.filePath) {
          return;
        }

        await getNodeFsPromises().writeFile(result.filePath, imageData.bytes);
        new Notice("Saved image");
      } catch {
        downloadImageData(imageData, fileName);
        new Notice("Started image download");
      }
    } catch (error) {
      new Notice(`Could not save image: ${getErrorMessage(error)}`);
    }
  }

  private async loadElectronImage(): Promise<ElectronNativeImage> {
    const nativeImage = getElectronNativeImage();
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

interface NodeFsPromises {
  writeFile(path: string, data: Uint8Array): Promise<void>;
}

interface ImageData {
  bytes: Uint8Array;
  extension: string;
}

async function loadImageData(src: string): Promise<ImageData> {
  const response = await fetch(src);
  if (!response.ok) {
    throw new Error(`Image request failed: ${response.status}`);
  }

  const blob = await response.blob();
  const bytes = new Uint8Array(await blob.arrayBuffer());
  return {
    bytes,
    extension: getImageExtension(blob.type),
  };
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
  const buffer = imageData.bytes.buffer.slice(
    imageData.bytes.byteOffset,
    imageData.bytes.byteOffset + imageData.bytes.byteLength,
  ) as ArrayBuffer;
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
  const linkEl = document.createElement("a");
  linkEl.href = url;
  linkEl.download = fileName;
  linkEl.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1000);
}

function sanitizeFileName(value: string): string {
  const sanitized = value.replace(/[<>:"/\\|?*\x00-\x1f]/g, "-").trim();
  return sanitized || "word-image";
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function getElectronClipboard(): ElectronClipboard | null {
  const electron = require("electron") as { clipboard?: ElectronClipboard };
  return electron.clipboard ?? null;
}

function getElectronDialog(): ElectronDialog | null {
  const electron = require("electron") as { dialog?: ElectronDialog };
  return electron.dialog ?? null;
}

function getElectronNativeImage(): ElectronNativeImageModule {
  const electron = require("electron") as {
    nativeImage: ElectronNativeImageModule;
  };
  return electron.nativeImage;
}

function getNodeFsPromises(): NodeFsPromises {
  return require("fs/promises") as NodeFsPromises;
}
