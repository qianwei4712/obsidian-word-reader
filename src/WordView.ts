import {
  FileView,
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

const ZOOM_LEVELS = [
  { label: "80%", value: 0.8 },
  { label: "100%", value: 1 },
  { label: "125%", value: 1.25 },
  { label: "150%", value: 1.5 },
];

export class WordView extends FileView {
  private readonly plugin: WordReaderPlugin;
  private rootEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private scrollEl: HTMLElement | null = null;
  private documentEl: HTMLElement | null = null;
  private searchInputEl: HTMLInputElement | null = null;
  private searchCountEl: HTMLElement | null = null;
  private buffer: ArrayBuffer | null = null;
  private renderToken = 0;
  private zoom = 1;
  private fitWidth = false;
  private searchQuery = "";

  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf);
    this.plugin = plugin;
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
    this.documentEl?.empty();
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

    const zoomSelectEl = toolbarEl.createEl("select", {
      attr: { "aria-label": "Zoom" },
    });
    for (const zoomLevel of ZOOM_LEVELS) {
      zoomSelectEl.createEl("option", {
        text: zoomLevel.label,
        value: String(zoomLevel.value),
      });
    }
    zoomSelectEl.value = String(this.zoom);
    zoomSelectEl.addEventListener("change", () => {
      this.zoom = Number(zoomSelectEl.value);
      this.applyDocumentOptions();
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

    this.searchCountEl = toolbarEl.createSpan({
      cls: "word-reader-search-count",
      text: "0",
    });

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
    this.documentEl = this.scrollEl.createDiv({ cls: "word-reader-document" });
    this.applyDocumentOptions();
  }

  private async loadDocx(file: TFile): Promise<void> {
    this.ensureLayout();

    const token = ++this.renderToken;
    this.buffer = null;
    this.documentEl?.empty();
    this.setStatus(`Loading ${file.name}...`);

    if (file.stat.size > 25 * 1024 * 1024) {
      new Notice("Large Word document. Rendering may take a while.");
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

      this.showError(error);
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
  }

  private setStatus(message: string, isError = false): void {
    if (!this.statusEl) {
      return;
    }

    this.statusEl.setText(message);
    this.statusEl.toggleClass("is-error", isError);
  }

  private showError(error: unknown): void {
    const message = error instanceof Error ? error.message : String(error);
    this.documentEl?.empty();
    this.documentEl?.createDiv({
      cls: "word-reader-empty",
      text: "This Word document could not be rendered. Use the external open button to view it in Word or WPS.",
    });
    this.setStatus(`Failed to open Word document: ${message}`, true);
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

    const query = this.searchQuery.trim();
    if (!query) {
      this.searchCountEl?.setText("0");
      return;
    }

    const count = highlightText(this.documentEl, query);
    this.searchCountEl?.setText(String(count));
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

function highlightText(rootEl: HTMLElement, query: string): number {
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

  let count = 0;
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

      count += 1;
      lastIndex = match.index + match[0].length;
    }

    if (lastIndex < text.length) {
      fragment.appendText(text.slice(lastIndex));
    }

    textNode.replaceWith(fragment);
  }

  return count;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
