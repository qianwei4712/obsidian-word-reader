import { setIcon } from "obsidian";

import {
  hasReaderCapability,
  type ReaderCapabilities,
  type ReaderCapability,
} from "./capabilities";
import type { ReaderStatus } from "./status";

export interface OfficeReaderShellLayout {
  rootEl: HTMLElement;
  toolbarEl: HTMLElement;
  statusEl: HTMLElement;
  bodyEl: HTMLElement;
}

export interface OfficeReaderShellOptions {
  rootClasses?: string;
  toolbarClasses?: string;
  statusClasses?: string;
  bodyClasses?: string;
}

export interface ReaderErrorView {
  title: string;
  body: string;
  tips: readonly string[];
  detailsLabel: string;
  details: string;
}

/**
 * Shared, format-neutral DOM shell. Format sessions own their inner body but
 * receive the same root, toolbar, status and error conventions.
 */
export class OfficeReaderShell {
  constructor(
    private readonly contentEl: HTMLElement,
    readonly capabilities: ReaderCapabilities,
  ) {}

  build(options: OfficeReaderShellOptions = {}): OfficeReaderShellLayout {
    this.contentEl.empty();
    const rootEl = this.contentEl.createDiv({
      cls: joinClasses("office-reader-root", options.rootClasses),
    });
    const toolbarEl = rootEl.createDiv({
      cls: joinClasses("office-reader-toolbar", options.toolbarClasses),
    });
    const statusEl = rootEl.createDiv({
      cls: joinClasses("office-reader-status", options.statusClasses),
      attr: {
        role: "status",
        "aria-live": "polite",
      },
    });
    const bodyEl = rootEl.createDiv({
      cls: joinClasses("office-reader-body", options.bodyClasses),
    });
    return { rootEl, toolbarEl, statusEl, bodyEl };
  }

  createToolbarButton(
    parentEl: HTMLElement,
    capability: ReaderCapability,
    icon: string,
    label: string,
    onClick: () => void,
  ): HTMLButtonElement | null {
    if (!hasReaderCapability(this.capabilities, capability)) {
      return null;
    }
    const buttonEl = parentEl.createEl("button", {
      cls: "office-reader-toolbar-button",
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

  renderStatus(statusEl: HTMLElement | null, status: ReaderStatus): void {
    if (!statusEl) {
      return;
    }
    statusEl.setText(status.message);
    statusEl.toggleClass("is-error", status.kind === "error");
    statusEl.toggleClass("is-loading", status.kind === "loading");
    statusEl.setAttribute("aria-busy", String(status.kind === "loading"));
  }

  renderError(
    parentEl: HTMLElement,
    error: ReaderErrorView,
  ): HTMLElement {
    const messageEl = parentEl.createDiv({
      cls: "office-reader-message office-reader-error",
    });
    messageEl.createDiv({
      cls: "office-reader-message-title word-reader-message-title",
      text: error.title,
    });
    messageEl.createDiv({
      cls: "office-reader-message-body word-reader-message-body",
      text: error.body,
    });
    if (error.tips.length > 0) {
      const listEl = messageEl.createEl("ul", {
        cls: "office-reader-message-list word-reader-message-list",
      });
      for (const tip of error.tips) {
        listEl.createEl("li", { text: tip });
      }
    }
    const detailsEl = messageEl.createEl("details", {
      cls: "office-reader-diagnostics word-reader-diagnostics",
    });
    detailsEl.createEl("summary", { text: error.detailsLabel });
    detailsEl.createEl("pre", {
      cls: "office-reader-diagnostics-details",
      text: error.details,
    });
    return messageEl;
  }
}

function joinClasses(
  required: string,
  optional: string | undefined,
): string {
  return optional ? `${required} ${optional}` : required;
}
