import type { App, TFile } from "obsidian";

import {
  supportsReaderFile,
  type OfficeReaderAdapter,
  type ReaderTextExtraction,
} from "../reader/adapters";
import { defineReaderCapabilities } from "../reader/capabilities";
import { extractDocxText } from "./docxText";

export const DOCX_VIEW_TYPE = "word-reader-view";

export class DocxAdapter
implements OfficeReaderAdapter<ArrayBuffer, HTMLElement> {
  readonly format = "docx" as const;
  readonly viewType = DOCX_VIEW_TYPE;
  readonly extensions = ["docx", "doc"] as const;
  readonly capabilities = defineReaderCapabilities({
    reload: true,
    zoom: true,
    fit: true,
    navigation: true,
    search: true,
    copyText: true,
    copyMarkdown: true,
    summaryNote: true,
    diagnostics: true,
    openExternal: true,
  });

  supports(file: Pick<TFile, "extension">): boolean {
    return supportsReaderFile(this.extensions, file);
  }

  open(app: App, file: TFile): Promise<ArrayBuffer> {
    return app.vault.readBinary(file);
  }

  extractText(source: HTMLElement): ReaderTextExtraction {
    return extractDocxText(source);
  }
}

export const DOCX_ADAPTER = new DocxAdapter();
