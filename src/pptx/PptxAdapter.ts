import type { App, TFile } from "obsidian";

import {
  supportsReaderFile,
  type OfficeReaderAdapter,
  type ReaderTextExtraction,
} from "../reader/adapters";
import { defineReaderCapabilities } from "../reader/capabilities";
import {
  formatSlideText,
  type PptxSlideMetadata,
} from "./pptxMetadata";
import { PptxPackage } from "./pptxPackage";

export const PPTX_VIEW_TYPE = "pptx-reader-view";

export class PptxAdapter
implements OfficeReaderAdapter<PptxPackage, PptxSlideMetadata> {
  readonly format = "pptx" as const;
  readonly viewType = PPTX_VIEW_TYPE;
  readonly extensions = ["pptx"] as const;
  readonly capabilities = defineReaderCapabilities({
    reload: true,
    zoom: true,
    fit: true,
    navigation: true,
    search: true,
    copyText: true,
    notes: true,
    summaryNote: true,
    diagnostics: true,
    fullscreen: true,
    paged: true,
    openExternal: true,
  });

  supports(file: Pick<TFile, "extension">): boolean {
    return supportsReaderFile(this.extensions, file);
  }

  async open(app: App, file: TFile): Promise<PptxPackage> {
    const buffer = await app.vault.readBinary(file);
    return PptxPackage.load(buffer);
  }

  extractText(source: PptxSlideMetadata): ReaderTextExtraction {
    return {
      plainText: formatSlideText(source),
    };
  }
}

export const PPTX_ADAPTER = new PptxAdapter();
