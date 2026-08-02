import type { App, TFile } from "obsidian";

import {
  supportsReaderFile,
  type OfficeReaderAdapter,
  type ReaderTextExtraction,
} from "../reader/adapters";
import { defineReaderCapabilities } from "../reader/capabilities";
import { XlsxPackage } from "./xlsxPackage";
import type { XlsxWorksheet } from "./xlsxWorksheet";

export const XLSX_VIEW_TYPE = "xlsx-reader-view";

export class XlsxAdapter
implements OfficeReaderAdapter<XlsxPackage, XlsxWorksheet> {
  readonly format = "xlsx" as const;
  readonly viewType = XLSX_VIEW_TYPE;
  readonly extensions = ["xlsx", "xls"] as const;
  readonly capabilities = defineReaderCapabilities({
    reload: true,
    zoom: true,
    fit: true,
    navigation: true,
    search: true,
    copyText: true,
    copyMarkdown: true,
    summaryNote: true,
    openExternal: true,
    diagnostics: true,
  });

  supports(file: Pick<TFile, "extension">): boolean {
    return supportsReaderFile(this.extensions, file);
  }

  async open(app: App, file: TFile): Promise<XlsxPackage> {
    const buffer = await app.vault.readBinary(file);
    return XlsxPackage.load(buffer);
  }

  extractText(source: XlsxWorksheet): ReaderTextExtraction {
    return {
      plainText: source
        .getPopulatedCells()
        .map((cell) => `${cell.ref}\t${cell.displayValue}`)
        .join("\n"),
    };
  }
}

export const XLSX_ADAPTER = new XlsxAdapter();
