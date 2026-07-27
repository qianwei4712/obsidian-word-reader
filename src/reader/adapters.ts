import type { App, TFile } from "obsidian";

import type { ReaderCapabilities } from "./capabilities";
import type { ReaderFormat } from "./readingState";

export interface ReaderTextExtraction {
  plainText: string;
  markdown?: string;
}

export interface OfficeReaderAdapter<OpenedDocument, TextSource> {
  readonly format: ReaderFormat;
  readonly viewType: string;
  readonly extensions: readonly string[];
  readonly capabilities: ReaderCapabilities;

  supports(file: Pick<TFile, "extension">): boolean;
  open(app: App, file: TFile): Promise<OpenedDocument>;
  extractText(source: TextSource): ReaderTextExtraction;
}

export function supportsReaderFile(
  extensions: readonly string[],
  file: Pick<TFile, "extension">,
): boolean {
  const extension = file.extension.toLowerCase();
  return extensions.some((candidate) => candidate === extension);
}
