import type { ReaderFormat } from "./readingState";

export interface ReaderDiagnosticFile {
  name: string;
  size: number;
  mtime: number;
}

export interface ReaderDiagnosticReport<Details = Record<string, unknown>> {
  product: "Office Reader";
  schemaVersion: 1;
  format: ReaderFormat;
  kind: "error" | "render";
  file: {
    name: string;
    sizeBytes: number;
    modifiedAt: string;
  };
  summary: string;
  details: Details;
}

export interface ReaderDiagnostics<Category extends string = string> {
  category: Category;
  fileName: string;
  fileSizeBytes: number;
  modifiedAt: string;
  errorSummary: string;
}

export function createReaderDiagnostics<Category extends string>(
  category: Category,
  errorSummary: string,
  file: ReaderDiagnosticFile,
): ReaderDiagnostics<Category> {
  return {
    category,
    fileName: file.name,
    fileSizeBytes: file.size,
    modifiedAt: new Date(file.mtime).toISOString(),
    errorSummary,
  };
}

export function createReaderDiagnosticReport<Details>(
  format: ReaderFormat,
  kind: "error" | "render",
  file: ReaderDiagnosticFile,
  summary: string,
  details: Details,
): ReaderDiagnosticReport<Details> {
  return {
    product: "Office Reader",
    schemaVersion: 1,
    format,
    kind,
    file: {
      name: file.name,
      sizeBytes: file.size,
      modifiedAt: new Date(file.mtime).toISOString(),
    },
    summary,
    details,
  };
}

export function formatReaderDiagnosticReport(
  report: ReaderDiagnosticReport<unknown>,
): string {
  return JSON.stringify(report, null, 2);
}

export function fingerprintMessage(message: string): string {
  let hash = 0x811c9dc5;
  for (const character of message) {
    hash ^= character.charCodeAt(0);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}
