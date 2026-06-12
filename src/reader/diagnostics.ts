export interface ReaderDiagnosticFile {
  name: string;
  size: number;
  mtime: number;
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

export function formatReaderDiagnostics(
  product: string,
  diagnostics: ReaderDiagnostics,
): string {
  return JSON.stringify(
    {
      product,
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

export function fingerprintMessage(message: string): string {
  let hash = 0x811c9dc5;
  for (const character of message) {
    hash ^= character.charCodeAt(0);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}
