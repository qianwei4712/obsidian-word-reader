import type { XlsxCell } from "./xlsxTypes";
import type { XlsxCellPosition } from "./xlsxSelection";
import type { XlsxWorksheet } from "./xlsxWorksheet";

const DEFAULT_SEARCH_CHUNK_SIZE = 1_000;

export interface XlsxSearchOptions {
  chunkSize?: number;
  isCancelled?: () => boolean;
  yieldControl?: () => Promise<void>;
}

export interface XlsxWorkbookSearchSource {
  sheets: readonly { name: string }[];
  getWorksheet: (
    index: number,
    options?: { isCancelled?: () => boolean },
  ) => Promise<Pick<XlsxWorksheet, "getPopulatedCells">>;
}

export interface XlsxWorkbookSearchResult extends XlsxCellPosition {
  sheetIndex: number;
  sheetName: string;
}

export interface XlsxWorkbookSearchOptions extends XlsxSearchOptions {
  onSheet?: (
    sheetName: string,
    sheetIndex: number,
    sheetCount: number,
  ) => void;
}

export class XlsxSearchCancelledError extends Error {
  constructor() {
    super("The XLSX search was cancelled.");
    this.name = "XlsxSearchCancelledError";
  }
}

export async function searchXlsxCells(
  cells: readonly XlsxCell[],
  query: string,
  options: XlsxSearchOptions = {},
): Promise<XlsxCellPosition[]> {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  if (!normalizedQuery) {
    return [];
  }

  const chunkSize = normalizeChunkSize(options.chunkSize);
  const results: XlsxCellPosition[] = [];
  for (let index = 0; index < cells.length; index += chunkSize) {
    throwIfCancelled(options);
    const chunkEnd = Math.min(index + chunkSize, cells.length);
    for (let cellIndex = index; cellIndex < chunkEnd; cellIndex += 1) {
      const cell = cells[cellIndex];
      const searchable = `${cell.displayValue}\n${cell.formula?.text ?? ""}`
        .toLocaleLowerCase();
      if (searchable.includes(normalizedQuery)) {
        results.push({ row: cell.row, column: cell.column });
      }
    }
    if (chunkEnd < cells.length && options.yieldControl) {
      await options.yieldControl();
    }
  }
  throwIfCancelled(options);
  return results;
}

export async function searchXlsxWorkbook(
  workbook: XlsxWorkbookSearchSource,
  query: string,
  options: XlsxWorkbookSearchOptions = {},
): Promise<XlsxWorkbookSearchResult[]> {
  const normalizedQuery = query.trim();
  if (!normalizedQuery) {
    return [];
  }
  const results: XlsxWorkbookSearchResult[] = [];
  for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex += 1) {
    throwIfCancelled(options);
    const descriptor = workbook.sheets[sheetIndex];
    options.onSheet?.(
      descriptor.name,
      sheetIndex,
      workbook.sheets.length,
    );
    const worksheet = await workbook.getWorksheet(sheetIndex, {
      isCancelled: options.isCancelled,
    });
    const matches = await searchXlsxCells(
      worksheet.getPopulatedCells(),
      normalizedQuery,
      options,
    );
    results.push(
      ...matches.map((match) => ({
        ...match,
        sheetIndex,
        sheetName: descriptor.name,
      })),
    );
    if (
      sheetIndex + 1 < workbook.sheets.length &&
      options.yieldControl
    ) {
      await options.yieldControl();
    }
  }
  throwIfCancelled(options);
  return results;
}

function normalizeChunkSize(chunkSize: number | undefined): number {
  if (!Number.isFinite(chunkSize) || !chunkSize || chunkSize < 1) {
    return DEFAULT_SEARCH_CHUNK_SIZE;
  }
  return Math.floor(chunkSize);
}

function throwIfCancelled(options: XlsxSearchOptions): void {
  if (options.isCancelled?.()) {
    throw new XlsxSearchCancelledError();
  }
}
