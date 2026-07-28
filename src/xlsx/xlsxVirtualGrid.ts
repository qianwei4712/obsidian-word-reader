import type { XlsxCell } from "./xlsxTypes";
import { XlsxWorksheet } from "./xlsxWorksheet";

const DEFAULT_OVERSCAN_ROWS = 4;
const DEFAULT_OVERSCAN_COLUMNS = 2;
const DEFAULT_MAX_MOUNTED_CELLS = 2_500;
const MAX_OVERSCAN = 50;

export interface XlsxGridViewport {
  scrollTop: number;
  scrollLeft: number;
  width: number;
  height: number;
}

export interface XlsxVirtualWindow {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  offsetTop: number;
  offsetLeft: number;
  width: number;
  height: number;
  mountedCellCount: number;
  estimatedDomNodeCount: number;
  populatedCells: readonly XlsxCell[];
}

export interface XlsxVirtualGridOptions {
  overscanRows?: number;
  overscanColumns?: number;
  maxMountedCells?: number;
}

export class XlsxVirtualGrid {
  private readonly rows: SparseAxis;
  private readonly columns: SparseAxis;
  private readonly overscanRows: number;
  private readonly overscanColumns: number;
  private readonly maxMountedCells: number;

  constructor(
    private readonly worksheet: XlsxWorksheet,
    options: XlsxVirtualGridOptions = {},
  ) {
    this.rows = new SparseAxis(
      worksheet.rowCount,
      worksheet.defaultRowHeight,
      worksheet.rowHeights,
    );
    this.columns = new SparseAxis(
      worksheet.columnCount,
      worksheet.defaultColumnWidth,
      worksheet.columnWidths,
    );
    this.overscanRows = readBoundedInteger(
      options.overscanRows,
      DEFAULT_OVERSCAN_ROWS,
      0,
      MAX_OVERSCAN,
    );
    this.overscanColumns = readBoundedInteger(
      options.overscanColumns,
      DEFAULT_OVERSCAN_COLUMNS,
      0,
      MAX_OVERSCAN,
    );
    this.maxMountedCells = readBoundedInteger(
      options.maxMountedCells,
      DEFAULT_MAX_MOUNTED_CELLS,
      1,
      100_000,
    );
  }

  get totalHeight(): number {
    return this.rows.totalSize;
  }

  get totalWidth(): number {
    return this.columns.totalSize;
  }

  calculate(viewport: XlsxGridViewport): XlsxVirtualWindow {
    const scrollTop = clamp(viewport.scrollTop, 0, this.totalHeight);
    const scrollLeft = clamp(viewport.scrollLeft, 0, this.totalWidth);
    const width = Math.max(0, viewport.width);
    const height = Math.max(0, viewport.height);
    const visibleStartRow = this.rows.indexAtOffset(scrollTop);
    const visibleEndRow = Math.min(
      this.worksheet.rowCount,
      this.rows.indexAtOffset(scrollTop + height) + 1,
    );
    const visibleStartColumn = this.columns.indexAtOffset(scrollLeft);
    const visibleEndColumn = Math.min(
      this.worksheet.columnCount,
      this.columns.indexAtOffset(scrollLeft + width) + 1,
    );

    let startRow = Math.max(0, visibleStartRow - this.overscanRows);
    let endRow = Math.min(
      this.worksheet.rowCount,
      visibleEndRow + this.overscanRows,
    );
    let startColumn = Math.max(
      0,
      visibleStartColumn - this.overscanColumns,
    );
    let endColumn = Math.min(
      this.worksheet.columnCount,
      visibleEndColumn + this.overscanColumns,
    );

    if (
      (endRow - startRow) * (endColumn - startColumn) >
      this.maxMountedCells
    ) {
      startRow = visibleStartRow;
      endRow = visibleEndRow;
      startColumn = visibleStartColumn;
      endColumn = visibleEndColumn;
    }
    if (
      (endRow - startRow) * (endColumn - startColumn) >
      this.maxMountedCells
    ) {
      const columnCount = Math.max(
        1,
        Math.min(endColumn - startColumn, this.maxMountedCells),
      );
      endColumn = startColumn + columnCount;
      const rowCount = Math.max(
        1,
        Math.floor(this.maxMountedCells / columnCount),
      );
      endRow = Math.min(endRow, startRow + rowCount);
    }

    const mountedRows = Math.max(0, endRow - startRow);
    const mountedColumns = Math.max(0, endColumn - startColumn);
    const mountedCellCount = mountedRows * mountedColumns;
    return {
      startRow,
      endRow,
      startColumn,
      endColumn,
      offsetTop: this.rows.offsetAt(startRow),
      offsetLeft: this.columns.offsetAt(startColumn),
      width:
        this.columns.offsetAt(endColumn) -
        this.columns.offsetAt(startColumn),
      height: this.rows.offsetAt(endRow) - this.rows.offsetAt(startRow),
      mountedCellCount,
      estimatedDomNodeCount:
        3 + mountedRows + mountedColumns + mountedCellCount,
      populatedCells: this.worksheet.getCellsInWindow(
        startRow,
        endRow,
        startColumn,
        endColumn,
      ),
    };
  }
}

class SparseAxis {
  private readonly overrides: ReadonlyArray<{
    index: number;
    cumulativeDelta: number;
  }>;

  constructor(
    readonly count: number,
    private readonly defaultSize: number,
    overrideSizes: ReadonlyMap<number, number>,
  ) {
    if (!Number.isInteger(count) || count < 1) {
      throw new RangeError("A virtual grid axis must contain at least one item.");
    }
    if (!Number.isFinite(defaultSize) || defaultSize <= 0) {
      throw new RangeError("A virtual grid axis size must be positive.");
    }
    let cumulativeDelta = 0;
    this.overrides = Array.from(overrideSizes.entries())
      .filter(
        ([index, size]) =>
          Number.isInteger(index) &&
          index >= 0 &&
          index < count &&
          Number.isFinite(size) &&
          size > 0,
      )
      .sort(([left], [right]) => left - right)
      .map(([index, size]) => {
        cumulativeDelta += size - defaultSize;
        return { index, cumulativeDelta };
      });
  }

  get totalSize(): number {
    return this.offsetAt(this.count);
  }

  offsetAt(index: number): number {
    const boundedIndex = clamp(Math.floor(index), 0, this.count);
    const overrideIndex = this.lastOverrideBefore(boundedIndex);
    const delta =
      overrideIndex < 0
        ? 0
        : this.overrides[overrideIndex].cumulativeDelta;
    return boundedIndex * this.defaultSize + delta;
  }

  indexAtOffset(offset: number): number {
    const boundedOffset = clamp(offset, 0, this.totalSize);
    let low = 0;
    let high = this.count;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (this.offsetAt(middle + 1) <= boundedOffset) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    return Math.min(this.count - 1, low);
  }

  private lastOverrideBefore(index: number): number {
    let low = 0;
    let high = this.overrides.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (this.overrides[middle].index < index) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    return low - 1;
  }
}

function readBoundedInteger(
  value: number | undefined,
  fallback: number,
  minimum: number,
  maximum: number,
): number {
  if (value === undefined) {
    return fallback;
  }
  if (!Number.isInteger(value) || value < minimum || value > maximum) {
    throw new RangeError(
      `Virtual grid option must be an integer from ${minimum} to ${maximum}.`,
    );
  }
  return value;
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}
