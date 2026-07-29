import { columnIndexToName } from "./xlsxReferences";
import type { XlsxCell } from "./xlsxTypes";
import type { XlsxWorksheet } from "./xlsxWorksheet";

export const MAX_XLSX_COPY_CELLS = 250_000;

export interface XlsxCellPosition {
  row: number;
  column: number;
}

export interface XlsxSelectionRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export type XlsxCopyMode = "display" | "formula";

export function normalizeXlsxSelection(
  anchor: XlsxCellPosition,
  focus: XlsxCellPosition,
): XlsxSelectionRange {
  return {
    startRow: Math.min(anchor.row, focus.row),
    endRow: Math.max(anchor.row, focus.row),
    startColumn: Math.min(anchor.column, focus.column),
    endColumn: Math.max(anchor.column, focus.column),
  };
}

export function xlsxSelectionContains(
  selection: XlsxSelectionRange,
  row: number,
  column: number,
): boolean {
  return (
    row >= selection.startRow &&
    row <= selection.endRow &&
    column >= selection.startColumn &&
    column <= selection.endColumn
  );
}

export function getXlsxSelectionCellCount(
  selection: XlsxSelectionRange,
): number {
  return (
    (selection.endRow - selection.startRow + 1) *
    (selection.endColumn - selection.startColumn + 1)
  );
}

export function xlsxSelectionToTsv(
  worksheet: Pick<XlsxWorksheet, "getCell">,
  selection: XlsxSelectionRange,
  mode: XlsxCopyMode,
): string {
  assertCopySize(selection);

  const lines: string[] = [];
  for (let row = selection.startRow; row <= selection.endRow; row += 1) {
    const values: string[] = [];
    for (
      let column = selection.startColumn;
      column <= selection.endColumn;
      column += 1
    ) {
      values.push(
        escapeTsvValue(
          getCellCopyValue(worksheet.getCell(row, column), mode),
        ),
      );
    }
    lines.push(values.join("\t"));
  }
  return lines.join("\n");
}

export function xlsxSelectionToMarkdown(
  worksheet: Pick<XlsxWorksheet, "getCell">,
  selection: XlsxSelectionRange,
): string {
  assertCopySize(selection);
  const headers: string[] = [];
  for (
    let column = selection.startColumn;
    column <= selection.endColumn;
    column += 1
  ) {
    headers.push(columnIndexToName(column));
  }
  const rows = [
    markdownRow(headers),
    markdownRow(headers.map(() => "---")),
  ];
  for (let row = selection.startRow; row <= selection.endRow; row += 1) {
    const values: string[] = [];
    for (
      let column = selection.startColumn;
      column <= selection.endColumn;
      column += 1
    ) {
      values.push(
        escapeMarkdownTableValue(
          getCellCopyValue(worksheet.getCell(row, column), "display"),
        ),
      );
    }
    rows.push(markdownRow(values));
  }
  return rows.join("\n");
}

export class XlsxSelectionTooLargeError extends Error {
  constructor(readonly cellCount: number) {
    super(
      `The selected range contains ${cellCount} cells, exceeding the ${MAX_XLSX_COPY_CELLS} cell copy limit.`,
    );
    this.name = "XlsxSelectionTooLargeError";
  }
}

function assertCopySize(selection: XlsxSelectionRange): void {
  const cellCount = getXlsxSelectionCellCount(selection);
  if (cellCount > MAX_XLSX_COPY_CELLS) {
    throw new XlsxSelectionTooLargeError(cellCount);
  }
}

function getCellCopyValue(
  cell: XlsxCell | undefined,
  mode: XlsxCopyMode,
): string {
  if (!cell) {
    return "";
  }
  if (mode === "formula" && cell.formula) {
    return `=${cell.formula.text}`;
  }
  return cell.displayValue;
}

function escapeTsvValue(value: string): string {
  if (!/[\t\n\r"]/.test(value)) {
    return value;
  }
  return `"${value.replace(/"/g, "\"\"")}"`;
}

function escapeMarkdownTableValue(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\\/g, "\\\\")
    .replace(/\|/g, "\\|")
    .replace(/\r\n|\r|\n/g, "<br>");
}

function markdownRow(values: readonly string[]): string {
  return `| ${values.join(" | ")} |`;
}
