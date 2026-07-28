import type { XlsxMergeRange } from "./xlsxTypes";

const MAX_XLSX_ROW = 1_048_576;
const MAX_XLSX_COLUMN = 16_384;

export interface XlsxCellPosition {
  row: number;
  column: number;
}

export function parseCellReference(reference: string): XlsxCellPosition {
  const match = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/i.exec(reference);
  if (!match) {
    throw new Error(`Invalid XLSX cell reference: ${reference}`);
  }
  const column = columnNameToIndex(match[1]);
  const row = Number(match[2]) - 1;
  if (
    row < 0 ||
    row >= MAX_XLSX_ROW ||
    column < 0 ||
    column >= MAX_XLSX_COLUMN
  ) {
    throw new Error(`XLSX cell reference is outside the worksheet: ${reference}`);
  }
  return { row, column };
}

export function columnNameToIndex(name: string): number {
  let value = 0;
  for (const character of name.toUpperCase()) {
    const digit = character.charCodeAt(0) - 64;
    if (digit < 1 || digit > 26) {
      throw new Error(`Invalid XLSX column name: ${name}`);
    }
    value = value * 26 + digit;
  }
  return value - 1;
}

export function columnIndexToName(index: number): string {
  if (!Number.isInteger(index) || index < 0 || index >= MAX_XLSX_COLUMN) {
    throw new RangeError(`XLSX column index is outside the worksheet: ${index}`);
  }
  let remaining = index + 1;
  let name = "";
  while (remaining > 0) {
    const digit = (remaining - 1) % 26;
    name = String.fromCharCode(65 + digit) + name;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return name;
}

export function makeCellReference(row: number, column: number): string {
  if (!Number.isInteger(row) || row < 0 || row >= MAX_XLSX_ROW) {
    throw new RangeError(`XLSX row index is outside the worksheet: ${row}`);
  }
  return `${columnIndexToName(column)}${row + 1}`;
}

export function parseRangeReference(reference: string): XlsxMergeRange {
  const [startReference, endReference = startReference] = reference.split(":");
  const start = parseCellReference(startReference);
  const end = parseCellReference(endReference);
  return {
    ref: reference,
    startRow: Math.min(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endRow: Math.max(start.row, end.row),
    endColumn: Math.max(start.column, end.column),
  };
}
