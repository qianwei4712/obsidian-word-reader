import type { OoxmlRelationship } from "../ooxml/relationships";
import { validateXmlStructure } from "../ooxml/xmlStructure";
import {
  makeCellReference,
  parseCellReference,
  parseRangeReference,
} from "./xlsxReferences";
import {
  excelSerialToDate,
  formatXlsxValue,
  XlsxStyleTable,
} from "./xlsxStyles";
import type {
  XlsxCell,
  XlsxCellValue,
  XlsxFrozenPane,
  XlsxHyperlink,
  XlsxImage,
  XlsxMergeRange,
  XlsxSheetDescriptor,
} from "./xlsxTypes";

const DEFAULT_ROW_HEIGHT_PX = 20;
const DEFAULT_COLUMN_WIDTH_PX = 96;
const CANCELLATION_CHECK_INTERVAL = 512;

export interface XlsxWorksheetParseOptions {
  descriptor: XlsxSheetDescriptor;
  styles: XlsxStyleTable;
  sharedStrings: readonly string[];
  relationships: ReadonlyMap<string, OoxmlRelationship>;
  images?: readonly XlsxImage[];
  date1904: boolean;
  isCancelled?: () => boolean;
}

export class XlsxWorksheet {
  private readonly cellsByRow = new Map<number, Map<number, XlsxCell>>();

  constructor(
    readonly descriptor: XlsxSheetDescriptor,
    readonly rowCount: number,
    readonly columnCount: number,
    readonly defaultRowHeight: number,
    readonly defaultColumnWidth: number,
    readonly rowHeights: ReadonlyMap<number, number>,
    readonly columnWidths: ReadonlyMap<number, number>,
    readonly merges: readonly XlsxMergeRange[],
    readonly frozenPane: XlsxFrozenPane | null,
    readonly hyperlinks: readonly XlsxHyperlink[],
    readonly images: readonly XlsxImage[],
    cells: readonly XlsxCell[],
  ) {
    for (const cell of cells) {
      let row = this.cellsByRow.get(cell.row);
      if (!row) {
        row = new Map();
        this.cellsByRow.set(cell.row, row);
      }
      row.set(cell.column, cell);
    }
  }

  get name(): string {
    return this.descriptor.name;
  }

  get state(): XlsxSheetDescriptor["state"] {
    return this.descriptor.state;
  }

  get populatedCellCount(): number {
    let count = 0;
    for (const row of this.cellsByRow.values()) {
      count += row.size;
    }
    return count;
  }

  getCell(row: number, column: number): XlsxCell | undefined {
    return this.cellsByRow.get(row)?.get(column);
  }

  getPopulatedCells(): XlsxCell[] {
    const cells: XlsxCell[] = [];
    for (const row of this.cellsByRow.values()) {
      cells.push(...row.values());
    }
    return cells.sort(
      (left, right) =>
        left.row - right.row || left.column - right.column,
    );
  }

  getCellsInWindow(
    startRow: number,
    endRow: number,
    startColumn: number,
    endColumn: number,
  ): XlsxCell[] {
    const cells: XlsxCell[] = [];
    for (let rowIndex = startRow; rowIndex < endRow; rowIndex += 1) {
      const row = this.cellsByRow.get(rowIndex);
      if (!row) {
        continue;
      }
      for (const [columnIndex, cell] of row) {
        if (columnIndex >= startColumn && columnIndex < endColumn) {
          cells.push(cell);
        }
      }
    }
    return cells;
  }
}

export function parseXlsxWorksheet(
  xml: string,
  options: XlsxWorksheetParseOptions,
): XlsxWorksheet {
  ensureSafeWorksheetXml(xml, options.descriptor.path);
  checkCancelled(options.isCancelled);

  const defaultRowHeight = parseDefaultRowHeight(xml);
  const defaultColumnWidth = parseDefaultColumnWidth(xml);
  const rowHeights = parseRowHeights(xml);
  const columnWidths = parseColumnWidths(xml);
  const cells: XlsxCell[] = [];
  let maximumRow = 0;
  let maximumColumn = 0;
  let rowSequence = 0;

  const rowPattern =
    /<(?:[A-Za-z_][\w.-]*:)?row\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?row\s*>/gi;
  for (const rowMatch of xml.matchAll(rowPattern)) {
    const rowAttributes = rowMatch[1];
    const rowXml = rowMatch[2];
    const explicitRow = readIntegerAttribute(rowAttributes, "r");
    const rowIndex = explicitRow === null ? rowSequence : explicitRow - 1;
    rowSequence = rowIndex + 1;
    let columnSequence = 0;

    const cellPattern =
      /<(?:[A-Za-z_][\w.-]*:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gi;
    for (const cellMatch of rowXml.matchAll(cellPattern)) {
      const cellAttributes = cellMatch[1];
      const cellXml = cellMatch[2] ?? "";
      const explicitReference = readXmlAttribute(cellAttributes, "r");
      const position = explicitReference
        ? parseCellReference(explicitReference)
        : { row: rowIndex, column: columnSequence };
      columnSequence = position.column + 1;
      const cell = parseCell(
        explicitReference ?? "",
        position.row,
        position.column,
        cellAttributes,
        cellXml,
        options,
      );
      cells.push(cell);
      maximumRow = Math.max(maximumRow, position.row);
      maximumColumn = Math.max(maximumColumn, position.column);
    }

    if (rowSequence % CANCELLATION_CHECK_INTERVAL === 0) {
      checkCancelled(options.isCancelled);
    }
  }

  const hyperlinks = parseHyperlinks(xml, options.relationships);
  const cellsByReference = new Map(cells.map((cell) => [cell.ref, cell]));
  for (const hyperlink of hyperlinks) {
    const cell = cellsByReference.get(hyperlink.ref);
    if (cell) {
      cell.hyperlink = hyperlink;
    }
  }

  const dimension = parseDimension(xml);
  const rowCount = Math.max(
    1,
    maximumRow + 1,
    dimension?.endRow === undefined ? 0 : dimension.endRow + 1,
  );
  const columnCount = Math.max(
    1,
    maximumColumn + 1,
    dimension?.endColumn === undefined ? 0 : dimension.endColumn + 1,
  );
  checkCancelled(options.isCancelled);

  return new XlsxWorksheet(
    options.descriptor,
    rowCount,
    columnCount,
    defaultRowHeight,
    defaultColumnWidth,
    rowHeights,
    columnWidths,
    parseMergeRanges(xml),
    parseFrozenPane(xml),
    hyperlinks,
    options.images ?? [],
    cells,
  );
}

function parseCell(
  explicitReference: string,
  row: number,
  column: number,
  attributes: string,
  xml: string,
  options: XlsxWorksheetParseOptions,
): XlsxCell {
  const ref = explicitReference || makeCellReference(row, column);
  const type = readXmlAttribute(attributes, "t") ?? "n";
  const styleId = readIntegerAttribute(attributes, "s") ?? 0;
  const style = options.styles.get(styleId);
  const rawValue = extractElementText(xml, "v");
  const inlineText =
    type === "inlineStr" ? extractAllElementText(xml, "t").join("") : null;
  const value = parseCellValue(
    type,
    rawValue,
    inlineText,
    style.isDate,
    options.sharedStrings,
    options.date1904,
  );
  const formulaText = extractElementText(xml, "f");
  const hasFormula =
    formulaText !== null ||
    /<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*\/>/i.test(xml);
  return {
    ref,
    row,
    column,
    value,
    displayValue: formatXlsxValue(value, style),
    styleId,
    style,
    formula:
      !hasFormula
        ? undefined
        : {
            text: formulaText ?? "",
            cachedValue: value,
            calculation: "cached-only",
            hasExternalReference: /\[[^\]]+\]/.test(formulaText ?? ""),
            requestsRemoteData:
              /\b(?:WEBSERVICE|FILTERXML|RTD|STOCKHISTORY)\s*\(/i.test(
                formulaText ?? "",
              ) || /\[[^\]]+\]/.test(formulaText ?? ""),
          },
  };
}

function parseCellValue(
  type: string,
  rawValue: string | null,
  inlineText: string | null,
  isDate: boolean,
  sharedStrings: readonly string[],
  date1904: boolean,
): XlsxCellValue {
  if (type === "inlineStr") {
    return inlineText ?? "";
  }
  if (rawValue === null) {
    return null;
  }
  switch (type) {
    case "s": {
      const index = Number(rawValue);
      return Number.isInteger(index) ? sharedStrings[index] ?? "" : "";
    }
    case "str":
    case "e":
      return rawValue;
    case "b":
      return rawValue === "1" || rawValue.toLowerCase() === "true";
    case "d": {
      const timestamp = Date.parse(rawValue);
      return Number.isFinite(timestamp) ? new Date(timestamp) : rawValue;
    }
    default: {
      const numericValue = Number(rawValue);
      if (!Number.isFinite(numericValue)) {
        return rawValue;
      }
      return isDate ? excelSerialToDate(numericValue, date1904) : numericValue;
    }
  }
}

function parseDimension(xml: string): XlsxMergeRange | null {
  const match =
    /<(?:[A-Za-z_][\w.-]*:)?dimension\b([^>]*?)\/?>/i.exec(xml);
  const reference = match ? readXmlAttribute(match[1], "ref") : null;
  if (!reference) {
    return null;
  }
  try {
    return parseRangeReference(reference);
  } catch {
    return null;
  }
}

function parseMergeRanges(xml: string): XlsxMergeRange[] {
  const ranges: XlsxMergeRange[] = [];
  const pattern =
    /<(?:[A-Za-z_][\w.-]*:)?mergeCell\b([^>]*?)\/?>/gi;
  for (const match of xml.matchAll(pattern)) {
    const reference = readXmlAttribute(match[1], "ref");
    if (reference) {
      ranges.push(parseRangeReference(reference));
    }
  }
  return ranges;
}

function parseFrozenPane(xml: string): XlsxFrozenPane | null {
  const match = /<(?:[A-Za-z_][\w.-]*:)?pane\b([^>]*?)\/?>/i.exec(xml);
  if (!match || readXmlAttribute(match[1], "state") !== "frozen") {
    return null;
  }
  const rows = readNumberAttribute(match[1], "ySplit") ?? 0;
  const columns = readNumberAttribute(match[1], "xSplit") ?? 0;
  return {
    rows: Math.max(0, Math.floor(rows)),
    columns: Math.max(0, Math.floor(columns)),
    topLeftCell: readXmlAttribute(match[1], "topLeftCell") ?? undefined,
  };
}

function parseHyperlinks(
  xml: string,
  relationships: ReadonlyMap<string, OoxmlRelationship>,
): XlsxHyperlink[] {
  const hyperlinks: XlsxHyperlink[] = [];
  const pattern =
    /<(?:[A-Za-z_][\w.-]*:)?hyperlink\b([^>]*?)\/?>/gi;
  for (const match of xml.matchAll(pattern)) {
    const ref = readXmlAttribute(match[1], "ref");
    if (!ref) {
      continue;
    }
    const relationshipId =
      readXmlAttribute(match[1], "r:id") ??
      readXmlAttribute(match[1], "id");
    const relationship = relationshipId
      ? relationships.get(relationshipId)
      : undefined;
    hyperlinks.push({
      ref,
      target: relationship?.target,
      location: readXmlAttribute(match[1], "location") ?? undefined,
      tooltip: readXmlAttribute(match[1], "tooltip") ?? undefined,
      external: relationship?.external ?? false,
    });
  }
  return hyperlinks;
}

function parseDefaultRowHeight(xml: string): number {
  const match =
    /<(?:[A-Za-z_][\w.-]*:)?sheetFormatPr\b([^>]*?)\/?>/i.exec(xml);
  const points = match
    ? readNumberAttribute(match[1], "defaultRowHeight")
    : null;
  return points === null ? DEFAULT_ROW_HEIGHT_PX : points * (96 / 72);
}

function parseDefaultColumnWidth(xml: string): number {
  const match =
    /<(?:[A-Za-z_][\w.-]*:)?sheetFormatPr\b([^>]*?)\/?>/i.exec(xml);
  const characters = match
    ? readNumberAttribute(match[1], "defaultColWidth")
    : null;
  return characters === null
    ? DEFAULT_COLUMN_WIDTH_PX
    : columnWidthToPixels(characters);
}

function parseRowHeights(xml: string): Map<number, number> {
  const heights = new Map<number, number>();
  const pattern = /<(?:[A-Za-z_][\w.-]*:)?row\b([^>]*)>/gi;
  for (const match of xml.matchAll(pattern)) {
    const row = readIntegerAttribute(match[1], "r");
    const height = readNumberAttribute(match[1], "ht");
    if (row !== null && height !== null) {
      heights.set(row - 1, height * (96 / 72));
    }
  }
  return heights;
}

function parseColumnWidths(xml: string): Map<number, number> {
  const widths = new Map<number, number>();
  const pattern = /<(?:[A-Za-z_][\w.-]*:)?col\b([^>]*?)\/?>/gi;
  for (const match of xml.matchAll(pattern)) {
    const start = readIntegerAttribute(match[1], "min");
    const end = readIntegerAttribute(match[1], "max");
    const width = readNumberAttribute(match[1], "width");
    if (start === null || end === null || width === null) {
      continue;
    }
    const pixelWidth = columnWidthToPixels(width);
    for (let column = start - 1; column < end; column += 1) {
      widths.set(column, pixelWidth);
    }
  }
  return widths;
}

function columnWidthToPixels(characters: number): number {
  return Math.max(1, Math.floor(characters * 7 + 5));
}

function ensureSafeWorksheetXml(xml: string, path: string): void {
  validateXmlStructure(xml, path, "worksheet");
}

function extractElementText(xml: string, name: string): string | null {
  const pattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${name}\\s*>`,
    "i",
  );
  const match = pattern.exec(xml);
  return match ? decodeXmlText(stripXmlTags(match[1])) : null;
}

function extractAllElementText(xml: string, name: string): string[] {
  const values: string[] = [];
  const pattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${name}\\s*>`,
    "gi",
  );
  for (const match of xml.matchAll(pattern)) {
    values.push(decodeXmlText(stripXmlTags(match[1])));
  }
  return values;
}

function stripXmlTags(value: string): string {
  return value.replace(/<[^>]*>/g, "");
}

export function decodeXmlText(value: string): string {
  return value.replace(
    /&(?:#x([0-9a-f]+)|#(\d+)|amp|apos|gt|lt|quot);/gi,
    (entity, hex: string | undefined, decimal: string | undefined) => {
      if (hex) {
        return String.fromCodePoint(Number.parseInt(hex, 16));
      }
      if (decimal) {
        return String.fromCodePoint(Number.parseInt(decimal, 10));
      }
      switch (entity.toLowerCase()) {
        case "&amp;":
          return "&";
        case "&apos;":
          return "'";
        case "&gt;":
          return ">";
        case "&lt;":
          return "<";
        case "&quot;":
          return "\"";
        default:
          return entity;
      }
    },
  );
}

export function readXmlAttribute(
  attributes: string,
  name: string,
): string | null {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(
    `(?:^|\\s)${escapedName}\\s*=\\s*("[^"]*"|'[^']*')`,
    "i",
  );
  const match = pattern.exec(attributes);
  return match
    ? decodeXmlText(match[1].slice(1, -1))
    : null;
}

function readIntegerAttribute(attributes: string, name: string): number | null {
  const value = readNumberAttribute(attributes, name);
  return value !== null && Number.isInteger(value) ? value : null;
}

function readNumberAttribute(attributes: string, name: string): number | null {
  const raw = readXmlAttribute(attributes, name);
  if (raw === null) {
    return null;
  }
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

function checkCancelled(isCancelled: (() => boolean) | undefined): void {
  if (isCancelled?.()) {
    throw new XlsxWorksheetCancelledError();
  }
}

export class XlsxWorksheetCancelledError extends Error {
  constructor() {
    super("XLSX worksheet parsing was cancelled.");
    this.name = "XlsxWorksheetCancelledError";
  }
}
