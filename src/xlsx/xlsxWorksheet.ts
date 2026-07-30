import type { OoxmlRelationship } from "../ooxml/relationships";
import {
  StreamingXmlStructureValidator,
} from "../ooxml/xmlStructure";
import { parseXml } from "../pptx/xml";
import {
  parseXlsxConditionalFormatting,
  XlsxConditionalFormattingIndex,
  type XlsxConditionalPresentation,
} from "./xlsxConditionalFormatting";
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
  XlsxChart,
  XlsxComment,
  XlsxConditionalFormattingRule,
  XlsxFrozenPane,
  XlsxHyperlink,
  XlsxImage,
  XlsxMergeRange,
  XlsxSheetDescriptor,
  XlsxWorksheetParseDiagnostics,
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
  charts?: readonly XlsxChart[];
  comments?: readonly XlsxComment[];
  date1904: boolean;
  isCancelled?: () => boolean;
}

export interface XlsxWorksheetRichContent {
  images?: readonly XlsxImage[];
  charts?: readonly XlsxChart[];
  comments?: readonly XlsxComment[];
}

export class XlsxWorksheet {
  private readonly cellsByRow = new Map<number, Map<number, XlsxCell>>();
  private readonly commentsByCell = new Map<string, XlsxComment>();
  private readonly conditionalFormattingIndex: XlsxConditionalFormattingIndex;

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
    readonly charts: readonly XlsxChart[],
    readonly comments: readonly XlsxComment[],
    readonly conditionalFormattingRules:
      readonly XlsxConditionalFormattingRule[],
    readonly parseDiagnostics: XlsxWorksheetParseDiagnostics,
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
    for (const comment of comments) {
      this.commentsByCell.set(cellKey(comment.row, comment.column), comment);
    }
    this.conditionalFormattingIndex = new XlsxConditionalFormattingIndex(
      conditionalFormattingRules,
      cells,
    );
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

  getComment(row: number, column: number): XlsxComment | undefined {
    return this.commentsByCell.get(cellKey(row, column));
  }

  getConditionalPresentation(
    row: number,
    column: number,
  ): XlsxConditionalPresentation | null {
    return this.conditionalFormattingIndex.resolve(
      this.getCell(row, column),
    );
  }

  getPopulatedCells(limit = Number.POSITIVE_INFINITY): XlsxCell[] {
    const normalizedLimit =
      Number.isFinite(limit) && limit >= 0
        ? Math.floor(limit)
        : Number.POSITIVE_INFINITY;
    const cells: XlsxCell[] = [];
    const rows = Array.from(this.cellsByRow.keys()).sort(
      (left, right) => left - right,
    );
    for (const rowIndex of rows) {
      const row = this.cellsByRow.get(rowIndex);
      if (!row) {
        continue;
      }
      const columns = Array.from(row.keys()).sort(
        (left, right) => left - right,
      );
      for (const columnIndex of columns) {
        const cell = row.get(columnIndex);
        if (cell) {
          cells.push(cell);
        }
        if (cells.length >= normalizedLimit) {
          return cells;
        }
      }
    }
    return cells;
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
  const parser = new XlsxWorksheetStreamParser(options);
  parser.push(xml);
  return parser.finish();
}

export class XlsxWorksheetStreamParser {
  private readonly validator: StreamingXmlStructureValidator;
  private readonly cells: XlsxCell[] = [];
  private readonly rowHeights = new Map<number, number>();
  private phase: "before" | "rows" | "after" = "before";
  private metadataXml = "";
  private rowBuffer = "";
  private maximumRow = 0;
  private maximumColumn = 0;
  private rowSequence = 0;
  private inputChunks = 0;
  private maximumSheetDataBufferCharacters = 0;
  private inputComplete = false;
  private built = false;

  constructor(private readonly options: XlsxWorksheetParseOptions) {
    this.validator = new StreamingXmlStructureValidator(
      options.descriptor.path,
      "worksheet",
    );
  }

  push(chunk: string): void {
    if (this.inputComplete) {
      throw new Error("XLSX worksheet parser has already finished.");
    }
    checkCancelled(this.options.isCancelled);
    this.inputChunks += 1;
    this.validator.push(chunk);
    this.consume(chunk);
  }

  completeInput(): string {
    if (this.inputComplete) {
      return this.metadataXml;
    }
    this.inputComplete = true;
    this.validator.finish();
    if (this.phase === "rows") {
      throw new Error(
        `Invalid XML in ${this.options.descriptor.path}: unclosed sheetData`,
      );
    }
    if (this.phase === "before") {
      this.metadataXml += this.rowBuffer;
      this.rowBuffer = "";
    }
    return this.metadataXml;
  }

  finish(
    richContent: XlsxWorksheetRichContent = {},
  ): XlsxWorksheet {
    if (this.built) {
      throw new Error("XLSX worksheet parser has already finished.");
    }
    this.built = true;
    this.completeInput();
    checkCancelled(this.options.isCancelled);
    const metadataDocument = parseXml(
      this.metadataXml,
      this.options.descriptor.path,
    );
    const metadataRoot = metadataDocument.documentElement;
    const hyperlinks = parseHyperlinks(
      this.metadataXml,
      this.options.relationships,
    );
    const cellsByReference = new Map(
      this.cells.map((cell) => [cell.ref, cell]),
    );
    for (const hyperlink of hyperlinks) {
      const cell = cellsByReference.get(hyperlink.ref);
      if (cell) {
        cell.hyperlink = hyperlink;
      }
    }
    const dimension = parseDimension(this.metadataXml);
    const images = richContent.images ?? this.options.images ?? [];
    const charts = richContent.charts ?? this.options.charts ?? [];
    const comments = richContent.comments ?? this.options.comments ?? [];
    const contentExtent = calculateRichContentExtent(
      images,
      charts,
      comments,
    );
    const rowCount = Math.max(
      1,
      this.maximumRow + 1,
      dimension?.endRow === undefined ? 0 : dimension.endRow + 1,
      contentExtent.rowCount,
    );
    const columnCount = Math.max(
      1,
      this.maximumColumn + 1,
      dimension?.endColumn === undefined ? 0 : dimension.endColumn + 1,
      contentExtent.columnCount,
    );
    checkCancelled(this.options.isCancelled);
    return new XlsxWorksheet(
      this.options.descriptor,
      rowCount,
      columnCount,
      parseDefaultRowHeight(this.metadataXml),
      parseDefaultColumnWidth(this.metadataXml),
      this.rowHeights,
      parseColumnWidths(this.metadataXml),
      parseMergeRanges(this.metadataXml),
      parseFrozenPane(this.metadataXml),
      hyperlinks,
      images,
      charts,
      comments,
      parseXlsxConditionalFormatting(
        metadataRoot,
        this.options.styles,
      ),
      {
        mode: "streamed",
        inputChunks: this.inputChunks,
        maximumSheetDataBufferCharacters:
          this.maximumSheetDataBufferCharacters,
        metadataCharacters: this.metadataXml.length,
      },
      this.cells,
    );
  }

  private consume(chunk: string): void {
    if (this.phase === "before") {
      this.rowBuffer += chunk;
      const open =
        /<(?:[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>/i.exec(this.rowBuffer);
      if (!open || open.index === undefined) {
        return;
      }
      this.metadataXml +=
        this.rowBuffer.slice(0, open.index) +
        "<sheetData/>";
      const opening = open[0];
      const remainder = this.rowBuffer.slice(open.index + opening.length);
      this.rowBuffer = "";
      if (opening.trimEnd().endsWith("/>")) {
        this.phase = "after";
        this.metadataXml += remainder;
        return;
      }
      this.phase = "rows";
      this.consumeRows(remainder);
      return;
    }
    if (this.phase === "rows") {
      this.consumeRows(chunk);
      return;
    }
    this.metadataXml += chunk;
  }

  private consumeRows(chunk: string): void {
    this.rowBuffer += chunk;
    this.maximumSheetDataBufferCharacters = Math.max(
      this.maximumSheetDataBufferCharacters,
      this.rowBuffer.length,
    );
    const close =
      /<\/(?:[A-Za-z_][\w.-]*:)?sheetData\s*>/i.exec(this.rowBuffer);
    if (close && close.index !== undefined) {
      const source = this.rowBuffer;
      this.parseCompleteRows(source.slice(0, close.index), true);
      this.phase = "after";
      this.metadataXml += source.slice(
        close.index + close[0].length,
      );
      this.rowBuffer = "";
      return;
    }
    this.parseCompleteRows(this.rowBuffer, false);
  }

  private parseCompleteRows(source: string, final: boolean): void {
    const rowPattern =
      /<(?:[A-Za-z_][\w.-]*:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?row\s*>)/gi;
    let consumed = 0;
    for (const match of source.matchAll(rowPattern)) {
      consumed = (match.index ?? 0) + match[0].length;
      this.parseRow(match[1], match[2] ?? "");
    }
    if (final) {
      if (source.slice(consumed).trim().length > 0) {
        throw new Error(
          `Invalid XML in ${this.options.descriptor.path}: malformed sheetData content`,
        );
      }
      this.rowBuffer = "";
      return;
    }
    this.rowBuffer = source.slice(consumed);
  }

  private parseRow(attributes: string, xml: string): void {
    const explicitRow = readIntegerAttribute(attributes, "r");
    const rowIndex =
      explicitRow === null ? this.rowSequence : explicitRow - 1;
    this.rowSequence = rowIndex + 1;
    this.maximumRow = Math.max(this.maximumRow, rowIndex);
    const height = readNumberAttribute(attributes, "ht");
    if (height !== null) {
      this.rowHeights.set(rowIndex, height * (96 / 72));
    }
    let columnSequence = 0;
    const cellPattern =
      /<(?:[A-Za-z_][\w.-]*:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gi;
    for (const cellMatch of xml.matchAll(cellPattern)) {
      const cellAttributes = cellMatch[1];
      const cellXml = cellMatch[2] ?? "";
      const explicitReference = readXmlAttribute(cellAttributes, "r");
      const position = explicitReference
        ? parseCellReference(explicitReference)
        : { row: rowIndex, column: columnSequence };
      columnSequence = position.column + 1;
      this.cells.push(
        parseCell(
          explicitReference ?? "",
          position.row,
          position.column,
          cellAttributes,
          cellXml,
          this.options,
        ),
      );
      this.maximumRow = Math.max(this.maximumRow, position.row);
      this.maximumColumn = Math.max(this.maximumColumn, position.column);
    }
    if (this.rowSequence % CANCELLATION_CHECK_INTERVAL === 0) {
      checkCancelled(this.options.isCancelled);
    }
  }
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

function calculateRichContentExtent(
  images: readonly XlsxImage[],
  charts: readonly XlsxChart[],
  comments: readonly XlsxComment[],
): { rowCount: number; columnCount: number } {
  let rowCount = 1;
  let columnCount = 1;
  for (const comment of comments) {
    rowCount = Math.max(rowCount, comment.row + 1);
    columnCount = Math.max(columnCount, comment.column + 1);
  }
  for (const drawing of [...images, ...charts]) {
    const end = drawing.anchor.to ?? drawing.anchor.from;
    rowCount = Math.max(rowCount, end.row + 1);
    columnCount = Math.max(columnCount, end.column + 1);
  }
  return { rowCount, columnCount };
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

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

export class XlsxWorksheetCancelledError extends Error {
  constructor() {
    super("XLSX worksheet parsing was cancelled.");
    this.name = "XlsxWorksheetCancelledError";
  }
}
