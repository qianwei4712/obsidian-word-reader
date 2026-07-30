import JSZip from "jszip";

import { enforceOoxmlPackagePolicy } from "../ooxml/packagePolicy";
import {
  DEFAULT_XLSX_ZIP_LIMITS,
  validateZipSafety,
  type ZipSafetyLimits,
  type ZipSafetySummary,
} from "../ooxml/packageSafety";
import { resolvePackagePath } from "../ooxml/packagePath";
import {
  readRelationships,
  resolveInternalRelationship,
  type OoxmlRelationship,
} from "../ooxml/relationships";
import {
  StreamingXmlStructureValidator,
} from "../ooxml/xmlStructure";
import {
  attribute,
  descendantsNamed,
  firstChildNamed,
  firstDescendantNamed,
  parseXml,
  textContent,
} from "../pptx/xml";
import { LruCache } from "../reader/lruCache";
import { parseXlsxChart } from "./xlsxCharts";
import { parseXlsxComments } from "./xlsxComments";
import { parseXlsxDefinedNameTarget } from "./xlsxDefinedNames";
import { XlsxStyleTable } from "./xlsxStyles";
import type {
  XlsxChart,
  XlsxDrawingAnchor,
  XlsxDrawingPosition,
  XlsxDefinedName,
  XlsxImage,
  XlsxPackageDiagnostics,
  XlsxSheetDescriptor,
  XlsxSheetVisibility,
} from "./xlsxTypes";
import {
  decodeXmlText,
  XlsxWorksheet,
  XlsxWorksheetCancelledError,
  XlsxWorksheetStreamParser,
} from "./xlsxWorksheet";

const WORKBOOK_PATH = "xl/workbook.xml";
const WORKSHEET_CACHE_CAPACITY = 2;
const IMAGE_CACHE_CAPACITY = 8;
const MAX_XLSX_IMAGE_BYTES = 16 * 1024 * 1024;
const MAX_XLSX_DRAWING_XML_BYTES = 8 * 1024 * 1024;
const MAX_XLSX_CHART_XML_BYTES = 8 * 1024 * 1024;
const MAX_XLSX_COMMENTS_XML_BYTES = 16 * 1024 * 1024;
const MAX_XLSX_IMAGES_PER_SHEET = 512;
const MAX_XLSX_CHARTS_PER_SHEET = 128;
const SAFE_IMAGE_MIME_TYPES = new Map([
  ["bmp", "image/bmp"],
  ["gif", "image/gif"],
  ["jpeg", "image/jpeg"],
  ["jpg", "image/jpeg"],
  ["png", "image/png"],
]);

export interface XlsxLoadOptions {
  limits?: ZipSafetyLimits;
  isCancelled?: () => boolean;
}

export interface XlsxWorksheetLoadOptions {
  isCancelled?: () => boolean;
  onProgress?: (percent: number) => void;
}

export class XlsxPackageError extends Error {
  constructor(
    readonly kind: "damaged" | "unsupported",
    message: string,
  ) {
    super(message);
    this.name = "XlsxPackageError";
  }
}

export class XlsxPackage {
  readonly diagnostics: XlsxPackageDiagnostics;

  private readonly worksheetCache =
    new LruCache<number, Promise<XlsxWorksheet>>(WORKSHEET_CACHE_CAPACITY);
  private readonly imageCache =
    new LruCache<string, Promise<Uint8Array | null>>(IMAGE_CACHE_CAPACITY);
  private sharedStringsPromise: Promise<readonly string[]> | null = null;

  private constructor(
    private readonly zip: JSZip,
    readonly sheets: readonly XlsxSheetDescriptor[],
    readonly definedNames: readonly XlsxDefinedName[],
    readonly styles: XlsxStyleTable,
    readonly date1904: boolean,
    readonly zipSummary: ZipSafetySummary,
    private readonly sharedStringsPath: string | null,
    ignoredExternalRelationships: number,
    ignoredDataConnections: number,
  ) {
    this.diagnostics = {
      ignoredExternalRelationships,
      ignoredDataConnections,
      formulaCalculation: "cached-only",
    };
  }

  static async load(
    buffer: ArrayBuffer,
    options: XlsxLoadOptions = {},
  ): Promise<XlsxPackage> {
    const zipSummary = validateZipSafety(
      buffer,
      options.limits ?? DEFAULT_XLSX_ZIP_LIMITS,
    );
    checkCancelled(options.isCancelled);

    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(buffer, {
        createFolders: false,
        checkCRC32: false,
      });
    } catch (error) {
      throw new XlsxPackageError(
        "damaged",
        `The XLSX ZIP package could not be opened: ${getErrorMessage(error)}`,
      );
    }
    checkCancelled(options.isCancelled);

    ensureRequiredEntry(zip, "[Content_Types].xml");
    ensureRequiredEntry(zip, WORKBOOK_PATH);
    await enforceOoxmlPackagePolicy(zip, "xlsx");
    checkCancelled(options.isCancelled);

    const workbook = parsePackageXml(
      await readZipText(zip, WORKBOOK_PATH),
      WORKBOOK_PATH,
    );
    const relationships = await readRelationships(zip, WORKBOOK_PATH);
    const sheets: XlsxSheetDescriptor[] = [];
    let ignoredExternalRelationships = 0;
    for (const element of descendantsNamed(
      workbook.documentElement,
      "sheet",
    )) {
      const name = attribute(element, "name");
      const relationshipId = attribute(element, "id");
      const relationship = relationshipId
        ? relationships.get(relationshipId)
        : undefined;
      if (!name || !relationshipId || !relationship) {
        continue;
      }
      if (relationship.external) {
        continue;
      }
      if (!relationship.type.endsWith("/worksheet")) {
        continue;
      }
      const path = resolvePackagePath(WORKBOOK_PATH, relationship.target);
      ensureRequiredEntry(zip, path);
      sheets.push({
        name,
        state: normalizeSheetState(attribute(element, "state")),
        path,
        relationshipId,
      });
    }
    if (sheets.length === 0) {
      throw new XlsxPackageError(
        "unsupported",
        "The workbook does not contain any readable worksheets.",
      );
    }

    const stylesPath = findInternalPartPath(
      WORKBOOK_PATH,
      relationships,
      "/styles",
    );
    const styles = stylesPath
      ? XlsxStyleTable.parse(await readZipText(zip, stylesPath), stylesPath)
      : XlsxStyleTable.empty();
    const sharedStringsPath = findInternalPartPath(
      WORKBOOK_PATH,
      relationships,
      "/sharedStrings",
    );
    const workbookProperties = firstDescendantNamed(
      workbook.documentElement,
      "workbookPr",
    );
    const date1904 = readBooleanAttribute(workbookProperties, "date1904");
    const ignoredDataConnections = Object.keys(zip.files).filter((name) => {
      const lowerName = name.toLowerCase();
      return (
        lowerName === "xl/connections.xml" ||
        lowerName.startsWith("xl/externallinks/")
      );
    }).length;
    ignoredExternalRelationships += Array.from(relationships.values()).filter(
      (relationship) => relationship.external,
    ).length;
    const definedNames = parseWorkbookDefinedNames(
      workbook.documentElement,
      sheets,
    );
    checkCancelled(options.isCancelled);

    return new XlsxPackage(
      zip,
      sheets,
      definedNames,
      styles,
      date1904,
      zipSummary,
      sharedStringsPath,
      ignoredExternalRelationships,
      ignoredDataConnections,
    );
  }

  get sheetCount(): number {
    return this.sheets.length;
  }

  async getWorksheet(
    index: number,
    options: XlsxWorksheetLoadOptions = {},
  ): Promise<XlsxWorksheet> {
    const descriptor = this.sheets[index];
    if (!descriptor) {
      throw new RangeError(`Worksheet index ${index} is outside the workbook.`);
    }
    const existing = this.worksheetCache.get(index);
    if (existing) {
      return existing;
    }

    const pending = this.readWorksheet(descriptor, options);
    void this.worksheetCache.set(index, pending);
    try {
      return await pending;
    } catch (error) {
      this.worksheetCache.delete(index);
      throw error;
    }
  }

  async getImageBinary(path: string): Promise<Uint8Array | null> {
    if (
      !safeImageMimeType(path) ||
      !this.isPartWithinLimit(path, MAX_XLSX_IMAGE_BYTES)
    ) {
      return null;
    }
    return this.imageCache.getOrCreate(path, () => {
      const entry = this.zip.file(path);
      return entry ? entry.async("uint8array") : Promise.resolve(null);
    });
  }

  getCacheDiagnostics(): {
    worksheets: number;
    images: number;
    worksheetLimit: number;
    imageLimit: number;
  } {
    return {
      worksheets: this.worksheetCache.size,
      images: this.imageCache.size,
      worksheetLimit: this.worksheetCache.capacity,
      imageLimit: this.imageCache.capacity,
    };
  }

  clearCaches(): void {
    this.worksheetCache.clear();
    this.imageCache.clear();
    this.sharedStringsPromise = null;
  }

  private async readWorksheet(
    descriptor: XlsxSheetDescriptor,
    options: XlsxWorksheetLoadOptions,
  ): Promise<XlsxWorksheet> {
    checkCancelled(options.isCancelled);
    const [sharedStrings, relationships] = await Promise.all([
      this.getSharedStrings(options.isCancelled),
      readRelationships(this.zip, descriptor.path),
    ]);
    checkCancelled(options.isCancelled);
    this.diagnostics.ignoredExternalRelationships += Array.from(
      relationships.values(),
    ).filter((relationship) => relationship.external).length;
    const parser = new XlsxWorksheetStreamParser({
      descriptor,
      styles: this.styles,
      sharedStrings,
      relationships,
      date1904: this.date1904,
      isCancelled: options.isCancelled,
    });
    const worksheetEntry = this.zip.file(descriptor.path);
    if (!worksheetEntry) {
      throw new XlsxPackageError(
        "damaged",
        `The XLSX package is missing ${descriptor.path}.`,
      );
    }
    await streamZipText(
      worksheetEntry,
      (chunk) => parser.push(chunk),
      options.isCancelled,
      options.onProgress,
    );
    const metadataXml = parser.completeInput();
    const drawings = await this.readWorksheetDrawings(
      descriptor.path,
      metadataXml,
      relationships,
      options.isCancelled,
    );
    const comments = await this.readWorksheetComments(
      descriptor.path,
      relationships,
      options.isCancelled,
    );
    checkCancelled(options.isCancelled);
    return parser.finish({
      images: drawings.images,
      charts: drawings.charts,
      comments,
    });
  }

  private getSharedStrings(
    isCancelled?: () => boolean,
  ): Promise<readonly string[]> {
    if (!this.sharedStringsPath) {
      return Promise.resolve([]);
    }
    if (this.sharedStringsPromise) {
      return this.sharedStringsPromise;
    }
    const path = this.sharedStringsPath;
    const entry = this.zip.file(path);
    if (!entry) {
      return Promise.reject(
        new XlsxPackageError(
          "damaged",
          `The XLSX package is missing ${path}.`,
        ),
      );
    }
    const parser = new XlsxSharedStringsStreamParser(path);
    const pending = streamZipText(
      entry,
      (chunk) => parser.push(chunk),
      isCancelled,
      undefined,
    ).then(() => parser.finish());
    this.sharedStringsPromise = pending;
    void pending.catch(() => {
      if (this.sharedStringsPromise === pending) {
        this.sharedStringsPromise = null;
      }
    });
    return this.sharedStringsPromise;
  }

  private async readWorksheetDrawings(
    worksheetPath: string,
    worksheetXml: string,
    worksheetRelationships: ReadonlyMap<string, OoxmlRelationship>,
    isCancelled: (() => boolean) | undefined,
  ): Promise<{ images: XlsxImage[]; charts: XlsxChart[] }> {
    const drawingRelationshipIds = readDrawingRelationshipIds(worksheetXml);
    const images: XlsxImage[] = [];
    const charts: XlsxChart[] = [];
    for (const relationshipId of drawingRelationshipIds) {
      if (
        images.length >= MAX_XLSX_IMAGES_PER_SHEET &&
        charts.length >= MAX_XLSX_CHARTS_PER_SHEET
      ) {
        break;
      }
      checkCancelled(isCancelled);
      const drawingRelationship = worksheetRelationships.get(relationshipId);
      const drawingPath = resolveInternalRelationship(
        worksheetPath,
        drawingRelationship,
      );
      if (!drawingPath || !drawingRelationship?.type.endsWith("/drawing")) {
        continue;
      }
      const drawingEntry = this.zip.file(drawingPath);
      if (!drawingEntry) {
        continue;
      }
      if (
        !this.isPartWithinLimit(
          drawingPath,
          MAX_XLSX_DRAWING_XML_BYTES,
        )
      ) {
        continue;
      }
      const [drawingXml, drawingRelationships] = await Promise.all([
        drawingEntry.async("string"),
        readRelationships(this.zip, drawingPath),
      ]);
      const drawingObjects = await parseDrawingObjects(
        this.zip,
        drawingPath,
        drawingXml,
        drawingRelationships,
        isCancelled,
        (path, maximumBytes) =>
          this.isPartWithinLimit(path, maximumBytes),
      );
      images.push(
        ...drawingObjects.images.slice(
          0,
          MAX_XLSX_IMAGES_PER_SHEET - images.length,
        ),
      );
      charts.push(
        ...drawingObjects.charts.slice(
          0,
          MAX_XLSX_CHARTS_PER_SHEET - charts.length,
        ),
      );
    }
    return { images, charts };
  }

  private async readWorksheetComments(
    worksheetPath: string,
    relationships: ReadonlyMap<string, OoxmlRelationship>,
    isCancelled: (() => boolean) | undefined,
  ) {
    const commentsPath = findInternalPartPath(
      worksheetPath,
      relationships,
      "/comments",
    );
    if (
      !commentsPath ||
      !this.zip.file(commentsPath) ||
      !this.isPartWithinLimit(
        commentsPath,
        MAX_XLSX_COMMENTS_XML_BYTES,
      )
    ) {
      return [];
    }
    checkCancelled(isCancelled);
    const comments = parseXlsxComments(
      await readZipText(this.zip, commentsPath),
      commentsPath,
    );
    checkCancelled(isCancelled);
    return comments;
  }

  private isPartWithinLimit(path: string, maximumBytes: number): boolean {
    const entry = this.zipSummary.entries.find(
      (candidate) => candidate.name === path,
    );
    return Boolean(
      entry &&
      entry.uncompressedBytes >= 0 &&
      entry.uncompressedBytes <= maximumBytes,
    );
  }
}

function parseWorkbookDefinedNames(
  workbook: Element,
  sheets: readonly XlsxSheetDescriptor[],
): XlsxDefinedName[] {
  const names: XlsxDefinedName[] = [];
  for (const element of descendantsNamed(workbook, "definedName")) {
    const name = attribute(element, "name")?.trim();
    if (
      !name ||
      name.toLocaleLowerCase().startsWith("_xlnm.") ||
      readBooleanAttribute(element, "hidden")
    ) {
      continue;
    }
    const target = textContent(element).trim();
    const localSheetIdValue = attribute(element, "localSheetId");
    const localSheetId = parseOptionalSheetIndex(
      localSheetIdValue,
      sheets.length,
    );
    if (localSheetIdValue !== null && localSheetId === undefined) {
      continue;
    }
    const parsed = parseXlsxDefinedNameTarget(
      target,
      sheets,
      localSheetId,
    );
    if (!parsed) {
      continue;
    }
    names.push({
      name,
      target,
      sheetIndex: parsed.sheetIndex,
      scopeSheetIndex: localSheetId,
      range: parsed.range,
    });
  }
  return names;
}

function parseOptionalSheetIndex(
  value: string | null,
  sheetCount: number,
): number | undefined {
  if (value === null) {
    return undefined;
  }
  const index = Number(value);
  return Number.isInteger(index) && index >= 0 && index < sheetCount
    ? index
    : undefined;
}

export function parseSharedStrings(xml: string, path: string): string[] {
  const parser = new XlsxSharedStringsStreamParser(path);
  parser.push(xml);
  return parser.finish();
}

class XlsxSharedStringsStreamParser {
  private readonly validator: StreamingXmlStructureValidator;
  private readonly strings: string[] = [];
  private buffer = "";
  private finished = false;

  constructor(private readonly path: string) {
    this.validator = new StreamingXmlStructureValidator(path, "sst");
  }

  push(chunk: string): void {
    if (this.finished) {
      throw new Error("XLSX shared-string parser has already finished.");
    }
    this.validator.push(chunk);
    this.buffer += chunk;
    const itemPattern =
      /<(?:[A-Za-z_][\w.-]*:)?si\b[^>]*?\/>|<(?:[A-Za-z_][\w.-]*:)?si\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?si\s*>/gi;
    let consumed = 0;
    for (const item of this.buffer.matchAll(itemPattern)) {
      consumed = (item.index ?? 0) + item[0].length;
      this.strings.push(parseSharedStringItem(item[1] ?? ""));
    }
    if (consumed > 0) {
      this.buffer = this.buffer.slice(consumed);
    }
  }

  finish(): string[] {
    if (this.finished) {
      throw new Error("XLSX shared-string parser has already finished.");
    }
    this.finished = true;
    this.validator.finish();
    return this.strings;
  }
}

function parseSharedStringItem(xml: string): string {
  const fragments: string[] = [];
  const textPattern =
    /<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t\s*>/gi;
  for (const text of xml.matchAll(textPattern)) {
    fragments.push(decodeXmlText(text[1].replace(/<[^>]*>/g, "")));
  }
  return fragments.join("");
}

async function parseDrawingObjects(
  zip: JSZip,
  drawingPath: string,
  xml: string,
  relationships: ReadonlyMap<string, OoxmlRelationship>,
  isCancelled: (() => boolean) | undefined,
  isPartWithinLimit: (path: string, maximumBytes: number) => boolean,
): Promise<{ images: XlsxImage[]; charts: XlsxChart[] }> {
  const document = parsePackageXml(xml, drawingPath);
  const images: XlsxImage[] = [];
  const charts: XlsxChart[] = [];
  for (const anchor of [
    ...descendantsNamed(document.documentElement, "twoCellAnchor"),
    ...descendantsNamed(document.documentElement, "oneCellAnchor"),
  ]) {
    checkCancelled(isCancelled);
    const drawingAnchor = parseDrawingAnchor(anchor);
    if (!drawingAnchor) {
      continue;
    }
    const blip = firstDescendantNamed(anchor, "blip");
    const relationshipId = attribute(blip, "embed");
    const relationship = relationshipId
      ? relationships.get(relationshipId)
      : undefined;
    const imagePath = resolveInternalRelationship(drawingPath, relationship);
    const mimeType = imagePath ? safeImageMimeType(imagePath) : null;
    if (
      images.length < MAX_XLSX_IMAGES_PER_SHEET &&
      imagePath &&
      mimeType &&
      relationship?.type.endsWith("/image") &&
      zip.file(imagePath) &&
      isPartWithinLimit(imagePath, MAX_XLSX_IMAGE_BYTES)
    ) {
      const properties = firstDescendantNamed(anchor, "cNvPr");
      images.push({
        path: imagePath,
        mimeType,
        row: drawingAnchor.from.row,
        column: drawingAnchor.from.column,
        anchor: drawingAnchor,
        name: attribute(properties, "name") ?? undefined,
        description: attribute(properties, "descr") ?? undefined,
      });
    }
    const chartElement = firstDescendantNamed(anchor, "chart");
    const chartRelationshipId = attribute(chartElement, "id");
    const chartRelationship = chartRelationshipId
      ? relationships.get(chartRelationshipId)
      : undefined;
    const chartPath = resolveInternalRelationship(
      drawingPath,
      chartRelationship,
    );
    if (
      charts.length < MAX_XLSX_CHARTS_PER_SHEET &&
      chartPath &&
      chartRelationship?.type.endsWith("/chart") &&
      zip.file(chartPath) &&
      isPartWithinLimit(chartPath, MAX_XLSX_CHART_XML_BYTES)
    ) {
      charts.push(
        parseXlsxChart(
          await readZipText(zip, chartPath),
          chartPath,
          drawingAnchor,
        ),
      );
    }
  }
  return { images, charts };
}

function parseDrawingAnchor(element: Element): XlsxDrawingAnchor | null {
  const from = parseDrawingPosition(firstChildNamed(element, "from"));
  if (!from) {
    return null;
  }
  const to = parseDrawingPosition(firstChildNamed(element, "to"));
  const extent = firstChildNamed(element, "ext");
  const width = Number(attribute(extent, "cx"));
  const height = Number(attribute(extent, "cy"));
  return {
    from,
    to: to ?? undefined,
    widthPx:
      Number.isFinite(width) && width > 0
        ? emuToPixels(width)
        : undefined,
    heightPx:
      Number.isFinite(height) && height > 0
        ? emuToPixels(height)
        : undefined,
  };
}

function parseDrawingPosition(
  element: Element | null,
): XlsxDrawingPosition | null {
  if (!element) {
    return null;
  }
  const row = Number(textContent(firstChildNamed(element, "row")));
  const column = Number(textContent(firstChildNamed(element, "col")));
  const rowOffset = Number(
    textContent(firstChildNamed(element, "rowOff")),
  );
  const columnOffset = Number(
    textContent(firstChildNamed(element, "colOff")),
  );
  if (
    !Number.isInteger(row) ||
    row < 0 ||
    !Number.isInteger(column) ||
    column < 0
  ) {
    return null;
  }
  return {
    row,
    column,
    rowOffsetPx:
      Number.isFinite(rowOffset) && rowOffset > 0
        ? emuToPixels(rowOffset)
        : 0,
    columnOffsetPx:
      Number.isFinite(columnOffset) && columnOffset > 0
        ? emuToPixels(columnOffset)
        : 0,
  };
}

function emuToPixels(value: number): number {
  return value / 9_525;
}

function readDrawingRelationshipIds(xml: string): string[] {
  const ids = new Set<string>();
  const pattern =
    /<(?:[A-Za-z_][\w.-]*:)?drawing\b([^>]*?)\/?>/gi;
  for (const match of xml.matchAll(pattern)) {
    const id = /\br:id\s*=\s*["']([^"']+)["']/i.exec(match[1])?.[1];
    if (id) {
      ids.add(id);
    }
  }
  return Array.from(ids);
}

function safeImageMimeType(path: string): string | null {
  const extension = path.toLowerCase().split(".").at(-1);
  return extension ? SAFE_IMAGE_MIME_TYPES.get(extension) ?? null : null;
}

function findInternalPartPath(
  sourcePath: string,
  relationships: ReadonlyMap<string, OoxmlRelationship>,
  typeSuffix: string,
): string | null {
  const relationship = Array.from(relationships.values()).find(
    (candidate) =>
      !candidate.external && candidate.type.endsWith(typeSuffix),
  );
  return resolveInternalRelationship(sourcePath, relationship);
}

function normalizeSheetState(value: string | null): XlsxSheetVisibility {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}

function readBooleanAttribute(
  element: Element | null,
  name: string,
): boolean {
  const value = attribute(element, name);
  return value === "1" || value === "true";
}

function ensureRequiredEntry(zip: JSZip, path: string): void {
  if (!zip.file(path)) {
    throw new XlsxPackageError(
      "damaged",
      `The XLSX package is missing ${path}.`,
    );
  }
}

async function readZipText(zip: JSZip, path: string): Promise<string> {
  const entry = zip.file(path);
  if (!entry) {
    throw new XlsxPackageError(
      "damaged",
      `The XLSX package is missing ${path}.`,
    );
  }
  return entry.async("string");
}

async function streamZipText(
  entry: JSZip.JSZipObject,
  onChunk: (chunk: string) => void,
  isCancelled: (() => boolean) | undefined,
  onProgress: ((percent: number) => void) | undefined,
): Promise<void> {
  const streamable = entry as JSZip.JSZipObject & {
    internalStream(
      type: "string",
    ): JSZip.JSZipStreamHelper<string>;
  };
  if (typeof streamable.internalStream !== "function") {
    onChunk(await entry.async("string"));
    onProgress?.(100);
    return;
  }
  await new Promise<void>((resolve, reject) => {
    const stream = streamable.internalStream("string");
    let settled = false;
    const rejectOnce = (error: unknown): void => {
      if (settled) {
        return;
      }
      settled = true;
      stream.pause();
      reject(error instanceof Error ? error : new Error(String(error)));
    };
    stream.on("data", (chunk, metadata) => {
      if (settled) {
        return;
      }
      try {
        checkCancelled(isCancelled);
        onChunk(chunk);
        onProgress?.(metadata.percent);
      } catch (error) {
        rejectOnce(error);
      }
    });
    stream.on("error", rejectOnce);
    stream.on("end", () => {
      if (!settled) {
        settled = true;
        onProgress?.(100);
        resolve();
      }
    });
    stream.resume();
  });
}

function parsePackageXml(xml: string, path: string): Document {
  try {
    return parseXml(xml, path);
  } catch (error) {
    throw new XlsxPackageError(
      "damaged",
      `The XLSX package contains invalid XML in ${path}: ${getErrorMessage(error)}`,
    );
  }
}

function checkCancelled(isCancelled: (() => boolean) | undefined): void {
  if (isCancelled?.()) {
    throw new XlsxWorksheetCancelledError();
  }
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
