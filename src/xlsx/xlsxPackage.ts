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
import { validateXmlStructure } from "../ooxml/xmlStructure";
import {
  attribute,
  descendantsNamed,
  firstDescendantNamed,
  parseXml,
  textContent,
} from "../pptx/xml";
import { LruCache } from "../reader/lruCache";
import { parseXlsxDefinedNameTarget } from "./xlsxDefinedNames";
import { XlsxStyleTable } from "./xlsxStyles";
import type {
  XlsxDefinedName,
  XlsxImage,
  XlsxPackageDiagnostics,
  XlsxSheetDescriptor,
  XlsxSheetVisibility,
} from "./xlsxTypes";
import {
  decodeXmlText,
  parseXlsxWorksheet,
  XlsxWorksheet,
  XlsxWorksheetCancelledError,
} from "./xlsxWorksheet";

const WORKBOOK_PATH = "xl/workbook.xml";
const WORKSHEET_CACHE_CAPACITY = 2;
const IMAGE_CACHE_CAPACITY = 8;
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
    if (!safeImageMimeType(path)) {
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
    const [xml, sharedStrings, relationships] = await Promise.all([
      readZipText(this.zip, descriptor.path),
      this.getSharedStrings(),
      readRelationships(this.zip, descriptor.path),
    ]);
    checkCancelled(options.isCancelled);
    this.diagnostics.ignoredExternalRelationships += Array.from(
      relationships.values(),
    ).filter((relationship) => relationship.external).length;
    const images = await this.readWorksheetImages(
      descriptor.path,
      xml,
      relationships,
      options.isCancelled,
    );
    checkCancelled(options.isCancelled);
    return parseXlsxWorksheet(xml, {
      descriptor,
      styles: this.styles,
      sharedStrings,
      relationships,
      images,
      date1904: this.date1904,
      isCancelled: options.isCancelled,
    });
  }

  private getSharedStrings(): Promise<readonly string[]> {
    if (!this.sharedStringsPath) {
      return Promise.resolve([]);
    }
    this.sharedStringsPromise ??= readZipText(
      this.zip,
      this.sharedStringsPath,
    ).then((xml) => parseSharedStrings(xml, this.sharedStringsPath ?? ""));
    return this.sharedStringsPromise;
  }

  private async readWorksheetImages(
    worksheetPath: string,
    worksheetXml: string,
    worksheetRelationships: ReadonlyMap<string, OoxmlRelationship>,
    isCancelled: (() => boolean) | undefined,
  ): Promise<XlsxImage[]> {
    const drawingRelationshipIds = readDrawingRelationshipIds(worksheetXml);
    const images: XlsxImage[] = [];
    for (const relationshipId of drawingRelationshipIds) {
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
      const [drawingXml, drawingRelationships] = await Promise.all([
        drawingEntry.async("string"),
        readRelationships(this.zip, drawingPath),
      ]);
      images.push(
        ...parseDrawingImages(
          drawingPath,
          drawingXml,
          drawingRelationships,
        ),
      );
    }
    return images;
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
  validateXmlStructure(xml, path, "sst");
  const strings: string[] = [];
  const itemPattern =
    /<(?:[A-Za-z_][\w.-]*:)?si\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?si\s*>/gi;
  for (const item of xml.matchAll(itemPattern)) {
    const fragments: string[] = [];
    const textPattern =
      /<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t\s*>/gi;
    for (const text of item[1].matchAll(textPattern)) {
      fragments.push(decodeXmlText(text[1].replace(/<[^>]*>/g, "")));
    }
    strings.push(fragments.join(""));
  }
  return strings;
}

function parseDrawingImages(
  drawingPath: string,
  xml: string,
  relationships: ReadonlyMap<string, OoxmlRelationship>,
): XlsxImage[] {
  const document = parsePackageXml(xml, drawingPath);
  const images: XlsxImage[] = [];
  for (const anchor of [
    ...descendantsNamed(document.documentElement, "twoCellAnchor"),
    ...descendantsNamed(document.documentElement, "oneCellAnchor"),
  ]) {
    const blip = firstDescendantNamed(anchor, "blip");
    const relationshipId = attribute(blip, "embed");
    const relationship = relationshipId
      ? relationships.get(relationshipId)
      : undefined;
    const imagePath = resolveInternalRelationship(drawingPath, relationship);
    const mimeType = imagePath ? safeImageMimeType(imagePath) : null;
    if (!imagePath || !mimeType || !relationship?.type.endsWith("/image")) {
      continue;
    }
    const from = firstDescendantNamed(anchor, "from");
    const column = Number(textContent(firstDescendantNamed(from, "col")));
    const row = Number(textContent(firstDescendantNamed(from, "row")));
    const properties = firstDescendantNamed(anchor, "cNvPr");
    images.push({
      path: imagePath,
      mimeType,
      row: Number.isInteger(row) ? row : 0,
      column: Number.isInteger(column) ? column : 0,
      name: attribute(properties, "name") ?? undefined,
      description: attribute(properties, "descr") ?? undefined,
    });
  }
  return images;
}

function readDrawingRelationshipIds(xml: string): string[] {
  const ids: string[] = [];
  const pattern =
    /<(?:[A-Za-z_][\w.-]*:)?drawing\b([^>]*?)\/?>/gi;
  for (const match of xml.matchAll(pattern)) {
    const id = /\br:id\s*=\s*["']([^"']+)["']/i.exec(match[1])?.[1];
    if (id) {
      ids.push(id);
    }
  }
  return ids;
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
