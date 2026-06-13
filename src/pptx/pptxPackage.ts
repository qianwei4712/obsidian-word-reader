import JSZip from "jszip";

import {
  attribute,
  descendantsNamed,
  namespacedAttribute,
  numberAttribute,
  parseXml,
} from "./xml";
import {
  DEFAULT_PPTX_ZIP_LIMITS,
  validateZipSafety,
  type ZipSafetyLimits,
  type ZipSafetySummary,
} from "./zipLimits";
import {
  extractSlideMetadata,
  type PptxSlideMetadata,
} from "./pptxMetadata";

const DEFAULT_SLIDE_WIDTH = 12_192_000;
const DEFAULT_SLIDE_HEIGHT = 6_858_000;
const OFFICE_RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export interface PptxRelationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export interface PptxSlideContext {
  slidePath: string;
  slide: Document;
  slideRelationships: Map<string, PptxRelationship>;
  layoutPath: string | null;
  layout: Document | null;
  layoutRelationships: Map<string, PptxRelationship>;
  masterPath: string | null;
  master: Document | null;
  masterRelationships: Map<string, PptxRelationship>;
  themePath: string | null;
  theme: Document | null;
}

export class PptxPackageError extends Error {
  constructor(
    readonly kind: "damaged" | "unsupported",
    message: string,
  ) {
    super(message);
    this.name = "PptxPackageError";
  }
}

export class PptxPackage {
  readonly slideWidth: number;
  readonly slideHeight: number;
  readonly slidePaths: string[];
  readonly zipSummary: ZipSafetySummary;

  private readonly xmlCache = new Map<string, Promise<Document>>();
  private readonly relationshipCache =
    new Map<string, Promise<Map<string, PptxRelationship>>>();
  private readonly metadataCache =
    new Map<number, Promise<PptxSlideMetadata>>();

  private constructor(
    private readonly zip: JSZip,
    presentation: Document,
    presentationRelationships: Map<string, PptxRelationship>,
    zipSummary: ZipSafetySummary,
  ) {
    const slideSize = descendantsNamed(
      presentation.documentElement,
      "sldSz",
    )[0];
    this.slideWidth = numberAttribute(
      slideSize,
      "cx",
      DEFAULT_SLIDE_WIDTH,
    );
    this.slideHeight = numberAttribute(
      slideSize,
      "cy",
      DEFAULT_SLIDE_HEIGHT,
    );
    this.slidePaths = descendantsNamed(
      presentation.documentElement,
      "sldId",
    )
      .map((slideId) =>
        namespacedAttribute(
          slideId,
          "id",
          OFFICE_RELATIONSHIP_NAMESPACE,
        ),
      )
      .map((relationshipId) =>
        relationshipId
          ? presentationRelationships.get(relationshipId)
          : undefined,
      )
      .filter(
        (relationship): relationship is PptxRelationship =>
          relationship !== undefined && !relationship.external,
      )
      .map((relationship) =>
        resolvePackagePath("ppt/presentation.xml", relationship.target),
      );
    this.zipSummary = zipSummary;
  }

  static async load(
    buffer: ArrayBuffer,
    limits: ZipSafetyLimits = DEFAULT_PPTX_ZIP_LIMITS,
  ): Promise<PptxPackage> {
    const zipSummary = validateZipSafety(buffer, limits);
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(buffer, {
        createFolders: false,
        checkCRC32: false,
      });
    } catch (error) {
      throw new PptxPackageError(
        "damaged",
        `The PowerPoint ZIP package could not be opened: ${getErrorMessage(error)}`,
      );
    }

    ensureRequiredEntry(zip, "[Content_Types].xml");
    ensureRequiredEntry(zip, "ppt/presentation.xml");
    if (zip.file("ppt/vbaProject.bin")) {
      throw new PptxPackageError(
        "unsupported",
        "Macro-enabled PowerPoint packages are not supported.",
      );
    }
    const presentationXml = await readZipText(zip, "ppt/presentation.xml");
    const presentation = parsePackageXml(
      presentationXml,
      "ppt/presentation.xml",
    );
    const relationships = await readRelationships(
      zip,
      "ppt/presentation.xml",
    );
    const pptx = new PptxPackage(
      zip,
      presentation,
      relationships,
      zipSummary,
    );
    if (pptx.slidePaths.length === 0) {
      throw new PptxPackageError(
        "unsupported",
        "The presentation does not contain any readable slides.",
      );
    }
    for (const slidePath of pptx.slidePaths) {
      ensureRequiredEntry(zip, slidePath);
    }
    return pptx;
  }

  get slideCount(): number {
    return this.slidePaths.length;
  }

  async getSlideContext(index: number): Promise<PptxSlideContext> {
    const slidePath = this.slidePaths[index];
    if (!slidePath) {
      throw new RangeError(`Slide index ${index} is outside the presentation.`);
    }

    const slide = await this.getXml(slidePath);
    const slideRelationships = await this.getRelationships(slidePath);
    const layoutRelationship = findRelationship(
      slideRelationships,
      "/slideLayout",
    );
    const layoutPath =
      layoutRelationship && !layoutRelationship.external
        ? resolvePackagePath(slidePath, layoutRelationship.target)
        : null;
    const layout = layoutPath ? await this.getOptionalXml(layoutPath) : null;
    const layoutRelationships = layoutPath
      ? await this.getRelationships(layoutPath)
      : new Map<string, PptxRelationship>();

    const masterRelationship = findRelationship(
      layoutRelationships,
      "/slideMaster",
    );
    const masterPath =
      layoutPath && masterRelationship && !masterRelationship.external
        ? resolvePackagePath(layoutPath, masterRelationship.target)
        : null;
    const master = masterPath ? await this.getOptionalXml(masterPath) : null;
    const masterRelationships = masterPath
      ? await this.getRelationships(masterPath)
      : new Map<string, PptxRelationship>();

    const themeRelationship = findRelationship(
      masterRelationships,
      "/theme",
    );
    const themePath =
      masterPath && themeRelationship && !themeRelationship.external
        ? resolvePackagePath(masterPath, themeRelationship.target)
        : null;
    const theme = themePath ? await this.getOptionalXml(themePath) : null;

    return {
      slidePath,
      slide,
      slideRelationships,
      layoutPath,
      layout,
      layoutRelationships,
      masterPath,
      master,
      masterRelationships,
      themePath,
      theme,
    };
  }

  getSlideMetadata(index: number): Promise<PptxSlideMetadata> {
    let pending = this.metadataCache.get(index);
    if (!pending) {
      pending = this.readSlideMetadata(index);
      this.metadataCache.set(index, pending);
    }
    return pending;
  }

  async getAllSlideMetadata(): Promise<PptxSlideMetadata[]> {
    return Promise.all(
      this.slidePaths.map((_, index) => this.getSlideMetadata(index)),
    );
  }

  async getBinary(path: string): Promise<Uint8Array | null> {
    const entry = this.zip.file(path);
    return entry ? entry.async("uint8array") : null;
  }

  getImageMimeType(path: string): string | null {
    const extension = path.split(".").at(-1)?.toLowerCase();
    switch (extension) {
      case "jpg":
      case "jpeg":
        return "image/jpeg";
      case "bmp":
        return "image/bmp";
      case "png":
        return "image/png";
      default:
        return null;
    }
  }

  private getXml(path: string): Promise<Document> {
    let pending = this.xmlCache.get(path);
    if (!pending) {
      pending = readZipText(this.zip, path).then((xml) =>
        parsePackageXml(xml, path),
      );
      this.xmlCache.set(path, pending);
    }
    return pending;
  }

  private async getOptionalXml(path: string): Promise<Document | null> {
    return this.zip.file(path) ? this.getXml(path) : null;
  }

  private getRelationships(
    path: string,
  ): Promise<Map<string, PptxRelationship>> {
    let pending = this.relationshipCache.get(path);
    if (!pending) {
      pending = readRelationships(this.zip, path);
      this.relationshipCache.set(path, pending);
    }
    return pending;
  }

  private async readSlideMetadata(index: number): Promise<PptxSlideMetadata> {
    const slidePath = this.slidePaths[index];
    if (!slidePath) {
      throw new RangeError(`Slide index ${index} is outside the presentation.`);
    }
    const slide = await this.getXml(slidePath);
    const relationships = await this.getRelationships(slidePath);
    const notesRelationship = findRelationship(relationships, "/notesSlide");
    const notesPath =
      notesRelationship && !notesRelationship.external
        ? resolvePackagePath(slidePath, notesRelationship.target)
        : null;
    const notes = notesPath ? await this.getOptionalXml(notesPath) : null;
    return extractSlideMetadata(
      index,
      slide,
      notes,
      "",
    );
  }
}

export function resolvePackagePath(basePath: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizePackagePath(target.slice(1));
  }
  const baseParts = basePath.split("/");
  baseParts.pop();
  return normalizePackagePath([...baseParts, ...target.split("/")].join("/"));
}

export function relationshipPartPath(path: string): string {
  const parts = path.split("/");
  const fileName = parts.pop();
  return [...parts, "_rels", `${fileName}.rels`].join("/");
}

function normalizePackagePath(path: string): string {
  const normalized: string[] = [];
  for (const part of path.replace(/\\/g, "/").split("/")) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      if (normalized.length === 0) {
        throw new PptxPackageError(
          "damaged",
          "A package relationship escapes the archive root.",
        );
      }
      normalized.pop();
      continue;
    }
    normalized.push(part);
  }
  return normalized.join("/");
}

async function readRelationships(
  zip: JSZip,
  sourcePath: string,
): Promise<Map<string, PptxRelationship>> {
  const path = relationshipPartPath(sourcePath);
  const entry = zip.file(path);
  if (!entry) {
    return new Map();
  }
  const document = parsePackageXml(await entry.async("string"), path);
  const relationships = new Map<string, PptxRelationship>();
  for (const element of descendantsNamed(
    document.documentElement,
    "Relationship",
  )) {
    const id = attribute(element, "Id");
    const type = attribute(element, "Type");
    const target = attribute(element, "Target");
    if (!id || !type || !target) {
      continue;
    }
    relationships.set(id, {
      id,
      type,
      target,
      external: attribute(element, "TargetMode") === "External",
    });
  }
  return relationships;
}

function findRelationship(
  relationships: Map<string, PptxRelationship>,
  typeSuffix: string,
): PptxRelationship | undefined {
  return Array.from(relationships.values()).find((relationship) =>
    relationship.type.endsWith(typeSuffix),
  );
}

function ensureRequiredEntry(zip: JSZip, path: string): void {
  if (!zip.file(path)) {
    throw new PptxPackageError(
      "damaged",
      `The PowerPoint package is missing ${path}.`,
    );
  }
}

async function readZipText(zip: JSZip, path: string): Promise<string> {
  const entry = zip.file(path);
  if (!entry) {
    throw new PptxPackageError(
      "damaged",
      `The PowerPoint package is missing ${path}.`,
    );
  }
  return entry.async("string");
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function parsePackageXml(xml: string, path: string): Document {
  try {
    return parseXml(xml, path);
  } catch (error) {
    throw new PptxPackageError(
      "damaged",
      `The PowerPoint package contains invalid XML in ${path}: ${getErrorMessage(error)}`,
    );
  }
}
