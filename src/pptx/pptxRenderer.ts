import {
  type PptxPackage,
  type PptxRelationship,
  type PptxSlideContext,
  resolvePackagePath,
} from "./pptxPackage";
import {
  attribute,
  childElements,
  childrenNamed,
  descendantsNamed,
  firstChildNamed,
  firstDescendantNamed,
  localName,
  namespacedAttribute,
  numberAttribute,
  textContent,
} from "./xml";

const NATURAL_SLIDE_WIDTH = 960;
const SVG_NAMESPACE = "http://www.w3.org/2000/svg";
const OFFICE_RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export interface RenderedPptxSlide {
  element: HTMLElement;
  resources: Set<string>;
  width: number;
  height: number;
}

interface ThemeColors {
  [key: string]: string;
}

interface CoordinateSpace {
  offsetX: number;
  offsetY: number;
  scaleX: number;
  scaleY: number;
}

interface ShapeTransform {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipHorizontal: boolean;
  flipVertical: boolean;
}

interface LayerContext {
  xmlDocument: Document;
  path: string;
  relationships: Map<string, PptxRelationship>;
}

interface RenderContext {
  ownerDocument: Document;
  pptx: PptxPackage;
  theme: ThemeColors;
  resources: Set<string>;
  slideWidth: number;
  slideHeight: number;
  pixelWidth: number;
  pixelHeight: number;
  placeholderFallbacks: Map<string, Element>;
}

const DEFAULT_THEME_COLORS: ThemeColors = {
  dk1: "#000000",
  lt1: "#ffffff",
  dk2: "#44546a",
  lt2: "#e7e6e6",
  accent1: "#4472c4",
  accent2: "#ed7d31",
  accent3: "#a5a5a5",
  accent4: "#ffc000",
  accent5: "#5b9bd5",
  accent6: "#70ad47",
  hlink: "#0563c1",
  folHlink: "#954f72",
  tx1: "#000000",
  tx2: "#44546a",
  bg1: "#ffffff",
  bg2: "#e7e6e6",
};

export async function renderPptxSlide(
  pptx: PptxPackage,
  slideContext: PptxSlideContext,
  ownerDocument: Document,
): Promise<RenderedPptxSlide> {
  const pixelWidth = NATURAL_SLIDE_WIDTH;
  const pixelHeight =
    NATURAL_SLIDE_WIDTH * (pptx.slideHeight / pptx.slideWidth);
  const stageEl = ownerDocument.createElement("div");
  stageEl.className = "pptx-reader-stage";
  stageEl.style.width = `${pixelWidth}px`;
  stageEl.style.height = `${pixelHeight}px`;
  const resources = new Set<string>();
  const theme = parseThemeColors(slideContext.theme);
  const placeholderFallbacks = buildPlaceholderFallbacks(slideContext);
  const context: RenderContext = {
    ownerDocument,
    pptx,
    theme,
    resources,
    slideWidth: pptx.slideWidth,
    slideHeight: pptx.slideHeight,
    pixelWidth,
    pixelHeight,
    placeholderFallbacks,
  };

  stageEl.style.background = findBackgroundColor(slideContext, theme);
  const rootSpace: CoordinateSpace = {
    offsetX: 0,
    offsetY: 0,
    scaleX: 1,
    scaleY: 1,
  };

  if (slideContext.master && slideContext.masterPath) {
    await renderLayer(
      stageEl,
      {
        xmlDocument: slideContext.master,
        path: slideContext.masterPath,
        relationships: slideContext.masterRelationships,
      },
      context,
      rootSpace,
      true,
    );
  }
  if (slideContext.layout && slideContext.layoutPath) {
    await renderLayer(
      stageEl,
      {
        xmlDocument: slideContext.layout,
        path: slideContext.layoutPath,
        relationships: slideContext.layoutRelationships,
      },
      context,
      rootSpace,
      true,
    );
  }
  await renderLayer(
    stageEl,
    {
      xmlDocument: slideContext.slide,
      path: slideContext.slidePath,
      relationships: slideContext.slideRelationships,
    },
    context,
    rootSpace,
    false,
  );

  return {
    element: stageEl,
    resources,
    width: pixelWidth,
    height: pixelHeight,
  };
}

async function renderLayer(
  stageEl: HTMLElement,
  layer: LayerContext,
  context: RenderContext,
  coordinateSpace: CoordinateSpace,
  skipPlaceholders: boolean,
): Promise<void> {
  const shapeTree = firstDescendantNamed(
    layer.xmlDocument.documentElement,
    "spTree",
  );
  if (!shapeTree) {
    return;
  }
  for (const shape of childElements(shapeTree)) {
    if (
      !["sp", "pic", "graphicFrame", "cxnSp", "grpSp"].includes(
        localName(shape),
      )
    ) {
      continue;
    }
    if (skipPlaceholders && getPlaceholderKey(shape)) {
      continue;
    }
    await renderShape(
      stageEl,
      shape,
      layer,
      context,
      coordinateSpace,
    );
  }
}

async function renderShape(
  parentEl: HTMLElement,
  shape: Element,
  layer: LayerContext,
  context: RenderContext,
  coordinateSpace: CoordinateSpace,
): Promise<void> {
  if (localName(shape) === "grpSp") {
    const groupTransform = parseGroupTransform(shape, coordinateSpace);
    if (!groupTransform) {
      return;
    }
    for (const child of childElements(shape)) {
      if (
        ["sp", "pic", "graphicFrame", "cxnSp", "grpSp"].includes(
          localName(child),
        )
      ) {
        await renderShape(parentEl, child, layer, context, groupTransform);
      }
    }
    return;
  }

  const fallbackShape = getFallbackShape(shape, context.placeholderFallbacks);
  const transform = parseShapeTransform(
    shape,
    fallbackShape,
    coordinateSpace,
  );
  if (!transform || transform.width <= 0 || transform.height <= 0) {
    return;
  }
  const wrapperEl = context.ownerDocument.createElement("div");
  wrapperEl.className = `pptx-reader-element pptx-reader-${localName(shape)}`;
  applyTransform(wrapperEl, transform, context);
  parentEl.appendChild(wrapperEl);

  switch (localName(shape)) {
    case "pic":
      await renderPicture(wrapperEl, shape, layer, context);
      break;
    case "graphicFrame":
      renderGraphicFrame(wrapperEl, shape, context);
      break;
    case "cxnSp":
    case "sp":
    default:
      renderShapeGeometry(wrapperEl, shape, fallbackShape, context);
      renderShapeText(wrapperEl, shape, fallbackShape, context);
      break;
  }
}

function renderShapeGeometry(
  wrapperEl: HTMLElement,
  shape: Element,
  fallbackShape: Element | null,
  context: RenderContext,
): void {
  const shapeProperties =
    getShapeProperties(shape) ?? getShapeProperties(fallbackShape);
  const presetGeometry = firstChildNamed(shapeProperties, "prstGeom");
  const preset = attribute(presetGeometry, "prst") ?? "rect";
  const isTextBox =
    attribute(firstDescendantNamed(shape, "cNvSpPr"), "txBox") === "1";
  const fill = getShapeFill(
    shape,
    fallbackShape,
    context.theme,
    isTextBox || firstChildNamed(shape, "txBody") !== null,
  );
  const line = getShapeLine(shape, fallbackShape, context.theme);
  if (fill === "none" && !line) {
    return;
  }

  const svgEl = context.ownerDocument.createElementNS(
    SVG_NAMESPACE,
    "svg",
  );
  svgEl.classList.add("pptx-reader-shape-svg");
  svgEl.setAttribute("viewBox", "0 0 100 100");
  svgEl.setAttribute("preserveAspectRatio", "none");
  const geometryEl = createPresetGeometry(
    context.ownerDocument,
    preset,
  );
  geometryEl.setAttribute("fill", fill);
  if (line) {
    geometryEl.setAttribute("stroke", line.color);
    geometryEl.setAttribute("stroke-width", String(line.width));
  } else {
    geometryEl.setAttribute("stroke", "none");
  }
  svgEl.appendChild(geometryEl);
  wrapperEl.appendChild(svgEl);
}

function renderShapeText(
  wrapperEl: HTMLElement,
  shape: Element,
  fallbackShape: Element | null,
  context: RenderContext,
): void {
  const textBody = firstChildNamed(shape, "txBody");
  if (!textBody || descendantsNamed(textBody, "t").length === 0) {
    return;
  }

  const bodyProperties = firstChildNamed(textBody, "bodyPr");
  const textEl = context.ownerDocument.createElement("div");
  textEl.className = "pptx-reader-text";
  const anchor = attribute(bodyProperties, "anchor");
  textEl.style.justifyContent =
    anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
  textEl.style.padding = `${emuToPixels(numberAttribute(bodyProperties, "tIns", 45_720), context)}px ${emuToPixels(numberAttribute(bodyProperties, "rIns", 91_440), context)}px ${emuToPixels(numberAttribute(bodyProperties, "bIns", 45_720), context)}px ${emuToPixels(numberAttribute(bodyProperties, "lIns", 91_440), context)}px`;

  for (const paragraph of childrenNamed(textBody, "p")) {
    const paragraphEl = context.ownerDocument.createElement("div");
    paragraphEl.className = "pptx-reader-paragraph";
    const paragraphProperties = firstChildNamed(paragraph, "pPr");
    const alignment = attribute(paragraphProperties, "algn");
    paragraphEl.style.textAlign =
      alignment === "ctr"
        ? "center"
        : alignment === "r"
          ? "right"
          : alignment === "just"
            ? "justify"
            : "left";
    const level = numberAttribute(paragraphProperties, "lvl", 0);
    if (level > 0) {
      paragraphEl.style.paddingLeft = `${level * 18}px`;
    }
    const bullet = getBulletText(paragraphProperties);
    if (bullet) {
      const bulletEl = context.ownerDocument.createElement("span");
      bulletEl.className = "pptx-reader-bullet";
      bulletEl.textContent = bullet;
      paragraphEl.appendChild(bulletEl);
    }

    for (const run of childElements(paragraph)) {
      const runName = localName(run);
      if (runName === "br") {
        paragraphEl.appendChild(context.ownerDocument.createElement("br"));
        continue;
      }
      if (runName !== "r" && runName !== "fld") {
        continue;
      }
      const value = textContent(firstChildNamed(run, "t"));
      if (!value) {
        continue;
      }
      const spanEl = context.ownerDocument.createElement("span");
      spanEl.textContent = value;
      applyRunStyle(
        spanEl,
        getDefaultRunProperties(textBody, paragraph, fallbackShape),
        context.theme,
      );
      applyRunStyle(
        spanEl,
        firstChildNamed(run, "rPr"),
        context.theme,
      );
      paragraphEl.appendChild(spanEl);
    }
    if (!paragraphEl.hasChildNodes()) {
      paragraphEl.appendChild(context.ownerDocument.createElement("br"));
    }
    textEl.appendChild(paragraphEl);
  }
  wrapperEl.appendChild(textEl);
}

async function renderPicture(
  wrapperEl: HTMLElement,
  shape: Element,
  layer: LayerContext,
  context: RenderContext,
): Promise<void> {
  const blip = firstDescendantNamed(shape, "blip");
  const relationshipId = namespacedAttribute(
    blip,
    "embed",
    OFFICE_RELATIONSHIP_NAMESPACE,
  );
  const relationship = relationshipId
    ? layer.relationships.get(relationshipId)
    : undefined;
  if (!relationship || relationship.external) {
    renderUnsupportedPlaceholder(wrapperEl, context.ownerDocument, "Image");
    return;
  }
  const imagePath = resolvePackagePath(layer.path, relationship.target);
  const bytes = await context.pptx.getBinary(imagePath);
  if (!bytes) {
    renderUnsupportedPlaceholder(wrapperEl, context.ownerDocument, "Image");
    return;
  }
  const mimeType = context.pptx.getImageMimeType(imagePath);
  if (!mimeType) {
    renderUnsupportedPlaceholder(
      wrapperEl,
      context.ownerDocument,
      "Unsupported image",
    );
    return;
  }
  const blob = createBlob(bytes, mimeType);
  const url = URL.createObjectURL(blob);
  context.resources.add(url);
  const imageEl = context.ownerDocument.createElement("img");
  imageEl.className = "pptx-reader-image";
  imageEl.src = url;
  imageEl.alt =
    attribute(firstDescendantNamed(shape, "cNvPr"), "descr") ??
    attribute(firstDescendantNamed(shape, "cNvPr"), "name") ??
    "";
  imageEl.loading = "lazy";
  imageEl.decoding = "async";
  wrapperEl.appendChild(imageEl);
}

function renderGraphicFrame(
  wrapperEl: HTMLElement,
  shape: Element,
  context: RenderContext,
): void {
  const table = firstDescendantNamed(shape, "tbl");
  if (!table) {
    renderUnsupportedPlaceholder(
      wrapperEl,
      context.ownerDocument,
      "Unsupported object",
    );
    return;
  }

  const tableEl = context.ownerDocument.createElement("table");
  tableEl.className = "pptx-reader-table";
  const gridColumns = descendantsNamed(
    firstChildNamed(table, "tblGrid") ?? table,
    "gridCol",
  ).map((column) => numberAttribute(column, "w", 1));
  const totalGridWidth = gridColumns.reduce((total, width) => total + width, 0);
  if (totalGridWidth > 0) {
    const columnGroupEl = context.ownerDocument.createElement("colgroup");
    for (const width of gridColumns) {
      const columnEl = context.ownerDocument.createElement("col");
      columnEl.style.width = `${(width / totalGridWidth) * 100}%`;
      columnGroupEl.appendChild(columnEl);
    }
    tableEl.appendChild(columnGroupEl);
  }

  for (const row of childrenNamed(table, "tr")) {
    const rowEl = context.ownerDocument.createElement("tr");
    for (const cell of childrenNamed(row, "tc")) {
      const cellEl = context.ownerDocument.createElement("td");
      const gridSpan = numberAttribute(cell, "gridSpan", 1);
      const rowSpan = numberAttribute(cell, "rowSpan", 1);
      if (gridSpan > 1) {
        cellEl.colSpan = gridSpan;
      }
      if (rowSpan > 1) {
        cellEl.rowSpan = rowSpan;
      }
      const cellProperties = firstChildNamed(cell, "tcPr");
      const fill = readFillColor(cellProperties, context.theme);
      if (fill && fill !== "none") {
        cellEl.style.background = fill;
      }
      const textBody = firstChildNamed(cell, "txBody");
      if (textBody) {
        renderCellText(cellEl, textBody, context);
      }
      rowEl.appendChild(cellEl);
    }
    tableEl.appendChild(rowEl);
  }
  wrapperEl.appendChild(tableEl);
}

function renderCellText(
  cellEl: HTMLTableCellElement,
  textBody: Element,
  context: RenderContext,
): void {
  for (const paragraph of childrenNamed(textBody, "p")) {
    const paragraphEl = context.ownerDocument.createElement("div");
    for (const run of childElements(paragraph)) {
      if (!["r", "fld"].includes(localName(run))) {
        continue;
      }
      const spanEl = context.ownerDocument.createElement("span");
      spanEl.textContent = textContent(firstChildNamed(run, "t"));
      applyRunStyle(
        spanEl,
        getDefaultRunProperties(textBody, paragraph, null),
        context.theme,
      );
      applyRunStyle(spanEl, firstChildNamed(run, "rPr"), context.theme);
      paragraphEl.appendChild(spanEl);
    }
    cellEl.appendChild(paragraphEl);
  }
}

function renderUnsupportedPlaceholder(
  wrapperEl: HTMLElement,
  ownerDocument: Document,
  label: string,
): void {
  const placeholderEl = ownerDocument.createElement("div");
  placeholderEl.className = "pptx-reader-unsupported-object";
  placeholderEl.textContent = label;
  wrapperEl.appendChild(placeholderEl);
}

function applyRunStyle(
  spanEl: HTMLSpanElement,
  runProperties: Element | null,
  theme: ThemeColors,
): void {
  if (!runProperties) {
    return;
  }
  const fontSize = numberAttribute(runProperties, "sz", 0);
  if (fontSize > 0) {
    spanEl.style.fontSize = `${fontSize / 100}pt`;
  }
  if (attribute(runProperties, "b") === "1") {
    spanEl.classList.add("is-bold");
  }
  if (attribute(runProperties, "i") === "1") {
    spanEl.classList.add("is-italic");
  }
  if (
    attribute(runProperties, "u") &&
    attribute(runProperties, "u") !== "none"
  ) {
    spanEl.classList.add("is-underlined");
  }
  const latin = firstChildNamed(runProperties, "latin");
  const typeface = attribute(latin, "typeface");
  if (typeface && !typeface.startsWith("+")) {
    spanEl.style.fontFamily = typeface;
  }
  const color = readFillColor(runProperties, theme);
  if (color && color !== "none") {
    spanEl.style.color = color;
  }
}

function getDefaultRunProperties(
  textBody: Element,
  paragraph: Element,
  fallbackShape: Element | null,
): Element | null {
  const paragraphProperties = firstChildNamed(paragraph, "pPr");
  const direct = firstChildNamed(paragraphProperties, "defRPr");
  if (direct) {
    return direct;
  }

  const level = numberAttribute(paragraphProperties, "lvl", 0) + 1;
  const levelName = `lvl${level}pPr`;
  const textLevel = firstChildNamed(
    firstChildNamed(textBody, "lstStyle"),
    levelName,
  );
  const textDefault = firstChildNamed(textLevel, "defRPr");
  if (textDefault) {
    return textDefault;
  }

  const fallbackTextBody = firstChildNamed(fallbackShape, "txBody");
  const fallbackLevel = firstChildNamed(
    firstChildNamed(fallbackTextBody, "lstStyle"),
    levelName,
  );
  return (
    firstChildNamed(fallbackLevel, "defRPr") ??
    firstChildNamed(paragraph, "endParaRPr")
  );
}

function getBulletText(paragraphProperties: Element | null): string {
  const bulletCharacter = firstChildNamed(paragraphProperties, "buChar");
  const character = attribute(bulletCharacter, "char");
  if (character) {
    return `${character} `;
  }
  if (firstChildNamed(paragraphProperties, "buAutoNum")) {
    return "1. ";
  }
  return "";
}

function getShapeFill(
  shape: Element,
  fallbackShape: Element | null,
  theme: ThemeColors,
  isTextBox: boolean,
): string {
  const current = readFillColor(getShapeProperties(shape), theme);
  if (current) {
    return current;
  }
  const fallback = readFillColor(getShapeProperties(fallbackShape), theme);
  if (fallback) {
    return fallback;
  }
  const styleFill = readStyleReferenceColor(shape, "fillRef", theme);
  if (styleFill) {
    return styleFill;
  }
  return isTextBox ? "none" : theme.accent1;
}

function getShapeLine(
  shape: Element,
  fallbackShape: Element | null,
  theme: ThemeColors,
): { color: string; width: number } | null {
  const currentProperties = getShapeProperties(shape);
  const fallbackProperties = getShapeProperties(fallbackShape);
  const line =
    firstChildNamed(currentProperties, "ln") ??
    firstChildNamed(fallbackProperties, "ln");
  if (line && firstChildNamed(line, "noFill")) {
    return null;
  }
  const color =
    readFillColor(line, theme) ??
    readStyleReferenceColor(shape, "lnRef", theme) ??
    (localName(shape) === "cxnSp" ? theme.dk1 : null);
  if (!color || color === "none") {
    return null;
  }
  return {
    color,
    width: Math.max(0.5, numberAttribute(line, "w", 12_700) / 12_700),
  };
}

function readFillColor(
  container: Element | null | undefined,
  theme: ThemeColors,
): string | null {
  if (!container) {
    return null;
  }
  if (firstChildNamed(container, "noFill")) {
    return "none";
  }
  const solidFill = firstChildNamed(container, "solidFill");
  if (!solidFill) {
    return null;
  }
  return readColorChoice(solidFill, theme);
}

function readStyleReferenceColor(
  shape: Element,
  referenceName: string,
  theme: ThemeColors,
): string | null {
  const style = firstChildNamed(shape, "style");
  const reference = firstChildNamed(style, referenceName);
  return reference ? readColorChoice(reference, theme) : null;
}

function readColorChoice(
  container: Element,
  theme: ThemeColors,
): string | null {
  const colorElement = childElements(container).find((element) =>
    ["srgbClr", "schemeClr", "sysClr", "prstClr"].includes(
      localName(element),
    ),
  );
  if (!colorElement) {
    return null;
  }
  let color: string | undefined;
  switch (localName(colorElement)) {
    case "srgbClr":
      color = normalizeHexColor(attribute(colorElement, "val"));
      break;
    case "schemeClr":
      color = theme[attribute(colorElement, "val") ?? ""];
      break;
    case "sysClr":
      color = normalizeHexColor(
        attribute(colorElement, "lastClr") ??
          attribute(colorElement, "val"),
      );
      break;
    case "prstClr":
      color = presetColor(attribute(colorElement, "val"));
      break;
  }
  return color ? applyColorModifiers(color, colorElement) : null;
}

function parseThemeColors(theme: Document | null): ThemeColors {
  const colors = { ...DEFAULT_THEME_COLORS };
  if (!theme) {
    return colors;
  }
  const scheme = firstDescendantNamed(theme.documentElement, "clrScheme");
  if (!scheme) {
    return colors;
  }
  for (const entry of childElements(scheme)) {
    const value = readColorChoice(entry, colors);
    if (value) {
      colors[localName(entry)] = value;
    }
  }
  colors.tx1 = colors.dk1;
  colors.tx2 = colors.dk2;
  colors.bg1 = colors.lt1;
  colors.bg2 = colors.lt2;
  return colors;
}

function findBackgroundColor(
  slideContext: PptxSlideContext,
  theme: ThemeColors,
): string {
  for (const document of [
    slideContext.slide,
    slideContext.layout,
    slideContext.master,
  ]) {
    if (!document) {
      continue;
    }
    const background = firstDescendantNamed(
      document.documentElement,
      "bg",
    );
    const backgroundProperties = firstChildNamed(background, "bgPr");
    const color = readFillColor(backgroundProperties, theme);
    if (color && color !== "none") {
      return color;
    }
    const backgroundReference = firstChildNamed(background, "bgRef");
    if (backgroundReference) {
      const referenceColor = readColorChoice(backgroundReference, theme);
      if (referenceColor) {
        return referenceColor;
      }
    }
  }
  return theme.bg1;
}

function buildPlaceholderFallbacks(
  slideContext: PptxSlideContext,
): Map<string, Element> {
  const result = new Map<string, Element>();
  for (const document of [slideContext.master, slideContext.layout]) {
    if (!document) {
      continue;
    }
    const shapeTree = firstDescendantNamed(document.documentElement, "spTree");
    if (!shapeTree) {
      continue;
    }
    for (const shape of childElements(shapeTree)) {
      const key = getPlaceholderKey(shape);
      if (key) {
        result.set(key, shape);
      }
    }
  }
  return result;
}

function getFallbackShape(
  shape: Element,
  fallbacks: Map<string, Element>,
): Element | null {
  const key = getPlaceholderKey(shape);
  return key ? fallbacks.get(key) ?? null : null;
}

function getPlaceholderKey(shape: Element): string | null {
  const placeholder = firstDescendantNamed(shape, "ph");
  if (!placeholder) {
    return null;
  }
  const index = attribute(placeholder, "idx");
  if (index) {
    return `idx:${index}`;
  }
  return `type:${attribute(placeholder, "type") ?? "body"}`;
}

function getShapeProperties(shape: Element | null): Element | null {
  if (!shape) {
    return null;
  }
  switch (localName(shape)) {
    case "grpSp":
      return firstChildNamed(shape, "grpSpPr");
    default:
      return firstChildNamed(shape, "spPr");
  }
}

function parseShapeTransform(
  shape: Element,
  fallbackShape: Element | null,
  coordinateSpace: CoordinateSpace,
): ShapeTransform | null {
  const transform =
    getTransform(shape) ??
    getTransform(fallbackShape);
  if (!transform) {
    return null;
  }
  const offset = firstChildNamed(transform, "off");
  const extent = firstChildNamed(transform, "ext");
  if (!offset || !extent) {
    return null;
  }
  const x =
    coordinateSpace.offsetX +
    numberAttribute(offset, "x") * coordinateSpace.scaleX;
  const y =
    coordinateSpace.offsetY +
    numberAttribute(offset, "y") * coordinateSpace.scaleY;
  return {
    x,
    y,
    width: numberAttribute(extent, "cx") * coordinateSpace.scaleX,
    height: numberAttribute(extent, "cy") * coordinateSpace.scaleY,
    rotation: numberAttribute(transform, "rot", 0) / 60_000,
    flipHorizontal: attribute(transform, "flipH") === "1",
    flipVertical: attribute(transform, "flipV") === "1",
  };
}

function parseGroupTransform(
  shape: Element,
  parentSpace: CoordinateSpace,
): CoordinateSpace | null {
  const transform = getTransform(shape);
  if (!transform) {
    return null;
  }
  const offset = firstChildNamed(transform, "off");
  const extent = firstChildNamed(transform, "ext");
  const childOffset = firstChildNamed(transform, "chOff");
  const childExtent = firstChildNamed(transform, "chExt");
  if (!offset || !extent || !childOffset || !childExtent) {
    return null;
  }
  const childWidth = numberAttribute(childExtent, "cx", 1);
  const childHeight = numberAttribute(childExtent, "cy", 1);
  const groupScaleX =
    (numberAttribute(extent, "cx") / Math.max(childWidth, 1)) *
    parentSpace.scaleX;
  const groupScaleY =
    (numberAttribute(extent, "cy") / Math.max(childHeight, 1)) *
    parentSpace.scaleY;
  return {
    offsetX:
      parentSpace.offsetX +
      numberAttribute(offset, "x") * parentSpace.scaleX -
      numberAttribute(childOffset, "x") * groupScaleX,
    offsetY:
      parentSpace.offsetY +
      numberAttribute(offset, "y") * parentSpace.scaleY -
      numberAttribute(childOffset, "y") * groupScaleY,
    scaleX: groupScaleX,
    scaleY: groupScaleY,
  };
}

function getTransform(shape: Element | null): Element | null {
  if (shape && localName(shape) === "graphicFrame") {
    return firstChildNamed(shape, "xfrm");
  }
  return firstChildNamed(getShapeProperties(shape), "xfrm");
}

function applyTransform(
  element: HTMLElement,
  transform: ShapeTransform,
  context: RenderContext,
): void {
  element.style.left = `${(transform.x / context.slideWidth) * context.pixelWidth}px`;
  element.style.top = `${(transform.y / context.slideHeight) * context.pixelHeight}px`;
  element.style.width = `${(transform.width / context.slideWidth) * context.pixelWidth}px`;
  element.style.height = `${(transform.height / context.slideHeight) * context.pixelHeight}px`;
  const transforms: string[] = [];
  if (transform.rotation !== 0) {
    transforms.push(`rotate(${transform.rotation}deg)`);
  }
  if (transform.flipHorizontal || transform.flipVertical) {
    transforms.push(
      `scale(${transform.flipHorizontal ? -1 : 1}, ${transform.flipVertical ? -1 : 1})`,
    );
  }
  if (transforms.length > 0) {
    element.style.transform = transforms.join(" ");
  }
}

function emuToPixels(value: number, context: RenderContext): number {
  return (value / context.slideWidth) * context.pixelWidth;
}

function createPresetGeometry(
  ownerDocument: Document,
  preset: string,
): SVGElement {
  switch (preset) {
    case "ellipse":
      return createSvgElement(ownerDocument, "ellipse", {
        cx: "50",
        cy: "50",
        rx: "49",
        ry: "49",
      });
    case "line":
    case "straightConnector1":
      return createSvgElement(ownerDocument, "line", {
        x1: "0",
        y1: "0",
        x2: "100",
        y2: "100",
      });
    case "triangle":
      return createSvgElement(ownerDocument, "polygon", {
        points: "50,0 100,100 0,100",
      });
    case "rtTriangle":
      return createSvgElement(ownerDocument, "polygon", {
        points: "0,0 100,100 0,100",
      });
    case "diamond":
      return createSvgElement(ownerDocument, "polygon", {
        points: "50,0 100,50 50,100 0,50",
      });
    case "parallelogram":
      return createSvgElement(ownerDocument, "polygon", {
        points: "20,0 100,0 80,100 0,100",
      });
    case "trapezoid":
      return createSvgElement(ownerDocument, "polygon", {
        points: "20,0 80,0 100,100 0,100",
      });
    case "hexagon":
      return createSvgElement(ownerDocument, "polygon", {
        points: "25,0 75,0 100,50 75,100 25,100 0,50",
      });
    case "octagon":
      return createSvgElement(ownerDocument, "polygon", {
        points: "30,0 70,0 100,30 100,70 70,100 30,100 0,70 0,30",
      });
    case "rightArrow":
      return createSvgElement(ownerDocument, "polygon", {
        points: "0,25 60,25 60,0 100,50 60,100 60,75 0,75",
      });
    case "leftArrow":
      return createSvgElement(ownerDocument, "polygon", {
        points: "100,25 40,25 40,0 0,50 40,100 40,75 100,75",
      });
    case "upArrow":
      return createSvgElement(ownerDocument, "polygon", {
        points: "25,100 25,40 0,40 50,0 100,40 75,40 75,100",
      });
    case "downArrow":
      return createSvgElement(ownerDocument, "polygon", {
        points: "25,0 25,60 0,60 50,100 100,60 75,60 75,0",
      });
    case "chevron":
      return createSvgElement(ownerDocument, "polygon", {
        points: "0,0 65,0 100,50 65,100 0,100 35,50",
      });
    case "roundRect":
      return createSvgElement(ownerDocument, "rect", {
        x: "1",
        y: "1",
        width: "98",
        height: "98",
        rx: "10",
        ry: "10",
      });
    case "rect":
    default:
      return createSvgElement(ownerDocument, "rect", {
        x: "1",
        y: "1",
        width: "98",
        height: "98",
      });
  }
}

function createSvgElement(
  ownerDocument: Document,
  name: string,
  attributes: Record<string, string>,
): SVGElement {
  const element = ownerDocument.createElementNS(SVG_NAMESPACE, name);
  for (const [attributeName, value] of Object.entries(attributes)) {
    element.setAttribute(attributeName, value);
  }
  return element;
}

function applyColorModifiers(color: string, colorElement: Element): string {
  let [red, green, blue] = hexToRgb(color);
  const tint = firstChildNamed(colorElement, "tint");
  const shade = firstChildNamed(colorElement, "shade");
  const luminanceModifier = firstChildNamed(colorElement, "lumMod");
  const luminanceOffset = firstChildNamed(colorElement, "lumOff");
  if (tint) {
    const amount = numberAttribute(tint, "val", 0) / 100_000;
    red += (255 - red) * amount;
    green += (255 - green) * amount;
    blue += (255 - blue) * amount;
  }
  if (shade) {
    const amount = numberAttribute(shade, "val", 100_000) / 100_000;
    red *= amount;
    green *= amount;
    blue *= amount;
  }
  if (luminanceModifier) {
    const amount =
      numberAttribute(luminanceModifier, "val", 100_000) / 100_000;
    red *= amount;
    green *= amount;
    blue *= amount;
  }
  if (luminanceOffset) {
    const amount =
      (numberAttribute(luminanceOffset, "val", 0) / 100_000) * 255;
    red += amount;
    green += amount;
    blue += amount;
  }
  return rgbToHex(red, green, blue);
}

function normalizeHexColor(value: string | null): string | undefined {
  if (!value || !/^[0-9a-f]{6}$/i.test(value)) {
    return undefined;
  }
  return `#${value.toLowerCase()}`;
}

function presetColor(value: string | null): string | undefined {
  const colors: Record<string, string> = {
    black: "#000000",
    white: "#ffffff",
    red: "#ff0000",
    green: "#008000",
    blue: "#0000ff",
    yellow: "#ffff00",
    gray: "#808080",
    grey: "#808080",
    orange: "#ffa500",
    purple: "#800080",
  };
  return value ? colors[value] : undefined;
}

function hexToRgb(color: string): [number, number, number] {
  return [
    Number.parseInt(color.slice(1, 3), 16),
    Number.parseInt(color.slice(3, 5), 16),
    Number.parseInt(color.slice(5, 7), 16),
  ];
}

function rgbToHex(red: number, green: number, blue: number): string {
  return `#${[red, green, blue]
    .map((value) =>
      Math.min(255, Math.max(0, Math.round(value)))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")}`;
}

function createBlob(bytes: Uint8Array, type: string): Blob {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return new Blob([buffer], { type });
}
