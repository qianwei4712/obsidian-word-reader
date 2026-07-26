export const DOCX_HEADING_LEVEL_ATTRIBUTE =
  "data-word-reader-heading-level";

const HEADING_LEVEL_STYLE_PROPERTY = `$${DOCX_HEADING_LEVEL_ATTRIBUTE}`;
const MIN_OUTLINE_LEVEL = 0;
const MAX_OUTLINE_LEVEL = 8;

type UnknownRecord = Record<string, unknown>;

export function annotateDocxHeadingLevels(document: unknown): number {
  const documentRecord = asRecord(document);
  const documentPart = asRecord(documentRecord?.documentPart);
  const body = asRecord(documentPart?.body);
  if (!body) {
    return 0;
  }

  const stylesPart = asRecord(documentRecord?.stylesPart);
  const styles = Array.isArray(stylesPart?.styles) ? stylesPart.styles : [];
  const resolveStyleOutlineLevel = createStyleOutlineLevelResolver(styles);
  let headingCount = 0;

  visitNode(body, (paragraph) => {
    const directOutlineLevel = readInteger(paragraph.outlineLevel);
    const styleName = readString(paragraph.styleName);
    const outlineLevel =
      directOutlineLevel ??
      (styleName ? resolveStyleOutlineLevel(styleName) : undefined);
    const headingLevel = toHeadingLevel(outlineLevel);
    if (headingLevel === undefined) {
      return;
    }

    let cssStyle = asRecord(paragraph.cssStyle);
    if (!cssStyle) {
      cssStyle = {};
      paragraph.cssStyle = cssStyle;
    }
    cssStyle[HEADING_LEVEL_STYLE_PROPERTY] = String(headingLevel);
    headingCount += 1;
  });

  return headingCount;
}

function visitNode(
  node: UnknownRecord,
  visitParagraph: (paragraph: UnknownRecord) => void,
): void {
  if (node.type === "paragraph") {
    visitParagraph(node);
  }

  if (!Array.isArray(node.children)) {
    return;
  }

  for (const child of node.children) {
    const childRecord = asRecord(child);
    if (childRecord) {
      visitNode(childRecord, visitParagraph);
    }
  }
}

function createStyleOutlineLevelResolver(
  styles: unknown[],
): (styleId: string) => number | undefined {
  const styleMap = new Map<string, UnknownRecord>();
  for (const style of styles) {
    const styleRecord = asRecord(style);
    const styleId = readString(styleRecord?.id);
    const target = readString(styleRecord?.target);
    if (styleRecord && styleId && (!target || target === "p")) {
      styleMap.set(styleId, styleRecord);
    }
  }

  const cache = new Map<string, number | null>();

  return (styleId: string): number | undefined => {
    return resolveStyleOutlineLevel(styleId, new Set());
  };

  function resolveStyleOutlineLevel(
    styleId: string,
    resolving: Set<string>,
  ): number | undefined {
    if (cache.has(styleId)) {
      return cache.get(styleId) ?? undefined;
    }
    if (resolving.has(styleId)) {
      return undefined;
    }

    const style = styleMap.get(styleId);
    if (!style) {
      cache.set(styleId, null);
      return undefined;
    }

    resolving.add(styleId);
    const paragraphProps = asRecord(style.paragraphProps);
    let outlineLevel = readInteger(paragraphProps?.outlineLevel);

    if (outlineLevel === undefined) {
      outlineLevel =
        inferOutlineLevel(readString(style.name)) ??
        inferOutlineLevel(readString(style.id));
    }

    if (outlineLevel === undefined) {
      const basedOn = readString(style.basedOn);
      if (basedOn) {
        outlineLevel = resolveStyleOutlineLevel(basedOn, resolving);
      }
    }

    resolving.delete(styleId);
    cache.set(styleId, outlineLevel ?? null);
    return outlineLevel;
  }
}

function inferOutlineLevel(styleName: string | undefined): number | undefined {
  if (!styleName) {
    return undefined;
  }

  const match = /^(?:heading|标题|標題)[\s_-]*([1-9])$/iu.exec(
    styleName.trim(),
  );
  return match ? Number(match[1]) - 1 : undefined;
}

function toHeadingLevel(outlineLevel: number | undefined): number | undefined {
  if (
    outlineLevel === undefined ||
    outlineLevel < MIN_OUTLINE_LEVEL ||
    outlineLevel > MAX_OUTLINE_LEVEL
  ) {
    return undefined;
  }
  return outlineLevel + 1;
}

function asRecord(value: unknown): UnknownRecord | undefined {
  return typeof value === "object" && value !== null
    ? (value as UnknownRecord)
    : undefined;
}

function readString(value: unknown): string | undefined {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function readInteger(value: unknown): number | undefined {
  return typeof value === "number" && Number.isInteger(value)
    ? value
    : undefined;
}
