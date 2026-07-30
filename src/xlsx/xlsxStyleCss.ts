import type {
  XlsxBorderEdge,
  XlsxCellStyle,
  XlsxColor,
  XlsxDifferentialStyle,
} from "./xlsxTypes";

export type XlsxCssProperties = Partial<
  Pick<
    CSSStyleDeclaration,
    | "backgroundColor"
    | "borderBottom"
    | "borderLeft"
    | "borderRight"
    | "borderTop"
    | "color"
    | "fontFamily"
    | "fontSize"
    | "fontStyle"
    | "fontWeight"
    | "textAlign"
    | "textDecoration"
    | "transform"
    | "verticalAlign"
    | "whiteSpace"
  >
>;

const THEME_COLORS = [
  "#ffffff",
  "#000000",
  "#e7e6e6",
  "#44546a",
  "#4472c4",
  "#ed7d31",
  "#a5a5a5",
  "#ffc000",
  "#5b9bd5",
  "#70ad47",
  "#0563c1",
  "#954f72",
] as const;

const INDEXED_COLORS = [
  "#000000",
  "#ffffff",
  "#ff0000",
  "#00ff00",
  "#0000ff",
  "#ffff00",
  "#ff00ff",
  "#00ffff",
] as const;

export function xlsxCellStyleToCss(
  style: XlsxCellStyle,
): XlsxCssProperties {
  const decorations: string[] = [];
  if (style.font.underline) {
    decorations.push("underline");
  }
  if (style.font.strike) {
    decorations.push("line-through");
  }

  const css: XlsxCssProperties = {
    fontWeight: style.font.bold ? "700" : "400",
    fontStyle: style.font.italic ? "italic" : "normal",
    textDecoration: decorations.join(" ") || "none",
    whiteSpace: style.alignment.wrapText ? "normal" : "nowrap",
  };
  const fontColor = xlsxColorToCss(style.font.color);
  if (fontColor) {
    css.color = fontColor;
  }
  const fillColor =
    style.fill.pattern === "solid"
      ? xlsxColorToCss(style.fill.foreground)
      : null;
  if (fillColor) {
    css.backgroundColor = fillColor;
  }
  if (style.font.name) {
    css.fontFamily = style.font.name;
  }
  if (
    style.font.size !== undefined &&
    Number.isFinite(style.font.size) &&
    style.font.size > 0
  ) {
    css.fontSize = `${Math.min(style.font.size, 96)}pt`;
  }
  if (isTextAlignment(style.alignment.horizontal)) {
    css.textAlign = style.alignment.horizontal;
  }
  if (isVerticalAlignment(style.alignment.vertical)) {
    css.verticalAlign = style.alignment.vertical;
  }
  if (style.alignment.textRotation !== 0) {
    const rotation = Math.min(
      90,
      Math.max(-90, style.alignment.textRotation),
    );
    css.transform = `rotate(${rotation}deg)`;
  }

  css.borderLeft = borderEdgeToCss(style.border.left);
  css.borderRight = borderEdgeToCss(style.border.right);
  css.borderTop = borderEdgeToCss(style.border.top);
  css.borderBottom = borderEdgeToCss(style.border.bottom);
  return css;
}

export function applyXlsxCellStyle(
  element: HTMLElement,
  style: XlsxCellStyle,
): void {
  Object.assign(element.style, xlsxCellStyleToCss(style));
}

export function xlsxDifferentialStyleToCss(
  style: XlsxDifferentialStyle | undefined,
): XlsxCssProperties {
  if (!style) {
    return {};
  }
  const css: XlsxCssProperties = {};
  if (style.font) {
    const decorations: string[] = [];
    if (style.font.underline) {
      decorations.push("underline");
    }
    if (style.font.strike) {
      decorations.push("line-through");
    }
    if (style.font.bold !== undefined) {
      css.fontWeight = style.font.bold ? "700" : "400";
    }
    if (style.font.italic !== undefined) {
      css.fontStyle = style.font.italic ? "italic" : "normal";
    }
    if (
      style.font.underline !== undefined ||
      style.font.strike !== undefined
    ) {
      css.textDecoration = decorations.join(" ") || "none";
    }
    const color = xlsxColorToCss(style.font.color);
    if (color) {
      css.color = color;
    }
    if (style.font.name) {
      css.fontFamily = style.font.name;
    }
    if (
      style.font.size !== undefined &&
      Number.isFinite(style.font.size) &&
      style.font.size > 0
    ) {
      css.fontSize = `${Math.min(style.font.size, 96)}pt`;
    }
  }
  if (style.fill?.pattern === "solid") {
    const color = xlsxColorToCss(style.fill.foreground);
    if (color) {
      css.backgroundColor = color;
    }
  }
  if (style.border) {
    if (style.border.left) {
      css.borderLeft = borderEdgeToCss(style.border.left);
    }
    if (style.border.right) {
      css.borderRight = borderEdgeToCss(style.border.right);
    }
    if (style.border.top) {
      css.borderTop = borderEdgeToCss(style.border.top);
    }
    if (style.border.bottom) {
      css.borderBottom = borderEdgeToCss(style.border.bottom);
    }
  }
  return css;
}

export function xlsxColorToCss(color: XlsxColor | undefined): string | null {
  if (!color) {
    return null;
  }
  if (color.rgb && /^[0-9a-f]{6}(?:[0-9a-f]{2})?$/i.test(color.rgb)) {
    const rgb = color.rgb.length === 8 ? color.rgb.slice(2) : color.rgb;
    return `#${rgb.toLowerCase()}`;
  }
  if (
    color.theme !== undefined &&
    Number.isInteger(color.theme) &&
    color.theme >= 0
  ) {
    return THEME_COLORS[color.theme] ?? null;
  }
  if (
    color.indexed !== undefined &&
    Number.isInteger(color.indexed) &&
    color.indexed >= 0
  ) {
    return INDEXED_COLORS[color.indexed] ?? null;
  }
  return null;
}

function borderEdgeToCss(edge: XlsxBorderEdge): string {
  if (!edge.style) {
    return "";
  }
  const width =
    edge.style === "medium"
      ? "2px"
      : edge.style === "thick"
        ? "3px"
        : "1px";
  const lineStyle =
    edge.style.includes("dash") || edge.style === "dotted"
      ? edge.style === "dotted"
        ? "dotted"
        : "dashed"
      : edge.style === "double"
        ? "double"
        : "solid";
  return `${width} ${lineStyle} ${xlsxColorToCss(edge.color) ?? "currentColor"}`;
}

function isTextAlignment(
  value: string | undefined,
): value is "left" | "right" | "center" | "justify" {
  return (
    value === "left" ||
    value === "right" ||
    value === "center" ||
    value === "justify"
  );
}

function isVerticalAlignment(
  value: string | undefined,
): value is "top" | "middle" | "bottom" {
  return value === "top" || value === "middle" || value === "bottom";
}
