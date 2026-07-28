import {
  attribute,
  childrenNamed,
  firstChildNamed,
  parseXml,
} from "../pptx/xml";
import type {
  XlsxAlignmentStyle,
  XlsxBorderEdge,
  XlsxBorderStyle,
  XlsxCellStyle,
  XlsxCellValue,
  XlsxColor,
  XlsxFillStyle,
  XlsxFontStyle,
} from "./xlsxTypes";

const BUILTIN_NUMBER_FORMATS = new Map<number, string>([
  [0, "General"],
  [1, "0"],
  [2, "0.00"],
  [3, "#,##0"],
  [4, "#,##0.00"],
  [9, "0%"],
  [10, "0.00%"],
  [14, "m/d/yy"],
  [15, "d-mmm-yy"],
  [16, "d-mmm"],
  [17, "mmm-yy"],
  [18, "h:mm AM/PM"],
  [19, "h:mm:ss AM/PM"],
  [20, "h:mm"],
  [21, "h:mm:ss"],
  [22, "m/d/yy h:mm"],
  [45, "mm:ss"],
  [46, "[h]:mm:ss"],
  [47, "mmss.0"],
  [49, "@"],
]);

const DEFAULT_FONT: XlsxFontStyle = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
};
const DEFAULT_FILL: XlsxFillStyle = {
  pattern: "none",
};
const DEFAULT_BORDER_EDGE: XlsxBorderEdge = {};
const DEFAULT_BORDER: XlsxBorderStyle = {
  left: DEFAULT_BORDER_EDGE,
  right: DEFAULT_BORDER_EDGE,
  top: DEFAULT_BORDER_EDGE,
  bottom: DEFAULT_BORDER_EDGE,
};
const DEFAULT_ALIGNMENT: XlsxAlignmentStyle = {
  wrapText: false,
  textRotation: 0,
};
const DEFAULT_STYLE: XlsxCellStyle = {
  id: 0,
  numberFormatId: 0,
  numberFormatCode: "General",
  isDate: false,
  font: DEFAULT_FONT,
  fill: DEFAULT_FILL,
  border: DEFAULT_BORDER,
  alignment: DEFAULT_ALIGNMENT,
};

export class XlsxStyleTable {
  constructor(
    private readonly styles: readonly XlsxCellStyle[],
  ) {}

  static empty(): XlsxStyleTable {
    return new XlsxStyleTable([DEFAULT_STYLE]);
  }

  static parse(xml: string, path = "xl/styles.xml"): XlsxStyleTable {
    const document = parseXml(xml, path);
    const root = document.documentElement;
    const numberFormats = new Map(BUILTIN_NUMBER_FORMATS);
    for (const numFmt of childrenNamed(
      firstChildNamed(root, "numFmts") ?? root,
      "numFmt",
    )) {
      const id = readIntegerAttribute(numFmt, "numFmtId", -1);
      const code = attribute(numFmt, "formatCode");
      if (id >= 0 && code) {
        numberFormats.set(id, code);
      }
    }

    const fonts = parseFonts(firstChildNamed(root, "fonts"));
    const fills = parseFills(firstChildNamed(root, "fills"));
    const borders = parseBorders(firstChildNamed(root, "borders"));
    const cellXfs = firstChildNamed(root, "cellXfs");
    const styles = cellXfs
      ? childrenNamed(cellXfs, "xf").map((xf, id) => {
          const numberFormatId = readIntegerAttribute(xf, "numFmtId", 0);
          const numberFormatCode =
            numberFormats.get(numberFormatId) ?? "General";
          return {
            id,
            numberFormatId,
            numberFormatCode,
            isDate: isDateNumberFormat(numberFormatId, numberFormatCode),
            font:
              fonts[readIntegerAttribute(xf, "fontId", 0)] ?? DEFAULT_FONT,
            fill:
              fills[readIntegerAttribute(xf, "fillId", 0)] ?? DEFAULT_FILL,
            border:
              borders[readIntegerAttribute(xf, "borderId", 0)] ??
              DEFAULT_BORDER,
            alignment: parseAlignment(firstChildNamed(xf, "alignment")),
          };
        })
      : [];
    return new XlsxStyleTable(styles.length > 0 ? styles : [DEFAULT_STYLE]);
  }

  get(styleId: number): XlsxCellStyle {
    return this.styles[styleId] ?? this.styles[0] ?? DEFAULT_STYLE;
  }

  get size(): number {
    return this.styles.length;
  }
}

export function formatXlsxValue(
  value: XlsxCellValue,
  style: XlsxCellStyle,
): string {
  if (value === null) {
    return "";
  }
  if (value instanceof Date) {
    return formatDate(value, style.numberFormatCode);
  }
  if (typeof value !== "number") {
    return String(value);
  }

  const code = stripNumberFormatLiterals(style.numberFormatCode);
  if (code.includes("%")) {
    const decimals = decimalPlaces(code);
    return `${(value * 100).toFixed(decimals)}%`;
  }
  if (code.includes("E+") || code.includes("E-")) {
    return value.toExponential(decimalPlaces(code));
  }
  const decimals = decimalPlaces(code);
  if (decimals > 0 || /[#0],[#0]/.test(code)) {
    return addThousandsSeparators(value.toFixed(decimals), code.includes(","));
  }
  return String(value);
}

export function excelSerialToDate(serial: number, date1904: boolean): Date {
  const millisecondsPerDay = 86_400_000;
  if (date1904) {
    return new Date(Date.UTC(1904, 0, 1) + serial * millisecondsPerDay);
  }
  const adjustedSerial = serial >= 60 ? serial - 1 : serial;
  return new Date(
    Date.UTC(1899, 11, 31) + adjustedSerial * millisecondsPerDay,
  );
}

function parseFonts(parent: Element | null): XlsxFontStyle[] {
  if (!parent) {
    return [DEFAULT_FONT];
  }
  return childrenNamed(parent, "font").map((font) => ({
    name: attribute(firstChildNamed(font, "name"), "val") ?? undefined,
    size: readOptionalNumber(firstChildNamed(font, "sz"), "val"),
    bold: firstChildNamed(font, "b") !== null,
    italic: firstChildNamed(font, "i") !== null,
    underline: firstChildNamed(font, "u") !== null,
    strike: firstChildNamed(font, "strike") !== null,
    color: parseColor(firstChildNamed(font, "color")),
  }));
}

function parseFills(parent: Element | null): XlsxFillStyle[] {
  if (!parent) {
    return [DEFAULT_FILL];
  }
  return childrenNamed(parent, "fill").map((fill) => {
    const pattern = firstChildNamed(fill, "patternFill");
    return {
      pattern: attribute(pattern, "patternType") ?? "none",
      foreground: parseColor(firstChildNamed(pattern, "fgColor")),
      background: parseColor(firstChildNamed(pattern, "bgColor")),
    };
  });
}

function parseBorders(parent: Element | null): XlsxBorderStyle[] {
  if (!parent) {
    return [DEFAULT_BORDER];
  }
  return childrenNamed(parent, "border").map((border) => ({
    left: parseBorderEdge(firstChildNamed(border, "left")),
    right: parseBorderEdge(firstChildNamed(border, "right")),
    top: parseBorderEdge(firstChildNamed(border, "top")),
    bottom: parseBorderEdge(firstChildNamed(border, "bottom")),
  }));
}

function parseBorderEdge(edge: Element | null): XlsxBorderEdge {
  return {
    style: attribute(edge, "style") ?? undefined,
    color: parseColor(firstChildNamed(edge, "color")),
  };
}

function parseAlignment(element: Element | null): XlsxAlignmentStyle {
  return {
    horizontal: attribute(element, "horizontal") ?? undefined,
    vertical: attribute(element, "vertical") ?? undefined,
    wrapText: readBooleanAttribute(element, "wrapText"),
    textRotation: readIntegerAttribute(element, "textRotation", 0),
  };
}

function parseColor(element: Element | null): XlsxColor | undefined {
  if (!element) {
    return undefined;
  }
  const rgb = attribute(element, "rgb") ?? undefined;
  const theme = readOptionalNumber(element, "theme");
  const indexed = readOptionalNumber(element, "indexed");
  return rgb || theme !== undefined || indexed !== undefined
    ? { rgb, theme, indexed }
    : undefined;
}

function readIntegerAttribute(
  element: Element | null,
  name: string,
  fallback: number,
): number {
  const value = Number(attribute(element, name));
  return Number.isInteger(value) ? value : fallback;
}

function readOptionalNumber(
  element: Element | null,
  name: string,
): number | undefined {
  const raw = attribute(element, name);
  if (raw === null) {
    return undefined;
  }
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function readBooleanAttribute(element: Element | null, name: string): boolean {
  const value = attribute(element, name);
  return value === "1" || value === "true";
}

function isDateNumberFormat(id: number, code: string): boolean {
  if (
    (id >= 14 && id <= 22) ||
    (id >= 27 && id <= 36) ||
    (id >= 45 && id <= 47) ||
    (id >= 50 && id <= 58)
  ) {
    return true;
  }
  const stripped = stripNumberFormatLiterals(code).toLowerCase();
  return /(?:^|[^a-z])[ymdhis]+(?:[^a-z]|$)/.test(stripped);
}

function stripNumberFormatLiterals(code: string): string {
  return code
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/\[(?!h+\]|m+\]|s+\])[^[]*]/gi, "");
}

function decimalPlaces(code: string): number {
  const decimal = /\.(0+|#+)/.exec(code);
  return decimal?.[1].length ?? 0;
}

function addThousandsSeparators(value: string, enabled: boolean): string {
  if (!enabled) {
    return value;
  }
  const [integer, fraction] = value.split(".");
  const sign = integer.startsWith("-") ? "-" : "";
  const digits = sign ? integer.slice(1) : integer;
  const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return fraction === undefined
    ? `${sign}${grouped}`
    : `${sign}${grouped}.${fraction}`;
}

function formatDate(date: Date, code: string): string {
  const year = String(date.getUTCFullYear()).padStart(4, "0");
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  const hours = String(date.getUTCHours()).padStart(2, "0");
  const minutes = String(date.getUTCMinutes()).padStart(2, "0");
  const seconds = String(date.getUTCSeconds()).padStart(2, "0");
  const includesTime = /[hs]/i.test(stripNumberFormatLiterals(code));
  return includesTime
    ? `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`
    : `${year}-${month}-${day}`;
}
