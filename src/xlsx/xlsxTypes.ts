export type XlsxSheetVisibility = "visible" | "hidden" | "veryHidden";

export interface XlsxColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
}

export interface XlsxFontStyle {
  name?: string;
  size?: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color?: XlsxColor;
}

export interface XlsxFillStyle {
  pattern: string;
  foreground?: XlsxColor;
  background?: XlsxColor;
}

export interface XlsxBorderEdge {
  style?: string;
  color?: XlsxColor;
}

export interface XlsxBorderStyle {
  left: XlsxBorderEdge;
  right: XlsxBorderEdge;
  top: XlsxBorderEdge;
  bottom: XlsxBorderEdge;
}

export interface XlsxAlignmentStyle {
  horizontal?: string;
  vertical?: string;
  wrapText: boolean;
  textRotation: number;
}

export interface XlsxCellStyle {
  id: number;
  numberFormatId: number;
  numberFormatCode: string;
  isDate: boolean;
  font: XlsxFontStyle;
  fill: XlsxFillStyle;
  border: XlsxBorderStyle;
  alignment: XlsxAlignmentStyle;
}

export interface XlsxFormula {
  text: string;
  cachedValue: XlsxCellValue;
  calculation: "cached-only";
  hasExternalReference: boolean;
  requestsRemoteData: boolean;
}

export interface XlsxHyperlink {
  ref: string;
  target?: string;
  location?: string;
  tooltip?: string;
  external: boolean;
}

export interface XlsxImage {
  path: string;
  mimeType: string;
  row: number;
  column: number;
  name?: string;
  description?: string;
}

export type XlsxCellValue =
  | string
  | number
  | boolean
  | Date
  | null;

export interface XlsxCell {
  ref: string;
  row: number;
  column: number;
  value: XlsxCellValue;
  displayValue: string;
  styleId: number;
  style: XlsxCellStyle;
  formula?: XlsxFormula;
  hyperlink?: XlsxHyperlink;
}

export interface XlsxMergeRange {
  ref: string;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

export interface XlsxFrozenPane {
  rows: number;
  columns: number;
  topLeftCell?: string;
}

export interface XlsxSheetDescriptor {
  name: string;
  state: XlsxSheetVisibility;
  path: string;
  relationshipId: string;
}

export interface XlsxPackageDiagnostics {
  ignoredExternalRelationships: number;
  ignoredDataConnections: number;
  formulaCalculation: "cached-only";
}
