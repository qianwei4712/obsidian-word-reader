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

export interface XlsxDifferentialStyle {
  font?: Partial<XlsxFontStyle>;
  fill?: Partial<XlsxFillStyle>;
  border?: Partial<XlsxBorderStyle>;
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
  anchor: XlsxDrawingAnchor;
  name?: string;
  description?: string;
}

export interface XlsxDrawingPosition {
  row: number;
  column: number;
  rowOffsetPx: number;
  columnOffsetPx: number;
}

export interface XlsxDrawingAnchor {
  from: XlsxDrawingPosition;
  to?: XlsxDrawingPosition;
  widthPx?: number;
  heightPx?: number;
}

export interface XlsxComment {
  ref: string;
  row: number;
  column: number;
  author: string;
  text: string;
}

export type XlsxChartKind =
  | "area"
  | "bar"
  | "line"
  | "pie"
  | "unsupported";

export interface XlsxChartSeries {
  name: string;
  categories: readonly string[];
  values: readonly number[];
}

export interface XlsxChart {
  path: string;
  kind: XlsxChartKind;
  title: string;
  anchor: XlsxDrawingAnchor;
  series: readonly XlsxChartSeries[];
  truncated: boolean;
}

export interface XlsxConditionalValue {
  type: "min" | "max" | "num" | "percent" | "percentile";
  value?: number;
}

export interface XlsxConditionalFormattingRuleBase {
  priority: number;
  ranges: readonly XlsxMergeRange[];
  stopIfTrue: boolean;
}

export interface XlsxCellIsRule
extends XlsxConditionalFormattingRuleBase {
  type: "cellIs";
  operator:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqual"
    | "lessThanOrEqual";
  values: readonly number[];
  style?: XlsxDifferentialStyle;
}

export interface XlsxColorScaleRule
extends XlsxConditionalFormattingRuleBase {
  type: "colorScale";
  thresholds: readonly XlsxConditionalValue[];
  colors: readonly XlsxColor[];
}

export interface XlsxDataBarRule
extends XlsxConditionalFormattingRuleBase {
  type: "dataBar";
  minimum: XlsxConditionalValue;
  maximum: XlsxConditionalValue;
  color: XlsxColor;
}

export type XlsxConditionalFormattingRule =
  | XlsxCellIsRule
  | XlsxColorScaleRule
  | XlsxDataBarRule;

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

export interface XlsxDefinedName {
  name: string;
  target: string;
  sheetIndex: number;
  scopeSheetIndex?: number;
  range: XlsxMergeRange;
}

export interface XlsxPackageDiagnostics {
  ignoredExternalRelationships: number;
  ignoredDataConnections: number;
  formulaCalculation: "cached-only";
}

export interface XlsxWorksheetParseDiagnostics {
  mode: "streamed";
  inputChunks: number;
  maximumSheetDataBufferCharacters: number;
  metadataCharacters: number;
}
