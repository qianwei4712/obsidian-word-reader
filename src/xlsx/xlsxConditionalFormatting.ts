import {
  attribute,
  childrenNamed,
  descendantsNamed,
  firstChildNamed,
} from "../pptx/xml";
import { parseRangeReference } from "./xlsxReferences";
import {
  xlsxColorToCss,
  xlsxDifferentialStyleToCss,
  type XlsxCssProperties,
} from "./xlsxStyleCss";
import type { XlsxStyleTable } from "./xlsxStyles";
import type {
  XlsxCell,
  XlsxCellIsRule,
  XlsxColor,
  XlsxConditionalFormattingRule,
  XlsxConditionalValue,
  XlsxMergeRange,
} from "./xlsxTypes";

const MAX_CONDITIONAL_FORMATTING_RULES = 256;
const MAX_RANGES_PER_RULE = 64;

export interface XlsxConditionalPresentation {
  css: XlsxCssProperties;
  dataBar?: {
    color: string;
    fraction: number;
  };
}

interface XlsxConditionalStatistics {
  minimum: number;
  maximum: number;
  percentiles: ReadonlyMap<number, number>;
}

export function parseXlsxConditionalFormatting(
  root: Element,
  styles: XlsxStyleTable,
): XlsxConditionalFormattingRule[] {
  const rules: XlsxConditionalFormattingRule[] = [];
  for (const group of descendantsNamed(root, "conditionalFormatting")) {
    const ranges = parseRanges(attribute(group, "sqref"));
    if (ranges.length === 0) {
      continue;
    }
    for (const rule of childrenNamed(group, "cfRule")) {
      if (rules.length >= MAX_CONDITIONAL_FORMATTING_RULES) {
        return sortRules(rules);
      }
      const parsed = parseRule(rule, ranges, styles);
      if (parsed) {
        rules.push(parsed);
      }
    }
  }
  return sortRules(rules);
}

export class XlsxConditionalFormattingIndex {
  private readonly statistics = new Map<
    XlsxConditionalFormattingRule,
    XlsxConditionalStatistics
  >();

  constructor(
    private readonly rules: readonly XlsxConditionalFormattingRule[],
    private readonly cells: readonly XlsxCell[],
  ) {}

  resolve(cell: XlsxCell | undefined): XlsxConditionalPresentation | null {
    if (!cell || typeof cell.value !== "number") {
      return null;
    }
    const presentation: XlsxConditionalPresentation = { css: {} };
    let applied = false;
    for (const rule of this.rules) {
      if (!contains(rule.ranges, cell.row, cell.column)) {
        continue;
      }
      if (rule.type === "cellIs") {
        if (!matchesCellRule(rule, cell.value)) {
          continue;
        }
        Object.assign(
          presentation.css,
          xlsxDifferentialStyleToCss(rule.style),
        );
        applied = true;
      } else {
        const statistics = this.getStatistics(rule);
        if (rule.type === "colorScale") {
          const color = resolveColorScale(rule, cell.value, statistics);
          if (!color) {
            continue;
          }
          presentation.css.backgroundColor = color;
          applied = true;
        } else {
          const minimum = resolveThreshold(rule.minimum, statistics);
          const maximum = resolveThreshold(rule.maximum, statistics);
          const span = maximum - minimum;
          const fraction =
            span <= 0 ? 1 : clamp((cell.value - minimum) / span, 0, 1);
          const color = xlsxColorToCss(rule.color);
          if (!color) {
            continue;
          }
          presentation.dataBar = { color, fraction };
          applied = true;
        }
      }
      if (rule.stopIfTrue) {
        break;
      }
    }
    return applied ? presentation : null;
  }

  private getStatistics(
    rule: XlsxConditionalFormattingRule,
  ): XlsxConditionalStatistics {
    const cached = this.statistics.get(rule);
    if (cached) {
      return cached;
    }
    const percentileRanks = new Set(
      conditionalValues(rule)
        .filter((threshold) => threshold.type === "percentile")
        .map((threshold) => clamp(threshold.value ?? 0, 0, 100)),
    );
    const values: number[] = [];
    let minimum = Number.POSITIVE_INFINITY;
    let maximum = Number.NEGATIVE_INFINITY;
    for (const cell of this.cells) {
      if (
        typeof cell.value === "number" &&
        contains(rule.ranges, cell.row, cell.column)
      ) {
        minimum = Math.min(minimum, cell.value);
        maximum = Math.max(maximum, cell.value);
        if (percentileRanks.size > 0) {
          values.push(cell.value);
        }
      }
    }
    const hasValues = Number.isFinite(minimum) && Number.isFinite(maximum);
    const percentiles = new Map<number, number>();
    if (hasValues && values.length > 0) {
      values.sort((left, right) => left - right);
      for (const rank of percentileRanks) {
        percentiles.set(rank, percentile(values, rank));
      }
    }
    const statistics: XlsxConditionalStatistics = {
      minimum: hasValues ? minimum : 0,
      maximum: hasValues ? maximum : 0,
      percentiles,
    };
    this.statistics.set(rule, statistics);
    return statistics;
  }
}

function parseRule(
  element: Element,
  ranges: readonly XlsxMergeRange[],
  styles: XlsxStyleTable,
): XlsxConditionalFormattingRule | null {
  const type = attribute(element, "type");
  const priority = readInteger(attribute(element, "priority"), 1);
  const stopIfTrue = readBoolean(attribute(element, "stopIfTrue"));
  if (type === "cellIs") {
    const operator = parseCellOperator(attribute(element, "operator"));
    const values = childrenNamed(element, "formula")
      .map((formula) => parseNumericFormula(formula.textContent ?? ""))
      .filter((value): value is number => value !== null);
    const requiredValues =
      operator === "between" || operator === "notBetween" ? 2 : 1;
    if (!operator || values.length < requiredValues) {
      return null;
    }
    const differentialStyleId = readOptionalInteger(
      attribute(element, "dxfId"),
    );
    return {
      type,
      priority,
      stopIfTrue,
      ranges,
      operator,
      values: values.slice(0, requiredValues),
      style:
        differentialStyleId === undefined
          ? undefined
          : styles.getDifferentialStyle(differentialStyleId),
    };
  }
  if (type === "colorScale") {
    const colorScale = firstChildNamed(element, "colorScale");
    if (!colorScale) {
      return null;
    }
    const thresholds = childrenNamed(colorScale, "cfvo")
      .map(parseConditionalValue)
      .filter((value): value is XlsxConditionalValue => value !== null);
    const colors = childrenNamed(colorScale, "color")
      .map(parseColor)
      .filter((color): color is XlsxColor => color !== null);
    if (
      thresholds.length < 2 ||
      colors.length < 2 ||
      thresholds.length !== colors.length
    ) {
      return null;
    }
    return {
      type,
      priority,
      stopIfTrue,
      ranges,
      thresholds: thresholds.slice(0, 3),
      colors: colors.slice(0, 3),
    };
  }
  if (type === "dataBar") {
    const dataBar = firstChildNamed(element, "dataBar");
    if (!dataBar) {
      return null;
    }
    const thresholds = childrenNamed(dataBar, "cfvo")
      .map(parseConditionalValue)
      .filter((value): value is XlsxConditionalValue => value !== null);
    const color = parseColor(firstChildNamed(dataBar, "color"));
    if (thresholds.length < 2 || !color) {
      return null;
    }
    return {
      type,
      priority,
      stopIfTrue,
      ranges,
      minimum: thresholds[0],
      maximum: thresholds[1],
      color,
    };
  }
  return null;
}

function parseRanges(value: string | null): XlsxMergeRange[] {
  if (!value) {
    return [];
  }
  const ranges: XlsxMergeRange[] = [];
  for (const reference of value.trim().split(/\s+/)) {
    if (ranges.length >= MAX_RANGES_PER_RULE) {
      break;
    }
    try {
      ranges.push(parseRangeReference(reference));
    } catch {
      // Unsupported references are ignored without affecting safe ranges.
    }
  }
  return ranges;
}

function parseCellOperator(
  value: string | null,
): XlsxCellIsRule["operator"] | null {
  switch (value) {
    case "between":
    case "notBetween":
    case "equal":
    case "notEqual":
    case "greaterThan":
    case "lessThan":
    case "greaterThanOrEqual":
    case "lessThanOrEqual":
      return value;
    default:
      return null;
  }
}

function parseNumericFormula(value: string): number | null {
  const normalized = value.trim().replace(/^=/, "");
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(normalized)) {
    return null;
  }
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : null;
}

function parseConditionalValue(
  element: Element,
): XlsxConditionalValue | null {
  const type = attribute(element, "type");
  if (
    type !== "min" &&
    type !== "max" &&
    type !== "num" &&
    type !== "percent" &&
    type !== "percentile"
  ) {
    return null;
  }
  const rawValue = attribute(element, "val");
  const value = rawValue === null ? undefined : Number(rawValue);
  return {
    type,
    value: Number.isFinite(value) ? value : undefined,
  };
}

function parseColor(element: Element | null): XlsxColor | null {
  if (!element) {
    return null;
  }
  const rgb = attribute(element, "rgb") ?? undefined;
  const theme = readOptionalInteger(attribute(element, "theme"));
  const indexed = readOptionalInteger(attribute(element, "indexed"));
  return rgb || theme !== undefined || indexed !== undefined
    ? { rgb, theme, indexed }
    : null;
}

function matchesCellRule(rule: XlsxCellIsRule, value: number): boolean {
  const first = rule.values[0] ?? 0;
  const second = rule.values[1] ?? first;
  switch (rule.operator) {
    case "between":
      return value >= Math.min(first, second) &&
        value <= Math.max(first, second);
    case "notBetween":
      return value < Math.min(first, second) ||
        value > Math.max(first, second);
    case "equal":
      return value === first;
    case "notEqual":
      return value !== first;
    case "greaterThan":
      return value > first;
    case "lessThan":
      return value < first;
    case "greaterThanOrEqual":
      return value >= first;
    case "lessThanOrEqual":
      return value <= first;
  }
}

function resolveColorScale(
  rule: Extract<XlsxConditionalFormattingRule, { type: "colorScale" }>,
  value: number,
  statistics: XlsxConditionalStatistics,
): string | null {
  const stops = rule.thresholds.map((threshold, index) => ({
    value: resolveThreshold(threshold, statistics),
    color: xlsxColorToCss(rule.colors[index]),
  }));
  if (stops.some((stop) => !stop.color)) {
    return null;
  }
  if (value <= stops[0].value) {
    return stops[0].color;
  }
  const final = stops.at(-1);
  if (!final) {
    return null;
  }
  if (value >= final.value) {
    return final.color;
  }
  for (let index = 1; index < stops.length; index += 1) {
    const right = stops[index];
    const left = stops[index - 1];
    if (value <= right.value) {
      const span = right.value - left.value;
      const fraction = span <= 0 ? 1 : (value - left.value) / span;
      return interpolateHex(left.color ?? "", right.color ?? "", fraction);
    }
  }
  return final.color;
}

function resolveThreshold(
  threshold: XlsxConditionalValue,
  statistics: XlsxConditionalStatistics,
): number {
  if (threshold.type === "min") {
    return statistics.minimum;
  }
  if (threshold.type === "max") {
    return statistics.maximum;
  }
  if (threshold.type === "num") {
    return threshold.value ?? 0;
  }
  const percentageRank = clamp(threshold.value ?? 0, 0, 100);
  if (threshold.type === "percentile") {
    return (
      statistics.percentiles.get(percentageRank) ??
      statistics.minimum
    );
  }
  const percentage = percentageRank / 100;
  return (
    statistics.minimum +
    (statistics.maximum - statistics.minimum) * percentage
  );
}

function conditionalValues(
  rule: XlsxConditionalFormattingRule,
): readonly XlsxConditionalValue[] {
  if (rule.type === "colorScale") {
    return rule.thresholds;
  }
  if (rule.type === "dataBar") {
    return [rule.minimum, rule.maximum];
  }
  return [];
}

function percentile(sortedValues: readonly number[], rank: number): number {
  if (sortedValues.length === 0) {
    return 0;
  }
  const position = (sortedValues.length - 1) * (rank / 100);
  const lowerIndex = Math.floor(position);
  const upperIndex = Math.ceil(position);
  const lower = sortedValues[lowerIndex] ?? sortedValues[0] ?? 0;
  const upper = sortedValues[upperIndex] ?? lower;
  return lower + (upper - lower) * (position - lowerIndex);
}

function interpolateHex(
  left: string,
  right: string,
  fraction: number,
): string | null {
  const leftRgb = parseHex(left);
  const rightRgb = parseHex(right);
  if (!leftRgb || !rightRgb) {
    return null;
  }
  const bounded = clamp(fraction, 0, 1);
  const channels = leftRgb.map((channel, index) =>
    Math.round(channel + ((rightRgb[index] ?? channel) - channel) * bounded),
  );
  return `#${channels.map((channel) => channel.toString(16).padStart(2, "0")).join("")}`;
}

function parseHex(value: string): [number, number, number] | null {
  const match = /^#([0-9a-f]{6})$/i.exec(value);
  if (!match) {
    return null;
  }
  return [
    Number.parseInt(match[1].slice(0, 2), 16),
    Number.parseInt(match[1].slice(2, 4), 16),
    Number.parseInt(match[1].slice(4, 6), 16),
  ];
}

function contains(
  ranges: readonly XlsxMergeRange[],
  row: number,
  column: number,
): boolean {
  return ranges.some(
    (range) =>
      row >= range.startRow &&
      row <= range.endRow &&
      column >= range.startColumn &&
      column <= range.endColumn,
  );
}

function sortRules(
  rules: readonly XlsxConditionalFormattingRule[],
): XlsxConditionalFormattingRule[] {
  return [...rules].sort(
    (left, right) => left.priority - right.priority,
  );
}

function readInteger(value: string | null, fallback: number): number {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : fallback;
}

function readOptionalInteger(value: string | null): number | undefined {
  if (value === null) {
    return undefined;
  }
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

function readBoolean(value: string | null): boolean {
  return value === "1" || value === "true";
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}
