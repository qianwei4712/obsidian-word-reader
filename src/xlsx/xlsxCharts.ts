import {
  descendantsNamed,
  firstDescendantNamed,
  parseXml,
  textContent,
} from "../pptx/xml";
import { validateXmlStructure } from "../ooxml/xmlStructure";
import type {
  XlsxChart,
  XlsxChartKind,
  XlsxChartSeries,
  XlsxDrawingAnchor,
} from "./xlsxTypes";

const MAX_CHART_SERIES = 16;
const MAX_CHART_POINTS = 512;

export function parseXlsxChart(
  xml: string,
  path: string,
  anchor: XlsxDrawingAnchor,
): XlsxChart {
  validateXmlStructure(xml, path, "chartSpace");
  const document = parseXml(xml, path);
  const root = document.documentElement;
  const chartElement =
    firstDescendantNamed(root, "barChart") ??
    firstDescendantNamed(root, "lineChart") ??
    firstDescendantNamed(root, "pieChart") ??
    firstDescendantNamed(root, "areaChart");
  const kind = chartElement
    ? chartKind(chartElement)
    : "unsupported";
  const allSeries = chartElement
    ? descendantsNamed(chartElement, "ser")
    : [];
  const series = allSeries
    .slice(0, MAX_CHART_SERIES)
    .map(parseSeries)
    .filter((candidate) => candidate.values.length > 0);
  const truncated =
    allSeries.length > MAX_CHART_SERIES ||
    series.some(
      (candidate) =>
        candidate.values.length >= MAX_CHART_POINTS ||
        candidate.categories.length >= MAX_CHART_POINTS,
    );
  return {
    path,
    kind,
    title: readChartTitle(root),
    anchor,
    series,
    truncated,
  };
}

function chartKind(element: Element): XlsxChartKind {
  switch (element.localName || element.nodeName.split(":").at(-1)) {
    case "barChart":
      return "bar";
    case "lineChart":
      return "line";
    case "pieChart":
      return "pie";
    case "areaChart":
      return "area";
    default:
      return "unsupported";
  }
}

function parseSeries(element: Element, index: number): XlsxChartSeries {
  const textElement = firstDescendantNamed(
    firstDescendantNamed(element, "tx"),
    "strCache",
  );
  const directName = textContent(
    firstDescendantNamed(firstDescendantNamed(element, "tx"), "v"),
  ).trim();
  const cachedName = readCachedText(textElement)[0];
  const categoriesElement = firstDescendantNamed(element, "cat");
  const valuesElement = firstDescendantNamed(element, "val");
  const categories = readCachedText(categoriesElement).slice(
    0,
    MAX_CHART_POINTS,
  );
  const values = readCachedNumbers(valuesElement).slice(
    0,
    MAX_CHART_POINTS,
  );
  return {
    name: directName || cachedName || `Series ${index + 1}`,
    categories,
    values,
  };
}

function readCachedText(element: Element | null): string[] {
  if (!element) {
    return [];
  }
  const points = descendantsNamed(element, "pt")
    .map((point) => ({
      index: readIndex(point),
      value: textContent(firstDescendantNamed(point, "v")).trim(),
    }))
    .sort((left, right) => left.index - right.index);
  return points.map((point) => point.value);
}

function readCachedNumbers(element: Element | null): number[] {
  if (!element) {
    return [];
  }
  return descendantsNamed(element, "pt")
    .map((point) => ({
      index: readIndex(point),
      value: Number(textContent(firstDescendantNamed(point, "v")).trim()),
    }))
    .filter((point) => Number.isFinite(point.value))
    .sort((left, right) => left.index - right.index)
    .map((point) => point.value);
}

function readChartTitle(root: Element): string {
  const title = firstDescendantNamed(root, "title");
  if (!title) {
    return "";
  }
  return descendantsNamed(title, "t")
    .map((element) => textContent(element))
    .join("")
    .trim();
}

function readIndex(point: Element): number {
  const raw = point.getAttribute("idx");
  const value = Number(raw);
  return Number.isInteger(value) && value >= 0 ? value : 0;
}
