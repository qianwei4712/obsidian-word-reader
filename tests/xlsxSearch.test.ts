import assert from "node:assert/strict";
import test from "node:test";

import {
  searchXlsxCells,
  searchXlsxWorkbook,
  XlsxSearchCancelledError,
} from "../src/xlsx/xlsxSearch";
import type {
  XlsxCell,
  XlsxCellStyle,
  XlsxFormula,
} from "../src/xlsx/xlsxTypes";

const DEFAULT_STYLE: XlsxCellStyle = {
  id: 0,
  numberFormatId: 0,
  numberFormatCode: "General",
  isDate: false,
  font: {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
  },
  fill: { pattern: "none" },
  border: { left: {}, right: {}, top: {}, bottom: {} },
  alignment: { wrapText: false, textRotation: 0 },
};

const remoteFormula: XlsxFormula = {
  text: 'WEBSERVICE("https://example.com")',
  cachedValue: "#VALUE!",
  calculation: "cached-only",
  hasExternalReference: false,
  requestsRemoteData: true,
};

const cells: XlsxCell[] = [
  makeCell("A1", 0, 0, "Quarterly result"),
  makeCell("D2", 1, 3, "#VALUE!", remoteFormula),
  makeCell("Z1000000", 999_999, 25, "Sparse tail"),
];

void test("XLSX search matches displayed values and formula text", async () => {
  assert.deepEqual(await searchXlsxCells(cells, "quarterly"), [
    { row: 0, column: 0 },
  ]);
  assert.deepEqual(await searchXlsxCells(cells, "webservice"), [
    { row: 1, column: 3 },
  ]);
  assert.deepEqual(await searchXlsxCells(cells, "sparse tail"), [
    { row: 999_999, column: 25 },
  ]);
});

void test("XLSX search yields between chunks and cancels promptly", async () => {
  let cancelled = false;
  let yieldCount = 0;
  await assert.rejects(
    searchXlsxCells(cells, "missing", {
      chunkSize: 1,
      isCancelled: () => cancelled,
      yieldControl: async () => {
        yieldCount += 1;
        cancelled = true;
      },
    }),
    XlsxSearchCancelledError,
  );
  assert.equal(yieldCount, 1);
});

void test("XLSX workbook search returns visible and hidden sheet coordinates", async () => {
  const progress: string[] = [];
  const workbook = {
    sheets: [{ name: "Summary" }, { name: "Hidden data" }],
    getWorksheet: async (index: number) => ({
      getPopulatedCells: () =>
        index === 0
          ? [makeCell("A1", 0, 0, "Quarterly result")]
          : [makeCell("C4", 3, 2, "Quarterly archive")],
    }),
  };
  assert.deepEqual(
    await searchXlsxWorkbook(workbook, "quarterly", {
      onSheet: (name) => progress.push(name),
    }),
    [
      {
        sheetIndex: 0,
        sheetName: "Summary",
        row: 0,
        column: 0,
      },
      {
        sheetIndex: 1,
        sheetName: "Hidden data",
        row: 3,
        column: 2,
      },
    ],
  );
  assert.deepEqual(progress, ["Summary", "Hidden data"]);
});

void test("XLSX workbook search cancels between worksheets", async () => {
  let cancelled = false;
  await assert.rejects(
    searchXlsxWorkbook(
      {
        sheets: [{ name: "One" }, { name: "Two" }],
        getWorksheet: async () => ({
          getPopulatedCells: () => cells,
        }),
      },
      "missing",
      {
        isCancelled: () => cancelled,
        yieldControl: async () => {
          cancelled = true;
        },
      },
    ),
    XlsxSearchCancelledError,
  );
});

function makeCell(
  ref: string,
  row: number,
  column: number,
  displayValue: string,
  formula?: XlsxFormula,
): XlsxCell {
  return {
    ref,
    row,
    column,
    value: displayValue,
    displayValue,
    styleId: 0,
    style: DEFAULT_STYLE,
    formula,
  };
}
