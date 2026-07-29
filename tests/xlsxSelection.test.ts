import assert from "node:assert/strict";
import test from "node:test";

import {
  MAX_XLSX_COPY_CELLS,
  getXlsxSelectionCellCount,
  normalizeXlsxSelection,
  xlsxSelectionContains,
  xlsxSelectionToMarkdown,
  xlsxSelectionToTsv,
  XlsxSelectionTooLargeError,
} from "../src/xlsx/xlsxSelection";
import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import { createRichXlsx } from "./xlsxFixture";

void test("XLSX selection normalizes drag direction and containment", () => {
  const selection = normalizeXlsxSelection(
    { row: 4, column: 3 },
    { row: 1, column: 1 },
  );
  assert.deepEqual(selection, {
    startRow: 1,
    endRow: 4,
    startColumn: 1,
    endColumn: 3,
  });
  assert.equal(getXlsxSelectionCellCount(selection), 12);
  assert.equal(xlsxSelectionContains(selection, 2, 2), true);
  assert.equal(xlsxSelectionContains(selection, 0, 2), false);
});

void test("XLSX selection copies displayed values by default and formulas separately", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  const selection = {
    startRow: 1,
    endRow: 1,
    startColumn: 1,
    endColumn: 4,
  };
  assert.equal(
    xlsxSelectionToTsv(sheet, selection, "display"),
    "12,345.68\t12.50%\t42\t12346.678",
  );
  assert.equal(
    xlsxSelectionToTsv(sheet, selection, "formula"),
    "12,345.68\t12.50%\t\"=WEBSERVICE(\"\"https://example.invalid/value\"\")\"\t=SUM(B2,1)",
  );
});

void test("XLSX selection exports all displayed rows as a Markdown table", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  assert.equal(
    xlsxSelectionToMarkdown(
      sheet,
      {
        startRow: 1,
        endRow: 1,
        startColumn: 1,
        endColumn: 3,
      },
    ),
    [
      "| B | C | D |",
      "| --- | --- | --- |",
      "| 12,345.68 | 12.50% | 42 |",
    ].join("\n"),
  );
});

void test("XLSX selection bounds clipboard materialization", () => {
  const oversized = {
    startRow: 0,
    endRow: MAX_XLSX_COPY_CELLS,
    startColumn: 0,
    endColumn: 0,
  };
  assert.throws(
    () =>
      xlsxSelectionToTsv(
        { getCell: () => undefined },
        oversized,
        "display",
      ),
    XlsxSelectionTooLargeError,
  );
  assert.throws(
    () =>
      xlsxSelectionToMarkdown(
        { getCell: () => undefined },
        oversized,
      ),
    XlsxSelectionTooLargeError,
  );
});
