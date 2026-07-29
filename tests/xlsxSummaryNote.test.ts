import assert from "node:assert/strict";
import test from "node:test";

import { getXlsxReaderText } from "../src/xlsx/xlsxI18n";
import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import {
  buildXlsxSummaryNote,
  collectXlsxWorkbookSummary,
  XlsxSummaryCancelledError,
} from "../src/xlsx/xlsxSummaryNote";
import { createRichXlsx } from "./xlsxFixture";

void test("XLSX summary notes contain workbook structure, names, and bounded content", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const summary = await collectXlsxWorkbookSummary(workbook);
  const note = buildXlsxSummaryNote(
    {
      basename: "Quarterly workbook",
      path: "Finance/Quarterly workbook.xlsx",
    },
    summary,
    0,
    "A1:E2",
    "| A | B |\n| --- | --- |\n| Merged title |  |",
    getXlsxReaderText("en"),
    new Date(2026, 6, 30),
  );

  assert.match(note, /type: spreadsheet-note/);
  assert.match(note, /reader_format: xlsx/);
  assert.match(note, /current_sheet: "Summary"/);
  assert.match(note, /selected_range: "A1:E2"/);
  assert.match(note, /sheet_count: 2/);
  assert.match(note, /named_range_count: 2/);
  assert.match(note, /ReportArea.*Summary.*A.*E.*2/);
  assert.match(note, /Hidden data \(hidden\)/);
  assert.match(note, /Sparse tail/);
  assert.match(note, /Hidden value/);
  assert.match(note, /## Selected range snapshot/);
  assert.match(note, /\| Merged title \|/);
});

void test("XLSX summary collection is cooperatively cancellable", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  let cancelled = false;
  await assert.rejects(
    collectXlsxWorkbookSummary(workbook, {
      isCancelled: () => cancelled,
      yieldControl: async () => {
        cancelled = true;
      },
    }),
    XlsxSummaryCancelledError,
  );
});
