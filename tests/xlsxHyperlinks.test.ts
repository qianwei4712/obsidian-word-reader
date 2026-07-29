import assert from "node:assert/strict";
import test from "node:test";

import {
  isSafeXlsxExternalTarget,
  parseXlsxWorkbookLocation,
} from "../src/xlsx/xlsxHyperlinks";

void test("XLSX workbook locations parse quoted sheets without resolving external workbooks", () => {
  assert.deepEqual(parseXlsxWorkbookLocation("'Hidden data'!$A$1"), {
    sheetName: "Hidden data",
    reference: "$A$1",
  });
  assert.deepEqual(parseXlsxWorkbookLocation("'Team''s sheet'!B12"), {
    sheetName: "Team's sheet",
    reference: "B12",
  });
  assert.deepEqual(parseXlsxWorkbookLocation("C3"), {
    sheetName: undefined,
    reference: "C3",
  });
  assert.equal(parseXlsxWorkbookLocation("[remote.xlsx]Sheet1!A1"), null);
  assert.equal(parseXlsxWorkbookLocation("'[remote.xlsx]Sheet1'!A1"), null);
});

void test("XLSX external links allow only explicit safe protocols", () => {
  assert.equal(isSafeXlsxExternalTarget("https://example.com"), true);
  assert.equal(isSafeXlsxExternalTarget("http://example.com"), true);
  assert.equal(isSafeXlsxExternalTarget("mailto:user@example.com"), true);
  assert.equal(isSafeXlsxExternalTarget("javascript:alert(1)"), false);
  assert.equal(isSafeXlsxExternalTarget("file:///C:/secret.txt"), false);
  assert.equal(isSafeXlsxExternalTarget("\\\\server\\share"), false);
});
