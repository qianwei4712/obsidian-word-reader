import assert from "node:assert/strict";
import test from "node:test";

import {
  parseXlsxDefinedNameTarget,
  resolveXlsxDefinedName,
} from "../src/xlsx/xlsxDefinedNames";
import type { XlsxDefinedName } from "../src/xlsx/xlsxTypes";

const sheets = [{ name: "Summary" }, { name: "Hidden data" }];

void test("XLSX named ranges resolve quoted, local, and absolute targets", () => {
  assert.deepEqual(
    parseXlsxDefinedNameTarget(
      "'Hidden data'!$B$2:$D$4",
      sheets,
    ),
    {
      sheetIndex: 1,
      range: {
        ref: "$B$2:$D$4",
        startRow: 1,
        startColumn: 1,
        endRow: 3,
        endColumn: 3,
      },
    },
  );
  assert.deepEqual(
    parseXlsxDefinedNameTarget("$A$1", sheets, 0),
    {
      sheetIndex: 0,
      range: {
        ref: "$A$1",
        startRow: 0,
        startColumn: 0,
        endRow: 0,
        endColumn: 0,
      },
    },
  );
  assert.equal(
    parseXlsxDefinedNameTarget("[remote.xlsx]Sheet1!A1", sheets),
    null,
  );
  assert.equal(parseXlsxDefinedNameTarget("OFFSET(A1,1,1)", sheets, 0), null);
});

void test("sheet-local names take precedence over workbook names", () => {
  const names: XlsxDefinedName[] = [
    makeName("Focus", 0, undefined),
    makeName("Focus", 1, 1),
  ];
  assert.equal(resolveXlsxDefinedName(names, "focus", 1)?.sheetIndex, 1);
  assert.equal(resolveXlsxDefinedName(names, "FOCUS", 0)?.sheetIndex, 0);
  assert.equal(resolveXlsxDefinedName(names, "missing", 0), null);
});

function makeName(
  name: string,
  sheetIndex: number,
  scopeSheetIndex: number | undefined,
): XlsxDefinedName {
  return {
    name,
    target: "A1",
    sheetIndex,
    scopeSheetIndex,
    range: {
      ref: "A1",
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 0,
    },
  };
}
