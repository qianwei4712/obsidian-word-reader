import assert from "node:assert/strict";
import test from "node:test";

import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import { XlsxVirtualGrid } from "../src/xlsx/xlsxVirtualGrid";
import { createRichXlsx } from "./xlsxFixture";

void test("XLSX virtual grid mounts only the viewport and bounded overscan", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  const grid = new XlsxVirtualGrid(sheet, {
    overscanRows: 4,
    overscanColumns: 2,
    maxMountedCells: 2_500,
  });
  const firstWindow = grid.calculate({
    scrollTop: 0,
    scrollLeft: 0,
    width: 1_024,
    height: 720,
  });

  assert.equal(firstWindow.startRow, 0);
  assert.equal(firstWindow.startColumn, 0);
  assert.ok(firstWindow.endRow < 100);
  assert.ok(firstWindow.mountedCellCount <= 2_500);
  assert.ok(firstWindow.populatedCells.some((cell) => cell.ref === "A1"));
  assert.ok(
    firstWindow.populatedCells.every((cell) => cell.ref !== "Z100000"),
  );
});

void test("100,000-row sparse XLSX scroll reaches the tail without full-grid DOM", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  const grid = new XlsxVirtualGrid(sheet);
  const tailWindow = grid.calculate({
    scrollTop: grid.totalHeight - 600,
    scrollLeft: grid.totalWidth - 600,
    width: 600,
    height: 600,
  });

  assert.ok(tailWindow.startRow > 99_900);
  assert.equal(tailWindow.endColumn, 26);
  assert.ok(tailWindow.mountedCellCount <= 2_500);
  assert.ok(tailWindow.populatedCells.some((cell) => cell.ref === "Z100000"));
  assert.ok(tailWindow.estimatedDomNodeCount < 3_000);
});

void test("XLSX virtual grid enforces finite overscan and mount budgets", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  assert.throws(
    () => new XlsxVirtualGrid(sheet, { overscanRows: 51 }),
    RangeError,
  );
  const grid = new XlsxVirtualGrid(sheet, { maxMountedCells: 100 });
  const window = grid.calculate({
    scrollTop: 0,
    scrollLeft: 0,
    width: 10_000,
    height: 10_000,
  });
  assert.ok(window.mountedCellCount <= 100);
});
