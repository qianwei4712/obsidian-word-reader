import assert from "node:assert/strict";
import test from "node:test";

import { OoxmlPolicyError } from "../src/ooxml/packagePolicy";
import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import { XlsxWorksheetCancelledError } from "../src/xlsx/xlsxWorksheet";
import {
  addOoxmlEntry,
  createRichXlsx,
} from "./xlsxFixture";

void test("XlsxPackage parses safe workbook structure and rich cell data", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  assert.equal(workbook.sheetCount, 2);
  assert.equal(workbook.sheets[0].name, "Summary");
  assert.equal(workbook.sheets[1].state, "hidden");
  assert.deepEqual(
    workbook.definedNames.map((name) => ({
      name: name.name,
      target: name.target,
      sheetIndex: name.sheetIndex,
      scopeSheetIndex: name.scopeSheetIndex,
      ref: name.range.ref,
    })),
    [
      {
        name: "ReportArea",
        target: "'Summary'!$A$1:$E$2",
        sheetIndex: 0,
        scopeSheetIndex: undefined,
        ref: "$A$1:$E$2",
      },
      {
        name: "LocalTail",
        target: "Summary!$Z$100000",
        sheetIndex: 0,
        scopeSheetIndex: 0,
        ref: "$Z$100000",
      },
    ],
  );
  assert.equal(workbook.date1904, false);
  assert.equal(workbook.diagnostics.formulaCalculation, "cached-only");
  assert.ok(workbook.diagnostics.ignoredDataConnections >= 2);

  const sheet = await workbook.getWorksheet(0);
  assert.equal(sheet.rowCount, 100_000);
  assert.equal(sheet.columnCount, 26);
  assert.equal(sheet.populatedCellCount, 12);
  assert.deepEqual(sheet.frozenPane, {
    rows: 1,
    columns: 1,
    topLeftCell: "B2",
  });
  assert.equal(sheet.merges[0].ref, "A1:B1");
  assert.equal(sheet.getCell(0, 0)?.value, "Merged title");
  assert.equal(sheet.getCell(0, 0)?.style.font.bold, true);
  assert.equal(sheet.getCell(0, 0)?.style.fill.pattern, "solid");
  assert.equal(sheet.getCell(1, 0)?.displayValue, "2024-01-01");
  assert.equal(sheet.getCell(1, 1)?.displayValue, "12,345.68");
  assert.equal(sheet.getCell(1, 2)?.displayValue, "12.50%");
  assert.equal(sheet.getCell(99_999, 25)?.value, "Sparse tail");
  assert.deepEqual(sheet.getComment(1, 1), {
    ref: "B2",
    row: 1,
    column: 1,
    author: "Ada",
    text: "Review cached revenue before publishing.",
  });
  assert.equal(sheet.getCell(2, 6), undefined);
  assert.deepEqual(sheet.getComment(2, 6), {
    ref: "G3",
    row: 2,
    column: 6,
    author: "Ada",
    text: "Blank cells keep comment markers.",
  });
  assert.equal(sheet.conditionalFormattingRules.length, 3);
  assert.deepEqual(sheet.getConditionalPresentation(1, 1)?.css, {
    fontWeight: "700",
    color: "#ffffff",
    backgroundColor: "#c00000",
  });
  assert.equal(
    sheet.getConditionalPresentation(2, 1)?.css.backgroundColor,
    "#f8696b",
  );
  assert.ok(
    Math.abs(
      (sheet.getConditionalPresentation(2, 2)?.dataBar?.fraction ?? 0) -
        3 / 7,
    ) < 0.0001,
  );
});

void test("XLSX formulas remain cached-only and links are metadata only", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  const remoteFormula = sheet.getCell(1, 3)?.formula;
  assert.equal(remoteFormula?.text, 'WEBSERVICE("https://example.invalid/value")');
  assert.equal(remoteFormula?.cachedValue, 42);
  assert.equal(remoteFormula?.calculation, "cached-only");
  assert.equal(remoteFormula?.requestsRemoteData, true);
  assert.equal(sheet.getCell(1, 4)?.formula?.cachedValue, 12_346.678);
  assert.equal(sheet.getCell(1, 5)?.formula?.text, "");
  assert.equal(sheet.getCell(1, 5)?.formula?.cachedValue, 99);

  const externalLink = sheet.getCell(1, 2)?.hyperlink;
  assert.equal(externalLink?.external, true);
  assert.equal(externalLink?.target, "https://example.invalid/docs");
  assert.equal(sheet.getCell(1, 4)?.hyperlink?.location, "'Hidden data'!A1");
  assert.ok(workbook.diagnostics.ignoredExternalRelationships >= 2);
});

void test("XLSX image metadata and raster bytes stay package-local", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  assert.deepEqual(sheet.images[0], {
    path: "xl/media/image1.png",
    mimeType: "image/png",
    row: 4,
    column: 3,
    anchor: {
      from: {
        row: 4,
        column: 3,
        rowOffsetPx: 0,
        columnOffsetPx: 0,
      },
      to: undefined,
      widthPx: 100,
      heightPx: 100,
    },
    name: "Sample image",
    description: "One pixel fixture",
  });
  assert.deepEqual(sheet.charts[0], {
    path: "xl/charts/chart1.xml",
    kind: "bar",
    title: "Quarterly sales",
    anchor: {
      from: {
        row: 4,
        column: 6,
        rowOffsetPx: 0,
        columnOffsetPx: 0,
      },
      to: {
        row: 12,
        column: 10,
        rowOffsetPx: 0,
        columnOffsetPx: 0,
      },
      widthPx: undefined,
      heightPx: undefined,
    },
    series: [
      {
        name: "Revenue",
        categories: ["Q1", "Q2", "Q3"],
        values: [10, 20, 15],
      },
    ],
    truncated: false,
  });
  const image = await workbook.getImageBinary(sheet.images[0].path);
  assert.equal(image?.[0], 0x89);
});

void test("XLSX worksheet parsing can be cancelled without caching a result", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  await assert.rejects(
    () => workbook.getWorksheet(0, { isCancelled: () => true }),
    XlsxWorksheetCancelledError,
  );
  assert.equal(workbook.getCacheDiagnostics().worksheets, 0);
  assert.equal((await workbook.getWorksheet(0)).rowCount, 100_000);
});

void test("XLSX worksheets report streamed decompression progress", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const progress: number[] = [];
  await workbook.getWorksheet(0, {
    onProgress: (percent) => progress.push(percent),
  });
  assert.ok(progress.length > 0);
  assert.equal(progress.at(-1), 100);
  assert.ok(progress.every((percent) => percent >= 0 && percent <= 100));
});

void test("XLSX packages reject macros, OLE objects, and script media", async () => {
  const base = await createRichXlsx();
  const cases: Array<[string, string | Uint8Array, OoxmlPolicyError["kind"]]> = [
    ["xl/vbaProject.bin", Uint8Array.from([1]), "active-content"],
    ["xl/macroSheets/sheet1.xml", "<worksheet/>", "active-content"],
    ["xl/dialogSheets/sheet1.xml", "<dialogsheet/>", "active-content"],
    ["xl/activeX/activeX1.xml", "<activeX/>", "active-content"],
    ["xl/embeddings/oleObject1.bin", Uint8Array.from([1]), "ole-object"],
    ["xl/media/active.svg", "<svg><script/></svg>", "script-media"],
  ];
  for (const [path, content, kind] of cases) {
    const buffer = await addOoxmlEntry(base, path, content);
    await assert.rejects(
      () => XlsxPackage.load(buffer),
      (error: unknown) =>
        error instanceof OoxmlPolicyError && error.kind === kind,
    );
  }
});

void test("XLSM macro-enabled content types remain rejected", async () => {
  const macroEnabled = await addOoxmlEntry(
    await createRichXlsx(),
    "[Content_Types].xml",
    `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
</Types>`,
  );
  await assert.rejects(
    () => XlsxPackage.load(macroEnabled),
    (error: unknown) =>
      error instanceof OoxmlPolicyError &&
      error.kind === "active-content",
  );
});

void test("XlsxPackage rejects missing workbook parts and encrypted containers", async () => {
  await assert.rejects(
    () => XlsxPackage.load(Uint8Array.from([1, 2, 3, 4]).buffer),
    /ZIP-based Office document/,
  );
  const encryptedHeader = Uint8Array.from([
    0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
    ...Array.from({ length: 24 }, () => 0),
  ]);
  await assert.rejects(
    () => XlsxPackage.load(encryptedHeader.buffer),
    /Encrypted or legacy binary Office containers/,
  );
});

void test("XLSX streaming parts reject structurally malformed XML", async () => {
  const base = await createRichXlsx();
  const malformedSheet = await addOoxmlEntry(
    base,
    "xl/worksheets/sheet1.xml",
    "<worksheet><sheetData></worksheet>",
  );
  const workbook = await XlsxPackage.load(malformedSheet);
  await assert.rejects(
    () => workbook.getWorksheet(0),
    /mismatched closing tag/,
  );

  const malformedStrings = await addOoxmlEntry(
    base,
    "xl/sharedStrings.xml",
    "<sst><si><t>broken</si></sst>",
  );
  const stringsWorkbook = await XlsxPackage.load(malformedStrings);
  await assert.rejects(
    () => stringsWorkbook.getWorksheet(0),
    /mismatched closing tag/,
  );
});
