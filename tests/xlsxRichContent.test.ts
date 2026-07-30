import assert from "node:assert/strict";
import test from "node:test";

import { parseXlsxChart } from "../src/xlsx/xlsxCharts";
import { parseXlsxComments } from "../src/xlsx/xlsxComments";
import { XlsxStyleTable } from "../src/xlsx/xlsxStyles";
import {
  parseXlsxWorksheet,
  XlsxWorksheetCancelledError,
  XlsxWorksheetStreamParser,
} from "../src/xlsx/xlsxWorksheet";

const descriptor = {
  name: "Streamed",
  state: "visible" as const,
  path: "xl/worksheets/sheet1.xml",
  relationshipId: "rId1",
};

void test("XLSX comments preserve authors and rich text safely", () => {
  const comments = parseXlsxComments(
    `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Ada &amp; Lin</author></authors>
  <commentList>
    <comment ref="C7" authorId="0"><text><r><t>First </t></r><r><t>note</t></r></text></comment>
    <comment ref="bad" authorId="0"><text><t>Ignored</t></text></comment>
  </commentList>
</comments>`,
    "xl/comments1.xml",
  );
  assert.deepEqual(comments, [
    {
      ref: "C7",
      row: 6,
      column: 2,
      author: "Ada & Lin",
      text: "First note",
    },
  ]);
});

void test("XLSX chart parser reads cached common-chart data only", () => {
  const chart = parseXlsxChart(
    `<c:chartSpace xmlns:c="chart"><c:chart>
      <c:title><c:tx><c:rich><c:p><c:r><c:t>Trend</c:t></c:r></c:p></c:rich></c:tx></c:title>
      <c:plotArea><c:lineChart><c:ser>
        <c:tx><c:v>Visits</c:v></c:tx>
        <c:cat><c:strRef><c:strCache><c:pt idx="1"><c:v>Tue</c:v></c:pt><c:pt idx="0"><c:v>Mon</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:f>External.xlsx!A1:A2</c:f><c:numCache><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser></c:lineChart></c:plotArea>
    </c:chart></c:chartSpace>`,
    "xl/charts/chart1.xml",
    {
      from: {
        row: 0,
        column: 0,
        rowOffsetPx: 0,
        columnOffsetPx: 0,
      },
      widthPx: 320,
      heightPx: 200,
    },
  );
  assert.equal(chart.kind, "line");
  assert.equal(chart.title, "Trend");
  assert.deepEqual(chart.series, [
    {
      name: "Visits",
      categories: ["Mon", "Tue"],
      values: [4, 9],
    },
  ]);
});

void test("XLSX chart parser identifies every supported chart subset and fallback", () => {
  const anchor = {
    from: {
      row: 0,
      column: 0,
      rowOffsetPx: 0,
      columnOffsetPx: 0,
    },
    widthPx: 320,
    heightPx: 200,
  };
  for (const [element, expected] of [
    ["barChart", "bar"],
    ["lineChart", "line"],
    ["areaChart", "area"],
    ["pieChart", "pie"],
    ["scatterChart", "unsupported"],
  ] as const) {
    const chart = parseXlsxChart(
      `<c:chartSpace xmlns:c="chart"><c:chart><c:plotArea>
        <c:${element}><c:ser><c:val><c:numCache>
          <c:pt idx="0"><c:v>7</c:v></c:pt>
        </c:numCache></c:val></c:ser></c:${element}>
      </c:plotArea></c:chart></c:chartSpace>`,
      `xl/charts/${element}.xml`,
      anchor,
    );
    assert.equal(chart.kind, expected);
    assert.equal(chart.series.length, expected === "unsupported" ? 0 : 1);
  }
});

void test("XLSX worksheet parser consumes split XML chunks without retaining sheetData", () => {
  const styles = XlsxStyleTable.parse(
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <cellXfs count="1"><xf numFmtId="0"/></cellXfs>
      <dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>
    </styleSheet>`,
  );
  const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B100000"/>
  <sheetData>
    <row r="1" ht="24"><c r="A1" t="inlineStr"><is><t>streamed</t></is></c></row>
    <row r="100000"><c r="B100000"><v>9</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="B100000"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>9</formula></cfRule></conditionalFormatting>
</worksheet>`;
  const parser = new XlsxWorksheetStreamParser({
    descriptor,
    styles,
    sharedStrings: [],
    relationships: new Map(),
    date1904: false,
  });
  for (let offset = 0; offset < xml.length; offset += 7) {
    parser.push(xml.slice(offset, offset + 7));
  }
  const metadata = parser.completeInput();
  assert.ok(metadata.includes("<sheetData/>"));
  assert.ok(!metadata.includes("streamed"));
  const worksheet = parser.finish();
  assert.equal(worksheet.rowCount, 100_000);
  assert.equal(worksheet.getCell(0, 0)?.displayValue, "streamed");
  assert.equal(worksheet.getCell(99_999, 1)?.value, 9);
  assert.ok((worksheet.rowHeights.get(0) ?? 0) > 31);
  assert.equal(
    worksheet.getConditionalPresentation(99_999, 1)?.css.backgroundColor,
    "#ff0000",
  );
});

void test("XLSX three-color scales resolve percentile thresholds from cell values", () => {
  const worksheet = parseXlsxWorksheet(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1"><c r="A1"><v>0</v></c></row>
        <row r="2"><c r="A2"><v>10</v></c></row>
        <row r="3"><c r="A3"><v>100</v></c></row>
      </sheetData>
      <conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1">
        <colorScale>
          <cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>
          <color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/>
        </colorScale>
      </cfRule></conditionalFormatting>
    </worksheet>`,
    {
      descriptor,
      styles: XlsxStyleTable.empty(),
      sharedStrings: [],
      relationships: new Map(),
      date1904: false,
    },
  );
  assert.equal(
    worksheet.getConditionalPresentation(1, 0)?.css.backgroundColor,
    "#ffff00",
  );
});

void test("XLSX streamed parsing cancels between chunks", () => {
  let cancelled = false;
  const parser = new XlsxWorksheetStreamParser({
    descriptor,
    styles: XlsxStyleTable.empty(),
    sharedStrings: [],
    relationships: new Map(),
    date1904: false,
    isCancelled: () => cancelled,
  });
  parser.push("<worksheet><sheetData>");
  cancelled = true;
  assert.throws(
    () => parser.push("<row r=\"1\"/></sheetData></worksheet>"),
    XlsxWorksheetCancelledError,
  );
});

void test("XLSX streamed structure validation rejects split mismatched tags", () => {
  const parser = new XlsxWorksheetStreamParser({
    descriptor,
    styles: XlsxStyleTable.empty(),
    sharedStrings: [],
    relationships: new Map(),
    date1904: false,
  });
  parser.push("<worksheet><sheet");
  assert.throws(
    () => parser.push("Data><row></sheetData></worksheet>"),
    /mismatched closing tag/,
  );
});
