import assert from "node:assert/strict";
import test from "node:test";

import {
  xlsxCellStyleToCss,
  xlsxColorToCss,
  xlsxDifferentialStyleToCss,
} from "../src/xlsx/xlsxStyleCss";
import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import { createRichXlsx } from "./xlsxFixture";

void test("XLSX style CSS preserves safe basic workbook styling", async () => {
  const workbook = await XlsxPackage.load(await createRichXlsx());
  const sheet = await workbook.getWorksheet(0);
  const style = sheet.getCell(0, 0)?.style;
  assert.ok(style);
  const css = xlsxCellStyleToCss(style);
  assert.equal(css.fontWeight, "700");
  assert.equal(css.fontStyle, "italic");
  assert.equal(css.color, "#336699");
  assert.equal(css.backgroundColor, "#ffcc00");
  assert.equal(css.textAlign, "center");
  assert.equal(css.whiteSpace, "normal");
  assert.match(css.borderLeft ?? "", /^1px solid /);
});

void test("XLSX colors reject malformed RGB and resolve bounded themes", () => {
  assert.equal(xlsxColorToCss({ rgb: "javascript:red" }), null);
  assert.equal(xlsxColorToCss({ rgb: "FF112233" }), "#112233");
  assert.equal(xlsxColorToCss({ theme: 4 }), "#4472c4");
  assert.equal(xlsxColorToCss({ theme: 100 }), null);
});

void test("XLSX differential styles override only declared properties", () => {
  assert.deepEqual(
    xlsxDifferentialStyleToCss({
      font: {
        bold: true,
        color: { rgb: "FFFFFFFF" },
      },
      fill: {
        pattern: "solid",
        foreground: { rgb: "FFC00000" },
      },
    }),
    {
      fontWeight: "700",
      color: "#ffffff",
      backgroundColor: "#c00000",
    },
  );
});
