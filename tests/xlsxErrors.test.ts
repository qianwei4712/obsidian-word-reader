import assert from "node:assert/strict";
import test from "node:test";

import { OoxmlPolicyError } from "../src/ooxml/packagePolicy";
import { ZipSafetyError } from "../src/ooxml/packageSafety";
import { classifyXlsxError } from "../src/xlsx/xlsxErrors";
import { getXlsxReaderText } from "../src/xlsx/xlsxI18n";
import { XlsxPackageError } from "../src/xlsx/xlsxPackage";

const text = getXlsxReaderText("en");

void test("XLSX errors distinguish package safety and unsupported content", () => {
  assert.equal(
    classifyXlsxError(
      new ZipSafetyError("encrypted", "encrypted"),
      text,
    ).title,
    text.errors.encryptedTitle,
  );
  assert.equal(
    classifyXlsxError(
      new ZipSafetyError("limit-exceeded", "too large"),
      text,
    ).title,
    text.errors.limitTitle,
  );
  assert.equal(
    classifyXlsxError(
      new XlsxPackageError("damaged", "missing workbook"),
      text,
    ).title,
    text.errors.damagedTitle,
  );
  assert.equal(
    classifyXlsxError(
      new OoxmlPolicyError("active-content", "macro"),
      text,
    ).title,
    text.errors.unsupportedTitle,
  );
});
