import assert from "node:assert/strict";
import test from "node:test";

import { classifyPptxError } from "../src/pptx/pptxErrors";
import { getPptxReaderText } from "../src/pptx/pptxI18n";
import { PptxPackageError } from "../src/pptx/pptxPackage";
import { ZipSafetyError } from "../src/pptx/zipLimits";

const text = getPptxReaderText("en");

void test("classifyPptxError distinguishes security and package failures", () => {
  assert.equal(
    classifyPptxError(
      new ZipSafetyError("encrypted", "encrypted"),
      text,
    ).title,
    text.errors.encryptedTitle,
  );
  assert.equal(
    classifyPptxError(
      new ZipSafetyError("limit-exceeded", "large"),
      text,
    ).title,
    text.errors.limitTitle,
  );
  assert.equal(
    classifyPptxError(
      new PptxPackageError("damaged", "bad xml"),
      text,
    ).title,
    text.errors.damagedTitle,
  );
  assert.equal(
    classifyPptxError(
      new PptxPackageError("unsupported", "no slides"),
      text,
    ).title,
    text.errors.unsupportedTitle,
  );
});
