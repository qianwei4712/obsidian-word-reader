import assert from "node:assert/strict";
import test from "node:test";

import {
  createPptxErrorDiagnosticReport,
  createPptxDiagnosticReport,
  formatPptxDiagnosticReport,
} from "../src/pptx/pptxDiagnostics";
import { PptxPackage } from "../src/pptx/pptxPackage";
import { createMinimalPptx } from "./pptxFixture";

void test("PPTX diagnostics contain local metrics without vault paths", async () => {
  const presentation = await PptxPackage.load(await createMinimalPptx());
  const report = createPptxDiagnosticReport(
    {
      name: "Quarterly review.pptx",
      stat: { size: 1234, mtime: 4567, ctime: 0 },
    },
    presentation,
    0,
    {
      durationMs: 12,
      yieldCount: 2,
      maxWorkSliceMs: 8,
      layerCount: 3,
      shapeCount: 4,
      textShapeCount: 1,
      imageCount: 1,
      tableCount: 1,
      chartCount: 0,
      smartArtCount: 0,
      unsupportedObjectCount: 0,
      resourceCount: 1,
      fontFamilies: ["Aptos"],
    },
    {
      mounted: 3,
      rendered: 2,
      rendering: 1,
      resourceCount: 2,
    },
  );
  const formatted = formatPptxDiagnosticReport(report);

  assert.equal(report.details.presentation.slideCount, 1);
  assert.ok(
    report.details.presentation.cache.xmlEntries <=
      report.details.presentation.cache.limits.xmlEntries,
  );
  assert.equal(report.file.name, "Quarterly review.pptx");
  assert.equal(report.schemaVersion, 1);
  assert.equal(report.format, "pptx");
  assert.equal(report.kind, "render");
  assert.equal(formatted.includes("vault"), false);
  assert.equal(formatted.includes("\\\\"), false);
  assert.deepEqual(JSON.parse(formatted), report);
});

void test("PPTX error diagnostics fingerprint raw errors without copying them", () => {
  const rawMessage =
    "C:\\Private\\Vault\\secret.pptx contains confidential slide text";
  const report = createPptxErrorDiagnosticReport(
    {
      name: "secret.pptx",
      stat: { size: 99, mtime: 4567, ctime: 0 },
    },
    "unknown",
    rawMessage,
  );
  const formatted = formatPptxDiagnosticReport(report);

  assert.equal(report.kind, "error");
  assert.equal(formatted.includes(rawMessage), false);
  assert.equal(formatted.includes("Private"), false);
  assert.match(report.details.fingerprint, /^[0-9a-f]{8}$/);
});
