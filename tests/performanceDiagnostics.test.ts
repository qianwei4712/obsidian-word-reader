import assert from "node:assert/strict";
import test from "node:test";

import {
  createReaderPerformanceDiagnosticReport,
  ReaderPerformanceTracker,
} from "../src/reader/performanceDiagnostics";

void test("performance diagnostics use one privacy-safe schema for every format", () => {
  for (const format of ["docx", "pptx", "xlsx"] as const) {
    const tracker = new ReaderPerformanceTracker();
    tracker.recordTiming("packageLoadMs", 12.345);
    tracker.recordTiming("parseMs", 20);
    tracker.recordTiming("firstReadableMs", 35);
    tracker.recordTiming("searchMs", 4);
    tracker.recordTiming("navigationMs", 1);
    tracker.recordScrollFrame(3);
    tracker.recordScrollFrame(7);
    tracker.recordCancellation(true);
    tracker.recordCleanup(true);
    const report = createReaderPerformanceDiagnosticReport(
      format,
      {
        name: "C:\\Private\\Vault\\report.office",
        size: 1234,
        mtime: Date.UTC(2026, 7, 2),
      },
      tracker.snapshot({
        domNodes: 42,
        cacheEntries: 2,
        cacheLimit: 8,
        retainedResources: 1,
      }),
    );
    const serialized = JSON.stringify(report);

    assert.equal(report.kind, "performance");
    assert.equal(report.file.name, "report.office");
    assert.equal(report.details.measurement, "obsidian-desktop-runtime");
    assert.equal(report.details.timings.packageLoadMs, 12.35);
    assert.equal(report.details.timings.scrollFrameP95Ms, 7);
    assert.deepEqual(report.details.privacy, {
      containsDocumentContent: false,
      containsSpeakerNotes: false,
      containsCellValues: false,
      containsInternalXml: false,
      containsVaultPath: false,
    });
    assert.equal(serialized.includes("Private"), false);
    assert.equal(serialized.includes("C:\\\\"), false);
    assert.equal(serialized.includes("document.xml"), false);
    assert.equal(serialized.includes("secret cell value"), false);
  }
});
