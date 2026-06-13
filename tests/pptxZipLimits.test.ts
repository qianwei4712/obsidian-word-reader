import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";

import {
  validateZipSafety,
  ZipSafetyError,
} from "../src/pptx/zipLimits";

void test("validateZipSafety accepts a normal PowerPoint-style package", async () => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  zip.file("ppt/presentation.xml", "<p:presentation/>");
  const buffer = await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });

  const summary = validateZipSafety(buffer);
  assert.ok(summary.fileCount >= 2);
  assert.ok(summary.totalUncompressedBytes > 0);
});

void test("validateZipSafety enforces file count and expanded size", async () => {
  const zip = new JSZip();
  zip.file("one.txt", "1234");
  zip.file("two.txt", "5678");
  const buffer = await zip.generateAsync({
    type: "arraybuffer",
    compression: "STORE",
  });

  assert.throws(
    () =>
      validateZipSafety(buffer, {
        maxFileCount: 1,
        maxEntryUncompressedBytes: 100,
        maxTotalUncompressedBytes: 100,
        maxCompressionRatio: 100,
      }),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "limit-exceeded",
  );
  assert.throws(
    () =>
      validateZipSafety(buffer, {
        maxFileCount: 10,
        maxEntryUncompressedBytes: 3,
        maxTotalUncompressedBytes: 100,
        maxCompressionRatio: 100,
      }),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "limit-exceeded",
  );
});

void test("validateZipSafety rejects non-ZIP and unsafe compression ratios", async () => {
  assert.throws(
    () => validateZipSafety(Uint8Array.from([1, 2, 3, 4]).buffer),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "format-mismatch",
  );

  const zip = new JSZip();
  zip.file("large.txt", "A".repeat(20_000));
  const buffer = await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  assert.throws(
    () =>
      validateZipSafety(buffer, {
        maxFileCount: 10,
        maxEntryUncompressedBytes: 100_000,
        maxTotalUncompressedBytes: 100_000,
        maxCompressionRatio: 2,
      }),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "limit-exceeded",
  );
});

void test("validateZipSafety identifies encrypted Office containers", () => {
  const encryptedHeader = Uint8Array.from([
    0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
    ...Array.from({ length: 24 }, () => 0),
  ]);
  assert.throws(
    () => validateZipSafety(encryptedHeader.buffer),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "encrypted",
  );
});
