import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";

import {
  DEFAULT_PPTX_ZIP_LIMITS,
  OOXML_ZIP_HARD_LIMITS,
  validateZipSafety,
  ZipSafetyError,
} from "../src/ooxml/packageSafety";

void test("OOXML safety limits keep global caps and stricter PPTX defaults", () => {
  assert.deepEqual(OOXML_ZIP_HARD_LIMITS, {
    maxFileCount: 10_000,
    maxEntryUncompressedBytes: 128 * 1024 * 1024,
    maxTotalUncompressedBytes: 512 * 1024 * 1024,
    maxCompressionRatio: 200,
  });
  assert.equal(Object.isFrozen(OOXML_ZIP_HARD_LIMITS), true);
  assert.deepEqual(DEFAULT_PPTX_ZIP_LIMITS, {
    maxFileCount: 2_000,
    maxEntryUncompressedBytes: 64 * 1024 * 1024,
    maxTotalUncompressedBytes: 256 * 1024 * 1024,
    maxCompressionRatio: 200,
  });
});

void test("OOXML safety refuses caller limits above a global hard cap", async () => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  assert.throws(
    () =>
      validateZipSafety(buffer, {
        ...OOXML_ZIP_HARD_LIMITS,
        maxFileCount: OOXML_ZIP_HARD_LIMITS.maxFileCount + 1,
      }),
    RangeError,
  );
});

void test("OOXML safety rejects damaged local headers and ZIP64 markers", async () => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  const valid = await zip.generateAsync({ type: "uint8array" });

  const damaged = valid.slice();
  damaged[0] = 0;
  assert.throws(
    () => validateZipSafety(toArrayBuffer(damaged)),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "damaged",
  );

  const endOffset = valid.length - 22;
  const zip64 = new Uint8Array(valid.length + 20);
  zip64.set(valid.subarray(0, endOffset), 0);
  new DataView(zip64.buffer).setUint32(endOffset, 0x07064b50, true);
  zip64.set(valid.subarray(endOffset), endOffset + 20);
  assert.throws(
    () => validateZipSafety(zip64.buffer),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "limit-exceeded",
  );
});

void test("OOXML safety rejects encrypted ZIP flags and unsafe entry paths", async () => {
  const zip = new JSZip();
  zip.file("safe.xml", "<safe/>");
  const valid = await zip.generateAsync({ type: "uint8array" });
  const encrypted = valid.slice();
  const centralOffset = findSignature(encrypted, 0x02014b50);
  new DataView(encrypted.buffer).setUint16(centralOffset + 8, 1, true);
  assert.throws(
    () => validateZipSafety(encrypted.buffer),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "encrypted",
  );

  const unsafe = new JSZip();
  unsafe.file("../escape.xml", "<unsafe/>");
  const unsafeBuffer = await unsafe.generateAsync({ type: "arraybuffer" });
  assert.throws(
    () => validateZipSafety(unsafeBuffer),
    (error: unknown) =>
      error instanceof ZipSafetyError && error.kind === "damaged",
  );
});

function findSignature(bytes: Uint8Array, signature: number): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = 0; offset <= bytes.length - 4; offset += 1) {
    if (view.getUint32(offset, true) === signature) {
      return offset;
    }
  }
  throw new Error("ZIP signature not found in fixture.");
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
}
