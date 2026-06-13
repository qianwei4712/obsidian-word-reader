import assert from "node:assert/strict";
import test from "node:test";

import { getWordReaderText } from "../src/i18n";
import {
  classifyWordError,
  createWordDiagnostics,
  fingerprintMessage,
  formatWordDiagnostics,
  hasZipSignature,
  isEncryptedOfficeBuffer,
} from "../src/wordErrors";

const text = getWordReaderText("en");

void test("hasZipSignature recognizes valid Office ZIP signatures", () => {
  assert.equal(hasZipSignature(Uint8Array.from([0x50, 0x4b, 0x03, 0x04]).buffer), true);
  assert.equal(hasZipSignature(Uint8Array.from([0x50, 0x4b, 0x05, 0x06]).buffer), true);
  assert.equal(hasZipSignature(Uint8Array.from([0x00, 0x01, 0x02, 0x03]).buffer), false);
});

void test("isEncryptedOfficeBuffer recognizes encrypted OLE containers", () => {
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const marker = Array.from("EncryptedPackage").flatMap((character) => [
    character.charCodeAt(0),
    0,
  ]);
  const encrypted = Uint8Array.from([...signature, ...marker]).buffer;

  assert.equal(isEncryptedOfficeBuffer(encrypted), true);
  assert.equal(
    isEncryptedOfficeBuffer(Uint8Array.from([0x50, 0x4b, 0x03, 0x04]).buffer),
    false,
  );
});

void test("classifyWordError covers the supported error categories", () => {
  const zip = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]).buffer;

  assert.equal(classifyWordError("Password required", text, zip).kind, "encrypted");
  assert.equal(
    classifyWordError(
      "Cannot open",
      text,
      Uint8Array.from([0x00, 0x01, 0x02, 0x03]).buffer,
    ).kind,
    "format-mismatch",
  );
  assert.equal(classifyWordError("CRC archive error", text, zip).kind, "zip-container");
  assert.equal(classifyWordError("Malformed XML", text, zip).kind, "xml-structure");
  assert.equal(
    classifyWordError("SmartArt is unsupported", text, zip).kind,
    "unsupported-structure",
  );
  assert.equal(classifyWordError("Renderer stopped", text, zip).kind, "unknown");
});

void test("diagnostics expose stable, privacy-safe metadata", () => {
  const message = "Raw error containing private document text";
  const diagnostics = createWordDiagnostics("unknown", message, {
    name: "Report.docx",
    size: 1234,
    mtime: Date.UTC(2026, 5, 7, 1, 2, 3),
  });
  const formatted = formatWordDiagnostics(diagnostics);
  const parsed = JSON.parse(formatted) as Record<string, unknown>;

  assert.equal(fingerprintMessage(message), fingerprintMessage(message));
  assert.match(diagnostics.errorSummary, /fingerprint: [0-9a-f]{8}/);
  assert.equal(formatted.includes(message), false);
  assert.deepEqual(parsed, {
    product: "Office Reader",
    category: "unknown",
    fileName: "Report.docx",
    fileSizeBytes: 1234,
    modifiedAt: "2026-06-07T01:02:03.000Z",
    errorSummary: diagnostics.errorSummary,
  });
});
