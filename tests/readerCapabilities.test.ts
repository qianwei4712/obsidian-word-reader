import assert from "node:assert/strict";
import test from "node:test";

import { DOCX_ADAPTER } from "../src/docx/DocxAdapter";
import { PPTX_ADAPTER } from "../src/pptx/PptxAdapter";
import {
  defineReaderCapabilities,
  hasReaderCapability,
} from "../src/reader/capabilities";

void test("reader capabilities default to disabled", () => {
  const capabilities = defineReaderCapabilities({
    reload: true,
  });
  assert.equal(hasReaderCapability(capabilities, "reload"), true);
  assert.equal(hasReaderCapability(capabilities, "notes"), false);
  assert.equal(Object.isFrozen(capabilities), true);
});

void test("format adapters own recognition and capability declarations", () => {
  assert.equal(DOCX_ADAPTER.supports({ extension: "DOCX" }), true);
  assert.equal(DOCX_ADAPTER.supports({ extension: "pptx" }), false);
  assert.equal(DOCX_ADAPTER.capabilities.copyMarkdown, true);
  assert.equal(DOCX_ADAPTER.capabilities.notes, false);

  assert.equal(PPTX_ADAPTER.supports({ extension: "PPTX" }), true);
  assert.equal(PPTX_ADAPTER.capabilities.notes, true);
  assert.equal(PPTX_ADAPTER.capabilities.paged, true);
  assert.equal(PPTX_ADAPTER.capabilities.copyMarkdown, false);
});
