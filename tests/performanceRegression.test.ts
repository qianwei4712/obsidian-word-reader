import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";

import { PptxPackage } from "../src/pptx/pptxPackage";
import { PptxNavigationWindow } from "../src/pptx/pptxNavigationWindow";
import {
  createLargeDocx,
  createLargePptx,
} from "./performanceFixtures";

void test("200-slide PPTX metadata indexing stays within cache budgets", async () => {
  const presentation = await PptxPackage.load(await createLargePptx(200));
  const metadata = await presentation.indexSlideMetadata({
    concurrency: 4,
    priorityIndex: 100,
  });
  const cache = presentation.getCacheDiagnostics();

  assert.equal(metadata.length, 200);
  assert.equal(metadata[100].title, "Performance slide 101");
  assert.ok(cache.xmlEntries <= cache.limits.xmlEntries);
  assert.ok(
    cache.relationshipEntries <= cache.limits.relationshipEntries,
  );
  assert.ok(cache.metadataEntries <= cache.limits.metadataEntries);
  presentation.clearCaches();
  assert.deepEqual(presentation.getCacheDiagnostics(), {
    xmlEntries: 0,
    relationshipEntries: 0,
    slideContextEntries: 0,
    binaryEntries: 0,
    metadataEntries: 0,
    limits: cache.limits,
  });
});

void test("1,000-slide PPTX uses a bounded navigation window", async () => {
  const presentation = await PptxPackage.load(await createLargePptx(1_000));
  const indices = Array.from(
    { length: presentation.slideCount },
    (_, index) => index,
  );
  const navigationWindow = new PptxNavigationWindow(196, 4, 60);
  const mounted = navigationWindow.calculate(indices, 96_000, 1_000);

  assert.equal(presentation.slideCount, 1_000);
  assert.ok(mounted.indices.length <= 60);
});

void test("100-page DOCX fixture preserves page and text scale", async () => {
  const zip = await JSZip.loadAsync(await createLargeDocx(100));
  const documentXml = await zip.file("word/document.xml")?.async("string");

  assert.ok(documentXml);
  assert.equal(documentXml.match(/w:type="page"/g)?.length, 99);
  assert.equal(documentXml.match(/Performance page \d+/g)?.length, 100);
});
