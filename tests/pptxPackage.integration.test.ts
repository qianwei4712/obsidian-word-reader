import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";

import {
  PptxPackage,
  relationshipPartPath,
  resolvePackagePath,
} from "../src/pptx/pptxPackage";
import { createMinimalPptx } from "./pptxFixture";

void test("PptxPackage resolves slide, layout, master, theme, and media", async () => {
  const pptx = await PptxPackage.load(await createMinimalPptx());
  assert.equal(pptx.slideCount, 1);
  assert.equal(pptx.slideWidth, 12_192_000);
  assert.equal(pptx.slideHeight, 6_858_000);
  assert.deepEqual(pptx.slidePaths, ["ppt/slides/slide1.xml"]);

  const context = await pptx.getSlideContext(0);
  assert.equal(context.layoutPath, "ppt/slideLayouts/slideLayout1.xml");
  assert.equal(context.masterPath, "ppt/slideMasters/slideMaster1.xml");
  assert.equal(context.themePath, "ppt/theme/theme1.xml");
  assert.equal(
    context.slideRelationships.get("rId2")?.target,
    "../media/image1.png",
  );
  assert.ok(await pptx.getBinary("ppt/media/image1.png"));
});

void test("PowerPoint relationship paths stay inside the package", () => {
  assert.equal(
    relationshipPartPath("ppt/slides/slide1.xml"),
    "ppt/slides/_rels/slide1.xml.rels",
  );
  assert.equal(
    resolvePackagePath(
      "ppt/slides/slide1.xml",
      "../slideLayouts/slideLayout1.xml",
    ),
    "ppt/slideLayouts/slideLayout1.xml",
  );
  assert.throws(() =>
    resolvePackagePath("ppt/presentation.xml", "../../../outside.xml"),
  );
});

void test("PptxPackage rejects macro-enabled packages", async () => {
  const zip = await JSZip.loadAsync(await createMinimalPptx());
  zip.file("ppt/vbaProject.bin", Uint8Array.from([1, 2, 3]));
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  await assert.rejects(
    () => PptxPackage.load(buffer),
    /Macro-enabled PowerPoint packages are not supported/,
  );
});
