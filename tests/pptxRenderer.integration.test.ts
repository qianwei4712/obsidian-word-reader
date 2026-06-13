import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";

import { PptxPackage } from "../src/pptx/pptxPackage";
import { renderPptxSlide } from "../src/pptx/pptxRenderer";
import { createMinimalPptx } from "./pptxFixture";

class MockClassList {
  readonly values = new Set<string>();

  add(...names: string[]): void {
    for (const name of names) {
      this.values.add(name);
    }
  }
}

class MockElement {
  readonly children: MockElement[] = [];
  readonly attributes = new Map<string, string>();
  readonly classList = new MockClassList();
  readonly style: Record<string, string> = {};
  className = "";
  textContent = "";
  src = "";
  alt = "";
  loading = "";
  decoding = "";
  colSpan = 1;
  rowSpan = 1;

  constructor(readonly tagName: string) {}

  appendChild(child: MockElement): MockElement {
    this.children.push(child);
    return child;
  }

  hasChildNodes(): boolean {
    return this.children.length > 0;
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }
}

class MockDocument {
  createElement(name: string): MockElement {
    return new MockElement(name);
  }

  createElementNS(_namespace: string, name: string): MockElement {
    return new MockElement(name);
  }
}

void test("renderPptxSlide renders theme shapes, text, image, and table", async () => {
  const pptx = await PptxPackage.load(await createMinimalPptx());
  const context = await pptx.getSlideContext(0);
  const rendered = await renderPptxSlide(
    pptx,
    context,
    new MockDocument() as unknown as Document,
  );
  const root = rendered.element as unknown as MockElement;
  const all = flatten(root);

  assert.equal(rendered.width, 960);
  assert.ok(rendered.height > 500);
  assert.ok(all.some((element) => element.textContent === "Hello PPTX"));
  assert.ok(all.some((element) => element.tagName === "img"));
  assert.ok(all.some((element) => element.tagName === "table"));
  assert.ok(all.some((element) => element.textContent === "A1"));
  assert.ok(
    all.some(
      (element) =>
        element.tagName === "rect" &&
        element.attributes.get("fill") === "#cc5500",
    ),
  );
  assert.equal(rendered.resources.size, 1);
  for (const resource of rendered.resources) {
    URL.revokeObjectURL(resource);
  }
});

void test("renderPptxSlide does not load embedded SVG resources", async () => {
  const zip = await JSZip.loadAsync(await createMinimalPptx());
  const relationships = await zip
    .file("ppt/slides/_rels/slide1.xml.rels")
    ?.async("string");
  assert.ok(relationships);
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    relationships.replace("image1.png", "image1.svg"),
  );
  zip.file(
    "ppt/media/image1.svg",
    '<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.com/remote.png"/></svg>',
  );
  const pptx = await PptxPackage.load(
    await zip.generateAsync({ type: "arraybuffer" }),
  );
  const context = await pptx.getSlideContext(0);
  const rendered = await renderPptxSlide(
    pptx,
    context,
    new MockDocument() as unknown as Document,
  );
  const all = flatten(rendered.element as unknown as MockElement);

  assert.equal(rendered.resources.size, 0);
  assert.ok(
    all.some((element) => element.textContent === "Unsupported image"),
  );
});

function flatten(root: MockElement): MockElement[] {
  return [root, ...root.children.flatMap(flatten)];
}
