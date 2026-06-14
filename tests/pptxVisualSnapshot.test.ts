import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import test from "node:test";

import { PptxPackage } from "../src/pptx/pptxPackage";
import { renderPptxSlide } from "../src/pptx/pptxRenderer";
import { createCompatibilityPptx } from "./pptxFixture";

class SnapshotClassList {
  readonly values = new Set<string>();

  add(...names: string[]): void {
    for (const name of names) {
      this.values.add(name);
    }
  }
}

class SnapshotElement {
  readonly children: SnapshotElement[] = [];
  readonly attributes = new Map<string, string>();
  readonly classList = new SnapshotClassList();
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

  appendChild(child: SnapshotElement): SnapshotElement {
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

class SnapshotDocument {
  createElement(name: string): SnapshotElement {
    return new SnapshotElement(name);
  }

  createElementNS(_namespace: string, name: string): SnapshotElement {
    return new SnapshotElement(name);
  }
}

void test("common PPTX layouts match the structural visual baseline", async () => {
  const pptx = await PptxPackage.load(await createCompatibilityPptx());
  const context = await pptx.getSlideContext(0);
  const rendered = await renderPptxSlide(
    pptx,
    context,
    new SnapshotDocument() as unknown as Document,
    { now: () => 0 },
  );
  const root = rendered.element as unknown as SnapshotElement;
  const snapshot = JSON.stringify(serializeElement(root));
  const hash = createHash("sha256").update(snapshot).digest("hex");

  assert.deepEqual(rendered.diagnostics, {
    durationMs: 0,
    yieldCount: 0,
    maxWorkSliceMs: 0,
    layerCount: 3,
    shapeCount: 9,
    textShapeCount: 2,
    imageCount: 1,
    tableCount: 1,
    chartCount: 1,
    smartArtCount: 1,
    unsupportedObjectCount: 2,
    resourceCount: 1,
    fontFamilies: ["Aptos"],
  });
  const labels = collectText(root);
  assert.ok(labels.includes("Chart"));
  assert.ok(labels.includes("SmartArt"));
  assert.equal(
    hash,
    "8286776424492aa0c05a6ef601e2974361b3e7adcfdb65de84caf57a448fd374",
  );

  for (const resource of rendered.resources) {
    URL.revokeObjectURL(resource);
  }
});

function collectText(element: SnapshotElement): string[] {
  return [
    element.textContent,
    ...element.children.flatMap((child) => collectText(child)),
  ].filter(Boolean);
}

function serializeElement(element: SnapshotElement): unknown {
  return {
    tag: element.tagName,
    classes: [
      ...element.className.split(/\s+/).filter(Boolean),
      ...element.classList.values,
    ].sort(),
    text: element.textContent,
    attributes: [...element.attributes].sort(([left], [right]) =>
      left.localeCompare(right)),
    style: Object.entries(element.style).sort(([left], [right]) =>
      left.localeCompare(right)),
    media:
      element.tagName === "img"
        ? {
            src: element.src ? "<blob>" : "",
            alt: element.alt,
            loading: element.loading,
            decoding: element.decoding,
          }
        : undefined,
    span:
      element.tagName === "td"
        ? { columns: element.colSpan, rows: element.rowSpan }
        : undefined,
    children: element.children.map(serializeElement),
  };
}
