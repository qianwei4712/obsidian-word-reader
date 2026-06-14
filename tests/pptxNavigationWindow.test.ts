import assert from "node:assert/strict";
import test from "node:test";

import { PptxNavigationWindow } from "../src/pptx/pptxNavigationWindow";

void test("PptxNavigationWindow bounds mounted rows for large decks", () => {
  const indices = Array.from({ length: 1_000 }, (_, index) => index);
  const window = new PptxNavigationWindow(200, 5, 60);
  const result = window.calculate(indices, 100_000, 1_000);

  assert.ok(result.indices.length <= 60);
  assert.ok(result.start > 0);
  assert.ok(result.end < 1_000);
  assert.equal(
    result.topSpacer +
      result.indices.length * window.rowHeight +
      result.bottomSpacer,
    indices.length * window.rowHeight,
  );
  assert.equal(window.scrollTopForIndex(indices, 500), 100_000);
});
