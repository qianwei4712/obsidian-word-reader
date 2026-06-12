import assert from "node:assert/strict";
import test from "node:test";

import {
  normalizeZoom,
  preserveZoomAnchor,
} from "../src/reader/zoom";

void test("normalizeZoom rounds and clamps zoom values", () => {
  const range = { min: 0.25, max: 4, step: 0.05 };
  assert.equal(normalizeZoom(1.234, range), 1.25);
  assert.equal(normalizeZoom(8, range), 4);
  assert.equal(normalizeZoom(Number.NaN, range), null);
});

void test("preserveZoomAnchor keeps the pointed content position stable", () => {
  assert.deepEqual(
    preserveZoomAnchor(
      { left: 100, top: 200 },
      { left: 50, top: 40 },
      1,
      2,
    ),
    { left: 250, top: 440 },
  );
});
