import assert from "node:assert/strict";
import test from "node:test";

import { PptxThumbnailWindow } from "../src/pptx/pptxThumbnailWindow";

void test("thumbnail window mounts visible slides with bounded overscan", () => {
  const window = new PptxThumbnailWindow(20, 1, 6);
  window.setVisible(5, true);
  window.setVisible(6, true);

  assert.deepEqual(window.update(5), {
    mount: [4, 5, 6, 7],
    unmount: [],
    mounted: [4, 5, 6, 7],
  });

  window.setVisible(5, false);
  window.setVisible(6, false);
  window.setVisible(12, true);
  assert.deepEqual(window.update(12), {
    mount: [11, 12, 13],
    unmount: [4, 5, 6, 7],
    mounted: [11, 12, 13],
  });
});

void test("thumbnail window prioritizes the current slide and resets cleanly", () => {
  const window = new PptxThumbnailWindow(4);
  assert.deepEqual(window.update(0).mounted, [0, 1]);
  assert.deepEqual(window.update(3), {
    mount: [2, 3],
    unmount: [0, 1],
    mounted: [2, 3],
  });
  assert.deepEqual(window.reset(), {
    mount: [],
    unmount: [2, 3],
    mounted: [],
  });
});

void test("thumbnail window keeps the priority slide when visible range exceeds the limit", () => {
  const window = new PptxThumbnailWindow(20, 1, 4);
  for (let index = 2; index <= 12; index += 1) {
    window.setVisible(index, true);
  }

  assert.deepEqual(window.update(10).mounted, [8, 9, 10, 11]);
});
