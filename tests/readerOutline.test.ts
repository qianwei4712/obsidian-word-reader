import assert from "node:assert/strict";
import test from "node:test";

import {
  buildReaderOutline,
  getVisibleOutlineId,
} from "../src/reader/outline";

void test("buildReaderOutline tracks hierarchy and collapsible parents", () => {
  const outline = buildReaderOutline([
    { element: "intro", level: 1, text: "Intro" },
    { element: "detail", level: 2, text: "Detail" },
    { element: "deep", level: 3, text: "Deep" },
    { element: "next", level: 1, text: "Next" },
  ]);

  assert.equal(outline[0].hasChildren, true);
  assert.equal(outline[1].hasChildren, true);
  assert.deepEqual(outline[2].ancestorIds, [outline[0].id, outline[1].id]);
  assert.deepEqual(outline[3].ancestorIds, []);
  assert.equal(
    getVisibleOutlineId(outline[2], new Set([outline[0].id])),
    outline[0].id,
  );
  assert.match(outline[0].id, /^0:1:[0-9a-f]{8}$/);
});
