import assert from "node:assert/strict";
import test from "node:test";

import { getPptxReaderText } from "../src/pptx/pptxI18n";
import { buildPresentationSummaryNote } from "../src/pptx/pptxSummaryNote";

void test("presentation notes include current and complete slide references", () => {
  const note = buildPresentationSummaryNote(
    {
      basename: "Quarterly review",
      path: "Presentations/Quarterly review.pptx",
    },
    [
      { index: 0, title: "Overview", text: "", notes: "" },
      { index: 1, title: "Results", text: "", notes: "" },
    ],
    1,
    getPptxReaderText("en"),
    new Date(2026, 5, 13),
  );

  assert.match(note, /type: presentation-note/);
  assert.match(note, /current_slide: 2/);
  assert.match(note, /Current slide: Slide 2 - Results/);
  assert.match(note, /- Slide 1 - Overview/);
  assert.match(note, /- Slide 2 - Results/);
});
