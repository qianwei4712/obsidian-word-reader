import assert from "node:assert/strict";
import test from "node:test";

import {
  formatSlideText,
  searchPptxSlides,
  type PptxSlideMetadata,
} from "../src/pptx/pptxMetadata";

const slides: PptxSlideMetadata[] = [
  {
    index: 0,
    title: "Quarterly results",
    text: "Revenue increased by 20 percent.",
    notes: "Explain the regional breakdown.",
  },
  {
    index: 1,
    title: "Next steps",
    text: "Launch the new product.",
    notes: "Revenue risk should be mentioned.",
  },
];

void test("searchPptxSlides searches visible text and speaker notes", () => {
  assert.deepEqual(
    searchPptxSlides(slides, "revenue").map((result) => ({
      slideIndex: result.slideIndex,
      matchCount: result.matchCount,
      matchedNotes: result.matchedNotes,
    })),
    [
      { slideIndex: 0, matchCount: 1, matchedNotes: false },
      { slideIndex: 1, matchCount: 1, matchedNotes: true },
    ],
  );
  assert.equal(searchPptxSlides(slides, "quarterly")[0]?.matchCount, 1);
  assert.equal(searchPptxSlides(slides, "missing").length, 0);
  assert.equal(searchPptxSlides(slides, "").length, 2);
});

void test("formatSlideText includes title and body without speaker notes", () => {
  assert.equal(
    formatSlideText(slides[0]),
    [
      "Quarterly results",
      "",
      "Revenue increased by 20 percent.",
    ].join("\n"),
  );
});
