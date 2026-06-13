import assert from "node:assert/strict";
import test from "node:test";

import {
  ReadingStateStore,
  normalizeReaderViewState,
} from "../src/reader/readingState";

const state = {
  zoom: 1.25,
  fitWidth: true,
  outlineVisible: false,
  scrollLeft: 12,
  scrollTop: 345,
  collapsedOutlineIds: ["0:1:Intro"],
};

void test("ReadingStateStore restores state and enforces LRU capacity", () => {
  const store = new ReadingStateStore(2);
  store.set("a.docx", state);
  store.set("b.docx", { ...state, scrollTop: 20 });
  assert.equal(store.get("a.docx")?.scrollTop, 345);

  store.set("c.docx", { ...state, scrollTop: 30 });

  assert.equal(store.size, 2);
  assert.equal(store.get("b.docx"), undefined);
  assert.equal(store.get("a.docx")?.zoom, 1.25);
  assert.equal(store.get("c.docx")?.scrollTop, 30);
});

void test("ReadingStateStore serializes and normalizes persisted input", () => {
  const original = new ReadingStateStore(3);
  original.set("report.docx", state);
  const restored = new ReadingStateStore(3, original.serialize());

  assert.deepEqual(restored.get("report.docx"), state);
  assert.deepEqual(normalizeReaderViewState(null), {
    zoom: 1,
    fitWidth: false,
    outlineVisible: true,
    scrollLeft: 0,
    scrollTop: 0,
    collapsedOutlineIds: [],
  });
});

void test("ReadingStateStore preserves a normalized presentation page", () => {
  const normalized = normalizeReaderViewState({
    ...state,
    page: 4.8,
  });
  assert.equal(normalized.page, 4);
});

void test("ReadingStateStore preserves presentation panel visibility", () => {
  const normalized = normalizeReaderViewState({
    ...state,
    notesVisible: true,
  });
  assert.equal(normalized.notesVisible, true);
});
