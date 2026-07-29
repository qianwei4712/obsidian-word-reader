import assert from "node:assert/strict";
import test from "node:test";

import {
  DEFAULT_READING_STATE_CAPACITY,
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

void test("ReadingStateStore defaults to a 100-file LRU", () => {
  const store = new ReadingStateStore();
  for (let index = 0; index < 101; index += 1) {
    store.set(`document-${index}.docx`, state);
  }
  assert.equal(DEFAULT_READING_STATE_CAPACITY, 100);
  assert.equal(store.size, 100);
  assert.equal(store.get("document-0.docx"), undefined);
  assert.equal(store.get("document-100.docx")?.zoom, state.zoom);
});

void test("ReadingStateStore serializes only identity, position, zoom, and navigation", () => {
  const store = new ReadingStateStore();
  store.set(
    {
      path: "deck.pptx",
      mtime: 123,
      format: "pptx",
    },
    {
      ...state,
      page: 7,
      notesVisible: true,
    },
  );
  const entry = store.serialize()[0];
  assert.deepEqual(Object.keys(entry).sort(), [
    "format",
    "lastAccessed",
    "mtime",
    "navigation",
    "path",
    "position",
    "zoom",
  ]);
  assert.equal(JSON.stringify(entry).includes("text"), false);
  assert.equal(entry.format, "pptx");
  assert.equal(entry.mtime, 123);
});

void test("mtime changes preserve zoom but invalidate position and navigation", () => {
  const store = new ReadingStateStore();
  store.set(
    {
      path: "changed.docx",
      mtime: 100,
      format: "docx",
    },
    state,
  );
  const restored = store.get({
    path: "changed.docx",
    mtime: 101,
    format: "docx",
  });
  assert.equal(restored?.zoom, state.zoom);
  assert.equal(restored?.fitWidth, state.fitWidth);
  assert.equal(restored?.scrollTop, 0);
  assert.equal(restored?.outlineVisible, true);
  assert.deepEqual(restored?.collapsedOutlineIds, []);
});

void test("legacy states adopt their first known mtime without losing position", () => {
  const store = new ReadingStateStore(100, [
    {
      path: "legacy.docx",
      lastAccessed: 1,
      state,
    },
  ]);
  const restored = store.get({
    path: "legacy.docx",
    mtime: 456,
    format: "docx",
  });
  assert.deepEqual(restored, state);
  assert.equal(store.serialize()[0].mtime, 456);
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

void test("ReadingStateStore preserves the active XLSX sheet without cell data", () => {
  const store = new ReadingStateStore();
  store.set(
    {
      path: "budget.xlsx",
      mtime: 800,
      format: "xlsx",
    },
    {
      ...state,
      outlineVisible: false,
      page: 3,
      scrollLeft: 640,
      scrollTop: 2_400,
    },
  );
  const serialized = store.serialize()[0];
  assert.equal(serialized.format, "xlsx");
  assert.equal(serialized.position.page, 3);
  assert.equal(serialized.position.scrollLeft, 640);
  assert.equal(serialized.position.scrollTop, 2_400);
  assert.equal(JSON.stringify(serialized).includes("cell"), false);
  assert.equal(JSON.stringify(serialized).includes("formula"), false);
});
