import assert from "node:assert/strict";
import test from "node:test";

import {
  ReaderStatusController,
  type ReaderStatus,
} from "../src/reader/status";

void test("ReaderStatusController preserves and refreshes the latest status", () => {
  const updates: ReaderStatus[] = [];
  const controller = new ReaderStatusController((status) => {
    updates.push(status);
  });

  controller.set("Reading report...", "loading");
  controller.refresh();

  assert.deepEqual(controller.value, {
    message: "Reading report...",
    kind: "loading",
  });
  assert.deepEqual(updates, [controller.value, controller.value]);
});
