import assert from "node:assert/strict";
import test from "node:test";

import { ReaderLifecycle } from "../src/reader/lifecycle";

void test("ReaderLifecycle invalidates stale work", () => {
  const lifecycle = new ReaderLifecycle();
  const first = lifecycle.begin();
  const second = lifecycle.begin();

  assert.equal(lifecycle.isCurrent(first), false);
  assert.equal(lifecycle.isCurrent(second), true);

  lifecycle.cancel();
  assert.equal(lifecycle.isCurrent(second), false);
});
