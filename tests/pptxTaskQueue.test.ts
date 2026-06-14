import assert from "node:assert/strict";
import test from "node:test";
import { setTimeout as delay } from "node:timers/promises";

import { PptxTaskQueue } from "../src/pptx/pptxTaskQueue";

void test("PptxTaskQueue limits concurrency and prioritizes queued work", async () => {
  const queue = new PptxTaskQueue(2);
  const started: number[] = [];
  const releases: Array<() => void> = [];
  let active = 0;
  let peak = 0;
  const schedule = (key: number, priority: number): void => {
    queue.schedule(key, priority, async () => {
      started.push(key);
      active += 1;
      peak = Math.max(peak, active);
      await new Promise<void>((resolve) => {
        releases.push(resolve);
      });
      active -= 1;
    });
  };

  schedule(1, 10);
  schedule(2, 20);
  schedule(3, 30);
  schedule(4, 5);
  await Promise.resolve();
  assert.deepEqual(started, [1, 2]);
  assert.equal(peak, 2);

  releases.shift()?.();
  await delay(0);
  assert.deepEqual(started, [1, 2, 4]);

  while (releases.length > 0) {
    releases.shift()?.();
    await delay(0);
  }
  assert.ok(peak <= 2);
});
