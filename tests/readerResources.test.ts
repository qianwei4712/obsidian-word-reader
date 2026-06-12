import assert from "node:assert/strict";
import test from "node:test";

import {
  RetainedResourceRegistry,
  releaseResources,
} from "../src/reader/resources";

void test("RetainedResourceRegistry delays retained resource release", () => {
  const released: string[] = [];
  const registry = new RetainedResourceRegistry((resource) => {
    released.push(resource);
  });
  registry.replace(["first", "second"]);
  const releaseFirst = registry.retain("first");

  registry.releaseActive();
  assert.deepEqual(released, ["second"]);

  releaseFirst();
  assert.deepEqual(released, ["second", "first"]);
});

void test("releaseResources releases every supplied resource", () => {
  const released: string[] = [];
  releaseResources(["a", "b"], (resource) => {
    released.push(resource);
  });
  assert.deepEqual(released, ["a", "b"]);
});

void test("replacing an active retained resource does not release it early", () => {
  const released: string[] = [];
  const registry = new RetainedResourceRegistry((resource) => {
    released.push(resource);
  });
  registry.replace(["shared"]);
  const releaseShared = registry.retain("shared");

  registry.replace(["shared"]);
  releaseShared();
  assert.deepEqual(released, []);

  registry.releaseActive();
  assert.deepEqual(released, ["shared"]);
});
