import assert from "node:assert/strict";
import test from "node:test";

import { validateDocxPackage } from "../src/docx/DocxAdapter";
import { OoxmlPolicyError } from "../src/ooxml/packagePolicy";
import { addOoxmlEntry } from "./xlsxFixture";
import { createLargeDocx } from "./performanceFixtures";

void test("DocxAdapter applies the shared OOXML safety gate", async () => {
  const buffer = await createLargeDocx(2);
  await assert.doesNotReject(() => validateDocxPackage(buffer));
});

void test("DocxAdapter rejects active, OLE, and script-capable parts", async () => {
  const base = await createLargeDocx(1);
  const cases: Array<[string, OoxmlPolicyError["kind"]]> = [
    ["word/vbaProject.bin", "active-content"],
    ["word/embeddings/oleObject1.bin", "ole-object"],
    ["word/media/remote.svg", "script-media"],
  ];
  for (const [path, kind] of cases) {
    const buffer = await addOoxmlEntry(base, path, Uint8Array.from([1]));
    await assert.rejects(
      () => validateDocxPackage(buffer),
      (error: unknown) =>
        error instanceof OoxmlPolicyError && error.kind === kind,
    );
  }
});
