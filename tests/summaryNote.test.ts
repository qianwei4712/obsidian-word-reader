import assert from "node:assert/strict";
import test from "node:test";

import { getWordReaderText } from "../src/i18n";
import { buildSummaryNote } from "../src/summaryNote";

void test("buildSummaryNote creates a deterministic localized note", () => {
  const note = buildSummaryNote(
    {
      basename: "Quarterly report",
      path: "Reports/Quarterly report.docx",
    },
    getWordReaderText("en"),
    new Date(2026, 5, 7),
  );

  assert.equal(
    note,
    [
      "---",
      'source: "Reports/Quarterly report.docx"',
      "type: word-note",
      "created: 2026-06-07",
      "---",
      "",
      "# Quarterly report",
      "",
      "Source: [[Reports/Quarterly report.docx]]",
      "",
      "## Summary",
      "",
      "## Key findings",
      "",
      "## Follow-ups",
      "",
      "## Quoted excerpts",
      "",
    ].join("\n"),
  );
});

void test("buildSummaryNote escapes YAML paths without changing the wiki link", () => {
  const note = buildSummaryNote(
    {
      basename: "Quoted",
      path: 'Folder\\A "quoted".docx',
    },
    getWordReaderText("zh-CN"),
    new Date(2026, 0, 2),
  );

  assert.match(note, /source: "Folder\\\\A \\"quoted\\"\.docx"/);
  assert.match(note, /原文: \[\[Folder\\A "quoted"\.docx\]\]/);
});
