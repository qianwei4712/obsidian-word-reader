import assert from "node:assert/strict";
import test from "node:test";

import {
  getWordReaderText,
  normalizeLanguage,
  type WordReaderLanguage,
} from "../src/i18n";

void test("normalizeLanguage accepts supported languages and falls back to Chinese", () => {
  assert.equal(normalizeLanguage("en"), "en");
  assert.equal(normalizeLanguage("zh-CN"), "zh-CN");
  assert.equal(normalizeLanguage("zh"), "zh-CN");
  assert.equal(normalizeLanguage(undefined), "zh-CN");
});

void test("getWordReaderText returns complete localized command labels", () => {
  const english = getWordReaderText("en");
  const chinese = getWordReaderText("zh-CN");
  const fallback = getWordReaderText("invalid" as WordReaderLanguage);

  assert.equal(english.commands.reload, "Reload current Word document");
  assert.equal(chinese.commands.reload, "重新加载当前 Word 文档");
  assert.equal(fallback, chinese);
});
