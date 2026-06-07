import assert from "node:assert/strict";
import test from "node:test";

import {
  DEFAULT_SETTINGS,
  normalizeSettings,
} from "../src/settingsModel";

void test("normalizeSettings returns defaults for missing or invalid input", () => {
  assert.deepEqual(normalizeSettings(null), DEFAULT_SETTINGS);
  assert.deepEqual(normalizeSettings("invalid"), DEFAULT_SETTINGS);
});

void test("normalizeSettings preserves supported values", () => {
  assert.deepEqual(
    normalizeSettings({
      language: "en",
      defaultZoomPercent: 135,
      defaultFitWidth: true,
      showOutlineByDefault: false,
      enableImagePreview: false,
      largeFileWarningMb: 64,
    }),
    {
      language: "en",
      defaultZoomPercent: 135,
      defaultFitWidth: true,
      showOutlineByDefault: false,
      enableImagePreview: false,
      largeFileWarningMb: 64,
    },
  );
});

void test("normalizeSettings clamps, rounds, and rejects invalid field types", () => {
  assert.deepEqual(
    normalizeSettings({
      language: "fr",
      defaultZoomPercent: 999,
      defaultFitWidth: "true",
      showOutlineByDefault: 0,
      enableImagePreview: true,
      largeFileWarningMb: 0.6,
    }),
    {
      language: "zh-CN",
      defaultZoomPercent: 400,
      defaultFitWidth: false,
      showOutlineByDefault: true,
      enableImagePreview: true,
      largeFileWarningMb: 1,
    },
  );

  assert.equal(
    normalizeSettings({ defaultZoomPercent: Number.NaN }).defaultZoomPercent,
    DEFAULT_SETTINGS.defaultZoomPercent,
  );
  assert.equal(
    normalizeSettings({ largeFileWarningMb: Number.POSITIVE_INFINITY })
      .largeFileWarningMb,
    DEFAULT_SETTINGS.largeFileWarningMb,
  );
});
