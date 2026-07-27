import assert from "node:assert/strict";
import test from "node:test";

import {
  DEFAULT_OFFICE_READER_SETTINGS,
  DEFAULT_SETTINGS,
  SETTINGS_SCHEMA_VERSION,
  migrateSettings,
  normalizeOfficeReaderSettings,
  normalizeSettings,
} from "../src/settingsModel";

void test("normalizeSettings returns defaults for missing or invalid input", () => {
  assert.deepEqual(normalizeSettings(null), DEFAULT_SETTINGS);
  assert.deepEqual(normalizeSettings("invalid"), DEFAULT_SETTINGS);
});

void test("migrateSettings converts flat data to the versioned format schema", () => {
  const migrated = migrateSettings({
    language: "en",
    defaultZoomPercent: 135,
    defaultFitWidth: true,
    showOutlineByDefault: false,
    enableImagePreview: false,
    largeFileWarningMb: 64,
    readingStates: [{ path: "ignored-by-settings.docx" }],
  });

  assert.deepEqual(migrated, {
    schemaVersion: SETTINGS_SCHEMA_VERSION,
    common: {
      language: "en",
      defaultZoomPercent: 135,
      largeFileWarningMb: 64,
    },
    docx: {
      defaultFitWidth: true,
      showOutlineByDefault: false,
      enableImagePreview: false,
    },
    pptx: DEFAULT_OFFICE_READER_SETTINGS.pptx,
    xlsx: DEFAULT_OFFICE_READER_SETTINGS.xlsx,
  });
  assert.deepEqual(normalizeOfficeReaderSettings(migrated), migrated);
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
