import {
  normalizeLanguage,
  type WordReaderLanguage,
} from "./i18n";

export type { WordReaderLanguage };

export const SETTINGS_SCHEMA_VERSION = 1 as const;

export interface WordReaderSettings {
  language: WordReaderLanguage;
  defaultZoomPercent: number;
  defaultFitWidth: boolean;
  showOutlineByDefault: boolean;
  enableImagePreview: boolean;
  largeFileWarningMb: number;
}

export interface CommonReaderSettings {
  language: WordReaderLanguage;
  defaultZoomPercent: number;
  largeFileWarningMb: number;
}

export interface DocxReaderSettings {
  defaultFitWidth: boolean;
  showOutlineByDefault: boolean;
  enableImagePreview: boolean;
}

export interface PptxReaderSettings {
  defaultFitWindow: boolean;
  showNavigationByDefault: boolean;
  showNotesByDefault: boolean;
}

export interface XlsxReaderSettings {
  defaultFitWidth: boolean;
}

export interface OfficeReaderSettings {
  schemaVersion: typeof SETTINGS_SCHEMA_VERSION;
  common: CommonReaderSettings;
  docx: DocxReaderSettings;
  pptx: PptxReaderSettings;
  xlsx: XlsxReaderSettings;
}

export const DEFAULT_SETTINGS: WordReaderSettings = {
  language: "zh-CN",
  defaultZoomPercent: 100,
  defaultFitWidth: false,
  showOutlineByDefault: true,
  enableImagePreview: true,
  largeFileWarningMb: 25,
};

export const DEFAULT_OFFICE_READER_SETTINGS: OfficeReaderSettings = {
  schemaVersion: SETTINGS_SCHEMA_VERSION,
  common: {
    language: DEFAULT_SETTINGS.language,
    defaultZoomPercent: DEFAULT_SETTINGS.defaultZoomPercent,
    largeFileWarningMb: DEFAULT_SETTINGS.largeFileWarningMb,
  },
  docx: {
    defaultFitWidth: DEFAULT_SETTINGS.defaultFitWidth,
    showOutlineByDefault: DEFAULT_SETTINGS.showOutlineByDefault,
    enableImagePreview: DEFAULT_SETTINGS.enableImagePreview,
  },
  pptx: {
    defaultFitWindow: true,
    showNavigationByDefault: true,
    showNotesByDefault: false,
  },
  xlsx: {
    defaultFitWidth: true,
  },
};

const MIN_ZOOM_PERCENT = 25;
const MAX_ZOOM_PERCENT = 400;
const MIN_LARGE_FILE_WARNING_MB = 1;
const MAX_LARGE_FILE_WARNING_MB = 500;

export function normalizeSettings(settings: unknown): WordReaderSettings {
  const source = getLegacyCompatibleSettingsSource(settings);
  return {
    language: normalizeLanguage(source.language),
    defaultZoomPercent: clampInteger(
      readNumber(source.defaultZoomPercent, DEFAULT_SETTINGS.defaultZoomPercent),
      MIN_ZOOM_PERCENT,
      MAX_ZOOM_PERCENT,
      DEFAULT_SETTINGS.defaultZoomPercent,
    ),
    defaultFitWidth: readBoolean(
      source.defaultFitWidth,
      DEFAULT_SETTINGS.defaultFitWidth,
    ),
    showOutlineByDefault: readBoolean(
      source.showOutlineByDefault,
      DEFAULT_SETTINGS.showOutlineByDefault,
    ),
    enableImagePreview: readBoolean(
      source.enableImagePreview,
      DEFAULT_SETTINGS.enableImagePreview,
    ),
    largeFileWarningMb: clampInteger(
      readNumber(
        source.largeFileWarningMb,
        DEFAULT_SETTINGS.largeFileWarningMb,
      ),
      MIN_LARGE_FILE_WARNING_MB,
      MAX_LARGE_FILE_WARNING_MB,
      DEFAULT_SETTINGS.largeFileWarningMb,
    ),
  };
}

export function normalizeOfficeReaderSettings(
  settings: unknown,
): OfficeReaderSettings {
  const source = isRecord(settings) ? settings : {};
  const common = isRecord(source.common) ? source.common : source;
  const docx = isRecord(source.docx) ? source.docx : source;
  const pptx = isRecord(source.pptx) ? source.pptx : {};
  const xlsx = isRecord(source.xlsx) ? source.xlsx : {};

  return {
    schemaVersion: SETTINGS_SCHEMA_VERSION,
    common: {
      language: normalizeLanguage(common.language),
      defaultZoomPercent: clampInteger(
        readNumber(
          common.defaultZoomPercent,
          DEFAULT_OFFICE_READER_SETTINGS.common.defaultZoomPercent,
        ),
        MIN_ZOOM_PERCENT,
        MAX_ZOOM_PERCENT,
        DEFAULT_OFFICE_READER_SETTINGS.common.defaultZoomPercent,
      ),
      largeFileWarningMb: clampInteger(
        readNumber(
          common.largeFileWarningMb,
          DEFAULT_OFFICE_READER_SETTINGS.common.largeFileWarningMb,
        ),
        MIN_LARGE_FILE_WARNING_MB,
        MAX_LARGE_FILE_WARNING_MB,
        DEFAULT_OFFICE_READER_SETTINGS.common.largeFileWarningMb,
      ),
    },
    docx: {
      defaultFitWidth: readBoolean(
        docx.defaultFitWidth,
        DEFAULT_OFFICE_READER_SETTINGS.docx.defaultFitWidth,
      ),
      showOutlineByDefault: readBoolean(
        docx.showOutlineByDefault,
        DEFAULT_OFFICE_READER_SETTINGS.docx.showOutlineByDefault,
      ),
      enableImagePreview: readBoolean(
        docx.enableImagePreview,
        DEFAULT_OFFICE_READER_SETTINGS.docx.enableImagePreview,
      ),
    },
    pptx: {
      defaultFitWindow: readBoolean(
        pptx.defaultFitWindow,
        DEFAULT_OFFICE_READER_SETTINGS.pptx.defaultFitWindow,
      ),
      showNavigationByDefault: readBoolean(
        pptx.showNavigationByDefault,
        DEFAULT_OFFICE_READER_SETTINGS.pptx.showNavigationByDefault,
      ),
      showNotesByDefault: readBoolean(
        pptx.showNotesByDefault,
        DEFAULT_OFFICE_READER_SETTINGS.pptx.showNotesByDefault,
      ),
    },
    xlsx: {
      defaultFitWidth: readBoolean(
        xlsx.defaultFitWidth,
        DEFAULT_OFFICE_READER_SETTINGS.xlsx.defaultFitWidth,
      ),
    },
  };
}

/**
 * Converts either the pre-2.4 flat settings object or a versioned settings
 * object into the current persisted schema. The function is intentionally
 * idempotent so callers can run it for every plugin load.
 */
export function migrateSettings(settings: unknown): OfficeReaderSettings {
  return normalizeOfficeReaderSettings(settings);
}

export function toLegacySettings(
  settings: OfficeReaderSettings,
): WordReaderSettings {
  return {
    language: settings.common.language,
    defaultZoomPercent: settings.common.defaultZoomPercent,
    defaultFitWidth: settings.docx.defaultFitWidth,
    showOutlineByDefault: settings.docx.showOutlineByDefault,
    enableImagePreview: settings.docx.enableImagePreview,
    largeFileWarningMb: settings.common.largeFileWarningMb,
  };
}

function getLegacyCompatibleSettingsSource(
  settings: unknown,
): Record<string, unknown> {
  if (!isRecord(settings)) {
    return {};
  }
  if (!isRecord(settings.common) && !isRecord(settings.docx)) {
    return settings;
  }
  const common = isRecord(settings.common) ? settings.common : {};
  const docx = isRecord(settings.docx) ? settings.docx : {};

  return {
    ...common,
    ...docx,
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function readNumber(value: unknown, fallback: number): number {
  return typeof value === "number" ? value : fallback;
}

function readBoolean(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function clampInteger(
  value: number,
  min: number,
  max: number,
  fallback: number,
): number {
  if (!Number.isFinite(value)) {
    return fallback;
  }

  return Math.min(Math.max(Math.round(value), min), max);
}
