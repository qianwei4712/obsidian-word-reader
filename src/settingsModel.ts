import {
  normalizeLanguage,
  type WordReaderLanguage,
} from "./i18n";

export type { WordReaderLanguage };

export interface WordReaderSettings {
  language: WordReaderLanguage;
  defaultZoomPercent: number;
  defaultFitWidth: boolean;
  showOutlineByDefault: boolean;
  enableImagePreview: boolean;
  largeFileWarningMb: number;
}

export const DEFAULT_SETTINGS: WordReaderSettings = {
  language: "zh-CN",
  defaultZoomPercent: 100,
  defaultFitWidth: false,
  showOutlineByDefault: true,
  enableImagePreview: true,
  largeFileWarningMb: 25,
};

const MIN_ZOOM_PERCENT = 25;
const MAX_ZOOM_PERCENT = 400;
const MIN_LARGE_FILE_WARNING_MB = 1;
const MAX_LARGE_FILE_WARNING_MB = 500;

export function normalizeSettings(settings: unknown): WordReaderSettings {
  const source = isRecord(settings) ? settings : {};
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
