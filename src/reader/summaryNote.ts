import type { ReaderFormat } from "./readingState";

export interface ReaderSummarySource {
  basename: string;
  path: string;
}

export interface ReaderSummaryFrontmatterOptions {
  format: ReaderFormat;
  legacyType: string;
  createdAt?: Date;
  fields?: Readonly<Record<string, string | number>>;
}

export function buildReaderSummaryFrontmatter(
  source: ReaderSummarySource,
  options: ReaderSummaryFrontmatterOptions,
): string[] {
  const fields = options.fields ?? {};
  return [
    "---",
    `source: "${escapeYamlString(source.path)}"`,
    `type: ${options.legacyType}`,
    "reader: office-reader",
    `reader_format: ${options.format}`,
    `created: ${formatLocalDate(options.createdAt ?? new Date())}`,
    ...Object.entries(fields).map(
      ([key, value]) => `${key}: ${formatYamlScalar(value)}`,
    ),
    "---",
  ];
}

function formatLocalDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function escapeYamlString(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function formatYamlScalar(value: string | number): string {
  return typeof value === "number"
    ? String(value)
    : `"${escapeYamlString(value)}"`;
}
