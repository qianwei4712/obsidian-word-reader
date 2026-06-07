import type { WordReaderText } from "./i18n";

export interface SummaryNoteSource {
  basename: string;
  path: string;
}

export function buildSummaryNote(
  source: SummaryNoteSource,
  text: WordReaderText,
  createdAt: Date = new Date(),
): string {
  const noteText = text.summaryNote;

  return [
    "---",
    `source: "${escapeYamlString(source.path)}"`,
    "type: word-note",
    `created: ${formatLocalDate(createdAt)}`,
    "---",
    "",
    `# ${source.basename}`,
    "",
    `${noteText.sourceLabel}: [[${source.path}]]`,
    "",
    `## ${noteText.summaryHeading}`,
    "",
    `## ${noteText.keyFindingsHeading}`,
    "",
    `## ${noteText.followUpsHeading}`,
    "",
    `## ${noteText.quotesHeading}`,
    "",
  ].join("\n");
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
