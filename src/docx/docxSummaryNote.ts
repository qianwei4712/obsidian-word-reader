import type { WordReaderText } from "../i18n";
import {
  buildReaderSummaryFrontmatter,
  type ReaderSummarySource,
} from "../reader/summaryNote";

export type SummaryNoteSource = ReaderSummarySource;

export function buildSummaryNote(
  source: SummaryNoteSource,
  text: WordReaderText,
  createdAt: Date = new Date(),
): string {
  const noteText = text.summaryNote;

  return [
    ...buildReaderSummaryFrontmatter(source, {
      format: "docx",
      legacyType: "word-note",
      createdAt,
    }),
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
