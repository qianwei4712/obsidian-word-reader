import { App, Notice, TFile, normalizePath } from "obsidian";
import { getWordReaderText, type WordReaderText } from "../i18n";

export async function createNoteFromDocx(
  app: App,
  file: TFile,
  text: WordReaderText = getWordReaderText("zh-CN"),
): Promise<void> {
  const notePath = normalizePath(
    file.path.replace(/\.docx$/i, ".md"),
  );

  const existing = app.vault.getAbstractFileByPath(notePath);
  if (existing instanceof TFile) {
    await app.workspace.getLeaf(true).openFile(existing);
    new Notice(text.notices.openedExistingSummaryNote);
    return;
  }

  const content = buildSummaryNote(file, text);
  const created = await app.vault.create(notePath, content);
  await app.workspace.getLeaf(true).openFile(created);
  new Notice(text.notices.createdSummaryNote);
}

function buildSummaryNote(file: TFile, text: WordReaderText): string {
  const title = file.basename;
  const sourcePath = file.path;
  const noteText = text.summaryNote;

  return [
    "---",
    `source: "${escapeYamlString(sourcePath)}"`,
    "type: word-note",
    `created: ${formatLocalDate(new Date())}`,
    "---",
    "",
    `# ${title}`,
    "",
    `${noteText.sourceLabel}: [[${sourcePath}]]`,
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
