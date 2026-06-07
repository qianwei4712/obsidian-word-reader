import { App, Notice, TFile, normalizePath } from "obsidian";
import { getWordReaderText, type WordReaderText } from "../i18n";
import { buildSummaryNote } from "../summaryNote";

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
