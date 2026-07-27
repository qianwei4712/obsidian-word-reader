import { App, TFile } from "obsidian";
import { getWordReaderText, type WordReaderText } from "../i18n";
import { createOrOpenReaderSummaryNote } from "../reader/summaryNoteFile";
import { buildSummaryNote } from "../docx/docxSummaryNote";

export async function createNoteFromDocx(
  app: App,
  file: TFile,
  text: WordReaderText = getWordReaderText("zh-CN"),
): Promise<void> {
  await createOrOpenReaderSummaryNote(
    app,
    file,
    ".docx",
    buildSummaryNote(file, text),
    {
      openedExisting: text.notices.openedExistingSummaryNote,
      created: text.notices.createdSummaryNote,
    },
  );
}
