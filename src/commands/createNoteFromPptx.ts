import { App, Notice, TFile, normalizePath } from "obsidian";

import type { PptxSlideMetadata } from "../pptx/pptxMetadata";
import type { PptxReaderText } from "../pptx/pptxI18n";
import { buildPresentationSummaryNote } from "../pptx/pptxSummaryNote";

export async function createNoteFromPptx(
  app: App,
  file: TFile,
  slides: readonly PptxSlideMetadata[],
  currentSlideIndex: number,
  text: PptxReaderText,
): Promise<void> {
  const notePath = normalizePath(file.path.replace(/\.pptx$/i, ".md"));
  const existing = app.vault.getAbstractFileByPath(notePath);
  if (existing instanceof TFile) {
    await app.workspace.getLeaf(true).openFile(existing);
    new Notice(text.notices.openedExistingSummaryNote);
    return;
  }

  const content = buildPresentationSummaryNote(
    file,
    slides,
    currentSlideIndex,
    text,
  );
  const created = await app.vault.create(notePath, content);
  await app.workspace.getLeaf(true).openFile(created);
  new Notice(text.notices.createdSummaryNote);
}
