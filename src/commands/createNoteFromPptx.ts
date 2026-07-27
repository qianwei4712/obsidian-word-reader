import { App, TFile } from "obsidian";

import type { PptxSlideMetadata } from "../pptx/pptxMetadata";
import type { PptxReaderText } from "../pptx/pptxI18n";
import { buildPresentationSummaryNote } from "../pptx/pptxSummaryNote";
import { createOrOpenReaderSummaryNote } from "../reader/summaryNoteFile";

export async function createNoteFromPptx(
  app: App,
  file: TFile,
  slides: readonly PptxSlideMetadata[],
  currentSlideIndex: number,
  text: PptxReaderText,
): Promise<void> {
  const content = buildPresentationSummaryNote(
    file,
    slides,
    currentSlideIndex,
    text,
  );
  await createOrOpenReaderSummaryNote(
    app,
    file,
    ".pptx",
    content,
    {
      openedExisting: text.notices.openedExistingSummaryNote,
      created: text.notices.createdSummaryNote,
    },
  );
}
