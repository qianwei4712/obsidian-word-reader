import { App, TFile } from "obsidian";

import { createOrOpenReaderSummaryNote } from "../reader/summaryNoteFile";
import type { XlsxReaderText } from "../xlsx/xlsxI18n";
import {
  buildXlsxSummaryNote,
  type XlsxWorkbookSummary,
} from "../xlsx/xlsxSummaryNote";

export async function createNoteFromXlsx(
  app: App,
  file: TFile,
  workbook: XlsxWorkbookSummary,
  currentSheetIndex: number,
  selectedRange: string,
  selectedRangeMarkdown: string,
  text: XlsxReaderText,
): Promise<void> {
  await createOrOpenReaderSummaryNote(
    app,
    file,
    ".xlsx",
    buildXlsxSummaryNote(
      file,
      workbook,
      currentSheetIndex,
      selectedRange,
      selectedRangeMarkdown,
      text,
    ),
    {
      openedExisting: text.notices.openedExistingSummaryNote,
      created: text.notices.createdSummaryNote,
    },
  );
}
