import {
  buildReaderSummaryFrontmatter,
  type ReaderSummarySource,
} from "../reader/summaryNote";
import type { XlsxReaderText } from "./xlsxI18n";
import type {
  XlsxDefinedName,
  XlsxSheetDescriptor,
} from "./xlsxTypes";
import type { XlsxWorksheet } from "./xlsxWorksheet";

export const MAX_XLSX_SUMMARY_PREVIEW_CELLS = 200;
const PREVIEW_CELLS_PER_SHEET = 12;

export interface XlsxSheetSummary {
  name: string;
  state: XlsxSheetDescriptor["state"];
  rowCount: number;
  columnCount: number;
  populatedCellCount: number;
  commentCount: number;
  imageCount: number;
  chartCount: number;
  conditionalFormattingRuleCount: number;
  preview: readonly {
    reference: string;
    displayValue: string;
  }[];
}

export interface XlsxWorkbookSummary {
  sheets: readonly XlsxSheetSummary[];
  definedNames: readonly XlsxDefinedName[];
}

export interface XlsxSummarySource {
  sheets: readonly XlsxSheetDescriptor[];
  definedNames: readonly XlsxDefinedName[];
  getWorksheet: (
    index: number,
    options?: { isCancelled?: () => boolean },
  ) => Promise<XlsxWorksheet>;
}

export interface XlsxSummaryCollectionOptions {
  isCancelled?: () => boolean;
  yieldControl?: () => Promise<void>;
  onSheet?: (
    sheetName: string,
    sheetIndex: number,
    sheetCount: number,
  ) => void;
}

export class XlsxSummaryCancelledError extends Error {
  constructor() {
    super("The XLSX summary was cancelled.");
    this.name = "XlsxSummaryCancelledError";
  }
}

export async function collectXlsxWorkbookSummary(
  workbook: XlsxSummarySource,
  options: XlsxSummaryCollectionOptions = {},
): Promise<XlsxWorkbookSummary> {
  const sheets: XlsxSheetSummary[] = [];
  let remainingPreviewCells = MAX_XLSX_SUMMARY_PREVIEW_CELLS;
  for (let index = 0; index < workbook.sheets.length; index += 1) {
    throwIfCancelled(options);
    const descriptor = workbook.sheets[index];
    options.onSheet?.(
      descriptor.name,
      index,
      workbook.sheets.length,
    );
    const worksheet = await workbook.getWorksheet(index, {
      isCancelled: options.isCancelled,
    });
    throwIfCancelled(options);
    const previewLimit = Math.min(
      PREVIEW_CELLS_PER_SHEET,
      remainingPreviewCells,
    );
    const preview = worksheet
      .getPopulatedCells(previewLimit)
      .map((cell) => ({
        reference: cell.ref,
        displayValue: cell.displayValue,
      }));
    remainingPreviewCells -= preview.length;
    sheets.push({
      name: descriptor.name,
      state: descriptor.state,
      rowCount: worksheet.rowCount,
      columnCount: worksheet.columnCount,
      populatedCellCount: worksheet.populatedCellCount,
      commentCount: worksheet.comments.length,
      imageCount: worksheet.images.length,
      chartCount: worksheet.charts.length,
      conditionalFormattingRuleCount:
        worksheet.conditionalFormattingRules.length,
      preview,
    });
    if (
      index + 1 < workbook.sheets.length &&
      options.yieldControl
    ) {
      await options.yieldControl();
    }
  }
  throwIfCancelled(options);
  return {
    sheets,
    definedNames: workbook.definedNames,
  };
}

export function buildXlsxSummaryNote(
  source: ReaderSummarySource,
  workbook: XlsxWorkbookSummary,
  currentSheetIndex: number,
  selectedRange: string,
  selectedRangeMarkdown: string,
  text: XlsxReaderText,
  createdAt: Date = new Date(),
): string {
  const noteText = text.summaryNote;
  const currentSheet = workbook.sheets[currentSheetIndex];
  const lines = [
    ...buildReaderSummaryFrontmatter(source, {
      format: "xlsx",
      legacyType: "spreadsheet-note",
      createdAt,
      fields: {
        current_sheet: currentSheet?.name ?? "",
        selected_range: selectedRange,
        sheet_count: workbook.sheets.length,
        named_range_count: workbook.definedNames.length,
        comment_count: workbook.sheets.reduce(
          (sum, sheet) => sum + sheet.commentCount,
          0,
        ),
        image_count: workbook.sheets.reduce(
          (sum, sheet) => sum + sheet.imageCount,
          0,
        ),
        chart_count: workbook.sheets.reduce(
          (sum, sheet) => sum + sheet.chartCount,
          0,
        ),
      },
    }),
    "",
    `# ${singleLine(source.basename)}`,
    "",
    `${noteText.sourceLabel}: [[${source.path}]]`,
    `${noteText.currentSheetLabel}: ${singleLine(currentSheet?.name ?? "")}`,
    `${noteText.selectedRangeLabel}: \`${escapeCode(selectedRange)}\``,
    "",
    `## ${noteText.workbookOverviewHeading}`,
    "",
    ...workbook.sheets.map((sheet) =>
      `- ${noteText.sheetSummary(
        singleLine(sheet.name),
        noteText.visibility[sheet.state],
        sheet.rowCount,
        sheet.columnCount,
        sheet.populatedCellCount,
        sheet.commentCount,
        sheet.imageCount,
        sheet.chartCount,
        sheet.conditionalFormattingRuleCount,
      )}`),
    "",
    `## ${noteText.namedRangesHeading}`,
    "",
    ...(workbook.definedNames.length > 0
      ? workbook.definedNames.map((name) =>
          `- \`${escapeCode(name.name)}\` → \`${escapeCode(name.target)}\``)
      : [noteText.noNamedRanges]),
    "",
    `## ${noteText.selectedRangeHeading}`,
    "",
    selectedRangeMarkdown || noteText.noPreview,
    "",
    `## ${noteText.worksheetPreviewsHeading}`,
    "",
  ];

  for (const sheet of workbook.sheets) {
    lines.push(
      `### ${singleLine(sheet.name)} · ${noteText.visibility[sheet.state]}`,
      "",
    );
    if (sheet.preview.length === 0) {
      lines.push(noteText.noPreview, "");
      continue;
    }
    lines.push(
      `| ${noteText.cellHeader} | ${noteText.displayedValueHeader} |`,
      "| --- | --- |",
      ...sheet.preview.map(
        (cell) =>
          `| ${escapeMarkdownCell(cell.reference)} | ${escapeMarkdownCell(cell.displayValue)} |`,
      ),
      "",
    );
  }

  lines.push(
    `## ${noteText.keyFindingsHeading}`,
    "",
    `## ${noteText.followUpsHeading}`,
    "",
  );
  return lines.join("\n");
}

function throwIfCancelled(options: XlsxSummaryCollectionOptions): void {
  if (options.isCancelled?.()) {
    throw new XlsxSummaryCancelledError();
  }
}

function singleLine(value: string): string {
  return value.replace(/\r\n|\r|\n/g, " ").trim();
}

function escapeCode(value: string): string {
  return singleLine(value).replace(/`/g, "'");
}

function escapeMarkdownCell(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\|/g, "\\|")
    .replace(/\r\n|\r|\n/g, "<br>");
}
