import type { TFile } from "obsidian";

import {
  createReaderDiagnosticReport,
  fingerprintMessage,
  formatReaderDiagnosticReport,
  type ReaderDiagnosticReport,
} from "../reader/diagnostics";
import type { PptxErrorKind } from "./pptxErrors";
import type { PptxPackage } from "./pptxPackage";
import type { PptxRenderDiagnostics } from "./pptxRenderer";

export interface PptxThumbnailDiagnostics {
  mounted: number;
  rendered: number;
  rendering: number;
  resourceCount: number;
}

export interface PptxDiagnosticDetails {
  presentation: {
    slide: number;
    slideCount: number;
    slideWidthEmu: number;
    slideHeightEmu: number;
    zipFileCount: number;
    zipCompressedBytes: number;
    zipExpandedBytes: number;
    cache: ReturnType<PptxPackage["getCacheDiagnostics"]>;
  };
  render: PptxRenderDiagnostics;
  thumbnails: PptxThumbnailDiagnostics;
}

export interface PptxErrorDiagnosticDetails {
  category: PptxErrorKind;
  fingerprint: string;
}

export type PptxDiagnosticReport =
  ReaderDiagnosticReport<PptxDiagnosticDetails>;
export type PptxErrorDiagnosticReport =
  ReaderDiagnosticReport<PptxErrorDiagnosticDetails>;

export function createPptxDiagnosticReport(
  file: Pick<TFile, "name" | "stat">,
  presentation: PptxPackage,
  slideIndex: number,
  render: PptxRenderDiagnostics,
  thumbnails: PptxThumbnailDiagnostics,
): PptxDiagnosticReport {
  const slide = slideIndex + 1;
  return createReaderDiagnosticReport(
    "pptx",
    "render",
    {
      name: file.name,
      size: file.stat.size,
      mtime: file.stat.mtime,
    },
    `Presentation render metrics for slide ${slide} of ${presentation.slideCount}`,
    {
      presentation: {
        slide,
        slideCount: presentation.slideCount,
        slideWidthEmu: presentation.slideWidth,
        slideHeightEmu: presentation.slideHeight,
        zipFileCount: presentation.zipSummary.fileCount,
        zipCompressedBytes: presentation.zipSummary.totalCompressedBytes,
        zipExpandedBytes: presentation.zipSummary.totalUncompressedBytes,
        cache: presentation.getCacheDiagnostics(),
      },
      render,
      thumbnails,
    },
  );
}

export function createPptxErrorDiagnosticReport(
  file: Pick<TFile, "name" | "stat">,
  category: PptxErrorKind,
  rawMessage: string,
): PptxErrorDiagnosticReport {
  const fingerprint = fingerprintMessage(rawMessage);
  return createReaderDiagnosticReport(
    "pptx",
    "error",
    {
      name: file.name,
      size: file.stat.size,
      mtime: file.stat.mtime,
    },
    `Presentation could not be opened (fingerprint: ${fingerprint})`,
    {
      category,
      fingerprint,
    },
  );
}

export function formatPptxDiagnosticReport(
  report: PptxDiagnosticReport | PptxErrorDiagnosticReport,
): string {
  return formatReaderDiagnosticReport(report);
}
