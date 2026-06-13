import type { TFile } from "obsidian";

import type { PptxPackage } from "./pptxPackage";
import type { PptxRenderDiagnostics } from "./pptxRenderer";

export interface PptxThumbnailDiagnostics {
  mounted: number;
  rendered: number;
  rendering: number;
  resourceCount: number;
}

export interface PptxDiagnosticReport {
  product: "Office Reader";
  version: 1;
  file: {
    name: string;
    sizeBytes: number;
    modifiedTime: number;
  };
  presentation: {
    slide: number;
    slideCount: number;
    slideWidthEmu: number;
    slideHeightEmu: number;
    zipFileCount: number;
    zipCompressedBytes: number;
    zipExpandedBytes: number;
  };
  render: PptxRenderDiagnostics;
  thumbnails: PptxThumbnailDiagnostics;
}

export function createPptxDiagnosticReport(
  file: Pick<TFile, "name" | "stat">,
  presentation: PptxPackage,
  slideIndex: number,
  render: PptxRenderDiagnostics,
  thumbnails: PptxThumbnailDiagnostics,
): PptxDiagnosticReport {
  return {
    product: "Office Reader",
    version: 1,
    file: {
      name: file.name,
      sizeBytes: file.stat.size,
      modifiedTime: file.stat.mtime,
    },
    presentation: {
      slide: slideIndex + 1,
      slideCount: presentation.slideCount,
      slideWidthEmu: presentation.slideWidth,
      slideHeightEmu: presentation.slideHeight,
      zipFileCount: presentation.zipSummary.fileCount,
      zipCompressedBytes: presentation.zipSummary.totalCompressedBytes,
      zipExpandedBytes: presentation.zipSummary.totalUncompressedBytes,
    },
    render,
    thumbnails,
  };
}

export function formatPptxDiagnosticReport(
  report: PptxDiagnosticReport,
): string {
  return JSON.stringify(report, null, 2);
}
