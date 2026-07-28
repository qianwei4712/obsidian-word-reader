// Executed in a dedicated Node.js process with --expose-gc.
import { performance } from "node:perf_hooks";

import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import { XlsxVirtualGrid } from "../src/xlsx/xlsxVirtualGrid";
import { XlsxWorksheetCancelledError } from "../src/xlsx/xlsxWorksheet";
import { createRichXlsx } from "../tests/xlsxFixture";

interface XlsxBenchmarkResult {
  fixture: string;
  firstPaintMs: number;
  scrollFrameP95Ms: number;
  peakHeapMiB: number;
  maximumDomNodes: number;
  cancellationPassed: boolean;
  populatedCells: number;
  rowCount: number;
}

void runBenchmark().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});

async function runBenchmark(): Promise<void> {
  const buffer = await createRichXlsx();
  runGarbageCollection();
  const baselineHeap = process.memoryUsage().heapUsed;
  let peakHeap = baselineHeap;
  const sampleHeap = (): void => {
    peakHeap = Math.max(peakHeap, process.memoryUsage().heapUsed);
  };

  const cancellationWorkbook = await XlsxPackage.load(buffer);
  let cancellationPassed = false;
  try {
    await cancellationWorkbook.getWorksheet(0, {
      isCancelled: () => true,
    });
  } catch (error) {
    cancellationPassed = error instanceof XlsxWorksheetCancelledError;
  }
  cancellationWorkbook.clearCaches();
  sampleHeap();

  const firstPaintStart = performance.now();
  const workbook = await XlsxPackage.load(buffer);
  sampleHeap();
  const worksheet = await workbook.getWorksheet(0);
  sampleHeap();
  const grid = new XlsxVirtualGrid(worksheet);
  const firstWindow = grid.calculate({
    scrollTop: 0,
    scrollLeft: 0,
    width: 1_280,
    height: 800,
  });
  const firstPaintMs = performance.now() - firstPaintStart;
  sampleHeap();

  const frameDurations: number[] = [];
  let maximumDomNodes = firstWindow.estimatedDomNodeCount;
  for (let frame = 0; frame < 240; frame += 1) {
    const progress = frame / 239;
    const frameStart = performance.now();
    const window = grid.calculate({
      scrollTop: progress * Math.max(0, grid.totalHeight - 800),
      scrollLeft: progress * Math.max(0, grid.totalWidth - 1_280),
      width: 1_280,
      height: 800,
    });
    frameDurations.push(performance.now() - frameStart);
    maximumDomNodes = Math.max(
      maximumDomNodes,
      window.estimatedDomNodeCount,
    );
    if (frame % 24 === 0) {
      sampleHeap();
    }
  }
  sampleHeap();

  frameDurations.sort((left, right) => left - right);
  const percentileIndex = Math.floor(frameDurations.length * 0.95);
  const result: XlsxBenchmarkResult = {
    fixture: "generated-rich-100000-row-sparse",
    firstPaintMs: round(firstPaintMs),
    scrollFrameP95Ms: round(frameDurations[percentileIndex] ?? 0),
    peakHeapMiB: round((peakHeap - baselineHeap) / (1024 * 1024)),
    maximumDomNodes,
    cancellationPassed,
    populatedCells: worksheet.populatedCellCount,
    rowCount: worksheet.rowCount,
  };

  console.log(`XLSX_BENCHMARK_JSON=${JSON.stringify(result)}`);
}

function runGarbageCollection(): void {
  const runtime = globalThis as typeof globalThis & {
    gc?: () => void;
  };
  runtime.gc?.();
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}
