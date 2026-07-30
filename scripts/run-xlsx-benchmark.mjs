import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(rootDir, ".benchmark-dist");
const outputPath = path.join(outputDir, "xlsx-benchmark.cjs");
const resultDir = path.join(rootDir, "benchmark-results");
const resultPath = path.join(resultDir, "xlsx-100k.latest.json");
const budget = JSON.parse(
  fs.readFileSync(
    path.join(rootDir, "benchmarks", "xlsx-100k-budget.json"),
    "utf8",
  ),
);

fs.rmSync(outputDir, { recursive: true, force: true });
fs.mkdirSync(outputDir, { recursive: true });

try {
  await build({
    entryPoints: [path.join(rootDir, "scripts", "xlsx-benchmark-runner.ts")],
    outfile: outputPath,
    bundle: true,
    packages: "external",
    platform: "node",
    format: "cjs",
    target: "node18",
    logLevel: "warning",
  });
  const run = spawnSync(
    process.execPath,
    ["--expose-gc", outputPath],
    {
      cwd: rootDir,
      encoding: "utf8",
    },
  );
  if (run.status !== 0) {
    process.stderr.write(run.stderr);
    process.stdout.write(run.stdout);
    process.exit(run.status ?? 1);
  }
  const resultLine = run.stdout
    .split(/\r?\n/)
    .find((line) => line.startsWith("XLSX_BENCHMARK_JSON="));
  if (!resultLine) {
    throw new Error("XLSX benchmark did not emit a result record.");
  }
  const result = JSON.parse(resultLine.slice("XLSX_BENCHMARK_JSON=".length));
  fs.mkdirSync(resultDir, { recursive: true });
  fs.writeFileSync(resultPath, `${JSON.stringify(result, null, 2)}\n`);

  const failures = [];
  for (const [metric, maximum] of Object.entries(budget.maximums)) {
    if (result[metric] > maximum) {
      failures.push(`${metric}=${result[metric]} exceeds ${maximum}`);
    }
  }
  for (const [metric, expected] of Object.entries(budget.requirements)) {
    if (result[metric] !== expected) {
      failures.push(`${metric}=${result[metric]} does not equal ${expected}`);
    }
  }
  if (failures.length > 0) {
    console.error("XLSX performance benchmark failed:");
    for (const failure of failures) {
      console.error(`- ${failure}`);
    }
    process.exit(1);
  }
  console.log(
    "XLSX performance benchmark passed: " +
      `first paint ${result.firstPaintMs} ms, ` +
      `scroll p95 ${result.scrollFrameP95Ms} ms, ` +
      `peak heap ${result.peakHeapMiB} MiB, ` +
      `${result.maximumDomNodes} max nodes, cancellation passed.`,
      ` Dense streamed parse ${result.denseParseMs} ms / ` +
      `${result.densePeakHeapMiB} MiB across ${result.streamedChunkCount} chunks, ` +
      `${result.maximumSheetDataBufferKiB} KiB maximum worksheet-data buffer.`,
  );
} finally {
  fs.rmSync(outputDir, { recursive: true, force: true });
}
