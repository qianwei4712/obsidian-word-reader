import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(rootDir, ".benchmark-dist");
const outputPath = path.join(outputDir, "performance-benchmark.cjs");
const resultDir = path.join(rootDir, "benchmark-results");
const resultPath = path.join(resultDir, "performance.latest.json");
const trendPath = path.join(resultDir, "performance-trend.json");
const summaryPath = path.join(resultDir, "performance-summary.md");
const formats = ["docx", "pptx", "xlsx"];
const iterations = 5;

fs.rmSync(outputDir, { recursive: true, force: true });
fs.mkdirSync(outputDir, { recursive: true });

try {
  await build({
    entryPoints: [
      path.join(rootDir, "scripts", "performance-benchmark-runner.ts"),
    ],
    outfile: outputPath,
    bundle: true,
    packages: "external",
    platform: "node",
    format: "cjs",
    target: "node18",
    logLevel: "warning",
  });

  const formatResults = {};
  for (const format of formats) {
    const samples = [];
    for (let iteration = 1; iteration <= iterations; iteration += 1) {
      process.stdout.write(
        `Running ${format.toUpperCase()} performance sample ` +
          `${iteration}/${iterations}...\n`,
      );
      samples.push(runSample(format));
    }
    formatResults[format] = aggregateSamples(samples);
  }

  const artifacts = collectArtifactSizes();
  const generatedAt = new Date().toISOString();
  const result = {
    product: "Office Reader",
    schemaVersion: 1,
    generatedAt,
    ci: {
      commit: process.env.GITHUB_SHA ?? null,
      runId: process.env.GITHUB_RUN_ID ?? null,
      runAttempt: process.env.GITHUB_RUN_ATTEMPT ?? null,
    },
    runtime: {
      node: process.version,
      platform: process.platform,
      architecture: process.arch,
      cpuCount: os.cpus().length,
    },
    iterations,
    artifacts,
    formats: formatResults,
  };

  const failures = validateBudgets(result);
  const summary = createSummary(result, failures);
  const trend = createTrendPoint(result, failures);
  fs.mkdirSync(resultDir, { recursive: true });
  fs.writeFileSync(resultPath, `${JSON.stringify(result, null, 2)}\n`);
  fs.writeFileSync(trendPath, `${JSON.stringify(trend, null, 2)}\n`);
  fs.writeFileSync(summaryPath, summary);

  process.stdout.write(summary);
  if (failures.length > 0) {
    process.stderr.write("Performance budgets failed:\n");
    for (const failure of failures) {
      process.stderr.write(`- ${failure}\n`);
    }
    process.exitCode = 1;
  }
} finally {
  fs.rmSync(outputDir, { recursive: true, force: true });
}

function runSample(format) {
  const run = spawnSync(
    process.execPath,
    ["--expose-gc", outputPath, format],
    {
      cwd: rootDir,
      encoding: "utf8",
      maxBuffer: 16 * 1024 * 1024,
    },
  );
  if (run.status !== 0) {
    process.stderr.write(run.stderr);
    process.stdout.write(run.stdout);
    process.exit(run.status ?? 1);
  }
  const resultLine = run.stdout
    .split(/\r?\n/)
    .find((line) => line.startsWith("OFFICE_READER_BENCHMARK_JSON="));
  if (!resultLine) {
    throw new Error(`${format.toUpperCase()} benchmark did not emit a sample.`);
  }
  const sample = JSON.parse(
    resultLine.slice("OFFICE_READER_BENCHMARK_JSON=".length),
  );
  if (sample.format !== format) {
    throw new Error(
      `${format.toUpperCase()} benchmark emitted format ${sample.format}.`,
    );
  }
  return sample;
}

function aggregateSamples(samples) {
  const metrics = new Set(
    samples.flatMap((sample) => Object.keys(sample.metrics)),
  );
  const median = {};
  const p95 = {};
  for (const metric of [...metrics].sort()) {
    const values = samples.map((sample) => sample.metrics[metric]);
    if (values.some((value) => typeof value !== "number")) {
      throw new Error(`Benchmark metric ${metric} is not numeric in every sample.`);
    }
    median[metric] = round(percentile(values, 0.5));
    p95[metric] = round(percentile(values, 0.95));
  }

  const checks = {};
  const checkNames = new Set(
    samples.flatMap((sample) => Object.keys(sample.checks)),
  );
  for (const check of [...checkNames].sort()) {
    checks[check] = samples.every((sample) => sample.checks[check] === true);
  }

  return {
    fixture: samples[0].fixture,
    context: samples[0].context,
    median,
    p95,
    checks,
    samples: samples.map((sample, index) => ({
      iteration: index + 1,
      metrics: sample.metrics,
      checks: sample.checks,
    })),
  };
}

function collectArtifactSizes() {
  const releaseFiles = ["main.js", "manifest.json", "styles.css"];
  const bundlePath = path.join(rootDir, "dist", "main.js");
  if (!fs.existsSync(bundlePath)) {
    throw new Error("dist/main.js does not exist; run the production build first.");
  }
  const missing = releaseFiles.filter(
    (file) => !fs.existsSync(path.join(rootDir, "dist", file)),
  );
  if (missing.length > 0) {
    throw new Error(`Missing release files: ${missing.join(", ")}`);
  }
  const releasePayloadBytes = releaseFiles.reduce(
    (total, file) => total + fs.statSync(path.join(rootDir, "dist", file)).size,
    0,
  );
  const zipHeaderBytes = releaseFiles.reduce(
    (total, file) => total + 30 + Buffer.byteLength(file) + 46 + Buffer.byteLength(file),
    22,
  );
  return {
    bundleBytes: fs.statSync(bundlePath).size,
    bundleLimitBytes: 500 * 1024,
    releaseZipBytes: releasePayloadBytes + zipHeaderBytes,
    releaseZipLimitBytes: 8 * 1024 * 1024,
    releaseZipMeasurement: "exact-size-for-current-stored-zip-packager",
  };
}

function validateBudgets(result) {
  const failures = [];
  if (result.artifacts.bundleBytes > result.artifacts.bundleLimitBytes) {
    failures.push(
      `artifacts.bundleBytes=${result.artifacts.bundleBytes} exceeds ` +
        result.artifacts.bundleLimitBytes,
    );
  }
  if (result.artifacts.releaseZipBytes > result.artifacts.releaseZipLimitBytes) {
    failures.push(
      `artifacts.releaseZipBytes=${result.artifacts.releaseZipBytes} exceeds ` +
        result.artifacts.releaseZipLimitBytes,
    );
  }

  for (const format of formats) {
    const budgetPath = path.join(
      rootDir,
      "benchmarks",
      `${format}-budget.json`,
    );
    const budget = JSON.parse(fs.readFileSync(budgetPath, "utf8"));
    const observed = result.formats[format];
    validateComparisons(format, observed, budget.maximums, "maximum", failures);
    validateComparisons(format, observed, budget.minimums, "minimum", failures);
    validateRequirements(format, observed, budget.requirements, failures);
  }
  return failures;
}

function validateComparisons(format, observed, entries = {}, kind, failures) {
  for (const [metricPath, threshold] of Object.entries(entries)) {
    const value = getPath(observed, metricPath);
    if (typeof value !== "number") {
      failures.push(`${format}.${metricPath} is missing or non-numeric`);
      continue;
    }
    if (kind === "maximum" && value > threshold) {
      failures.push(`${format}.${metricPath}=${value} exceeds ${threshold}`);
    }
    if (kind === "minimum" && value < threshold) {
      failures.push(`${format}.${metricPath}=${value} is below ${threshold}`);
    }
  }
}

function validateRequirements(format, observed, entries = {}, failures) {
  for (const [valuePath, expected] of Object.entries(entries)) {
    const value = getPath(observed, valuePath);
    if (value !== expected) {
      failures.push(
        `${format}.${valuePath}=${JSON.stringify(value)} does not equal ` +
          JSON.stringify(expected),
      );
    }
  }
}

function getPath(value, dottedPath) {
  return dottedPath.split(".").reduce(
    (current, part) => current?.[part],
    value,
  );
}

function createTrendPoint(result, failures) {
  return {
    product: result.product,
    schemaVersion: result.schemaVersion,
    generatedAt: result.generatedAt,
    ci: result.ci,
    runtime: result.runtime,
    passed: failures.length === 0,
    failures,
    artifacts: result.artifacts,
    formats: Object.fromEntries(
      formats.map((format) => [format, {
        fixture: result.formats[format].fixture,
        median: result.formats[format].median,
        p95: result.formats[format].p95,
        checks: result.formats[format].checks,
      }]),
    ),
  };
}

function createSummary(result, failures) {
  const lines = [
    "# Office Reader performance baseline",
    "",
    `Status: ${failures.length === 0 ? "passed" : "failed"}`,
    "",
    `Iterations per format: ${result.iterations} independent processes`,
    "",
    "| Format | First readable p95 | Parse p95 | Search p95 | Navigation p95 | Scroll p95 | Peak heap p95 | Max DOM p95 |",
    "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
  ];
  for (const format of formats) {
    const metrics = result.formats[format].p95;
    lines.push(
      `| ${format.toUpperCase()} | ${metrics.firstReadableMs} ms | ` +
        `${metrics.parseMs} ms | ${metrics.searchMs} ms | ` +
        `${metrics.navigationMs} ms | ${metrics.scrollFrameP95Ms} ms | ` +
        `${metrics.peakHeapMiB} MiB | ${metrics.maximumDomNodes} |`,
    );
  }
  lines.push(
    "",
    `Bundle: ${result.artifacts.bundleBytes} / ` +
      `${result.artifacts.bundleLimitBytes} bytes`,
    "",
    `Release zip: ${result.artifacts.releaseZipBytes} / ` +
      `${result.artifacts.releaseZipLimitBytes} bytes`,
    "",
  );
  if (failures.length > 0) {
    lines.push("## Failures", "", ...failures.map((failure) => `- ${failure}`), "");
  }
  return `${lines.join("\n")}\n`;
}

function percentile(values, quantile) {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  const index = Math.min(
    sorted.length - 1,
    Math.max(0, Math.ceil(sorted.length * quantile) - 1),
  );
  return sorted[index];
}

function round(value) {
  return Math.round(value * 100) / 100;
}
