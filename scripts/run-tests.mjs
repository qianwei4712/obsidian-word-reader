import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const testsDir = path.join(rootDir, "tests");
const outputDir = path.join(rootDir, ".test-dist");

function findTestFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      return findTestFiles(fullPath);
    }
    return entry.isFile() && entry.name.endsWith(".test.ts") ? [fullPath] : [];
  });
}

const testFiles = findTestFiles(testsDir).sort();
if (testFiles.length === 0) {
  console.error("No test files found.");
  process.exit(1);
}

fs.rmSync(outputDir, { recursive: true, force: true });

try {
  const result = await build({
    entryPoints: testFiles,
    outdir: outputDir,
    outbase: testsDir,
    bundle: true,
    packages: "external",
    platform: "node",
    format: "cjs",
    target: "node18",
    sourcemap: "inline",
    logLevel: "warning",
  });

  if (result.errors.length > 0) {
    process.exitCode = 1;
  } else {
    const compiledTests = testFiles.map((file) => {
      const relativePath = path.relative(testsDir, file);
      return path.join(outputDir, relativePath.replace(/\.ts$/, ".js"));
    });
    const testRun = spawnSync(
      process.execPath,
      ["--test", ...compiledTests],
      {
        cwd: rootDir,
        stdio: "inherit",
      },
    );
    process.exitCode = testRun.status ?? 1;
  }
} finally {
  fs.rmSync(outputDir, { recursive: true, force: true });
}
