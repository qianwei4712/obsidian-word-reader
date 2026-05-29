import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const pkg = JSON.parse(fs.readFileSync(path.join(rootDir, "package.json"), "utf8"));
const version = process.argv[2] ?? pkg.version;
const changelog = fs.readFileSync(path.join(rootDir, "CHANGELOG.md"), "utf8");
const lines = changelog.split(/\r?\n/);
const versionHeading = new RegExp(`^##\\s+${version.replaceAll(".", "\\.")}\\s+-\\s+\\d{4}-\\d{2}-\\d{2}`);
const anyVersionHeading = /^##\s+\d+\.\d+\.\d+\s+-\s+\d{4}-\d{2}-\d{2}/;
const start = lines.findIndex((line) => versionHeading.test(line));

if (start === -1) {
  console.error(`Could not find CHANGELOG.md section for ${version}.`);
  process.exit(1);
}

const next = lines.findIndex((line, index) => index > start && anyVersionHeading.test(line));
const section = lines.slice(start, next === -1 ? undefined : next).join("\n").trimEnd() + "\n";

const releaseDir = path.join(rootDir, "release");
fs.mkdirSync(releaseDir, { recursive: true });

const outputPath = path.join(releaseDir, `CHANGELOG-${version}.md`);
fs.writeFileSync(outputPath, section, "utf8");
console.log(`Created ${path.relative(rootDir, outputPath)}.`);
