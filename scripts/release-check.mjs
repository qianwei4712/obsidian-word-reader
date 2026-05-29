import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function readJson(relativePath) {
  return JSON.parse(fs.readFileSync(path.join(rootDir, relativePath), "utf8"));
}

function hasFile(relativePath) {
  return fs.existsSync(path.join(rootDir, relativePath));
}

function parseTagArg() {
  const index = process.argv.indexOf("--tag");
  if (index === -1) return null;
  return process.argv[index + 1] ?? "";
}

const pkg = readJson("package.json");
const manifest = readJson("manifest.json");
const lockfile = readJson("package-lock.json");
const changelog = fs.readFileSync(path.join(rootDir, "CHANGELOG.md"), "utf8");
const tagName = parseTagArg();
const failures = [];

function expect(condition, message) {
  if (!condition) failures.push(message);
}

expect(/^\d+\.\d+\.\d+$/.test(pkg.version), `package.json version must be plain semver, got ${pkg.version}`);
expect(pkg.version === manifest.version, `package.json version ${pkg.version} must match manifest.json version ${manifest.version}`);
expect(pkg.version === lockfile.version, `package-lock.json root version ${lockfile.version} must match package.json version ${pkg.version}`);
expect(
  pkg.version === lockfile.packages?.[""]?.version,
  `package-lock.json packages[""].version ${lockfile.packages?.[""]?.version} must match package.json version ${pkg.version}`,
);
expect(
  new RegExp(`^##\\s+${pkg.version.replaceAll(".", "\\.")}\\s+-\\s+\\d{4}-\\d{2}-\\d{2}`, "m").test(changelog),
  `CHANGELOG.md must contain a dated section for ${pkg.version}`,
);

if (tagName !== null) {
  expect(tagName === `v${pkg.version}`, `Git tag ${tagName || "(empty)"} must match package version v${pkg.version}`);
}

for (const file of ["dist/main.js", "dist/manifest.json", "dist/styles.css"]) {
  expect(hasFile(file), `${file} must exist after build`);
}

if (hasFile("dist/manifest.json")) {
  const distManifest = readJson("dist/manifest.json");
  expect(
    distManifest.version === pkg.version,
    `dist/manifest.json version ${distManifest.version} must match package version ${pkg.version}`,
  );
}

if (failures.length > 0) {
  console.error("Release check failed:");
  for (const failure of failures) {
    console.error(`- ${failure}`);
  }
  process.exit(1);
}

console.log(`Release check passed for ${pkg.name} v${pkg.version}.`);
