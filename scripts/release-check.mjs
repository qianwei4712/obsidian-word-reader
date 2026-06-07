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

function readZipEntryNames(relativePath) {
  const archive = fs.readFileSync(path.join(rootDir, relativePath));
  const endSignature = 0x06054b50;
  let endOffset = -1;

  for (let offset = archive.length - 22; offset >= 0; offset -= 1) {
    if (archive.readUInt32LE(offset) === endSignature) {
      endOffset = offset;
      break;
    }
  }

  if (endOffset === -1) {
    throw new Error(`${relativePath} does not contain a ZIP end record`);
  }

  const entryCount = archive.readUInt16LE(endOffset + 10);
  let offset = archive.readUInt32LE(endOffset + 16);
  const names = [];

  for (let index = 0; index < entryCount; index += 1) {
    if (archive.readUInt32LE(offset) !== 0x02014b50) {
      throw new Error(`${relativePath} has an invalid central directory`);
    }

    const nameLength = archive.readUInt16LE(offset + 28);
    const extraLength = archive.readUInt16LE(offset + 30);
    const commentLength = archive.readUInt16LE(offset + 32);
    names.push(archive.subarray(offset + 46, offset + 46 + nameLength).toString("utf8"));
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return names;
}

function parseTagArg() {
  const index = process.argv.indexOf("--tag");
  if (index === -1) return null;
  return process.argv[index + 1] ?? "";
}

const pkg = readJson("package.json");
const manifest = readJson("manifest.json");
const lockfile = readJson("package-lock.json");
const versions = readJson("versions.json");
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
expect(
  versions[pkg.version] === manifest.minAppVersion,
  `versions.json must map ${pkg.version} to minAppVersion ${manifest.minAppVersion}`,
);
expect(
  /^[a-z]+(?:-[a-z]+)*$/.test(manifest.id),
  `manifest.json id must contain lowercase letters and hyphens only, got ${manifest.id}`,
);
expect(
  !manifest.id.includes("obsidian") && !manifest.id.endsWith("plugin"),
  `manifest.json id must not contain "obsidian" or end with "plugin", got ${manifest.id}`,
);
expect(
  typeof manifest.description === "string" &&
    manifest.description.length <= 250 &&
    manifest.description.endsWith("."),
  "manifest.json description must be at most 250 characters and end with a period",
);
expect(
  typeof manifest.isDesktopOnly === "boolean",
  "manifest.json isDesktopOnly must be a boolean",
);

for (const file of [
  "README.md",
  "LICENSE",
  "manifest.json",
  "versions.json",
  "COMMUNITY_PLUGIN_CHECKLIST.md",
]) {
  expect(hasFile(file), `${file} must exist in the repository root`);
}

if (tagName !== null) {
  expect(tagName === pkg.version, `Git tag ${tagName || "(empty)"} must match package version ${pkg.version}`);
}

for (const file of ["dist/main.js", "dist/manifest.json", "dist/styles.css"]) {
  expect(hasFile(file), `${file} must exist after build`);
}

const releaseZip = `release/${pkg.name}-${pkg.version}.zip`;
expect(hasFile(releaseZip), `${releaseZip} must exist after packaging`);
if (hasFile(releaseZip)) {
  try {
    const actualEntries = readZipEntryNames(releaseZip).sort();
    const expectedEntries = ["main.js", "manifest.json", "styles.css"].sort();
    expect(
      JSON.stringify(actualEntries) === JSON.stringify(expectedEntries),
      `${releaseZip} must contain only ${expectedEntries.join(", ")}, got ${actualEntries.join(", ")}`,
    );
  } catch (error) {
    failures.push(error instanceof Error ? error.message : String(error));
  }
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

console.log(`Release check passed for ${pkg.name} ${pkg.version}.`);
