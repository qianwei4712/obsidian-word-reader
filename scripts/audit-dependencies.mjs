import fs from "node:fs";
import path from "node:path";
import { builtinModules } from "node:module";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const pkg = readJson("package.json");
const lockfile = readJson("package-lock.json");
const manifest = readJson("manifest.json");
const licenseInventoryPath = path.join(rootDir, "THIRD_PARTY_LICENSES.md");
const maximumBundleBytes = 500 * 1024;
const maximumReleaseBytes = 8 * 1024 * 1024;
const releaseMajor = Number.parseInt(pkg.version.split(".")[0] ?? "", 10);
const allowedLicenses = new Set([
  "Apache-2.0",
  "ISC",
  "MIT",
  "(MIT AND Zlib)",
  "(MIT OR GPL-3.0-or-later)",
]);
const failures = [];

function readJson(relativePath) {
  return JSON.parse(fs.readFileSync(path.join(rootDir, relativePath), "utf8"));
}

function expect(condition, message) {
  if (!condition) failures.push(message);
}

function getRuntimePackages() {
  return Object.entries(lockfile.packages ?? {})
    .filter(
      ([key, value]) =>
        key.startsWith("node_modules/") && value.dev !== true,
    )
    .map(([key, value]) => {
      const name = packageNameFromNodeModulesPath(key);
      const directSpec = pkg.dependencies?.[name];
      if (directSpec?.startsWith("file:")) {
        const vendorPackage = readJson(
          path.posix.join(
            directSpec.slice("file:".length).replaceAll("\\", "/"),
            "package.json",
          ),
        );
        return {
          name,
          version: vendorPackage.version,
          license: vendorPackage.license,
          source: directSpec,
        };
      }
      return {
        name,
        version: value.version,
        license: value.license,
        source: `npm:${name}@${value.version}`,
      };
    })
    .sort((left, right) => left.name.localeCompare(right.name));
}

function packageNameFromNodeModulesPath(key) {
  const marker = "node_modules/";
  const tail = key.slice(key.lastIndexOf(marker) + marker.length);
  const parts = tail.split("/");
  return tail.startsWith("@") ? parts.slice(0, 2).join("/") : parts[0];
}

function auditExactDirectVersions(runtimePackages) {
  const packageByName = new Map(
    runtimePackages.map((entry) => [entry.name, entry]),
  );
  for (const [name, spec] of Object.entries(pkg.dependencies ?? {})) {
    const installed = packageByName.get(name);
    expect(Boolean(installed), `Runtime dependency ${name} is missing from the lockfile.`);
    if (spec.startsWith("file:")) {
      const vendorPath = path.join(rootDir, spec.slice("file:".length));
      expect(
        fs.existsSync(path.join(vendorPath, "package.json")),
        `Vendored dependency ${name} is missing ${spec}/package.json.`,
      );
      continue;
    }
    expect(
      /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(spec),
      `Top-level runtime dependency ${name} must use an exact version, got ${spec}.`,
    );
    expect(
      installed?.version === spec,
      `Top-level runtime dependency ${name}@${spec} does not match lockfile ${installed?.version}.`,
    );
  }
}

async function getBundledPackageNames() {
  const result = await build({
    entryPoints: [path.join(rootDir, "src", "main.ts")],
    bundle: true,
    alias: {
      jszip: "jszip/lib/index.js",
    },
    external: [
      "obsidian",
      "electron",
      "@codemirror/autocomplete",
      "@codemirror/collab",
      "@codemirror/commands",
      "@codemirror/language",
      "@codemirror/lint",
      "@codemirror/search",
      "@codemirror/state",
      "@codemirror/view",
      "@lezer/common",
      "@lezer/highlight",
      "@lezer/lr",
      ...builtinModules,
      ...builtinModules.map((name) => `node:${name}`),
    ],
    format: "cjs",
    target: "es2018",
    write: false,
    metafile: true,
    logLevel: "silent",
  });
  const names = new Set();
  const inputs = Object.keys(result.metafile.inputs);
  for (const input of inputs) {
    const normalized = input.replaceAll("\\", "/");
    const nodeModulesIndex = normalized.lastIndexOf("node_modules/");
    if (nodeModulesIndex >= 0) {
      names.add(
        packageNameFromNodeModulesPath(normalized.slice(nodeModulesIndex)),
      );
      continue;
    }
    const vendorMatch = /(?:^|\/)vendor\/([^/]+)\//.exec(normalized);
    if (vendorMatch) {
      names.add(vendorMatch[1]);
    }
  }
  const bundlesXlsx = inputs.some((input) =>
    /(?:^|\/)src\/xlsx\//.test(input.replaceAll("\\", "/")),
  );
  if (releaseMajor < 3) {
    expect(
      !bundlesXlsx,
      "Unreleased XLSX research modules must not be bundled before 3.0.0.",
    );
  } else {
    expect(
      bundlesXlsx,
      "The 3.x public plugin must bundle the reviewed XLSX reader.",
    );
  }
  return names;
}

function auditLicenses(runtimePackages, bundledNames) {
  const runtimeNames = new Set(runtimePackages.map((entry) => entry.name));
  for (const name of bundledNames) {
    expect(
      runtimeNames.has(name),
      `Bundled package ${name} is absent from the runtime license inventory.`,
    );
  }
  for (const entry of runtimePackages) {
    expect(Boolean(entry.version), `Runtime package ${entry.name} has no exact version.`);
    expect(Boolean(entry.license), `Runtime package ${entry.name} has no license metadata.`);
    expect(
      allowedLicenses.has(entry.license),
      `Runtime package ${entry.name}@${entry.version} uses unreviewed license ${entry.license}.`,
    );
  }
}

function auditArtifactSizes() {
  const bundlePath = path.join(rootDir, "dist", "main.js");
  expect(fs.existsSync(bundlePath), "dist/main.js is missing; run the build first.");
  if (fs.existsSync(bundlePath)) {
    expect(
      fs.statSync(bundlePath).size <= maximumBundleBytes,
      `dist/main.js exceeds the ${maximumBundleBytes} byte budget.`,
    );
  }
  const distPaths = ["main.js", "manifest.json", "styles.css"]
    .map((name) => path.join(rootDir, "dist", name))
    .filter((file) => fs.existsSync(file));
  const distBytes = distPaths.reduce(
    (total, file) => total + fs.statSync(file).size,
    0,
  );
  expect(
    distBytes <= maximumReleaseBytes,
    `Build artifacts total ${distBytes} bytes, exceeding the ${maximumReleaseBytes} byte 3.0 target.`,
  );
  const releasePath = path.join(
    rootDir,
    "release",
    `${pkg.name}-${pkg.version}.zip`,
  );
  if (fs.existsSync(releasePath)) {
    expect(
      fs.statSync(releasePath).size <= maximumReleaseBytes,
      `${path.relative(rootDir, releasePath)} exceeds the ${maximumReleaseBytes} byte release target.`,
    );
  }
}

function generateLicenseInventory(runtimePackages, bundledNames) {
  const rows = runtimePackages.map((entry) => {
    const bundled = bundledNames.has(entry.name) ? "yes" : "runtime closure";
    return `| \`${entry.name}\` | \`${entry.version}\` | ${entry.license} | ${bundled} | \`${entry.source}\` |`;
  });
  return `# Third-party runtime licenses

Generated by \`npm run dependencies:licenses\` from the locked production
dependency graph. Top-level registry dependencies use exact versions; local
polyfills use exact vendored package versions. "Bundled" means the package was
observed in an esbuild metafile for \`src/main.ts\`.

| Package | Version | License | Bundled | Locked source |
| --- | --- | --- | --- | --- |
${rows.join("\n")}

## License choices and distribution

- JSZip is used under its MIT option.
- Pako includes both its MIT and zlib notices.
- Apache-2.0, ISC, MIT and zlib terms in this inventory were reviewed as
  compatible with this MIT-licensed, offline read-only plugin.
- The release artifact remains \`main.js\`, \`manifest.json\`, and
  \`styles.css\`; esbuild preserves legal comments from bundled dependencies.
`;
}

const runtimePackages = getRuntimePackages();
auditExactDirectVersions(runtimePackages);
const bundledNames = await getBundledPackageNames();
auditLicenses(runtimePackages, bundledNames);
auditArtifactSizes();
const mainSource = fs.readFileSync(path.join(rootDir, "src", "main.ts"), "utf8");
const settingsSource = fs.readFileSync(
  path.join(rootDir, "src", "settings.ts"),
  "utf8",
);
const hasXlsxView = fs.existsSync(path.join(rootDir, "src", "XlsxView.ts"));
const registersXlsx =
  /registerExtensions\s*\(\s*\[\.\.\.XLSX_ADAPTER\.extensions\]/s.test(
    mainSource,
  );
if (releaseMajor < 3) {
  expect(
    !manifest.description.toLowerCase().includes("xlsx"),
    "The public manifest must not advertise XLSX before 3.0.0.",
  );
  expect(
    !hasXlsxView,
    "The pre-3.0 research line must not provide a public XlsxView.",
  );
  expect(
    !registersXlsx,
    "The pre-3.0 research line must not register the .xlsx extension.",
  );
} else {
  expect(
    manifest.description.toLowerCase().includes("xlsx"),
    "The 3.x public manifest must advertise XLSX support.",
  );
  expect(hasXlsxView, "The 3.x public plugin must provide XlsxView.");
  expect(
    registersXlsx,
    "The 3.x public plugin must register the .xlsx extension.",
  );
}
expect(
  !/\b(?:enable\w*xlsx|xlsx\w*enabled)\b/i.test(settingsSource),
  "XLSX availability must not be controlled by an undocumented feature flag.",
);

const inventory = generateLicenseInventory(runtimePackages, bundledNames);
if (process.argv.includes("--write")) {
  fs.writeFileSync(licenseInventoryPath, inventory);
} else {
  expect(
    fs.existsSync(licenseInventoryPath),
    "THIRD_PARTY_LICENSES.md is missing; run npm run dependencies:licenses.",
  );
  if (fs.existsSync(licenseInventoryPath)) {
    expect(
      fs.readFileSync(licenseInventoryPath, "utf8") === inventory,
      "THIRD_PARTY_LICENSES.md is stale; run npm run dependencies:licenses.",
    );
  }
}

if (failures.length > 0) {
  console.error("Dependency audit failed:");
  for (const failure of failures) {
    console.error(`- ${failure}`);
  }
  process.exit(1);
}

console.log(
  `Dependency audit passed: ${runtimePackages.length} locked runtime packages, ` +
    `${bundledNames.size} bundled packages, 500 KiB/8 MiB size gates.`,
);
