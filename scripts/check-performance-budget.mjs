import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const bundlePath = path.join(rootDir, "dist", "main.js");
const maximumBundleBytes = 500 * 1024;
const failures = [];

if (!fs.existsSync(bundlePath)) {
  failures.push("dist/main.js does not exist; run the production build first.");
} else {
  const bundleBytes = fs.statSync(bundlePath).size;
  if (bundleBytes > maximumBundleBytes) {
    failures.push(
      `dist/main.js is ${bundleBytes} bytes; budget is ${maximumBundleBytes} bytes.`,
    );
  } else {
    console.log(
      `Performance budget passed: dist/main.js is ${bundleBytes} bytes ` +
        `(limit ${maximumBundleBytes}).`,
    );
  }
}

if (failures.length > 0) {
  console.error("Performance budget failed:");
  for (const failure of failures) {
    console.error(`- ${failure}`);
  }
  process.exit(1);
}
