import esbuild from "esbuild";
import process from "process";
import fs from "fs/promises";
import builtins from "builtin-modules";

const production = process.argv[2] === "production";
const outputDir = "dist";

async function copyPluginAssets() {
  await fs.mkdir(outputDir, { recursive: true });
  await Promise.all([
    fs.copyFile("manifest.json", `${outputDir}/manifest.json`),
    fs.copyFile("styles.css", `${outputDir}/styles.css`),
  ]);
}

const copyAssetsPlugin = {
  name: "copy-plugin-assets",
  setup(build) {
    build.onEnd(async (result) => {
      if (result.errors.length === 0) {
        await copyPluginAssets();
      }
    });
  },
};

const context = await esbuild.context({
  banner: {
    js: "/* Obsidian Word Reader */",
  },
  entryPoints: ["src/main.ts"],
  bundle: true,
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
    ...builtins,
  ],
  format: "cjs",
  target: "es2018",
  logLevel: "info",
  sourcemap: production ? false : "inline",
  treeShaking: true,
  outfile: `${outputDir}/main.js`,
  minify: production,
  plugins: [copyAssetsPlugin],
});

if (production) {
  await context.rebuild();
  process.exit(0);
}

await context.watch();
