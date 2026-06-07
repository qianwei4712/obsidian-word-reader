import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ts from "typescript";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourceDir = path.join(rootDir, "src");
const bundlePath = path.join(rootDir, "dist", "main.js");

function findSourceFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      return findSourceFiles(fullPath);
    }
    return entry.isFile() && /\.(?:ts|tsx|js|mjs)$/.test(entry.name)
      ? [fullPath]
      : [];
  });
}

function getMemberName(expression) {
  if (ts.isIdentifier(expression)) {
    return expression.text;
  }
  if (ts.isPropertyAccessExpression(expression)) {
    return expression.name.text;
  }
  if (
    ts.isElementAccessExpression(expression) &&
    expression.argumentExpression &&
    ts.isStringLiteralLike(expression.argumentExpression)
  ) {
    return expression.argumentExpression.text;
  }
  return null;
}

function getStaticString(expression) {
  return expression && ts.isStringLiteralLike(expression)
    ? expression.text
    : null;
}

function scanSourceText(sourceText, fileName) {
  const sourceFile = ts.createSourceFile(
    fileName,
    sourceText,
    ts.ScriptTarget.Latest,
    true,
    fileName.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
  const findings = [];

  function report(node, rule, message) {
    const position = sourceFile.getLineAndCharacterOfPosition(node.getStart());
    findings.push({
      fileName,
      line: position.line + 1,
      column: position.character + 1,
      rule,
      message,
    });
  }

  function visit(node) {
    if (ts.isCallExpression(node)) {
      const memberName = getMemberName(node.expression);
      const firstArgument = node.arguments[0];
      const firstString = getStaticString(firstArgument);

      if (memberName === "eval" || memberName === "Function") {
        report(node, "dynamic-code", `Disallowed ${memberName}() call.`);
      }

      if (
        (memberName === "createElement" || memberName === "createEl") &&
        firstString?.toLowerCase() === "script"
      ) {
        report(node, "script-element", "Creating script elements is disallowed.");
      }

      if (
        memberName === "insertAdjacentHTML" ||
        memberName === "importScripts"
      ) {
        report(node, "html-or-script-injection", `Disallowed ${memberName}() call.`);
      }

      if (
        (memberName === "write" || memberName === "writeln") &&
        ts.isPropertyAccessExpression(node.expression) &&
        getMemberName(node.expression.expression) === "document"
      ) {
        report(node, "document-write", `Disallowed document.${memberName}() call.`);
      }

      if (memberName === "fetch" || memberName === "requestUrl") {
        report(node, "network-request", `Disallowed ${memberName}() call.`);
      }

      if (
        (memberName === "setTimeout" || memberName === "setInterval") &&
        firstArgument &&
        !ts.isArrowFunction(firstArgument) &&
        !ts.isFunctionExpression(firstArgument) &&
        (ts.isStringLiteralLike(firstArgument) ||
          !ts.isIdentifier(firstArgument))
      ) {
        report(node, "string-timer", `${memberName}() must receive a function.`);
      }

      if (node.expression.kind === ts.SyntaxKind.ImportKeyword) {
        if (firstString === null) {
          report(node, "dynamic-import", "Dynamic import paths must be static strings.");
        } else if (/^(?:https?:)?\/\//i.test(firstString)) {
          report(node, "remote-import", "Remote module imports are disallowed.");
        }
      }

      if (
        memberName === "setAttribute" &&
        (firstString === "src" || firstString === "href")
      ) {
        const value = getStaticString(node.arguments[1]);
        if (value && /^(?:https?:)?\/\//i.test(value)) {
          report(node, "remote-resource", "Remote DOM resources are disallowed.");
        }
      }
    }

    if (ts.isNewExpression(node)) {
      const constructorName = getMemberName(node.expression);
      if (constructorName === "Function") {
        report(node, "dynamic-code", "Disallowed new Function() expression.");
      }
      if (
        constructorName === "WebSocket" ||
        constructorName === "EventSource" ||
        constructorName === "Worker" ||
        constructorName === "SharedWorker"
      ) {
        report(node, "remote-execution", `Disallowed ${constructorName} construction.`);
      }
    }

    if (
      ts.isBinaryExpression(node) &&
      node.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
      (ts.isPropertyAccessExpression(node.left) ||
        ts.isElementAccessExpression(node.left))
    ) {
      const propertyName = getMemberName(node.left);
      if (propertyName === "innerHTML" || propertyName === "outerHTML") {
        report(node, "html-injection", `Assignment to ${propertyName} is disallowed.`);
      }
    }

    ts.forEachChild(node, visit);
  }

  visit(sourceFile);
  return findings;
}

function validateScannerRules() {
  const cases = [
    ["eval(userCode)", "dynamic-code"],
    ["new Function(userCode)", "dynamic-code"],
    ['document.createElement("script")', "script-element"],
    ['target.insertAdjacentHTML("beforeend", html)', "html-or-script-injection"],
    ["element.innerHTML = html", "html-injection"],
    ['fetch("https://example.com")', "network-request"],
    ["import(moduleUrl)", "dynamic-import"],
  ];

  for (const [sourceText, expectedRule] of cases) {
    const findings = scanSourceText(sourceText, "scanner-self-test.ts");
    if (!findings.some((finding) => finding.rule === expectedRule)) {
      throw new Error(`Security scanner self-test failed for ${expectedRule}.`);
    }
  }
}

function scanBundle() {
  if (!fs.existsSync(bundlePath)) {
    return [{
      fileName: bundlePath,
      line: 1,
      column: 1,
      rule: "missing-bundle",
      message: "dist/main.js is missing. Run the production build first.",
    }];
  }

  const sourceText = fs.readFileSync(bundlePath, "utf8");
  const patterns = [
    {
      rule: "bundled-script-element",
      pattern: /create(?:Element|El)\s*\(\s*["']script["']/i,
      message: "The production bundle creates a script element.",
    },
    {
      rule: "bundled-script-markup",
      pattern: /<script(?:\s|>)/i,
      message: "The production bundle contains script markup.",
    },
    {
      rule: "bundled-import-scripts",
      pattern: /\bimportScripts\s*\(/,
      message: "The production bundle calls importScripts().",
    },
  ];

  return patterns.flatMap(({ rule, pattern, message }) => {
    const match = pattern.exec(sourceText);
    if (!match) {
      return [];
    }
    const prefix = sourceText.slice(0, match.index);
    const lines = prefix.split("\n");
    return [{
      fileName: bundlePath,
      line: lines.length,
      column: lines.at(-1).length + 1,
      rule,
      message,
    }];
  });
}

validateScannerRules();

const sourceFiles = findSourceFiles(sourceDir);
const findings = sourceFiles.flatMap((fileName) =>
  scanSourceText(fs.readFileSync(fileName, "utf8"), fileName)
);
findings.push(...scanBundle());

if (findings.length > 0) {
  console.error("Security scan failed:");
  for (const finding of findings) {
    const relativePath = path.relative(rootDir, finding.fileName);
    console.error(
      `- ${relativePath}:${finding.line}:${finding.column} ` +
      `[${finding.rule}] ${finding.message}`,
    );
  }
  process.exit(1);
}

console.log(
  `Security scan passed: ${sourceFiles.length} source files and dist/main.js checked.`,
);
