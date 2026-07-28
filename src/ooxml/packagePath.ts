export class OoxmlPathError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "OoxmlPathError";
  }
}

export function resolvePackagePath(basePath: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizePackagePath(target.slice(1));
  }
  const baseParts = basePath.split("/");
  baseParts.pop();
  return normalizePackagePath([...baseParts, ...target.split("/")].join("/"));
}

export function relationshipPartPath(path: string): string {
  const parts = path.split("/");
  const fileName = parts.pop();
  if (!fileName) {
    throw new OoxmlPathError("A relationship source must name a package part.");
  }
  return [...parts, "_rels", `${fileName}.rels`].join("/");
}

export function normalizePackagePath(path: string): string {
  const normalized: string[] = [];
  for (const part of path.replace(/\\/g, "/").split("/")) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      if (normalized.length === 0) {
        throw new OoxmlPathError(
          "A package relationship escapes the archive root.",
        );
      }
      normalized.pop();
      continue;
    }
    normalized.push(part);
  }
  return normalized.join("/");
}
