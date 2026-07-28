import JSZip from "jszip";

import {
  attribute,
  descendantsNamed,
  parseXml,
} from "../pptx/xml";
import {
  relationshipPartPath,
  resolvePackagePath,
} from "./packagePath";

export interface OoxmlRelationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export async function readRelationships(
  zip: JSZip,
  sourcePath: string,
): Promise<Map<string, OoxmlRelationship>> {
  const path = relationshipPartPath(sourcePath);
  const entry = zip.file(path);
  if (!entry) {
    return new Map();
  }
  const document = parseXml(await entry.async("string"), path);
  const relationships = new Map<string, OoxmlRelationship>();
  for (const element of descendantsNamed(
    document.documentElement,
    "Relationship",
  )) {
    const id = attribute(element, "Id");
    const type = attribute(element, "Type");
    const target = attribute(element, "Target");
    if (!id || !type || !target) {
      continue;
    }
    relationships.set(id, {
      id,
      type,
      target,
      external: attribute(element, "TargetMode") === "External",
    });
  }
  return relationships;
}

export function resolveInternalRelationship(
  sourcePath: string,
  relationship: OoxmlRelationship | undefined,
): string | null {
  if (!relationship || relationship.external) {
    return null;
  }
  return resolvePackagePath(sourcePath, relationship.target);
}
