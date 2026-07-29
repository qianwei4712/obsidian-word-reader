import { parseQualifiedRangeReference } from "./xlsxReferences";
import type {
  XlsxDefinedName,
  XlsxSheetDescriptor,
} from "./xlsxTypes";

export interface XlsxDefinedNameTarget {
  sheetIndex: number;
  range: XlsxDefinedName["range"];
}

export function parseXlsxDefinedNameTarget(
  target: string,
  sheets: readonly Pick<XlsxSheetDescriptor, "name">[],
  localSheetIndex?: number,
): XlsxDefinedNameTarget | null {
  const parsed = parseQualifiedRangeReference(target);
  if (!parsed) {
    return null;
  }
  const sheetIndex = parsed.sheetName
    ? sheets.findIndex((sheet) => sheet.name === parsed.sheetName)
    : localSheetIndex;
  if (
    sheetIndex === undefined ||
    sheetIndex < 0 ||
    sheetIndex >= sheets.length
  ) {
    return null;
  }
  return {
    sheetIndex,
    range: parsed.range,
  };
}

export function resolveXlsxDefinedName(
  names: readonly XlsxDefinedName[],
  query: string,
  activeSheetIndex: number,
): XlsxDefinedName | null {
  const normalized = query.trim().toLocaleLowerCase();
  if (!normalized) {
    return null;
  }
  return (
    names.find(
      (name) =>
        name.name.toLocaleLowerCase() === normalized &&
        name.scopeSheetIndex === activeSheetIndex,
    ) ??
    names.find(
      (name) =>
        name.name.toLocaleLowerCase() === normalized &&
        name.scopeSheetIndex === undefined,
    ) ??
    null
  );
}
