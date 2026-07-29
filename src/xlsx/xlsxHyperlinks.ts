export interface XlsxWorkbookLocation {
  sheetName?: string;
  reference: string;
}

export function parseXlsxWorkbookLocation(
  value: string,
): XlsxWorkbookLocation | null {
  const match =
    /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(\$?[A-Z]{1,3}\$?\d{1,7})$/i.exec(
      value.trim(),
    );
  if (!match) {
    return null;
  }
  const sheetName = (match[1] ?? match[2])?.replace(/''/g, "'");
  if (sheetName?.includes("[") || sheetName?.includes("]")) {
    return null;
  }
  return {
    sheetName,
    reference: match[3],
  };
}

export function isSafeXlsxExternalTarget(value: string): boolean {
  try {
    const url = new URL(value);
    return (
      url.protocol === "https:" ||
      url.protocol === "http:" ||
      url.protocol === "mailto:"
    );
  } catch {
    return false;
  }
}
