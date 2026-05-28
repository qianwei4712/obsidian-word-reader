import * as mammoth from "mammoth";

export async function extractPlainText(buffer: ArrayBuffer): Promise<string> {
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return result.value.trim();
}
