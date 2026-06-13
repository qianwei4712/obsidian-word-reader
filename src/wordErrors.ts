import type { WordReaderText } from "./i18n";
import {
  createReaderDiagnostics,
  fingerprintMessage,
  formatReaderDiagnostics,
  type ReaderDiagnosticFile,
  type ReaderDiagnostics,
} from "./reader/diagnostics";

export { fingerprintMessage };

export type WordErrorKind =
  | "encrypted"
  | "format-mismatch"
  | "zip-container"
  | "xml-structure"
  | "unsupported-structure"
  | "unknown";

export interface WordErrorInfo {
  kind: WordErrorKind;
  title: string;
  body: string;
  status: string;
  tips: string[];
}

export type WordDiagnostics = ReaderDiagnostics<WordErrorKind>;
export type WordDiagnosticFile = ReaderDiagnosticFile;

export function classifyWordError(
  message: string,
  text: WordReaderText,
  buffer: ArrayBuffer | null,
): WordErrorInfo {
  const normalizedMessage = message.toLowerCase();

  if (
    normalizedMessage.includes("password") ||
    normalizedMessage.includes("encrypt") ||
    normalizedMessage.includes("protected") ||
    (buffer && isEncryptedOfficeBuffer(buffer))
  ) {
    return {
      kind: "encrypted",
      title: text.errors.encryptedTitle,
      body: text.errors.encryptedBody,
      status: text.errors.encryptedStatus,
      tips: text.errors.encryptedTips,
    };
  }

  if (buffer && !hasZipSignature(buffer)) {
    return {
      kind: "format-mismatch",
      title: text.errors.formatMismatchTitle,
      body: text.errors.formatMismatchBody,
      status: text.errors.formatMismatchStatus,
      tips: text.errors.formatMismatchTips,
    };
  }

  if (
    normalizedMessage.includes("corrupt") ||
    normalizedMessage.includes("zip") ||
    normalizedMessage.includes("central directory") ||
    normalizedMessage.includes("archive") ||
    normalizedMessage.includes("crc") ||
    normalizedMessage.includes("unexpected end")
  ) {
    return {
      kind: "zip-container",
      title: text.errors.zipContainerTitle,
      body: text.errors.zipContainerBody,
      status: text.errors.zipContainerStatus,
      tips: text.errors.zipContainerTips,
    };
  }

  if (
    normalizedMessage.includes("xml") ||
    normalizedMessage.includes("malformed") ||
    normalizedMessage.includes("parse error") ||
    normalizedMessage.includes("invalid character") ||
    normalizedMessage.includes("unclosed") ||
    normalizedMessage.includes("namespace")
  ) {
    return {
      kind: "xml-structure",
      title: text.errors.xmlStructureTitle,
      body: text.errors.xmlStructureBody,
      status: text.errors.xmlStructureStatus,
      tips: text.errors.xmlStructureTips,
    };
  }

  if (
    normalizedMessage.includes("unsupported") ||
    normalizedMessage.includes("not supported") ||
    normalizedMessage.includes("not implemented") ||
    normalizedMessage.includes("smartart") ||
    normalizedMessage.includes("custom xml") ||
    normalizedMessage.includes("drawing")
  ) {
    return {
      kind: "unsupported-structure",
      title: text.errors.unsupportedStructureTitle,
      body: text.errors.unsupportedStructureBody,
      status: text.errors.unsupportedStructureStatus,
      tips: text.errors.unsupportedStructureTips,
    };
  }

  return {
    kind: "unknown",
    title: text.errors.genericTitle,
    body: text.errors.genericBody,
    status: text.errors.genericStatus,
    tips: text.errors.genericTips,
  };
}

export function hasZipSignature(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength < 4) {
    return false;
  }

  const bytes = new Uint8Array(buffer, 0, 4);
  return (
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    ((bytes[2] === 0x03 && bytes[3] === 0x04) ||
      (bytes[2] === 0x05 && bytes[3] === 0x06) ||
      (bytes[2] === 0x07 && bytes[3] === 0x08))
  );
}

export function isEncryptedOfficeBuffer(buffer: ArrayBuffer): boolean {
  const bytes = new Uint8Array(buffer);
  const oleSignature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (
    bytes.length < oleSignature.length ||
    !oleSignature.every((value, index) => bytes[index] === value)
  ) {
    return false;
  }

  return (
    containsUtf16LeText(bytes, "EncryptedPackage") ||
    containsUtf16LeText(bytes, "EncryptionInfo")
  );
}

export function createWordDiagnostics(
  category: WordErrorKind,
  message: string,
  file: WordDiagnosticFile,
): WordDiagnostics {
  return createReaderDiagnostics(
    category,
    createSafeDiagnosticSummary(category, message),
    file,
  );
}

export function createSafeDiagnosticSummary(
  category: WordErrorKind,
  message: string,
): string {
  const summaries: Record<WordErrorKind, string> = {
    encrypted: "Encrypted or protected Office container detected",
    "format-mismatch": "File content does not match the .docx container format",
    "zip-container": "The ZIP-based Office package could not be read",
    "xml-structure": "An internal document structure could not be parsed",
    "unsupported-structure": "The renderer reported an unsupported structure",
    unknown: "The renderer returned an unclassified error",
  };
  return `${summaries[category]} (fingerprint: ${fingerprintMessage(message)})`;
}

export function formatWordDiagnostics(
  diagnostics: WordDiagnostics,
): string {
  return formatReaderDiagnostics("Office Reader", diagnostics);
}

function containsUtf16LeText(bytes: Uint8Array, text: string): boolean {
  const pattern = Array.from(text).flatMap((character) => [
    character.charCodeAt(0),
    0,
  ]);

  for (let start = 0; start <= bytes.length - pattern.length; start += 1) {
    let matches = true;
    for (let offset = 0; offset < pattern.length; offset += 1) {
      if (bytes[start + offset] !== pattern[offset]) {
        matches = false;
        break;
      }
    }
    if (matches) {
      return true;
    }
  }

  return false;
}
