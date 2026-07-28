import { PptxPackageError } from "./pptxPackage";
import type { PptxReaderText } from "./pptxI18n";
import { ZipSafetyError } from "./zipLimits";

export interface PptxErrorInfo {
  kind: PptxErrorKind;
  title: string;
  body: string;
  tips: string[];
  details: string;
}

export type PptxErrorKind =
  | "format-mismatch"
  | "encrypted"
  | "limit-exceeded"
  | "damaged"
  | "unsupported"
  | "unknown";

export function classifyPptxError(
  error: unknown,
  text: PptxReaderText,
): PptxErrorInfo {
  const details = error instanceof Error ? error.message : String(error);
  const tips = [text.errors.tryAgain, text.errors.openExternally];

  if (error instanceof ZipSafetyError) {
    switch (error.kind) {
      case "format-mismatch":
        return {
          kind: "format-mismatch",
          title: text.errors.formatMismatchTitle,
          body: text.errors.formatMismatchBody,
          tips,
          details,
        };
      case "encrypted":
        return {
          kind: "encrypted",
          title: text.errors.encryptedTitle,
          body: text.errors.encryptedBody,
          tips: [text.errors.openExternally],
          details,
        };
      case "limit-exceeded":
        return {
          kind: "limit-exceeded",
          title: text.errors.limitTitle,
          body: text.errors.limitBody,
          tips: [text.errors.openExternally],
          details,
        };
      case "unsupported":
        return {
          kind: "unsupported",
          title: text.errors.unsupportedTitle,
          body: text.errors.unsupportedBody,
          tips,
          details,
        };
      case "damaged":
      default:
        return {
          kind: "damaged",
          title: text.errors.damagedTitle,
          body: text.errors.damagedBody,
          tips,
          details,
        };
    }
  }

  if (error instanceof PptxPackageError) {
    return {
      kind: error.kind === "unsupported" ? "unsupported" : "damaged",
      title:
        error.kind === "unsupported"
          ? text.errors.unsupportedTitle
          : text.errors.damagedTitle,
      body:
        error.kind === "unsupported"
          ? text.errors.unsupportedBody
          : text.errors.damagedBody,
      tips,
      details,
    };
  }

  return {
    kind: "unknown",
    title: text.errors.unknownTitle,
    body: text.errors.unknownBody,
    tips,
    details,
  };
}
