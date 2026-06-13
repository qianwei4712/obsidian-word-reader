import { PptxPackageError } from "./pptxPackage";
import type { PptxReaderText } from "./pptxI18n";
import { ZipSafetyError } from "./zipLimits";

export interface PptxErrorInfo {
  title: string;
  body: string;
  tips: string[];
  details: string;
}

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
          title: text.errors.formatMismatchTitle,
          body: text.errors.formatMismatchBody,
          tips,
          details,
        };
      case "encrypted":
        return {
          title: text.errors.encryptedTitle,
          body: text.errors.encryptedBody,
          tips: [text.errors.openExternally],
          details,
        };
      case "limit-exceeded":
        return {
          title: text.errors.limitTitle,
          body: text.errors.limitBody,
          tips: [text.errors.openExternally],
          details,
        };
      case "damaged":
      default:
        return {
          title: text.errors.damagedTitle,
          body: text.errors.damagedBody,
          tips,
          details,
        };
    }
  }

  if (error instanceof PptxPackageError) {
    return {
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
    title: text.errors.unknownTitle,
    body: text.errors.unknownBody,
    tips,
    details,
  };
}
