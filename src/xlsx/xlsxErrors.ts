import { OoxmlPolicyError } from "../ooxml/packagePolicy";
import { ZipSafetyError } from "../ooxml/packageSafety";
import { XlsxPackageError } from "./xlsxPackage";
import type { XlsxReaderText } from "./xlsxI18n";

export interface XlsxErrorInfo {
  title: string;
  body: string;
  tips: string[];
  details: string;
}

export function classifyXlsxError(
  error: unknown,
  text: XlsxReaderText,
): XlsxErrorInfo {
  const details = error instanceof Error ? error.message : String(error);
  const commonTips = [text.errors.tryAgain, text.errors.openExternally];

  if (error instanceof ZipSafetyError) {
    switch (error.kind) {
      case "format-mismatch":
        return {
          title: text.errors.formatMismatchTitle,
          body: text.errors.formatMismatchBody,
          tips: commonTips,
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
      case "unsupported":
        return {
          title: text.errors.unsupportedTitle,
          body: text.errors.unsupportedBody,
          tips: commonTips,
          details,
        };
      case "damaged":
      default:
        return {
          title: text.errors.damagedTitle,
          body: text.errors.damagedBody,
          tips: commonTips,
          details,
        };
    }
  }

  if (error instanceof XlsxPackageError) {
    const unsupported = error.kind === "unsupported";
    return {
      title: unsupported
        ? text.errors.unsupportedTitle
        : text.errors.damagedTitle,
      body: unsupported
        ? text.errors.unsupportedBody
        : text.errors.damagedBody,
      tips: commonTips,
      details,
    };
  }

  if (error instanceof OoxmlPolicyError) {
    return {
      title: text.errors.unsupportedTitle,
      body: text.errors.unsupportedBody,
      tips: [text.errors.openExternally],
      details,
    };
  }

  return {
    title: text.errors.unknownTitle,
    body: text.errors.unknownBody,
    tips: commonTips,
    details,
  };
}
