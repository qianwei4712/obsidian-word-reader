// SECURITY: This module uses docx-preview for safe local document rendering.
// No external scripts are loaded. All content comes from trusted local .docx files.
import {
  parseAsync,
  renderDocument,
  type Options,
} from "docx-preview";

import { annotateDocxHeadingLevels } from "./docxOutline";

const DOCX_RENDER_OPTIONS: Partial<Options> = {
  className: "docx",
  inWrapper: true,
  ignoreWidth: false,
  ignoreHeight: false,
  ignoreFonts: false,
  breakPages: true,
  // Blob URLs avoid duplicating large embedded images as base64 strings.
  // DocxSession tracks and revokes every generated URL when the preview changes.
  useBase64URL: false,
  experimental: true,
  trimXmlDeclaration: true,
};

/**
 * Renders a .docx file buffer into the target DOM element.
 * SECURITY: Only creates structural HTML elements (div, span, img, etc.) for display.
 * No <script> elements are created or executed during rendering.
 */
export async function renderDocx(
  buffer: ArrayBuffer,
  targetEl: HTMLElement,
): Promise<void> {
  const document: unknown = await parseAsync(buffer, DOCX_RENDER_OPTIONS);
  annotateDocxHeadingLevels(document);
  await renderDocument(
    document,
    targetEl,
    undefined,
    DOCX_RENDER_OPTIONS,
  );
}
