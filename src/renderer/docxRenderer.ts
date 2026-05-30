// SECURITY: This module uses docx-preview for safe local document rendering.
// No external scripts are loaded. All content comes from trusted local .docx files.
import { renderAsync } from "docx-preview";

interface RenderDocxOptions {
  fitWidth: boolean;
}

/**
 * Renders a .docx file buffer into the target DOM element.
 * SECURITY: Only creates structural HTML elements (div, span, img, etc.) for display.
 * No <script> elements are created or executed during rendering.
 */
export async function renderDocx(
  buffer: ArrayBuffer,
  targetEl: HTMLElement,
  options: RenderDocxOptions,
): Promise<void> {
  await renderAsync(buffer, targetEl, undefined, {
    className: "docx",
    inWrapper: true,
    ignoreWidth: options.fitWidth,
    ignoreHeight: false,
    ignoreFonts: false,
    breakPages: true,
    useBase64URL: true,
    experimental: true,
    trimXmlDeclaration: true,
  });
}
