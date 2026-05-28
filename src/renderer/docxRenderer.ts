import { renderAsync } from "docx-preview";

interface RenderDocxOptions {
  fitWidth: boolean;
}

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
