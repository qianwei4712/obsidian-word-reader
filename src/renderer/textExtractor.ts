// SECURITY: This module uses mammoth for safe local text extraction.
// No external scripts are loaded. All content comes from trusted local .docx files.
import * as mammoth from "mammoth";

/**
 * Extracts plain text from a .docx file buffer.
 * SECURITY: Only reads text content, no script execution or network access.
 */
export async function extractPlainText(buffer: ArrayBuffer): Promise<string> {
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return result.value.trim();
}

/**
 * Converts .docx content to Markdown format.
 * SECURITY: Only converts document structure to Markdown, no script execution or network access.
 */
export async function extractMarkdown(buffer: ArrayBuffer): Promise<string> {
  const result = await (
    mammoth as typeof mammoth & {
      convertToMarkdown: typeof mammoth.convertToHtml;
    }
  ).convertToMarkdown(
    { arrayBuffer: buffer },
    {
      styleMap: [
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Subtitle'] => p:fresh",
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "p[style-name='Heading 5'] => h5:fresh",
        "p[style-name='Heading 6'] => h6:fresh",
      ],
    },
  );
  return result.value.trim();
}
