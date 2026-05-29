import * as mammoth from "mammoth";

export async function extractPlainText(buffer: ArrayBuffer): Promise<string> {
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return result.value.trim();
}

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
