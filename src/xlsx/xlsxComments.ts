import {
  attribute,
  childrenNamed,
  descendantsNamed,
  firstChildNamed,
  parseXml,
  textContent,
} from "../pptx/xml";
import { validateXmlStructure } from "../ooxml/xmlStructure";
import { parseCellReference } from "./xlsxReferences";
import type { XlsxComment } from "./xlsxTypes";

const MAX_COMMENTS_PER_SHEET = 50_000;
const MAX_COMMENT_TEXT_LENGTH = 32_768;
const MAX_COMMENT_AUTHOR_LENGTH = 256;

export function parseXlsxComments(
  xml: string,
  path: string,
): XlsxComment[] {
  validateXmlStructure(xml, path, "comments");
  const document = parseXml(xml, path);
  const root = document.documentElement;
  const authors = childrenNamed(
    firstChildNamed(root, "authors") ?? root,
    "author",
  ).map((author) =>
    truncate(textContent(author).trim(), MAX_COMMENT_AUTHOR_LENGTH),
  );
  const comments: XlsxComment[] = [];
  for (const element of descendantsNamed(root, "comment")) {
    if (comments.length >= MAX_COMMENTS_PER_SHEET) {
      break;
    }
    const ref = attribute(element, "ref");
    if (!ref) {
      continue;
    }
    let position: { row: number; column: number };
    try {
      position = parseCellReference(ref);
    } catch {
      continue;
    }
    const authorId = Number(attribute(element, "authorId"));
    const author =
      Number.isInteger(authorId) && authorId >= 0
        ? authors[authorId] ?? ""
        : "";
    const textElement = firstChildNamed(element, "text");
    const text = truncate(
      textContent(textElement).trim(),
      MAX_COMMENT_TEXT_LENGTH,
    );
    comments.push({
      ref,
      row: position.row,
      column: position.column,
      author,
      text,
    });
  }
  return comments;
}

function truncate(value: string, limit: number): string {
  return value.length <= limit ? value : `${value.slice(0, limit - 1)}…`;
}
