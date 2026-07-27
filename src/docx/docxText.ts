import type { ReaderTextExtraction } from "../reader/adapters";

export function extractDocxText(rootEl: HTMLElement): ReaderTextExtraction {
  return {
    plainText: getRenderedPlainText(rootEl),
    markdown: markdownForChildNodes(rootEl)
      .replace(/\n{3,}/g, "\n\n")
      .trim(),
  };
}

export function docxFragmentToMarkdown(fragment: DocumentFragment): string {
  return markdownForChildNodes(fragment).replace(/\n{3,}/g, "\n\n").trim();
}

export function markdownForChildNodes(parent: Node): string {
  return Array.from(parent.childNodes).map(markdownForNode).join("");
}

function markdownForNode(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent ?? "";
  }

  if (!node.instanceOf(HTMLElement)) {
    return markdownForChildNodes(node);
  }

  const tagName = node.tagName.toLowerCase();
  const content = markdownForChildNodes(node).trim();

  switch (tagName) {
    case "h1":
    case "h2":
    case "h3":
    case "h4":
    case "h5":
    case "h6":
      return `${"#".repeat(Number(tagName.slice(1)))} ${content}\n\n`;
    case "p":
    case "div":
      return content ? `${content}\n\n` : "";
    case "strong":
    case "b":
      return content ? `**${content}**` : "";
    case "em":
    case "i":
      return content ? `*${content}*` : "";
    case "code":
      return content ? `\`${content.replace(/`/g, "\\`")}\`` : "";
    case "br":
      return "\n";
    case "a": {
      const href = node.getAttribute("href");
      return href && content ? `[${content}](${href})` : content;
    }
    case "ul":
      return `${markdownForList(node, "-")}\n`;
    case "ol":
      return `${markdownForList(node, "1.")}\n`;
    case "li":
      return content;
    case "table":
      return `${markdownForTable(node)}\n\n`;
    case "tr":
    case "tbody":
    case "thead":
    case "span":
    default:
      return markdownForChildNodes(node);
  }
}

function markdownForList(listEl: HTMLElement, marker: string): string {
  return Array.from(listEl.children)
    .filter((child): child is HTMLElement => child.instanceOf(HTMLElement))
    .filter((child) => child.tagName.toLowerCase() === "li")
    .map((itemEl) => `${marker} ${markdownForChildNodes(itemEl).trim()}`)
    .join("\n");
}

function markdownForTable(tableEl: HTMLElement): string {
  const rows = Array.from(tableEl.querySelectorAll("tr"))
    .map((rowEl) =>
      Array.from(rowEl.children).map((cellEl) =>
        normalizeWhitespace(cellEl.textContent ?? "").replace(/\|/g, "\\|"),
      ),
    )
    .filter((row) => row.length > 0);

  if (rows.length === 0) {
    return "";
  }

  const header = rows[0];
  const divider = header.map(() => "---");
  return [header, divider, ...rows.slice(1)]
    .map((row) => `| ${row.join(" | ")} |`)
    .join("\n");
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function getRenderedPlainText(rootEl: HTMLElement): string {
  return (rootEl.innerText || rootEl.textContent || "")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
