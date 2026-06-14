export interface TextSearchMatch {
  node: Text;
  start: number;
  end: number;
}

interface IndexedTextNode {
  node: Text;
  normalizedText: string;
}

export class TextSearchIndex {
  private constructor(private readonly entries: IndexedTextNode[]) {}

  static fromTextNodes(nodes: readonly Text[]): TextSearchIndex {
    const entries: IndexedTextNode[] = [];
    for (const node of nodes) {
      const value = node.nodeValue ?? "";
      if (value) {
        entries.push({
          node,
          normalizedText: value.toLocaleLowerCase(),
        });
      }
    }
    return new TextSearchIndex(entries);
  }

  static build(rootEl: HTMLElement): TextSearchIndex {
    const ownerDocument = rootEl.ownerDocument;
    const nodeFilter = ownerDocument.defaultView?.NodeFilter;
    const showText = nodeFilter?.SHOW_TEXT ?? 4;
    const filterAccept = nodeFilter?.FILTER_ACCEPT ?? 1;
    const filterReject = nodeFilter?.FILTER_REJECT ?? 2;
    const walker = ownerDocument.createTreeWalker(rootEl, showText, {
      acceptNode(node) {
        const parentEl = node.parentElement;
        return !parentEl || parentEl.closest("style, script")
          ? filterReject
          : filterAccept;
      },
    });
    const nodes: Text[] = [];
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      nodes.push(node);
    }
    return TextSearchIndex.fromTextNodes(nodes);
  }

  get nodeCount(): number {
    return this.entries.length;
  }

  async find(
    query: string,
    options: {
      shouldContinue?: () => boolean;
      yieldControl?: () => Promise<void>;
      chunkSize?: number;
    } = {},
  ): Promise<TextSearchMatch[]> {
    const normalizedQuery = query.trim().toLocaleLowerCase();
    if (!normalizedQuery) {
      return [];
    }
    const shouldContinue = options.shouldContinue ?? (() => true);
    const yieldControl = options.yieldControl ?? (async () => undefined);
    const chunkSize = Math.max(1, options.chunkSize ?? 500);
    const matches: TextSearchMatch[] = [];

    for (let index = 0; index < this.entries.length; index += 1) {
      if (!shouldContinue()) {
        return [];
      }
      const entry = this.entries[index];
      let offset = 0;
      while (offset < entry.normalizedText.length) {
        const matchIndex = entry.normalizedText.indexOf(
          normalizedQuery,
          offset,
        );
        if (matchIndex === -1) {
          break;
        }
        matches.push({
          node: entry.node,
          start: matchIndex,
          end: matchIndex + normalizedQuery.length,
        });
        offset = matchIndex + normalizedQuery.length;
      }
      if ((index + 1) % chunkSize === 0 && index + 1 < this.entries.length) {
        await yieldControl();
      }
    }
    return matches;
  }
}
