import {
  attribute,
  descendantsNamed,
  firstDescendantNamed,
  textContent,
} from "./xml";

export interface PptxSlideMetadata {
  index: number;
  title: string;
  text: string;
  notes: string;
}

export interface PptxSearchResult {
  slideIndex: number;
  title: string;
  snippet: string;
  matchCount: number;
  matchedNotes: boolean;
}

interface IndexedSlideMetadata {
  metadata: PptxSlideMetadata;
  visibleText: string;
  normalizedVisibleText: string;
  normalizedNotes: string;
}

const EXCLUDED_NOTES_PLACEHOLDERS = new Set([
  "dt",
  "ftr",
  "hdr",
  "sldImg",
  "sldNum",
]);

export function extractSlideMetadata(
  index: number,
  slide: Document,
  notes: Document | null,
  fallbackTitle: string,
): PptxSlideMetadata {
  const shapes = descendantsNamed(slide.documentElement, "sp");
  const titleShape = shapes.find((shape) => {
    const type = attribute(firstDescendantNamed(shape, "ph"), "type");
    return type === "title" || type === "ctrTitle";
  });
  const title = normalizeText(extractElementParagraphs(titleShape))[0]
    ?? fallbackTitle;
  const paragraphs = normalizeText(
    extractElementParagraphs(slide.documentElement),
  );
  const noteParagraphs = notes
    ? normalizeText(
        descendantsNamed(notes.documentElement, "sp").flatMap((shape) => {
          const type = attribute(firstDescendantNamed(shape, "ph"), "type");
          return type && EXCLUDED_NOTES_PLACEHOLDERS.has(type)
            ? []
            : extractElementParagraphs(shape);
        }),
      )
    : [];

  return {
    index,
    title,
    text: paragraphs.join("\n"),
    notes: noteParagraphs.join("\n"),
  };
}

export function searchPptxSlides(
  slides: readonly PptxSlideMetadata[],
  query: string,
): PptxSearchResult[] {
  return new PptxSearchIndex(slides).search(query);
}

export class PptxSearchIndex {
  private readonly slides = new Map<number, IndexedSlideMetadata>();

  constructor(slides: readonly PptxSlideMetadata[] = []) {
    for (const slide of slides) {
      this.set(slide);
    }
  }

  set(slide: PptxSlideMetadata): void {
    const visibleText = combineTitleAndText(slide);
    this.slides.set(slide.index, {
      metadata: slide,
      visibleText,
      normalizedVisibleText: visibleText.toLocaleLowerCase(),
      normalizedNotes: slide.notes.toLocaleLowerCase(),
    });
  }

  clear(): void {
    this.slides.clear();
  }

  search(
    query: string,
    slideIndices: readonly number[] = [...this.slides.keys()].sort(
      (left, right) => left - right,
    ),
  ): PptxSearchResult[] {
    const normalizedQuery = query.trim().toLocaleLowerCase();
    if (!normalizedQuery) {
      return slideIndices.map((index) => {
        const slide = this.slides.get(index)?.metadata ?? emptyMetadata(index);
        return {
          slideIndex: slide.index,
          title: slide.title,
          snippet: firstUsefulLine(slide.text, slide.title),
          matchCount: 0,
          matchedNotes: false,
        };
      });
    }

    return slideIndices.flatMap((index) => {
      const indexed = this.slides.get(index);
      if (!indexed) {
        return [];
      }
      const visibleMatches = countNormalizedMatches(
        indexed.normalizedVisibleText,
        normalizedQuery,
      );
      const noteMatches = countNormalizedMatches(
        indexed.normalizedNotes,
        normalizedQuery,
      );
      const matchCount = visibleMatches + noteMatches;
      if (matchCount === 0) {
        return [];
      }
      const source =
        visibleMatches > 0
          ? indexed.visibleText
          : indexed.metadata.notes;
      return [{
        slideIndex: indexed.metadata.index,
        title: indexed.metadata.title,
        snippet: createSnippet(source, normalizedQuery),
        matchCount,
        matchedNotes: visibleMatches === 0 && noteMatches > 0,
      }];
    });
  }
}

export function formatSlideText(metadata: PptxSlideMetadata): string {
  const sections = metadata.title ? [metadata.title] : [];
  const body = bodyLinesWithoutFirstTitle(metadata).join("\n").trim();
  if (body) {
    sections.push(body);
  }
  return sections.join("\n\n");
}

function combineTitleAndText(metadata: PptxSlideMetadata): string {
  return [
    metadata.title,
    ...bodyLinesWithoutFirstTitle(metadata),
  ]
    .filter(Boolean)
    .join("\n");
}

function bodyLinesWithoutFirstTitle(
  metadata: PptxSlideMetadata,
): string[] {
  const lines = metadata.text.split("\n");
  if (!metadata.title) {
    return lines;
  }
  const titleIndex = lines.indexOf(metadata.title);
  if (titleIndex >= 0) {
    lines.splice(titleIndex, 1);
  }
  return lines;
}

function extractElementParagraphs(
  element: Element | null | undefined,
): string[] {
  if (!element) {
    return [];
  }
  const paragraphs = descendantsNamed(element, "p");
  if (paragraphs.length === 0) {
    return normalizeText([textContent(element)]);
  }
  return paragraphs.map((paragraph) =>
    descendantsNamed(paragraph, "t")
      .map((textElement) => textContent(textElement))
      .join(""),
  );
}

function normalizeText(values: readonly string[]): string[] {
  return values
    .map((value) => value.replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

function countNormalizedMatches(
  normalizedValue: string,
  normalizedQuery: string,
): number {
  let count = 0;
  let offset = 0;
  while (offset < normalizedValue.length) {
    const index = normalizedValue.indexOf(normalizedQuery, offset);
    if (index === -1) {
      break;
    }
    count += 1;
    offset = index + normalizedQuery.length;
  }
  return count;
}

function emptyMetadata(index: number): PptxSlideMetadata {
  return {
    index,
    title: "",
    text: "",
    notes: "",
  };
}

function createSnippet(value: string, normalizedQuery: string): string {
  const compact = value.replace(/\s+/g, " ").trim();
  const index = compact.toLocaleLowerCase().indexOf(normalizedQuery);
  if (index === -1) {
    return compact.slice(0, 120);
  }
  const start = Math.max(0, index - 40);
  const end = Math.min(compact.length, index + normalizedQuery.length + 60);
  return `${start > 0 ? "..." : ""}${compact.slice(start, end)}${
    end < compact.length ? "..." : ""
  }`;
}

function firstUsefulLine(value: string, title: string): string {
  return value
    .split("\n")
    .find((line) => line && line !== title)
    ?? "";
}
