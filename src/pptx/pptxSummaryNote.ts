import type { PptxSlideMetadata } from "./pptxMetadata";
import type { PptxReaderText } from "./pptxI18n";
import {
  buildReaderSummaryFrontmatter,
  type ReaderSummarySource,
} from "../reader/summaryNote";

export type PresentationSummarySource = ReaderSummarySource;

export function buildPresentationSummaryNote(
  source: PresentationSummarySource,
  slides: readonly PptxSlideMetadata[],
  currentSlideIndex: number,
  text: PptxReaderText,
  createdAt: Date = new Date(),
): string {
  const currentSlide = slides[currentSlideIndex];
  const reference = currentSlide
    ? text.summaryNote.slideReference(
        currentSlide.index + 1,
        currentSlide.title,
      )
    : text.summaryNote.slideReference(currentSlideIndex + 1, "");

  return [
    ...buildReaderSummaryFrontmatter(source, {
      format: "pptx",
      legacyType: "presentation-note",
      createdAt,
      fields: {
        current_slide: currentSlideIndex + 1,
      },
    }),
    "",
    `# ${source.basename}`,
    "",
    `${text.summaryNote.sourceLabel}: [[${source.path}]]`,
    `${text.summaryNote.currentSlideLabel}: ${reference}`,
    "",
    `## ${text.summaryNote.summaryHeading}`,
    "",
    `## ${text.summaryNote.keySlidesHeading}`,
    "",
    `- [ ] ${reference}`,
    "",
    `## ${text.summaryNote.followUpsHeading}`,
    "",
    `## ${text.summaryNote.slideReferencesHeading}`,
    "",
    ...slides.map((slide) =>
      `- ${text.summaryNote.slideReference(slide.index + 1, slide.title)}`),
    "",
  ].join("\n");
}
