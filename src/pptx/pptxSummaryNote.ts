import type { PptxSlideMetadata } from "./pptxMetadata";
import type { PptxReaderText } from "./pptxI18n";

export interface PresentationSummarySource {
  basename: string;
  path: string;
}

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
    "---",
    `source: "${escapeYamlString(source.path)}"`,
    "type: presentation-note",
    `created: ${formatLocalDate(createdAt)}`,
    `current_slide: ${currentSlideIndex + 1}`,
    "---",
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

function formatLocalDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function escapeYamlString(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}
