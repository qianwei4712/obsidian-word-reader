import assert from "node:assert/strict";
import test from "node:test";

import {
  annotateDocxHeadingLevels,
  DOCX_HEADING_LEVEL_ATTRIBUTE,
} from "../src/renderer/docxOutline";

const HEADING_ATTRIBUTE_KEY = `$${DOCX_HEADING_LEVEL_ATTRIBUTE}`;

void test("annotates direct and style-based DOCX outline levels", () => {
  const directStyle: Record<string, unknown> = {};
  const localizedStyle: Record<string, unknown> = {};
  const inheritedStyle: Record<string, unknown> = {};
  const bodyTextStyle: Record<string, unknown> = {};
  const deepStyle: Record<string, unknown> = {};
  const document = {
    stylesPart: {
      styles: [
        {
          id: "1",
          name: "标题 1",
          target: "p",
          paragraphProps: { outlineLevel: 0 },
        },
        {
          id: "Heading2",
          name: "heading 2",
          target: "p",
          paragraphProps: { outlineLevel: 1 },
        },
        {
          id: "CustomSection",
          name: "章节标题",
          target: "p",
          basedOn: "Heading2",
        },
        {
          id: "BodyOverride",
          target: "p",
          basedOn: "Heading2",
          paragraphProps: { outlineLevel: 9 },
        },
        {
          id: "Heading9",
          target: "p",
        },
      ],
    },
    documentPart: {
      body: {
        type: "document",
        children: [
          {
            type: "paragraph",
            outlineLevel: 2,
            cssStyle: directStyle,
          },
          {
            type: "paragraph",
            styleName: "1",
            cssStyle: localizedStyle,
          },
          {
            type: "table",
            children: [
              {
                type: "paragraph",
                styleName: "CustomSection",
                cssStyle: inheritedStyle,
              },
            ],
          },
          {
            type: "paragraph",
            styleName: "BodyOverride",
            cssStyle: bodyTextStyle,
          },
          {
            type: "paragraph",
            styleName: "Heading9",
            cssStyle: deepStyle,
          },
        ],
      },
    },
  };

  assert.equal(annotateDocxHeadingLevels(document), 4);
  assert.equal(directStyle[HEADING_ATTRIBUTE_KEY], "3");
  assert.equal(localizedStyle[HEADING_ATTRIBUTE_KEY], "1");
  assert.equal(inheritedStyle[HEADING_ATTRIBUTE_KEY], "2");
  assert.equal(bodyTextStyle[HEADING_ATTRIBUTE_KEY], undefined);
  assert.equal(deepStyle[HEADING_ATTRIBUTE_KEY], "9");
});

void test("handles missing parts and cyclic style inheritance", () => {
  assert.equal(annotateDocxHeadingLevels({}), 0);

  const paragraphStyle: Record<string, unknown> = {};
  const document = {
    stylesPart: {
      styles: [
        { id: "CycleA", target: "p", basedOn: "CycleB" },
        { id: "CycleB", target: "p", basedOn: "CycleA" },
      ],
    },
    documentPart: {
      body: {
        children: [
          {
            type: "paragraph",
            styleName: "CycleA",
            cssStyle: paragraphStyle,
          },
        ],
      },
    },
  };

  assert.equal(annotateDocxHeadingLevels(document), 0);
  assert.equal(paragraphStyle[HEADING_ATTRIBUTE_KEY], undefined);
});
