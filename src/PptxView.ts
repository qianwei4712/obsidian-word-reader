import type { WorkspaceLeaf } from "obsidian";

import type WordReaderPlugin from "./main";
import {
  PPTX_VIEW_TYPE,
} from "./pptx/PptxAdapter";
import { PptxSession } from "./pptx/PptxSession";

export const VIEW_TYPE_PPTX_READER = PPTX_VIEW_TYPE;

/**
 * Stable Obsidian view entry point. Presentation behavior is isolated in the
 * PPTX session and selected through its adapter capabilities.
 */
export class PptxView extends PptxSession {
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf, plugin);
  }
}
