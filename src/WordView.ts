import type { WorkspaceLeaf } from "obsidian";

import type WordReaderPlugin from "./main";
import {
  DOCX_VIEW_TYPE,
} from "./docx/DocxAdapter";
import { DocxSession } from "./docx/DocxSession";

export const VIEW_TYPE_WORD_READER = DOCX_VIEW_TYPE;

/**
 * Stable Obsidian view entry point. DOCX behavior lives in DocxSession so the
 * compatibility view type can remain unchanged while new formats share the
 * same adapter/session contract.
 */
export class WordView extends DocxSession {
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf, plugin);
  }
}
