import type { WorkspaceLeaf } from "obsidian";

import type WordReaderPlugin from "./main";
import {
  XLSX_VIEW_TYPE,
} from "./xlsx/XlsxAdapter";
import { XlsxSession } from "./xlsx/XlsxSession";

export const VIEW_TYPE_XLSX_READER = XLSX_VIEW_TYPE;

/**
 * Stable Obsidian entry point for the public read-only spreadsheet session.
 */
export class XlsxView extends XlsxSession {
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf, plugin);
  }
}
