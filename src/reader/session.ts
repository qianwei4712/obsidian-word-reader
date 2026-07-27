import type { TFile } from "obsidian";

import type { OfficeReaderAdapter } from "./adapters";
import type { ReaderCapabilities } from "./capabilities";

export interface ReaderSession {
  readonly adapter: OfficeReaderAdapter<unknown, unknown>;
  readonly capabilities: ReaderCapabilities;
  readonly file: TFile | null;

  onOpen: () => Promise<void>;
  onClose: () => Promise<void>;
  onLoadFile: (file: TFile) => Promise<void>;
  onUnloadFile: (file: TFile) => Promise<void>;
  reload: () => Promise<void>;
  refreshInterfaceLanguage: () => void;
  copyText?: () => Promise<void>;
  copyMarkdown?: () => Promise<void>;
  createSummaryNote?: () => Promise<void>;
  copyDiagnostics?: () => Promise<void>;
  openExternal?: () => Promise<void>;
  previousPage?: () => Promise<void>;
  nextPage?: () => Promise<void>;
  toggleNotes?: () => void;
  toggleFullscreen?: () => Promise<void>;
  focusSearch?: () => void;
}

export type ReaderCommand =
  | "reload"
  | "copyText"
  | "copyMarkdown"
  | "createSummaryNote"
  | "copyDiagnostics"
  | "openExternal"
  | "previousPage"
  | "nextPage"
  | "toggleNotes"
  | "toggleFullscreen"
  | "focusSearch";

export function isReaderSession(value: unknown): value is ReaderSession {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const candidate = value as Partial<ReaderSession>;
  return (
    typeof candidate.reload === "function" &&
    typeof candidate.refreshInterfaceLanguage === "function" &&
    typeof candidate.adapter === "object" &&
    candidate.adapter !== null &&
    typeof candidate.capabilities === "object" &&
    candidate.capabilities !== null
  );
}
