export type ReaderCapability =
  | "reload"
  | "zoom"
  | "fit"
  | "navigation"
  | "search"
  | "copyText"
  | "copyMarkdown"
  | "notes"
  | "summaryNote"
  | "diagnostics"
  | "fullscreen"
  | "paged"
  | "openExternal";

export type ReaderCapabilities = Readonly<
  Record<ReaderCapability, boolean>
>;

const NO_CAPABILITIES: ReaderCapabilities = {
  reload: false,
  zoom: false,
  fit: false,
  navigation: false,
  search: false,
  copyText: false,
  copyMarkdown: false,
  notes: false,
  summaryNote: false,
  diagnostics: false,
  fullscreen: false,
  paged: false,
  openExternal: false,
};

export function defineReaderCapabilities(
  capabilities: Partial<ReaderCapabilities>,
): ReaderCapabilities {
  return Object.freeze({
    ...NO_CAPABILITIES,
    ...capabilities,
  });
}

export function hasReaderCapability(
  capabilities: ReaderCapabilities,
  capability: ReaderCapability,
): boolean {
  return capabilities[capability];
}
