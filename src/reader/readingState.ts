export const DEFAULT_READING_STATE_CAPACITY = 100;
const MAX_COLLAPSED_OUTLINE_IDS = 200;

export type ReaderFormat = "docx" | "pptx" | "xlsx";

export interface ReaderFileIdentity {
  path: string;
  mtime: number;
  format: ReaderFormat;
}

/**
 * Runtime state used by reader sessions. Persisted data is deliberately
 * converted to the narrower position/zoom/navigation shape below.
 */
export interface ReaderViewState {
  zoom: number;
  fitWidth: boolean;
  outlineVisible: boolean;
  scrollLeft: number;
  scrollTop: number;
  collapsedOutlineIds: string[];
  page?: number;
  notesVisible?: boolean;
}

export interface PersistedReaderViewState {
  path: string;
  mtime: number;
  format: ReaderFormat;
  lastAccessed: number;
  position: {
    scrollLeft: number;
    scrollTop: number;
    page?: number;
  };
  zoom: {
    scale: number;
    fit: boolean;
  };
  navigation: {
    panelVisible: boolean;
    collapsedIds: string[];
    notesVisible?: boolean;
  };
}

export class ReadingStateStore {
  private readonly entries = new Map<string, PersistedReaderViewState>();
  private accessCounter = 0;

  constructor(
    private readonly capacity = DEFAULT_READING_STATE_CAPACITY,
    serialized?: unknown,
  ) {
    for (const entry of normalizeStoredEntries(serialized)) {
      this.entries.set(entry.path, entry);
      this.accessCounter = Math.max(this.accessCounter, entry.lastAccessed);
    }
    this.enforceCapacity();
  }

  get(identity: ReaderFileIdentity): ReaderViewState | undefined;
  get(path: string): ReaderViewState | undefined;
  get(identityOrPath: ReaderFileIdentity | string): ReaderViewState | undefined {
    const identity =
      typeof identityOrPath === "string"
        ? null
        : normalizeIdentity(identityOrPath);
    const path =
      typeof identityOrPath === "string"
        ? identityOrPath
        : identityOrPath.path;
    const entry = this.entries.get(path);
    if (!entry) {
      return undefined;
    }

    entry.lastAccessed = this.nextAccess();
    if (identity && entry.mtime === 0 && entry.format === identity.format) {
      entry.mtime = identity.mtime;
      return stateFromEntry(entry);
    }
    if (
      identity &&
      (entry.mtime !== identity.mtime || entry.format !== identity.format)
    ) {
      const preservedPreference = stateFromEntry(entry);
      entry.mtime = identity.mtime;
      entry.format = identity.format;
      entry.position = {
        scrollLeft: 0,
        scrollTop: 0,
      };
      entry.navigation = {
        panelVisible: defaultPanelVisibility(identity.format),
        collapsedIds: [],
      };
      return {
        ...createDefaultState(identity.format),
        zoom: preservedPreference.zoom,
        fitWidth: preservedPreference.fitWidth,
      };
    }

    return stateFromEntry(entry);
  }

  set(
    identity: ReaderFileIdentity,
    state: ReaderViewState,
  ): void;
  set(path: string, state: ReaderViewState): void;
  set(
    identityOrPath: ReaderFileIdentity | string,
    state: ReaderViewState,
  ): void {
    const identity =
      typeof identityOrPath === "string"
        ? inferIdentity(identityOrPath)
        : normalizeIdentity(identityOrPath);
    if (!identity.path) {
      return;
    }

    this.entries.set(
      identity.path,
      entryFromState(identity, state, this.nextAccess()),
    );
    this.enforceCapacity();
  }

  serialize(): PersistedReaderViewState[] {
    return Array.from(this.entries.values())
      .sort((left, right) => left.lastAccessed - right.lastAccessed)
      .map(cloneEntry);
  }

  get size(): number {
    return this.entries.size;
  }

  private nextAccess(): number {
    this.accessCounter += 1;
    return this.accessCounter;
  }

  private enforceCapacity(): void {
    while (this.entries.size > Math.max(1, this.capacity)) {
      let oldest: PersistedReaderViewState | null = null;
      for (const entry of this.entries.values()) {
        if (!oldest || entry.lastAccessed < oldest.lastAccessed) {
          oldest = entry;
        }
      }
      if (!oldest) {
        return;
      }
      this.entries.delete(oldest.path);
    }
  }
}

export function normalizeReaderViewState(value: unknown): ReaderViewState {
  const source = isRecord(value) ? value : {};
  const state: ReaderViewState = {
    zoom: readFiniteNumber(source.zoom, 1),
    fitWidth: readBoolean(source.fitWidth, false),
    outlineVisible: readBoolean(source.outlineVisible, true),
    scrollLeft: Math.max(0, readFiniteNumber(source.scrollLeft, 0)),
    scrollTop: Math.max(0, readFiniteNumber(source.scrollTop, 0)),
    collapsedOutlineIds: readStringArray(source.collapsedOutlineIds).slice(
      0,
      MAX_COLLAPSED_OUTLINE_IDS,
    ),
  };
  const page = readFiniteNumber(source.page, 0);
  if (page >= 1) {
    state.page = Math.floor(page);
  }
  if (typeof source.notesVisible === "boolean") {
    state.notesVisible = source.notesVisible;
  }
  return state;
}

function normalizeStoredEntries(value: unknown): PersistedReaderViewState[] {
  if (!Array.isArray(value)) {
    return [];
  }

  const entries: PersistedReaderViewState[] = [];
  for (const item of value) {
    if (!isRecord(item) || typeof item.path !== "string" || !item.path) {
      continue;
    }

    const format = normalizeFormat(
      item.format,
      inferFormatFromPath(item.path),
    );
    const lastAccessed = readFiniteNumber(
      item.lastAccessed,
      entries.length + 1,
    );
    if (
      isRecord(item.position) &&
      isRecord(item.zoom) &&
      isRecord(item.navigation)
    ) {
      const normalizedState = normalizeReaderViewState({
        zoom: item.zoom.scale,
        fitWidth: item.zoom.fit,
        outlineVisible: item.navigation.panelVisible,
        scrollLeft: item.position.scrollLeft,
        scrollTop: item.position.scrollTop,
        collapsedOutlineIds: item.navigation.collapsedIds,
        page: item.position.page,
        notesVisible: item.navigation.notesVisible,
      });
      entries.push(
        entryFromState(
          {
            path: item.path,
            mtime: Math.max(0, readFiniteNumber(item.mtime, 0)),
            format,
          },
          normalizedState,
          lastAccessed,
        ),
      );
      continue;
    }

    // Pre-2.4 entries stored an unrestricted view state under `state`.
    entries.push(
      entryFromState(
        {
          path: item.path,
          mtime: Math.max(0, readFiniteNumber(item.mtime, 0)),
          format,
        },
        normalizeReaderViewState(item.state),
        lastAccessed,
      ),
    );
  }
  return entries;
}

function entryFromState(
  identity: ReaderFileIdentity,
  state: ReaderViewState,
  lastAccessed: number,
): PersistedReaderViewState {
  const normalized = normalizeReaderViewState(state);
  const position: PersistedReaderViewState["position"] = {
    scrollLeft: normalized.scrollLeft,
    scrollTop: normalized.scrollTop,
  };
  if (normalized.page !== undefined) {
    position.page = normalized.page;
  }
  const navigation: PersistedReaderViewState["navigation"] = {
    panelVisible: normalized.outlineVisible,
    collapsedIds: [...normalized.collapsedOutlineIds],
  };
  if (normalized.notesVisible !== undefined) {
    navigation.notesVisible = normalized.notesVisible;
  }

  return {
    path: identity.path,
    mtime: identity.mtime,
    format: identity.format,
    lastAccessed,
    position,
    zoom: {
      scale: normalized.zoom,
      fit: normalized.fitWidth,
    },
    navigation,
  };
}

function stateFromEntry(entry: PersistedReaderViewState): ReaderViewState {
  return normalizeReaderViewState({
    zoom: entry.zoom.scale,
    fitWidth: entry.zoom.fit,
    outlineVisible: entry.navigation.panelVisible,
    scrollLeft: entry.position.scrollLeft,
    scrollTop: entry.position.scrollTop,
    collapsedOutlineIds: entry.navigation.collapsedIds,
    page: entry.position.page,
    notesVisible: entry.navigation.notesVisible,
  });
}

function cloneEntry(
  entry: PersistedReaderViewState,
): PersistedReaderViewState {
  return {
    ...entry,
    position: { ...entry.position },
    zoom: { ...entry.zoom },
    navigation: {
      ...entry.navigation,
      collapsedIds: [...entry.navigation.collapsedIds],
    },
  };
}

function normalizeIdentity(
  identity: ReaderFileIdentity,
): ReaderFileIdentity {
  return {
    path: identity.path,
    mtime: Math.max(0, readFiniteNumber(identity.mtime, 0)),
    format: normalizeFormat(
      identity.format,
      inferFormatFromPath(identity.path),
    ),
  };
}

function inferIdentity(path: string): ReaderFileIdentity {
  return {
    path,
    mtime: 0,
    format: inferFormatFromPath(path),
  };
}

function inferFormatFromPath(path: string): ReaderFormat {
  const extension = path.split(".").pop()?.toLowerCase();
  if (extension === "pptx") {
    return "pptx";
  }
  if (extension === "xlsx") {
    return "xlsx";
  }
  return "docx";
}

function normalizeFormat(
  value: unknown,
  fallback: ReaderFormat,
): ReaderFormat {
  return value === "docx" || value === "pptx" || value === "xlsx"
    ? value
    : fallback;
}

function defaultPanelVisibility(format: ReaderFormat): boolean {
  return format !== "xlsx";
}

function createDefaultState(format: ReaderFormat): ReaderViewState {
  return {
    zoom: 1,
    fitWidth: format !== "docx",
    outlineVisible: defaultPanelVisibility(format),
    scrollLeft: 0,
    scrollTop: 0,
    collapsedOutlineIds: [],
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function readFiniteNumber(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function readBoolean(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function readStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.filter((item): item is string => typeof item === "string");
}
