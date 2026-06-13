export const DEFAULT_READING_STATE_CAPACITY = 50;
const MAX_COLLAPSED_OUTLINE_IDS = 200;

export interface ReaderViewState {
  zoom: number;
  fitWidth: boolean;
  outlineVisible: boolean;
  scrollLeft: number;
  scrollTop: number;
  collapsedOutlineIds: string[];
  page?: number;
}

export interface PersistedReaderViewState {
  path: string;
  lastAccessed: number;
  state: ReaderViewState;
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

  get(path: string): ReaderViewState | undefined {
    const entry = this.entries.get(path);
    if (!entry) {
      return undefined;
    }

    entry.lastAccessed = this.nextAccess();
    return cloneState(entry.state);
  }

  set(path: string, state: ReaderViewState): void {
    if (!path) {
      return;
    }

    this.entries.set(path, {
      path,
      lastAccessed: this.nextAccess(),
      state: normalizeReaderViewState(state),
    });
    this.enforceCapacity();
  }

  serialize(): PersistedReaderViewState[] {
    return Array.from(this.entries.values())
      .sort((left, right) => left.lastAccessed - right.lastAccessed)
      .map((entry) => ({
        path: entry.path,
        lastAccessed: entry.lastAccessed,
        state: cloneState(entry.state),
      }));
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

    entries.push({
      path: item.path,
      lastAccessed: readFiniteNumber(item.lastAccessed, entries.length + 1),
      state: normalizeReaderViewState(item.state),
    });
  }
  return entries;
}

function cloneState(state: ReaderViewState): ReaderViewState {
  return {
    ...state,
    collapsedOutlineIds: [...state.collapsedOutlineIds],
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
