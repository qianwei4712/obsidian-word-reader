export class LruCache<Key, Value> {
  private readonly values = new Map<Key, Value>();

  constructor(readonly capacity: number) {
    if (!Number.isInteger(capacity) || capacity < 1) {
      throw new RangeError("LRU cache capacity must be a positive integer.");
    }
  }

  get size(): number {
    return this.values.size;
  }

  get(key: Key): Value | undefined {
    const value = this.values.get(key);
    if (value === undefined) {
      return undefined;
    }
    this.values.delete(key);
    this.values.set(key, value);
    return value;
  }

  set(key: Key, value: Value): Value {
    this.values.delete(key);
    this.values.set(key, value);
    while (this.values.size > this.capacity) {
      const oldest = this.values.keys().next().value as Key | undefined;
      if (oldest === undefined) {
        break;
      }
      this.values.delete(oldest);
    }
    return value;
  }

  getOrCreate(key: Key, factory: () => Value): Value {
    const existing = this.get(key);
    return existing === undefined ? this.set(key, factory()) : existing;
  }

  delete(key: Key): boolean {
    return this.values.delete(key);
  }

  clear(): void {
    this.values.clear();
  }
}
