export type PptxQueuedTask = (
  isCancelled: () => boolean,
) => Promise<void>;

interface QueueEntry {
  key: number;
  priority: number;
  sequence: number;
  run: PptxQueuedTask;
  cancelled: boolean;
}

export class PptxTaskQueue {
  private readonly queued = new Map<number, QueueEntry>();
  private readonly active = new Map<number, QueueEntry>();
  private sequence = 0;
  private paused = false;

  constructor(readonly concurrency = 2) {
    if (!Number.isInteger(concurrency) || concurrency < 1) {
      throw new RangeError("Task queue concurrency must be a positive integer.");
    }
  }

  get activeCount(): number {
    return this.active.size;
  }

  get queuedCount(): number {
    return this.queued.size;
  }

  schedule(key: number, priority: number, run: PptxQueuedTask): void {
    if (this.active.has(key)) {
      return;
    }
    const existing = this.queued.get(key);
    if (existing) {
      existing.priority = priority;
      existing.run = run;
    } else {
      this.queued.set(key, {
        key,
        priority,
        sequence: this.sequence,
        run,
        cancelled: false,
      });
      this.sequence += 1;
    }
    this.pump();
  }

  cancel(key: number): void {
    const queued = this.queued.get(key);
    if (queued) {
      queued.cancelled = true;
      this.queued.delete(key);
    }
    const active = this.active.get(key);
    if (active) {
      active.cancelled = true;
    }
  }

  setPaused(paused: boolean, cancelActive = false): void {
    this.paused = paused;
    if (paused && cancelActive) {
      for (const entry of this.active.values()) {
        entry.cancelled = true;
      }
    }
    if (!paused) {
      this.pump();
    }
  }

  clear(): void {
    for (const entry of this.queued.values()) {
      entry.cancelled = true;
    }
    for (const entry of this.active.values()) {
      entry.cancelled = true;
    }
    this.queued.clear();
  }

  private pump(): void {
    while (
      !this.paused &&
      this.active.size < this.concurrency &&
      this.queued.size > 0
    ) {
      const next = [...this.queued.values()].sort(
        (left, right) =>
          left.priority - right.priority ||
          left.sequence - right.sequence,
      )[0];
      this.queued.delete(next.key);
      if (next.cancelled) {
        continue;
      }
      this.active.set(next.key, next);
      void next.run(() => next.cancelled)
        .catch(() => undefined)
        .finally(() => {
          this.active.delete(next.key);
          this.pump();
        });
    }
  }
}
