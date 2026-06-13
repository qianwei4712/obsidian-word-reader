export interface PptxThumbnailWindowChange {
  mount: number[];
  unmount: number[];
  mounted: number[];
}

export class PptxThumbnailWindow {
  private readonly visible = new Set<number>();
  private readonly mounted = new Set<number>();

  constructor(
    private slideCount: number,
    private readonly overscan = 1,
    private readonly maxMounted = 8,
  ) {}

  setSlideCount(slideCount: number): PptxThumbnailWindowChange {
    this.slideCount = Math.max(0, slideCount);
    for (const index of [...this.visible]) {
      if (!this.isValid(index)) {
        this.visible.delete(index);
      }
    }
    return this.reconcile();
  }

  setVisible(index: number, visible: boolean): void {
    if (!this.isValid(index)) {
      return;
    }
    if (visible) {
      this.visible.add(index);
    } else {
      this.visible.delete(index);
    }
  }

  update(priorityIndex?: number): PptxThumbnailWindowChange {
    return this.reconcile(priorityIndex);
  }

  reset(): PptxThumbnailWindowChange {
    this.visible.clear();
    const unmount = [...this.mounted].sort((a, b) => a - b);
    this.mounted.clear();
    return { mount: [], unmount, mounted: [] };
  }

  get mountedIndices(): number[] {
    return [...this.mounted].sort((a, b) => a - b);
  }

  private reconcile(priorityIndex?: number): PptxThumbnailWindowChange {
    const desired = this.buildDesired(priorityIndex);
    const mount = [...desired]
      .filter((index) => !this.mounted.has(index))
      .sort((a, b) => a - b);
    const unmount = [...this.mounted]
      .filter((index) => !desired.has(index))
      .sort((a, b) => a - b);

    for (const index of unmount) {
      this.mounted.delete(index);
    }
    for (const index of mount) {
      this.mounted.add(index);
    }

    return {
      mount,
      unmount,
      mounted: this.mountedIndices,
    };
  }

  private buildDesired(priorityIndex?: number): Set<number> {
    const candidates = new Set<number>();
    for (const index of this.visible) {
      this.addRange(candidates, index);
    }
    if (priorityIndex !== undefined && this.isValid(priorityIndex)) {
      this.addRange(candidates, priorityIndex);
    }

    const anchors =
      this.visible.size > 0
        ? [...this.visible]
        : priorityIndex !== undefined && this.isValid(priorityIndex)
          ? [priorityIndex]
          : [];
    const hasPriority =
      priorityIndex !== undefined && this.isValid(priorityIndex);
    const ranked = [...candidates].sort((left, right) => {
      const leftDistance = hasPriority
        ? Math.abs(left - priorityIndex)
        : nearestDistance(left, anchors);
      const rightDistance = hasPriority
        ? Math.abs(right - priorityIndex)
        : nearestDistance(right, anchors);
      return leftDistance - rightDistance || left - right;
    });
    return new Set(ranked.slice(0, this.maxMounted));
  }

  private addRange(target: Set<number>, center: number): void {
    for (
      let index = center - this.overscan;
      index <= center + this.overscan;
      index += 1
    ) {
      if (this.isValid(index)) {
        target.add(index);
      }
    }
  }

  private isValid(index: number): boolean {
    return Number.isInteger(index) && index >= 0 && index < this.slideCount;
  }
}

function nearestDistance(index: number, anchors: readonly number[]): number {
  return anchors.reduce(
    (distance, anchor) => Math.min(distance, Math.abs(index - anchor)),
    Number.POSITIVE_INFINITY,
  );
}
