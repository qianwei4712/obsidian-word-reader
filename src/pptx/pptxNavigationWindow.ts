export interface PptxNavigationWindowResult {
  indices: number[];
  start: number;
  end: number;
  topSpacer: number;
  bottomSpacer: number;
}

export class PptxNavigationWindow {
  constructor(
    readonly rowHeight = 196,
    readonly overscan = 4,
    readonly maxMounted = 60,
  ) {}

  calculate(
    filteredIndices: readonly number[],
    scrollTop: number,
    viewportHeight: number,
  ): PptxNavigationWindowResult {
    if (filteredIndices.length === 0) {
      return {
        indices: [],
        start: 0,
        end: 0,
        topSpacer: 0,
        bottomSpacer: 0,
      };
    }

    const visibleStart = Math.max(
      0,
      Math.floor(Math.max(0, scrollTop) / this.rowHeight),
    );
    const visibleCount = Math.max(
      1,
      Math.ceil(Math.max(this.rowHeight, viewportHeight) / this.rowHeight),
    );
    let start = Math.max(0, visibleStart - this.overscan);
    let end = Math.min(
      filteredIndices.length,
      visibleStart + visibleCount + this.overscan,
    );

    if (end - start > this.maxMounted) {
      end = start + this.maxMounted;
    }
    if (end === filteredIndices.length) {
      start = Math.max(0, end - this.maxMounted);
    }

    return {
      indices: filteredIndices.slice(start, end),
      start,
      end,
      topSpacer: start * this.rowHeight,
      bottomSpacer: (filteredIndices.length - end) * this.rowHeight,
    };
  }

  scrollTopForIndex(
    filteredIndices: readonly number[],
    slideIndex: number,
  ): number | null {
    const position = filteredIndices.indexOf(slideIndex);
    return position < 0 ? null : position * this.rowHeight;
  }
}
