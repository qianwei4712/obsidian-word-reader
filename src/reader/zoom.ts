export interface ZoomRange {
  min: number;
  max: number;
  step: number;
}

export interface ScrollPoint {
  left: number;
  top: number;
}

export function normalizeZoom(value: number, range: ZoomRange): number | null {
  if (!Number.isFinite(value)) {
    return null;
  }

  const rounded = Math.round(value / range.step) * range.step;
  return Math.min(Math.max(rounded, range.min), range.max);
}

export function preserveZoomAnchor(
  scroll: ScrollPoint,
  pointer: ScrollPoint,
  previousZoom: number,
  nextZoom: number,
): ScrollPoint {
  if (previousZoom <= 0 || nextZoom <= 0) {
    return scroll;
  }

  const contentLeft = (scroll.left + pointer.left) / previousZoom;
  const contentTop = (scroll.top + pointer.top) / previousZoom;
  return {
    left: contentLeft * nextZoom - pointer.left,
    top: contentTop * nextZoom - pointer.top,
  };
}
