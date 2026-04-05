/** Crop stored as insets from preview container edges (same coordinate system as RN layout). */

export type CropInsets = {
  top: number;
  left: number;
  right: number;
  bottom: number;
};

export function clampNumber(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Clamp crop insets so the rectangle stays inside the letterboxed image rect
 * and respects a minimum size (preview coordinates).
 */
export function clampCropToLetterbox(
  c: CropInsets,
  cw: number,
  ch: number,
  ox: number,
  oy: number,
  vw: number,
  vh: number,
  minSize: number,
): CropInsets {
  let x1 = c.left;
  let y1 = c.top;
  let x2 = cw - c.right;
  let y2 = ch - c.bottom;

  const imgRight = ox + vw;
  const imgBottom = oy + vh;

  x1 = clampNumber(x1, ox, imgRight - minSize);
  y1 = clampNumber(y1, oy, imgBottom - minSize);
  x2 = clampNumber(x2, x1 + minSize, imgRight);
  y2 = clampNumber(y2, y1 + minSize, imgBottom);

  if (x2 - x1 < minSize) {
    x2 = Math.min(imgRight, x1 + minSize);
    x1 = x2 - minSize;
    if (x1 < ox) {
      x1 = ox;
      x2 = Math.min(imgRight, x1 + minSize);
    }
  }
  if (y2 - y1 < minSize) {
    y2 = Math.min(imgBottom, y1 + minSize);
    y1 = y2 - minSize;
    if (y1 < oy) {
      y1 = oy;
      y2 = Math.min(imgBottom, y1 + minSize);
    }
  }

  return {
    top: y1,
    left: x1,
    right: cw - x2,
    bottom: ch - y2,
  };
}
