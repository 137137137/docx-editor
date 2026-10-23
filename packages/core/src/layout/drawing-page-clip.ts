import type { LayoutBox } from './semantic-records.ts';

/** The frame fields a page clip reads: a structural subset of the anchor frame context. */
export interface PageClipFrame {
  readonly pageWidth: number;
  readonly pageHeight?: number | undefined;
  readonly marginLeft: number;
  readonly contentInsetTop: number;
  readonly contentInsetBottom: number;
  readonly contentBandHeight: number;
}

/**
 * Full physical page clip, including margins and all columns. A continuous section's
 * content band can end before the sheet does; it must not crop page-relative artwork.
 * Callers without physical height retain the full-band fallback.
 */
export function pageClipRegion(frameBase: PageClipFrame): LayoutBox {
  const height = frameBase.pageHeight;
  return Object.freeze({
    x: -frameBase.marginLeft,
    y: -frameBase.contentInsetTop,
    width: frameBase.pageWidth,
    height:
      height !== undefined && Number.isFinite(height) && height > 0
        ? height
        : frameBase.contentBandHeight + frameBase.contentInsetTop + frameBase.contentInsetBottom,
  });
}

/**
 * The horizontal room a body inline picture in one column may paint into. Beside another column
 * it stops at the middle of the gap; at the outside of the column set it reaches the paper edge.
 * Physical neighbours are used, so right-to-left column order needs no special case.
 */
export function columnPaperClip(
  paper: LayoutBox,
  columns: { readonly lefts: readonly number[]; readonly widths: readonly number[] },
  index: number
): LayoutBox {
  const left = columns.lefts[index];
  const width = columns.widths[index];
  if (left === undefined || width === undefined || columns.lefts.length < 2) return paper;
  const right = left + width;
  let from = paper.x;
  let to = paper.x + paper.width;
  for (const [other, otherLeft] of columns.lefts.entries()) {
    if (other === index) continue;
    const otherRight = otherLeft + (columns.widths[other] ?? 0);
    if (otherRight <= left) from = Math.max(from, (otherRight + left) / 2);
    if (otherLeft >= right) to = Math.min(to, (right + otherLeft) / 2);
  }
  return Object.freeze({ ...paper, x: from, width: Math.max(0, to - from) });
}
