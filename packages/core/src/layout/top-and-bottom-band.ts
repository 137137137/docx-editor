// Where a wrapTopAndBottom band applies, and where it sits relative to its anchor.

import type { DrawingProjection, ImageWrapTarget } from '../store/package/drawing-projection.ts';
import { emuToPointsSafe } from './drawing-layout.ts';

/**
 * Whether `behindDoc` takes the drawing out of the text flow. It orders paint only
 * (§20.4.2.3): Word still clears a `wrapTopAndBottom` band for a drawing behind the text,
 * the way a masthead box set behind its page title pushes the title below it.
 */
export function behindDocumentSkipsWrap(
  behindDocument: boolean | undefined,
  wrap: ImageWrapTarget
): boolean {
  return behindDocument === true && wrap !== 'topAndBottom';
}

/**
 * Where a topAndBottom band starts below its paragraph's origin.
 *
 * A drawing offset from its paragraph or line (`positionV relativeFrom="paragraph"` or
 * `"line"`, by `posOffset`) is published at that offset, so its band is too: the lines
 * above the band keep their place and only the ones it crosses move below it. Any other
 * frame, or an aligned position, keeps the band at the top of the anchor's line.
 */
export function topAndBottomBandOffset(projection: DrawingProjection, lineTop: number): number {
  const vertical = projection.position?.vertical;
  if (!vertical || projection.anchor?.simplePos || vertical.align !== null) return lineTop;
  const offset = emuToPointsSafe(vertical.offsetEmu ?? 0);
  if (offset === null) return lineTop;
  if (vertical.relativeFrom === 'paragraph') return offset;
  if (vertical.relativeFrom === 'line') return lineTop + offset;
  return lineTop;
}
