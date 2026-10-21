// Minimum cuts can differ from the contextual segments used for unwrapped text.
import type { OoxmlProperty } from '@docx-editor.dev/core/store';
import { wordBoundaries } from './cjk-line-break.ts';
import { cjkParagraphBreaks, hasCjkText } from './cjk-paragraph-breaks.ts';
import {
  eastAsianLanguage,
  resolveCjkTypography,
  type CjkTypographySettings,
} from './cjk-typography.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import type { ResolvedTabStops } from './paragraph-tabs.ts';
import { withoutLineEndWhitespace } from './trailing-spaces.ts';

/** CJK minimums use paragraph analysis; other scripts retain ordinary word cuts. */
export function cjkAutofitBreaks(
  pieces: readonly FieldAwarePiece[],
  props: readonly OoxmlProperty[],
  settings?: CjkTypographySettings
) {
  const policy = resolveCjkTypography(props, settings);
  const visible = pieces.filter((piece) => !piece.style.hidden);
  if (
    !policy.characterWrap &&
    !visible.some((piece) => hasCjkText(piece.text) || eastAsianLanguage(piece.props))
  )
    return null;
  // This qualification selects the East Asian analysis, whose decisions ignore text context.
  // Ordinary cross-run dash cuts remain the caller's paragraphWordCuts, for both min and max.
  const breaks = cjkParagraphBreaks(visible, policy)!;
  return {
    opens: (piece: FieldAwarePiece) => breaks.decision(piece, 0, '', piece.text) === 'opens',
    boundaries: (piece: FieldAwarePiece) =>
      breaks
        .boundaries(piece)
        .filter(
          (to) =>
            to === piece.text.length ||
            breaks.decision(piece, to, piece.text.slice(0, to), piece.text.slice(to)) === 'opens'
        ),
  };
}

/** The next unwrapped stop, excluding stops reserved for the numbering marker. */
export function nextAutofitTabStop(tabs: ResolvedTabStops, at: number): number {
  for (const stop of tabs.stops)
    if (!stop.numberingOnly && stop.positionPt > at + 1e-6) return stop.positionPt;
  const interval = tabs.defaultIntervalPt;
  return (Math.floor(at / interval + 1e-6) + 1) * interval;
}

/** Preserve upstream contextual word cuts, including paragraph-wide dash boundaries. */
export function measureUnwrappedAutofitText(
  text: string,
  boundaries: readonly number[] | undefined,
  measure: (text: string) => number,
  start: number,
  nextTabStop: (at: number) => number
): { width: number; trailing: number } {
  let width = 0;
  let trailing = 0;
  let from = 0;
  for (const to of boundaries ?? wordBoundaries(text, false)) {
    const candidate = text.slice(from, to);
    from = to;
    if (candidate.length === 0) continue;
    let advance: number;
    let spaces = 0;
    if (candidate === '\t') {
      advance = Math.max(0, nextTabStop(start + width) - (start + width));
    } else {
      const ink = withoutLineEndWhitespace(candidate);
      advance = measure(candidate);
      if (ink.length < candidate.length) spaces = advance - (ink ? measure(ink) : 0);
    }
    width += advance;
    trailing = spaces < advance ? spaces : trailing + spaces;
  }
  return { width, trailing };
}
