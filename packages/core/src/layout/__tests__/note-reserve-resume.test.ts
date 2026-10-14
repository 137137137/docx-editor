// A pass whose footnote reserves moved resumes from the last checkpoint those reserves left
// valid, and reuses the previous tail only where the reserves moved with it. Every resumed
// pass must equal a fresh pass over the same inputs.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlPart } from '../../store/package/ooxml-tree.ts';
import {
  createFixedMeasurer,
  createLayoutSession,
  layoutSemanticDocument,
  type SemanticLayout,
} from '../index.ts';
import type { LayoutSession } from '../layout-session.ts';
import { takeMovedSeedReserves } from '../note-reserves.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFixedMeasurer(6, 14);
const sectPr =
  '<w:sectPr><w:pgSz w:w="4000" w:h="2400"/>' +
  '<w:pgMar w:top="200" w:right="200" w:bottom="200" w:left="200" w:header="100" w:footer="100"/>' +
  '</w:sectPr>';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const sectionEnd = `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;

function partOf(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const filler = (count: number, tag: string) =>
  Array.from({ length: count }, (_, index) => paragraph(`${tag} ${index}`)).join('');

/** The page whose text holds `text` on one line. */
function pageWith(layout: SemanticLayout, text: string): number {
  return layout.pages.findIndex((page) =>
    page.fragments.some(
      (fragment) =>
        fragment.kind === 'paragraph' &&
        fragment.lines.some((line) => line.spans.map((span) => span.text).join('') === text)
    )
  );
}

/** Every fragment's paragraph, text, page and position. */
function signature(layout: SemanticLayout): string[] {
  return layout.pages.flatMap((page, pageIndex) =>
    page.fragments.map((fragment) =>
      fragment.kind === 'paragraph'
        ? `${pageIndex}|${fragment.paragraphId}|${fragment.lines
            .flatMap((line) => line.spans.map((span) => span.text))
            .join('')}|${fragment.box.y}|${fragment.lines.length}`
        : `${pageIndex}|table`
    )
  );
}

function lay(
  part: OoxmlPart,
  reserves: ReadonlyMap<number, number>,
  session?: LayoutSession,
  fromPreviousPass = false
): SemanticLayout {
  return layoutSemanticDocument(part, 1, {
    measurer,
    producer: 'test',
    pageBottomReserves: reserves,
    ...(fromPreviousPass ? { pageBottomReservesFromPreviousPass: true } : {}),
    ...(session ? { session } : {}),
  });
}

describe('a section pass whose reserves moved', () => {
  test('resumes before the first page whose reserve moved, and matches a fresh pass', () => {
    const part = partOf(filler(120, 'p') + sectPr);
    const session = createLayoutSession();
    const before = new Map([
      [3, 30],
      [8, 20],
    ]);
    lay(part, before, session);
    for (const moved of [
      new Map([...before, [10, 25]]),
      new Map([
        [3, 30],
        [8, 40],
      ]),
      new Map([[1, 15]]),
    ]) {
      const resumed = lay(part, moved, session);
      expect(signature(resumed)).toEqual(signature(lay(part, moved)));
      // Pages before the first moved reserve were not laid out again. A flow that comes back
      // into step after it reuses the previous tail too.
      const first = Math.min(
        ...[...moved.keys(), ...before.keys()].filter(
          (page) => moved.get(page) !== before.get(page)
        )
      );
      expect(session.stats.reusedPages).toBeGreaterThanOrEqual(first - 1);
      expect(session.stats.placed).toBeLessThan(session.stats.total);
      before.clear();
      for (const [page, height] of moved) before.set(page, height);
    }
  });

  test('a page break reuses the tail only where the reserves moved with it', () => {
    const plain = partOf(filler(40, 'a') + filler(80, 'b') + sectPr);
    const broken = partOf(filler(40, 'a') + pageBreak + filler(80, 'b') + sectPr);
    const reserves = new Map([
      [2, 20],
      [14, 30],
      [20, 10],
    ]);
    const fresh = lay(plain, reserves);
    const breakPage = pageWith(fresh, 'b 0');
    expect(breakPage).toBeGreaterThan(2);
    const shifted = new Map(
      [...reserves].map(([page, height]) => [page > breakPage ? page + 1 : page, height] as const)
    );
    for (const next of [shifted, reserves]) {
      const session = createLayoutSession();
      lay(plain, reserves, session);
      const resumed = lay(broken, next, session);
      expect(signature(resumed)).toEqual(signature(lay(broken, next)));
    }
  });
});

describe('a seeded notes pass over several sections', () => {
  test("moves each later section's seed with its first page", () => {
    const body = (extra: string) =>
      filler(15, 's0') +
      extra +
      filler(15, 's0 tail') +
      sectionEnd +
      filler(30, 's1') +
      sectionEnd +
      filler(30, 's2') +
      sectPr;
    const before = partOf(body(''));
    const after = partOf(body(pageBreak));
    const session = createLayoutSession();
    const first = lay(before, new Map(), session);
    const s1Start = pageWith(first, 's1 0');
    expect(s1Start).toBeGreaterThan(0);
    // Reserves on the second and third sections, at the pages of the previous pass.
    const seed = new Map([
      [s1Start + 1, 20],
      [first.pages.length - 2, 30],
    ]);
    lay(before, seed, session);
    takeMovedSeedReserves(session);

    const resumed = lay(after, seed, session, true);
    const moved = takeMovedSeedReserves(session);
    expect(moved).toBeDefined();
    const shift = resumed.pages.length - first.pages.length;
    expect(shift).toBe(1);
    expect([...moved!]).toEqual([...seed].map(([page, height]) => [page + shift, height]));
    expect(signature(resumed)).toEqual(signature(lay(after, moved!)));
  });
});
