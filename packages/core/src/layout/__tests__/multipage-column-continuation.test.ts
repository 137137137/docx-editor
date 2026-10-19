import { test, expect } from 'bun:test';
import { readOoxmlPart } from '../../store/package/ooxml-tree.ts';
import {
  layoutSemanticDocument,
  createFixedMeasurer,
  createLayoutSession,
} from '../semantic-layout.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFixedMeasurer(6, 14);
const sect = (columns = 1, type = 'continuous') =>
  `<w:sectPr><w:type w:val="${type}"/><w:pgSz w:w="6000" w:h="2400"/><w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200" w:header="0" w:footer="0"/><w:cols w:num="${columns}" w:space="200"/></w:sectPr>`;
const p = (text: string, height = 30, section = '') =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${height * 20}" w:lineRule="exact"/>${section}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const load = (body: string) => {
  const value = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/xml',
  });
  if (!value.ok) throw Error(value.reason);
  return value.part;
};
const part = (count = 10, columns = 2, followHeight = 30, type = 'continuous') =>
  load(
    Array.from({ length: count }, (_, i) => p(`B${i}`)).join('') +
      p('', 1, sect(columns)) +
      p('FOLLOW', followHeight) +
      sect(1, type)
  );
const lay = (
  value: ReturnType<typeof load>,
  session?: ReturnType<typeof createLayoutSession>,
  revision = 1
) => layoutSemanticDocument(value, revision, { measurer, ...(session ? { session } : {}) });
const rows = (layout: ReturnType<typeof lay>, name: string) =>
  layout.pages.flatMap((page, pageIndex) =>
    page.fragments.flatMap((f) =>
      f.kind === 'paragraph' && f.lines.flatMap((l) => l.spans.map((s) => s.text)).join('') === name
        ? [{ pageIndex, fragment: f }]
        : []
    )
  );
const only = (layout: ReturnType<typeof lay>, name: string) => {
  const found = rows(layout, name);
  expect(found).toHaveLength(1);
  return found[0]!;
};
const safe = (layout: ReturnType<typeof lay>) => {
  const follow = only(layout, 'FOLLOW');
  for (const f of layout.pages[follow.pageIndex]!.fragments) {
    if (
      f.kind === 'paragraph' &&
      !f.outOfFlow &&
      f.paragraphId !== follow.fragment.paragraphId &&
      f.lines.some((l) => l.spans.some((s) => s.text.trim()))
    )
      expect(f.box.y + f.box.height).toBeLessThanOrEqual(follow.fragment.box.y + 0.000001);
  }
};

test('a short full-width successor clears both columns on a multi-page host', () => {
  const value = lay(part(10, 2, 8));
  expect(value.pages).toHaveLength(2);
  const follow = only(value, 'FOLLOW');
  expect(follow.pageIndex).toBe(1);
  expect(follow.fragment.box.y).toBe(90);
  safe(value);
});
test('a three-column host clears its tallest last-page column', () => {
  const value = lay(part(13, 3, 8));
  expect(value.pages).toHaveLength(2);
  const follow = only(value, 'FOLLOW');
  expect(follow.pageIndex).toBe(1);
  expect(follow.fragment.box.y).toBe(90);
  safe(value);
});
test('continuation uses the final sheet rather than the first full sheet', () => {
  const value = lay(part(14, 2, 8));
  expect(value.pages).toHaveLength(3);
  const follow = only(value, 'FOLLOW');
  expect(follow.pageIndex).toBe(2);
  expect(follow.fragment.box.y).toBe(60);
  safe(value);
});
test('equal-height final columns do not create an extra page', () => {
  const value = lay(part(12, 2, 8));
  expect(value.pages).toHaveLength(2);
  expect(only(value, 'FOLLOW').fragment.box.y).toBe(90);
  safe(value);
});
test('an explicit next-page break keeps the successor at the new page origin', () => {
  const value = lay(part(10, 2, 8, 'nextPage'));
  expect(value.pages).toHaveLength(3);
  const follow = only(value, 'FOLLOW');
  expect(follow.pageIndex).toBe(2);
  expect(follow.fragment.box.y).toBe(0);
  safe(value);
});
test('single-column sections retain natural continuation', () => {
  const value = lay(part(4, 1, 8));
  expect(value.pages).toHaveLength(2);
  const follow = only(value, 'FOLLOW');
  expect(follow.pageIndex).toBe(1);
  expect(follow.fragment.box.y).toBe(30);
  safe(value);
});
test('a one-page balanced section retains its compact continuation', () => {
  const value = lay(part(4, 2, 8));
  expect(value.pages).toHaveLength(1);
  safe(value);
});
test('warm and cold continuation agree after the preceding section grows', () => {
  const session = createLayoutSession();
  lay(part(10, 2, 8), session, 1);
  const changed = part(11, 2, 8),
    warm = lay(changed, session, 2),
    cold = lay(changed, undefined, 2);
  expect(warm.pages).toEqual(cold.pages);
  expect(only(warm, 'FOLLOW').fragment.box.y).toBe(90);
  safe(warm);
});
