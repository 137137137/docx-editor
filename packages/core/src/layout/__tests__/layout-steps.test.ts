// Layout as steps that pause between blocks.
//
// The steps are the same statements as the plain call, so what they lay out must be the same
// pages, however the pauses fall: drained at once, stepped one pause at a time, or interleaved
// with another run on a session of its own. What these tests pin down, on a document with two
// sections and a table longer than a page: the steps pause between blocks and between table
// rows, they lay out exactly what the plain call lays out, and a cut to the first blocks lays
// out the pages before the cut as the whole body does.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { readOoxmlPackage, type OoxmlPart } from '@docx-editor.dev/core/store';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { createLayoutSession } from '../layout-session.ts';
import { drainLayoutSteps, type LayoutSteps } from '../layout-steps.ts';
import { layoutSemanticDocument, layoutSemanticDocumentSteps } from '../semantic-layout.ts';
import type { BlockFragmentRecord, PageRecord, SemanticLayout } from '../semantic-records.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const paragraph = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const prose = (from: number, count: number) =>
  Array.from({ length: count }, (_, index) =>
    paragraph(`paragraph ${from + index} with enough words to wrap across one line `.repeat(3))
  ).join('');
const table = (rows: number) =>
  '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  Array.from(
    { length: rows },
    (_, index) =>
      `<w:tr><w:tc>${paragraph(`row ${index} left`)}</w:tc><w:tc>${paragraph(`row ${index} right`)}</w:tc></w:tr>`
  ).join('') +
  '</w:tbl>';
const sectionBreak =
  '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr></w:p>';

/** Two sections: prose, a table longer than a page, and prose again under a narrower page. */
function documentPart(): OoxmlPart {
  const body =
    prose(0, 60) +
    table(90) +
    sectionBreak +
    prose(60, 60) +
    '<w:sectPr><w:pgSz w:w="10000" w:h="15840"/></w:sectPr>';
  const loaded = readOoxmlPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
      ),
    })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.package.parts.get(loaded.package.mainDocumentPart)!;
}

/** What a page shows: its geometry, and each fragment's kind, owner, place and lines. */
function pageShape(page: PageRecord) {
  const fragment = (entry: BlockFragmentRecord) => ({
    kind: entry.kind,
    id:
      entry.kind === 'paragraph' ? entry.paragraphId : entry.kind === 'table' ? entry.tableId : '',
    box: entry.box,
    lines: entry.kind === 'paragraph' ? entry.lines.map((line) => line.text) : [],
    rows: entry.kind === 'table' ? entry.rows.length : 0,
  });
  return { index: page.index, box: page.box, fragments: page.fragments.map(fragment) };
}

const shapes = (layout: SemanticLayout) => layout.pages.map(pageShape);

function layoutOptions() {
  return { measurer: createFixedMeasurer(), session: createLayoutSession() };
}

/** Run `steps` one pause at a time, counting the pauses. */
function stepped<T>(steps: LayoutSteps<T>): { value: T; pauses: number } {
  let pauses = 0;
  for (;;) {
    const next = steps.next();
    if (next.done) return { value: next.value, pauses };
    pauses += 1;
  }
}

describe('layout steps', () => {
  const part = documentPart();
  const plain = layoutSemanticDocument(part, 1, layoutOptions());

  test('the fixture spans both sections and splits its table across pages', () => {
    expect(plain.pages.length).toBeGreaterThan(8);
    const tablePages = plain.pages.filter((page) =>
      page.fragments.some((fragment) => fragment.kind === 'table')
    );
    expect(tablePages.length).toBeGreaterThan(1);
  });

  test('pause between blocks and between table rows, and lay out what the plain call does', () => {
    const { value, pauses } = stepped(layoutSemanticDocumentSteps(part, 1, layoutOptions()));
    // 121 paragraphs and a table, each prepared and placed, and 90 rows.
    expect(pauses).toBeGreaterThan(2 * 122 + 90);
    expect(shapes(value)).toEqual(shapes(plain));
  });

  test('two runs interleaved on sessions of their own each lay out what they would alone', () => {
    const first = layoutSemanticDocumentSteps(part, 1, layoutOptions());
    const second = layoutSemanticDocumentSteps(part, 1, layoutOptions());
    let a: SemanticLayout | null = null;
    let b: SemanticLayout | null = null;
    while (!a || !b) {
      if (!a) {
        const next = first.next();
        if (next.done) a = next.value;
      }
      if (!b) {
        for (let turn = 0; turn < 3 && !b; turn += 1) {
          const next = second.next();
          if (next.done) b = next.value;
        }
      }
    }
    expect(shapes(a)).toEqual(shapes(plain));
    expect(shapes(b)).toEqual(shapes(plain));
  });

  test('a drained run of a session laid out before resumes as the plain call does', () => {
    const options = layoutOptions();
    layoutSemanticDocument(part, 1, options);
    const again = drainLayoutSteps(layoutSemanticDocumentSteps(part, 1, options));
    expect(again.pages).toEqual(options.session.previous!.pages);
    expect(shapes(again)).toEqual(shapes(plain));
  });

  test('a cut to the first blocks lays out the pages before its last page as the body does', () => {
    for (const limit of [30, 45, 61, 70, 100]) {
      const cut = layoutSemanticDocument(part, 1, { ...layoutOptions(), bodyBlockLimit: limit });
      let frontier = cut.pages.length - 1;
      while (frontier > 0 && cut.pages[frontier]!.fragments.length === 0) frontier -= 1;
      expect(frontier).toBeGreaterThan(0);
      expect(shapes(cut).slice(0, frontier)).toEqual(shapes(plain).slice(0, frontier));
      expect(cut.pages.length).toBeLessThan(plain.pages.length);
    }
  });
});
