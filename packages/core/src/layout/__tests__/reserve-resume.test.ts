// A reserve round that resumes below its first moved reserve lays out what a full round does.
import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { readOoxmlPackage } from '../../store/package/ooxml-package.ts';
import { resolveNotesPart } from '../../store/package/note-references.ts';
import {
  resolveEndnoteProperties,
  resolveFootnoteProperties,
} from '../../store/package/note-properties.ts';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { createLayoutSession } from '../layout-session.ts';
import type { NotesLayoutInput } from '../note-pagination.ts';
import { suspendReserveResumeForTest } from '../reserve-resume.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SECT =
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';

/** Words in a deterministic but uneven rhythm, so paragraphs and notes break differently. */
const words = (seed: number, count: number): string =>
  Array.from(
    { length: count },
    (_, i) => ['law', 'review', 'citation', 'a', 'footnote'][(seed + i * 7) % 5]
  ).join(' ');

function footnoteDoc(input: {
  readonly paragraphs: number;
  readonly sectionBreakEvery?: number;
  readonly tableEvery?: number;
}): Uint8Array {
  let body = '';
  let noteId = 0;
  const notes: string[] = [];
  for (let i = 0; i < input.paragraphs; i += 1) {
    let runs = `<w:r><w:t xml:space="preserve">${words(i, 30 + ((i * 13) % 50))} </w:t></w:r>`;
    // Two references in most paragraphs, notes from one line to several.
    for (let ref = 0; ref < (i % 4 === 0 ? 1 : 2); ref += 1) {
      noteId += 1;
      runs += `<w:r><w:footnoteReference w:id="${noteId}"/></w:r><w:r><w:t xml:space="preserve"> ${words(noteId, 6)} </w:t></w:r>`;
      notes.push(
        `<w:footnote w:id="${noteId}"><w:p><w:r><w:t>${words(noteId * 3, 8 + ((noteId * 17) % 70))}</w:t></w:r></w:p></w:footnote>`
      );
    }
    body += `<w:p>${runs}</w:p>`;
    if (input.tableEvery && i % input.tableEvery === input.tableEvery - 1) {
      const cell = (n: number) =>
        `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${words(n, 12)}</w:t></w:r></w:p></w:tc>`;
      body += `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${Array.from({ length: 4 }, (_, r) => `<w:tr>${cell(i + r)}${cell(i + r + 1)}</w:tr>`).join('')}</w:tbl>`;
    }
    if (input.sectionBreakEvery && i % input.sectionBreakEvery === input.sectionBreakEvery - 1) {
      body += `<w:p><w:pPr><w:sectPr>${SECT}</w:sectPr></w:pPr></w:p>`;
    }
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rIdFn" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>${SECT}</w:sectPr></w:body></w:document>`
    ),
    'word/footnotes.xml': strToU8(
      `<w:footnotes xmlns:w="${W}">` +
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
        `${notes.join('')}</w:footnotes>`
    ),
  });
}

function layOut(bytes: Uint8Array, resume: boolean) {
  const loaded = readOoxmlPackage(bytes);
  if (!loaded.ok) throw new Error(loaded.reason);
  const part = loaded.package.parts.get(loaded.package.mainDocumentPart)!;
  const documentFootnoteProps = resolveFootnoteProperties(undefined, undefined);
  const documentEndnoteProps = resolveEndnoteProperties(undefined, undefined);
  const notes: NotesLayoutInput = {
    footnotesPart: resolveNotesPart(loaded.package, 'footnote'),
    endnotesPart: null,
    footnotePropsBySection: [documentFootnoteProps],
    endnotePropsBySection: [documentEndnoteProps],
    documentFootnoteProps,
    documentEndnoteProps,
    measurer: createFixedMeasurer(),
    producer: 'reserve-resume-test',
  };
  const session = createLayoutSession();
  const restore = resume ? () => {} : suspendReserveResumeForTest();
  try {
    const layout = layoutSemanticDocument(part, 1, {
      measurer: notes.measurer,
      notes,
      session,
      producer: 'reserve-resume-test',
    });
    return { layout, session };
  } finally {
    restore();
  }
}

describe('reserve rounds resumed below the first moved reserve', () => {
  const cases = [
    { name: 'one section', doc: { paragraphs: 60 } },
    { name: 'tables between paragraphs', doc: { paragraphs: 48, tableEvery: 6 } },
    { name: 'several sections', doc: { paragraphs: 60, sectionBreakEvery: 20 } },
  ];
  for (const { name, doc } of cases) {
    test(`${name}: the same pages as rounds laid out from the first page`, () => {
      const bytes = footnoteDoc(doc);
      const full = layOut(bytes, false);
      const resumed = layOut(bytes, true);
      expect(full.layout.pages.length).toBeGreaterThan(5);
      expect(JSON.stringify(resumed.layout)).toBe(JSON.stringify(full.layout));
      expect(resumed.session.notePageBottomReserves).toEqual(full.session.notePageBottomReserves);
      // The search took rounds, and the resumed ones were not laid out from the first page.
      expect(full.session.stats.fullPasses).toBeGreaterThan(2);
      expect(resumed.session.stats.fullPasses).toBeLessThan(full.session.stats.fullPasses);
    });
  }
});
