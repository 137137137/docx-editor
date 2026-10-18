// A large open lays its body out in slices: a growing prefix per task, each pass resuming
// from the last, then the normal full layout. The result must be exactly the layout an
// unsliced mount produces, and nothing may hand out the surface before it is complete.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { BlockFragmentRecord, SemanticLayout } from '../../layout/semantic-records.ts';
import { createDocxEditor } from '../docx-editor.ts';
import { mountPaginatedSurface } from '../paginated-surface.ts';
import {
  continueProgressiveOpen,
  progressiveOpenPending,
  progressiveOpenStepCount,
} from '../surface-progressive-open.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const NUMREL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';

const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
  '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

const paragraph = (text: string, extra = '') =>
  `<w:p><w:pPr>${extra}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const item = (text: string) =>
  paragraph(text, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');

function table(rows: number, label: string): string {
  const cells: string[] = [];
  for (let row = 0; row < rows; row += 1) {
    cells.push(
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${paragraph(`${label} ${row}`)}</w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${item(`${label} cell ${row}`)}</w:tc></w:tr>`
    );
  }
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    cells.join('') +
    '</w:tbl>'
  );
}

/** Several hundred blocks: lists, tables, keep-with-next, page and section breaks. */
function body(): string {
  const parts: string[] = [];
  for (let index = 0; index < 360; index += 1) {
    if (index % 40 === 0) parts.push(paragraph(`Heading ${index}`, '<w:keepNext/>'));
    if (index % 90 === 45) parts.push(table(25, `Table ${index}`));
    if (index === 150) parts.push(paragraph('Break', '<w:pageBreakBefore/>'));
    if (index === 240) {
      parts.push(
        paragraph(
          'Section end',
          '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>'
        )
      );
    }
    parts.push(index % 3 === 0 ? item(`Item ${index}`) : paragraph(`Body paragraph ${index}`));
  }
  return parts.join('') + '<w:sectPr/>';
}

/** Past the open-yield threshold through a filler part, so the open is deferred. */
function docx(): Uint8Array {
  const lines: string[] = [];
  for (let line = 1, total = 0; total < 700 * 1024; line += 1) {
    const text = `<l>filler line ${line} carrying a little ordinary sentence text.</l>`;
    lines.push(text);
    total += text.length;
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId9" Type="${NUMREL}" Target="numbering.xml"/></Relationships>`
    ),
    'word/numbering.xml': strToU8(NUMBERING),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body()}</w:body></w:document>`
    ),
    'customXml/item1.xml': strToU8(`<filler>${lines.join('')}</filler>`),
  });
}

/** Every fragment's identity, text, marker and position, in page order. */
function signature(layout: SemanticLayout): string[] {
  const rows: string[] = [];
  const visit = (blocks: readonly BlockFragmentRecord[], page: number): void => {
    for (const block of blocks) {
      if (block.kind === 'table') {
        rows.push(`${page}|table|${block.tableId}|${block.rows.length}|${block.box.y}`);
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks, page);
        continue;
      }
      const text = block.lines.flatMap((line) => line.spans.map((span) => span.text)).join('');
      rows.push(
        `${page}|${block.paragraphId}|${text}|${block.marker?.text ?? ''}|${block.box.y}|${block.lines.length}`
      );
    }
  };
  layout.pages.forEach((page, index) => visit(page.fragments, index));
  return rows;
}

async function until(check: () => boolean, timeoutMs = 20000): Promise<void> {
  const start = Date.now();
  while (!check()) {
    if (Date.now() - start > timeoutMs) throw new Error('condition not reached in time');
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

function unslicedSignature(bytes: Uint8Array): string[] {
  const container = document.createElement('div');
  const mounted = mountPaginatedSurface(container, bytes);
  if (!mounted.ok) throw new Error(mounted.reason);
  try {
    return signature(mounted.surface.layout());
  } finally {
    mounted.surface.destroy();
  }
}

describe('a large open laid out in slices', () => {
  test('mounts on its first pages, then finishes with exactly the layout of an unsliced mount', async () => {
    const bytes = docx();
    const stepsBefore = progressiveOpenStepCount();
    const editor = createDocxEditor({ document: bytes });
    const container = document.createElement('div');
    let paginatedEvents = 0;
    editor.on('paginated', () => {
      paginatedEvents += 1;
    });
    // Recorded at the event: the slices are queued tasks that can run before a poll.
    const seen: { opening: boolean; paginating: boolean; pages: number }[] = [];
    editor.on('selectionChange', (snapshot) => {
      seen.push({
        opening: snapshot.isOpening === true,
        paginating: snapshot.isPaginating === true,
        pages: editor.surface?.layout().pages.length ?? 0,
      });
    });
    editor.attach(container);
    expect(editor.snapshot().isOpening).toBe(true);
    expect(editor.surface).toBeNull();
    await until(() => editor.surface !== null && !editor.snapshot().isOpening);
    // The open ended on the first pages: the rest were still to be laid out.
    const firstShown = seen.find((entry) => !entry.opening && entry.pages > 0);
    expect(firstShown?.paginating).toBe(true);
    // The live pages are not muted behind the preview mark.
    expect(container.hasAttribute('data-docx-opening-preview')).toBe(false);
    await until(() => editor.snapshot().isPaginating !== true && paginatedEvents > 0);
    expect(progressiveOpenStepCount() - stepsBefore).toBeGreaterThanOrEqual(1);
    const sliced = signature(editor.surface!.layout());
    expect(sliced.length).toBeGreaterThan(400);
    expect(sliced).toEqual(unslicedSignature(bytes));
    expect(firstShown!.pages).toBeLessThan(editor.surface!.layout().pages.length);
    // Once per document.
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(paginatedEvents).toBe(1);
    editor.destroy();
  }, 60_000);

  test('a save while paginating saves, and the pages still end as an unsliced mount lays them out', async () => {
    const bytes = docx();
    const editor = createDocxEditor({ document: bytes });
    // Saved from the first report of the open ending, while its later pages are pending: the
    // slices are queued tasks a polling timer can miss.
    let saving: Promise<Uint8Array> | null = null;
    let paginatingAtSave = false;
    editor.on('selectionChange', (snapshot) => {
      if (saving || snapshot.isOpening || !snapshot.isPaginating) return;
      paginatingAtSave = true;
      saving = editor.save();
    });
    editor.attach(document.createElement('div'));
    await until(() => saving !== null);
    const saved = await saving!;
    expect(paginatingAtSave).toBe(true);
    expect(saved.byteLength).toBeGreaterThan(0);
    expect(editor.snapshot().isOpening).toBe(false);
    expect(editor.surface).not.toBeNull();
    await until(() => editor.snapshot().isPaginating !== true);
    expect(signature(editor.surface!.layout())).toEqual(unslicedSignature(bytes));
    editor.destroy();
  }, 60_000);

  /** A body whose long table is one block, after enough text to fill a few pages. */
  function longTableDocx(): Uint8Array {
    const parts: string[] = [];
    for (let index = 0; index < 80; index += 1) parts.push(paragraph(`Lead paragraph ${index}`));
    parts.push(table(600, 'Long'));
    for (let index = 0; index < 40; index += 1) parts.push(paragraph(`Tail paragraph ${index}`));
    return zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId9" Type="${NUMREL}" Target="numbering.xml"/></Relationships>`
      ),
      'word/numbering.xml': strToU8(NUMBERING),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}"><w:body>${parts.join('')}<w:sectPr/></w:body></w:document>`
      ),
    });
  }

  function slicedMount(bytes: Uint8Array) {
    const container = document.createElement('div');
    const mounted = mountPaginatedSurface(container, bytes, {
      progressiveOpen: true,
    } as Parameters<typeof mountPaginatedSurface>[2]);
    if (!mounted.ok) throw new Error(mounted.reason);
    return mounted.surface;
  }

  /** Sections whose text ends in a page break before an empty section mark. */
  function breakSectionsDocx(): Uint8Array {
    const parts: string[] = [];
    for (let section = 0; section < 4; section += 1) {
      for (let index = 0; index < 150; index += 1) {
        parts.push(paragraph(`Section ${section} paragraph ${index}`));
      }
      parts.push(`<w:p><w:r><w:t>Section ${section} end</w:t><w:br w:type="page"/></w:r></w:p>`);
      if (section < 3) parts.push('<w:p><w:pPr><w:sectPr/></w:pPr></w:p>');
    }
    return zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}"><w:body>${parts.join('')}<w:sectPr/></w:body></w:document>`
      ),
    });
  }

  /** Sections whose footer shows `Page N of M` through PAGE and NUMPAGES fields. */
  function pageFooterDocx(sectionCount: number): Uint8Array {
    const FOOTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const sectPr = `<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr>`;
    const parts: string[] = [];
    for (let section = 0; section < sectionCount; section += 1) {
      for (let index = 0; index < 120; index += 1) {
        parts.push(paragraph(`Section ${section} paragraph ${index}`));
      }
      if (section < sectionCount - 1) parts.push(`<w:p><w:pPr>${sectPr}</w:pPr></w:p>`);
    }
    const field = (instr: string) =>
      `<w:fldSimple w:instr=" ${instr} "><w:r><w:t>1</w:t></w:r></w:fldSimple>`;
    return zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rIdF" Type="${FOOTER_REL}" Target="footer1.xml"/></Relationships>`
      ),
      'word/footer1.xml': strToU8(
        `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field('PAGE')}` +
          `<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field('NUMPAGES')}</w:p></w:ftr>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${parts.join('')}${sectPr}</w:body></w:document>`
      ),
    });
  }

  function footerTexts(layout: SemanticLayout): string[] {
    return layout.pages.map((page) =>
      (page.footer?.fragments ?? [])
        .flatMap((fragment) =>
          fragment.kind === 'paragraph'
            ? fragment.lines.flatMap((line) => line.spans.map((span) => span.text))
            : []
        )
        .join('')
    );
  }

  for (const sectionCount of [1, 3]) {
    test(`page fields show the final page count after the slices (${sectionCount} sections)`, () => {
      const bytes = pageFooterDocx(sectionCount);
      const surface = slicedMount(bytes);
      try {
        let slices = 0;
        while (!continueProgressiveOpen(surface, 0)) slices += 1;
        expect(slices).toBeGreaterThanOrEqual(1);
        const texts = footerTexts(surface.layout());
        const total = texts.length;
        expect(total).toBeGreaterThan(4);
        expect(texts).toEqual(texts.map((_, index) => `Page ${index + 1} of ${total}`));
        const container = document.createElement('div');
        const unsliced = mountPaginatedSurface(container, bytes);
        if (!unsliced.ok) throw new Error(unsliced.reason);
        try {
          expect(footerTexts(unsliced.surface.layout())).toEqual(texts);
        } finally {
          unsliced.surface.destroy();
        }
      } finally {
        surface.destroy();
      }
    }, 60_000);
  }

  test('a section the slices completed is not laid out again when the next one starts', () => {
    const bytes = breakSectionsDocx();
    const surface = slicedMount(bytes);
    try {
      while (!continueProgressiveOpen(surface, 0)) {
        // One prefix pass per step.
      }
      // The first slice is the open's one full pass. Once the next section exists, the first
      // section's closing break keeps its empty mark on the break's sheet; that used to lay
      // the whole section out again from its first block, a second full pass.
      expect(surface.state().perf.fullPasses).toBe(1);
      expect(signature(surface.layout())).toEqual(unslicedSignature(bytes));
    } finally {
      surface.destroy();
    }
  }, 60_000);

  test('a long table is laid out across many slices, pausing between its rows', () => {
    const bytes = longTableDocx();
    const surface = slicedMount(bytes);
    try {
      // A zero budget runs one step per slice: one pause point at a time.
      let slices = 0;
      while (!continueProgressiveOpen(surface, 0)) slices += 1;
      // 600 rows pause every few rows, so the table alone took far more slices than the
      // block prefixes around it.
      expect(slices).toBeGreaterThan(50);
      expect(progressiveOpenPending(surface)).toBe(false);
      expect(signature(surface.layout())).toEqual(unslicedSignature(bytes));
    } finally {
      surface.destroy();
    }
  }, 60_000);

  test('a layout while a slice is paused in a table still ends in the right layout', () => {
    const bytes = longTableDocx();
    const surface = slicedMount(bytes);
    try {
      // Run until a slice is paused inside the long table.
      for (let slice = 0; slice < 30; slice += 1) continueProgressiveOpen(surface, 0);
      expect(progressiveOpenPending(surface)).toBe(true);
      // A zoom lays the document out at once, on the same session.
      surface.setScale(surface.state().scale * 1.25);
      while (!continueProgressiveOpen(surface, 0)) {
        // Drain whatever the zoom left.
      }
      const reference = slicedMount(bytes);
      try {
        while (!continueProgressiveOpen(reference, Infinity)) {
          // An unpaused open at the same scale.
        }
        reference.setScale(surface.state().scale);
        expect(signature(surface.layout())).toEqual(signature(reference.layout()));
      } finally {
        reference.destroy();
      }
    } finally {
      surface.destroy();
    }
  }, 60_000);
});
