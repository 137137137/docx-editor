// A large document opens on a preview of its first pages, and lays the rest out in slices.
//
// What these tests pin down, on a body with footnotes long enough that their reserves take
// several body passes to settle: the preview pages are the pages a plain mount lays out, and
// so is every page once the slices end; the preview is editable while the rest lays out, and
// an edit starts the full layout again for the new revision; a reveal past the preview lays
// out up to its target; teardown stops the slices; and the editor reports the state through
// `snapshot().isPaginating` and the `paginated` event.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { PageRecord, SemanticLayout } from '@docx-editor.dev/core/layout';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { updateSurfaceMeasurement } from '../surface-measurement.ts';
import { createFixedMeasurer } from '../../layout/index.ts';
import { suspendReserveResumeForTest } from '../../layout/reserve-resume.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const FN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes';

const BODY_PARAGRAPHS = 240;

/** Every third paragraph cites a footnote long enough to take several lines. */
function body(): string {
  const paragraphs: string[] = [];
  for (let index = 0; index < BODY_PARAGRAPHS; index += 1) {
    const cite =
      index % 3 === 0
        ? `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${index / 3 + 1}"/></w:r>`
        : '';
    paragraphs.push(
      `<w:p><w:r><w:t xml:space="preserve">Body paragraph ${index} carries ordinary prose across a line. </w:t></w:r>${cite}</w:p>`
    );
  }
  return paragraphs.join('');
}

function footnotes(): string {
  const notes = [
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
  ];
  for (let id = 1; id <= BODY_PARAGRAPHS / 3; id += 1) {
    const text = `Note ${id} cites a source at length, with a parenthetical that runs on. `.repeat(
      1 + (id % 4)
    );
    notes.push(
      `<w:footnote w:id="${id}"><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:footnote>`
    );
  }
  return `<w:footnotes xmlns:w="${W}">${notes.join('')}</w:footnotes>`;
}

/** The footnoted document, padded past the open-yield threshold with a part layout never reads. */
function docx(padding = 0): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Default Extension="txt" ContentType="text/plain"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId5" Type="${FN}" Target="footnotes.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body()}</w:body></w:document>`
    ),
    'word/footnotes.xml': strToU8(footnotes()),
    ...(padding > 0 ? { 'word/padding.txt': strToU8(paddingText(padding)) } : {}),
  });
}

/** Ordinary text, so the package passes the reader's compression-ratio bound. */
function paddingText(length: number): string {
  const lines: string[] = [];
  for (let total = 0, line = 1; total < length; line += 1) {
    const next = `padding line ${line} with a little ordinary sentence text.\n`;
    lines.push(next);
    total += next.length;
  }
  return lines.join('');
}

const BYTES = docx();

/** What a page shows, by text: the body lines, then the note lines. */
function pageText(page: PageRecord): string {
  const lines = (fragments: PageRecord['fragments']) =>
    fragments.flatMap((fragment) =>
      fragment.kind === 'paragraph'
        ? fragment.lines.map((line) => line.spans.map((span) => span.text).join(''))
        : []
    );
  const notes = (page.footnotes?.notes ?? []).flatMap((note) => lines(note.fragments));
  return [...lines(page.fragments), '--', ...notes].join('\n');
}

const texts = (layout: SemanticLayout) => layout.pages.map(pageText);

const surfaces: PaginatedSurface[] = [];
const editors: DocxEditorInstance[] = [];
afterEach(() => {
  for (const surface of surfaces.splice(0)) surface.destroy();
  for (const editor of editors.splice(0)) editor.destroy();
});

function mount(bytes: Uint8Array, progressive: boolean): PaginatedSurface {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const result = mountPaginatedSurface(container, bytes, {
    scale: 1,
    ...({ progressiveLayout: progressive } as object),
  });
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
  surfaces.push(result.surface);
  return result.surface;
}

describe('a progressive open', () => {
  const plain = mount(BYTES, false);
  const plainTexts = texts(plain.layout());

  test('the fixture needs several body passes to settle its footnote reserves', () => {
    expect(plain.paginating()).toBe(false);
    expect(plainTexts.length).toBeGreaterThan(6);
    // Counted with every reserve round laid out from the first page: a round that resumes
    // below the first reserve it moved is not a full pass.
    const restore = suspendReserveResumeForTest();
    try {
      expect(mount(BYTES, false).state().perf.fullPasses).toBeGreaterThan(2);
    } finally {
      restore();
    }
  });

  test('previews the first pages, then publishes every page a plain mount lays out', async () => {
    const surface = mount(BYTES, true);
    expect(surface.paginating()).toBe(true);
    const preview = texts(surface.layout());
    expect(preview.length).toBeGreaterThan(1);
    expect(preview.length).toBeLessThan(plainTexts.length);
    expect(preview).toEqual(plainTexts.slice(0, preview.length));

    await surface.whenPaginated();
    expect(surface.paginating()).toBe(false);
    expect(texts(surface.layout())).toEqual(plainTexts);
  });

  test('an edit on a previewed page lays the preview out again and restarts the rest', async () => {
    const surface = mount(BYTES, true);
    const first = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: first, offset: 0 },
      head: { paragraphId: first, offset: 0 },
    });
    surface.type('Edited ');
    expect(surface.paginating()).toBe(true);
    expect(texts(surface.layout())[0]).toContain('Edited Body paragraph 0');

    await surface.whenPaginated();
    const edited = mount(new Uint8Array(await surface.save()), false);
    expect(texts(surface.layout())).toEqual(texts(edited.layout()));
  });

  test('a reveal past the preview lays out up to its target', () => {
    const surface = mount(BYTES, true);
    const ids = surface.session.paragraphIds();
    const near = ids[Math.floor(ids.length * 0.3)]!;
    surface.revealParagraph(near);
    expect(surface.layout().pages.some((page) => pageText(page).includes('paragraph 72 '))).toBe(
      true
    );

    const far = ids[ids.length - 1]!;
    surface.revealParagraph(far);
    expect(surface.paginating()).toBe(false);
    expect(texts(surface.layout())).toEqual(plainTexts);
  });

  test('a relayout with no edit during the opening is one more pass, as after a plain open', async () => {
    const after = mount(BYTES, false);
    updateSurfaceMeasurement(after, {});
    const surface = mount(BYTES, true);
    updateSurfaceMeasurement(surface, {});
    await surface.whenPaginated();
    expect(texts(surface.layout())).toEqual(texts(after.layout()));
    expect(surface.state().perf.fullPasses).toBe(after.state().perf.fullPasses);
  });

  test('finishPagination lays out the rest in the same task', () => {
    const surface = mount(BYTES, true);
    surface.finishPagination();
    expect(surface.paginating()).toBe(false);
    expect(texts(surface.layout())).toEqual(plainTexts);
  });

  test('teardown stops the slices and settles the wait', async () => {
    const surface = mount(BYTES, true);
    const waited = surface.whenPaginated();
    surface.destroy();
    await waited;
    expect(surface.paginating()).toBe(false);
  });

  test('a short body opens at once', () => {
    const container = document.createElement('div');
    const result = mountPaginatedSurface(
      container,
      zipSync({
        '[Content_Types].xml': strToU8(
          `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        ),
        '_rels/.rels': strToU8(
          `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
        ),
        'word/document.xml': strToU8(
          `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>short</w:t></w:r></w:p></w:body></w:document>`
        ),
      }),
      { scale: 1, ...({ progressiveLayout: true } as object) }
    );
    if (!result.ok) throw new Error(result.reason);
    surfaces.push(result.surface);
    expect(result.surface.paginating()).toBe(false);
  });
});

// The facade resolves a large document's fonts before its mount. The open without that wait
// laid it out on the fixed measurer and then on the shaped one, on the same session, and the
// second pass starts its footnote reserve search where the first one settled; so the surface
// runs the fixed pass first (`fixedFirst`), and the pages are that open's.
describe('an open whose fonts are resolved before its mount', () => {
  const host = createFixedMeasurer(6.6, 15);

  function mountShaped(progressive: boolean): PaginatedSurface {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const result = mountPaginatedSurface(container, BYTES, {
      scale: 1,
      measurer: host,
      producer: 'host',
      ...({ progressiveLayout: progressive, fixedFirst: true } as object),
    });
    if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
    surfaces.push(result.surface);
    return result.surface;
  }

  // The open without the wait: the fixed measurer at the mount, the host's when it arrives.
  const twoPass = mount(BYTES, false);
  updateSurfaceMeasurement(twoPass, { measurer: host, producer: 'host' });
  const twoPassTexts = texts(twoPass.layout());

  test('lays out the pages of the open that measured on the fixed measurer first', () => {
    const surface = mountShaped(false);
    expect(surface.state().perf.fullPasses).toBeGreaterThan(1);
    expect(texts(surface.layout())).toEqual(twoPassTexts);
  });

  test('so does a progressive open, its preview included', async () => {
    const surface = mountShaped(true);
    expect(surface.paginating()).toBe(true);
    const preview = texts(surface.layout());
    expect(preview.length).toBeGreaterThan(1);
    expect(preview).toEqual(twoPassTexts.slice(0, preview.length));

    await surface.whenPaginated();
    expect(texts(surface.layout())).toEqual(twoPassTexts);
  });

  test('a deferred small document runs the pass at the zoom it held before its fit', async () => {
    // The synchronous open: mounted at the held zoom, fitted, then the fonts.
    const held = mount(BYTES, false) as PaginatedSurface & { setScale(scale: number): boolean };
    held.setScale(0.8);
    updateSurfaceMeasurement(held, { measurer: host, producer: 'host' });

    const container = document.createElement('div');
    document.body.appendChild(container);
    const result = mountPaginatedSurface(container, BYTES, {
      scale: 0.8,
      measurer: host,
      producer: 'host',
      ...({ progressiveLayout: true, fixedFirst: true, unfittedScale: 1 } as object),
    });
    if (!result.ok) throw new Error(result.reason);
    surfaces.push(result.surface);
    await result.surface.whenPaginated();
    expect(texts(result.surface.layout())).toEqual(texts(held.layout()));
    // The held zoom's pass is one more than an open already at its fit runs.
    const fitted = mountShaped(false);
    expect(result.surface.state().perf.fullPasses).toBeGreaterThan(fitted.state().perf.fullPasses);
  });

  test('a progressive open on the fixed measurer whose fonts land part way runs both passes', async () => {
    const surface = mount(BYTES, true);
    expect(surface.paginating()).toBe(true);
    updateSurfaceMeasurement(surface, { measurer: host, producer: 'host' });
    await surface.whenPaginated();
    expect(texts(surface.layout())).toEqual(twoPassTexts);
  });
});

describe('the editor while a large document paginates', () => {
  test('reports isPaginating, then fires paginated once every page is laid out', async () => {
    const editor = createDocxEditor({ document: docx(1024 * 1024) });
    editors.push(editor);
    let paginated = 0;
    editor.on('paginated', () => {
      paginated += 1;
    });
    const container = document.createElement('div');
    document.body.appendChild(container);
    editor.attach(container);
    // The open yields a frame before it mounts.
    const deadline = Date.now() + 5000;
    while (editor.surface === null && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 10));
    }
    expect(editor.snapshot().isPaginating).toBe(true);
    const previewTotal = editor.snapshot().page.total;

    await editor.surface!.whenPaginated();
    await Promise.resolve();
    expect(editor.snapshot().isPaginating).toBe(false);
    expect(editor.snapshot().page.total).toBeGreaterThan(previewTotal);
    expect(paginated).toBe(1);
  });

  test('a small document is paginated when it mounts', async () => {
    const editor = createDocxEditor({ document: BYTES });
    editors.push(editor);
    let paginated = 0;
    editor.on('paginated', () => {
      paginated += 1;
    });
    editor.attach(document.createElement('div'));
    expect(editor.snapshot().isPaginating).toBe(false);
    await Promise.resolve();
    await Promise.resolve();
    expect(paginated).toBe(1);
  });
});
