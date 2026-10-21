// A parsed document this page issued opens as its bytes do, without reading them again, and
// any other parsed document opens from its bytes.
//
// For real files: the session opened from an issued read holds the same package as the session
// opened from the bytes — same trees, same key order, same frozen state — projects the same body
// and saves the same bytes; and an editor handed the parsed document lays out the same pages and
// saves the same file as an editor handed the bytes, with no package read of its own. Editors
// are compared one at a time, each destroyed before the next opens. The worker tests run a real
// Worker over `serveDocumentParsing`.
//
// More files can be run through the same comparison by pointing PARSED_DOCUMENT_SAMPLES at a
// directory of .docx files.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { PageRecord } from '@docx-editor.dev/core/layout';
import { openTreeSession, openTreeSessionFromPackage } from '../../binding/tree-session.ts';
import { ooxmlPackageReadCount } from '../../store/package/ooxml-package.ts';
import {
  issueParsedDocument,
  parseDocument,
  parsedDocumentRead,
  parsedDocumentTransferables,
  type ParsedDocument,
} from '../../store/package/parsed-document.ts';
import { decodeValueGraph, encodeValueGraph } from '../../store/package/parsed-document-codec.ts';
import { sameGraph } from '../../store/__tests__/same-graph.ts';
import {
  holdParsedRead,
  prepareOpenSteps,
  takePreparedOpen,
} from '../docx-editor-prepared-open.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { createDocumentParser, type DocumentParser } from '../document-parser.ts';

const FIXTURES = join(import.meta.dir, '../../../../../e2e/fixtures');
const FILES = [
  'demo.docx',
  'issue-319-sections.docx',
  'issue-68-large-comments-suggestions.docx',
  'footer-textbox-page-fields.docx',
  'float-wrap-comprehensive-test.docx',
].map((name) => join(FIXTURES, name));

const extra = process.env.PARSED_DOCUMENT_SAMPLES;
const EXTRA_FILES = extra
  ? readdirSync(extra, { recursive: true, encoding: 'utf8' })
      .filter((name) => name.endsWith('.docx'))
      .sort()
      .map((name) => join(extra, name))
  : [];

const bytesOf = (path: string) => new Uint8Array(readFileSync(path));

/** A worker's answer as it arrives, before this page's parser accepts it. */
function cloned(bytes: Uint8Array): ParsedDocument {
  const parsed = parseDocument(bytes.slice());
  return structuredClone(parsed, { transfer: parsedDocumentTransferables(parsed) });
}

/** A worker's answer as this page's parser hands it on. */
const posted = (bytes: Uint8Array): ParsedDocument => issueParsedDocument(cloned(bytes))!;

function compareSessions(path: string): void {
  const bytes = bytesOf(path);
  const inPlace = openTreeSession(bytes);
  const read = parsedDocumentRead(posted(bytes));
  expect(read).not.toBeNull();
  const fromRead = openTreeSessionFromPackage(read!);
  expect(inPlace.ok).toBe(true);
  expect(fromRead.ok).toBe(true);
  if (!inPlace.ok || !fromRead.ok) return;
  expect(sameGraph(fromRead.session.currentPackage(), inPlace.session.currentPackage())).toBeNull();
  expect(fromRead.session.projectDoc().toJSON()).toEqual(inPlace.session.projectDoc().toJSON());
  expect(fromRead.session.paragraphIds()).toEqual(inPlace.session.paragraphIds());
  expect(Buffer.from(fromRead.session.save()).equals(Buffer.from(inPlace.session.save()))).toBe(
    true
  );
}

describe('a session opened from an issued read is the session opened from its bytes', () => {
  for (const path of [...FILES, ...EXTRA_FILES]) {
    test(path.slice(path.lastIndexOf('/') + 1), () => compareSessions(path));
  }
});

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

const editors: DocxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

/** An attached editor on `document`, once its pages have settled. */
async function opened(document: Uint8Array | ParsedDocument): Promise<DocxEditorInstance> {
  const editor = createDocxEditor({ document });
  editors.push(editor);
  const container = window.document.createElement('div');
  window.document.body.appendChild(container);
  editor.attach(container);
  // Done when the open, the fonts and the passes after them have settled: nothing pending, and
  // the same pages on consecutive looks.
  const deadline = Date.now() + 90_000;
  let last = '';
  let steady = 0;
  while (steady < 3 && Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 10));
    const settling =
      editor.surface === null ||
      editor.snapshot().isOpening ||
      editor.fontMeasurement().resolving ||
      container.hasAttribute('data-docx-layout-pending');
    const now = settling ? '' : JSON.stringify(editor.surface!.layout().pages.map(pageText));
    steady = now !== '' && now === last ? steady + 1 : 0;
    last = now;
  }
  if (!editor.surface) throw new Error(`no surface: ${editor.snapshot().parseError ?? ''}`);
  return editor;
}

/**
 * What an editor on `document` shows and saves, and how many package reads its open took. The
 * editor is destroyed before this answers: two live editors in one process can lay the same
 * bytes out differently, so each comparison sees one at a time.
 */
async function shown(document: Uint8Array | ParsedDocument) {
  const before = ooxmlPackageReadCount();
  const editor = await opened(document);
  const reads = ooxmlPackageReadCount() - before;
  const pages = editor.surface!.layout().pages.map(pageText);
  const saved = Buffer.from(await editor.save());
  editors.splice(editors.indexOf(editor), 1);
  editor.destroy();
  return { pages, saved, reads };
}

describe('an editor opened on a parsed document', () => {
  for (const path of [...FILES.slice(0, 3), ...EXTRA_FILES]) {
    test(`${path.slice(path.lastIndexOf('/') + 1)}: same pages, same save, no read`, async () => {
      const bytes = bytesOf(path);
      const plain = await shown(bytes);
      const fromParsed = await shown(posted(bytes));
      expect(fromParsed.reads).toBe(0);
      expect(fromParsed.pages).toEqual(plain.pages);
      expect(fromParsed.saved.equals(plain.saved)).toBe(true);
    }, 240_000);
  }

  test('a parsed document of a refused file is refused as its bytes are', async () => {
    const garbage = new TextEncoder().encode('not a zip');
    const plain = createDocxEditor({ document: garbage });
    const fromParsed = createDocxEditor({ document: posted(garbage) });
    editors.push(plain, fromParsed);
    for (const editor of [plain, fromParsed]) {
      editor.attach(window.document.createElement('div'));
    }
    expect(fromParsed.snapshot().parseError).toBe(plain.snapshot().parseError);
    expect(plain.snapshot().parseError).not.toBeNull();
  });

  test('load() takes a parsed document too', async () => {
    const editor = await opened(bytesOf(FILES[1]!));
    const parsed = posted(bytesOf(FILES[0]!));
    const reads = ooxmlPackageReadCount();
    editor.load(parsed);
    const deadline = Date.now() + 20_000;
    while (editor.snapshot().isOpening && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 5));
    }
    expect(ooxmlPackageReadCount()).toBe(reads);
    expect(editor.snapshot().parseError).toBeNull();
  });
});

/** `path`'s parsed document with every text run's text replaced and every part doubled. */
function forgedOf(path: string): ParsedDocument {
  const parsed = parseDocument(bytesOf(path));
  const result = decodeValueGraph(parsed.read as never)!.value as {
    ok: true;
    package: { parts: Map<string, unknown> };
  };
  const parts = new Map<string, unknown>();
  for (const [name, part] of result.package.parts) {
    const text = JSON.stringify(part).replace(/"text":"[^"]+"/g, '"text":"FORGED"');
    parts.set(name, JSON.parse(text));
    parts.set(`${name}.copy`, JSON.parse(text));
  }
  const read = encodeValueGraph({ ok: true, package: { ...result.package, parts } });
  return { ...structuredClone(parsed), read };
}

describe('a parsed document this page did not issue', () => {
  test('opens from its bytes: a forged read with changed text and duplicate ids is not used', async () => {
    const plain = await shown(bytesOf(FILES[0]!));
    const forged = await shown(forgedOf(FILES[0]!));
    expect(forged.reads).toBe(1);
    expect(forged.pages).toEqual(plain.pages);
    expect(forged.pages.join('\n')).not.toContain('FORGED');
  }, 240_000);

  test('a clone of an issued document opens from its bytes', async () => {
    const clone = structuredClone(posted(bytesOf(FILES[1]!)));
    expect((await shown(clone)).reads).toBe(1);
  }, 240_000);
});

describe('a document parsed on a worker', () => {
  const WORKER = new URL('./fixtures/parse-worker.ts', import.meta.url).href;
  const parsers: { parser: DocumentParser; worker: Worker }[] = [];
  afterEach(() => {
    for (const { parser, worker } of parsers.splice(0)) {
      parser.dispose();
      worker.terminate();
    }
  });
  const started = (): DocumentParser => {
    const worker = new Worker(WORKER);
    const parser = createDocumentParser(worker);
    parsers.push({ parser, worker });
    return parser;
  };

  test('opens without a package read, and leaves the host its bytes', async () => {
    const bytes = bytesOf(FILES[1]!);
    const parsed = await started().parse(bytes);
    expect(parsed).not.toBeNull();
    expect(bytes.byteLength).toBeGreaterThan(0);
    const plain = await shown(bytes);
    const fromWorker = await shown(parsed!);
    expect(fromWorker.reads).toBe(0);
    expect(fromWorker.pages).toEqual(plain.pages);
    expect(fromWorker.saved.equals(plain.saved)).toBe(true);
  }, 240_000);

  test('an aborted request settles null, and only the newest pick is answered', async () => {
    const parser = started();
    const first = new AbortController();
    const older = parser.parse(bytesOf(FILES[0]!), { signal: first.signal });
    first.abort();
    const newer = parser.parse(bytesOf(FILES[1]!));
    expect(await older).toBeNull();
    const parsed = await newer;
    expect(Buffer.from(parsed!.bytes).equals(Buffer.from(bytesOf(FILES[1]!)))).toBe(true);
    expect(await parser.parse(bytesOf(FILES[0]!), { signal: AbortSignal.abort() })).toBeNull();
  }, 240_000);

  test('dispose settles what is pending', async () => {
    const parser = started();
    const pending = parser.parse(bytesOf(FILES[0]!));
    parser.dispose();
    expect(await pending).toBeNull();
  });

  test('takes only a Worker', () => {
    const channel = new MessageChannel();
    expect(() => createDocumentParser(channel.port1 as unknown as Worker)).toThrow(TypeError);
    channel.port1.close();
  });
});

describe("the open scheduler's prepare steps", () => {
  test('a parsed document is prepared without a package read, once', () => {
    const bytes = bytesOf(FILES[0]!);
    const parsed = posted(bytes);
    holdParsedRead(parsed, parsed.bytes);
    const reads = ooxmlPackageReadCount();
    let step: ReturnType<typeof prepareOpenSteps> | void = prepareOpenSteps(
      parsed.bytes,
      {},
      () => undefined
    );
    while (typeof step === 'function') step = step();
    expect(ooxmlPackageReadCount()).toBe(reads);
    const taken = takePreparedOpen(parsed.bytes);
    expect(taken.openedSession?.ok).toBe(true);
    expect(takePreparedOpen(parsed.bytes)).toEqual({});
  });
});
