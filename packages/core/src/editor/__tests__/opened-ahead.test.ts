// A parsed document opened ahead of its editor (`prepareParsedDocument`) opens as it would have.
//
// The editor handed a prepared document takes the session prepared for it, once, and lays out
// the same pages and saves the same file as an editor that opened the document itself, with no
// package read of its own. An editor with a review module or a content-lock policy opens its
// own from the read, and an editor handed plain bytes never looks for a prepared session.
//
// More files can be run through the same comparison by pointing OPENED_AHEAD_SAMPLES at a
// directory of .docx files.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { PageRecord } from '@docx-editor.dev/core/layout';
import { ooxmlPackageReadCount } from '../../store/package/ooxml-package.ts';
import {
  parseDocument,
  parsedDocumentTransferables,
  type ParsedDocument,
} from '../../store/package/parsed-document.ts';
import {
  holdParsedRead,
  prepareOpenSteps,
  takePreparedOpen,
} from '../docx-editor-prepared-open.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { openedAheadWaiting, prepareParsedDocument } from '../opened-ahead.ts';
import { stubReviewModule } from './review-test-module.ts';

const FIXTURES = join(import.meta.dir, '../../../../../e2e/fixtures');
const FILES = ['demo.docx', 'issue-319-sections.docx', 'footer-textbox-page-fields.docx'].map(
  (name) => join(FIXTURES, name)
);

const extra = process.env.OPENED_AHEAD_SAMPLES;
const EXTRA_FILES = extra
  ? readdirSync(extra, { recursive: true, encoding: 'utf8' })
      .filter((name) => name.endsWith('.docx'))
      .sort()
      .map((name) => join(extra, name))
  : [];

const bytesOf = (path: string) => new Uint8Array(readFileSync(path));

/** What a worker's answer is on arrival. */
function posted(bytes: Uint8Array): ParsedDocument {
  const parsed = parseDocument(bytes.slice());
  return structuredClone(parsed, { transfer: parsedDocumentTransferables(parsed) });
}

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

async function opened(
  document: Uint8Array | ParsedDocument,
  config: Partial<Parameters<typeof createDocxEditor>[0]> = {}
): Promise<DocxEditorInstance> {
  const editor = createDocxEditor({ ...config, document });
  editors.push(editor);
  const container = window.document.createElement('div');
  window.document.body.appendChild(container);
  editor.attach(container);
  const deadline = Date.now() + 20_000;
  while (
    (editor.surface === null || editor.snapshot().isOpening || editor.snapshot().isPaginating) &&
    Date.now() < deadline
  ) {
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  if (!editor.surface) throw new Error(`no surface: ${editor.snapshot().parseError ?? ''}`);
  return editor;
}

describe('an editor opened on a document prepared ahead', () => {
  for (const path of [...FILES, ...EXTRA_FILES]) {
    test(`${path.slice(path.lastIndexOf('/') + 1)}: takes the session, same pages, same save`, async () => {
      const plain = await opened(posted(bytesOf(path)));
      const parsed = posted(bytesOf(path));
      expect(prepareParsedDocument(parsed)).toBe(true);
      expect(openedAheadWaiting(parsed.bytes)).toBe(true);
      const reads = ooxmlPackageReadCount();
      const ahead = await opened(parsed);
      expect(ooxmlPackageReadCount()).toBe(reads);
      expect(openedAheadWaiting(parsed.bytes)).toBe(false);
      expect(ahead.surface!.layout().pages.map(pageText)).toEqual(
        plain.surface!.layout().pages.map(pageText)
      );
      const [left, right] = await Promise.all([ahead.save(), plain.save()]);
      expect(Buffer.from(left).equals(Buffer.from(right))).toBe(true);
    }, 60_000);
  }

  test('a session is taken once: a second editor on the same document opens its own', async () => {
    const parsed = posted(bytesOf(FILES[0]!));
    prepareParsedDocument(parsed);
    const first = await opened(parsed);
    const second = await opened(parsed);
    expect(second.surface!.session).not.toBe(first.surface!.session);
    expect(second.surface!.layout().pages.map(pageText)).toEqual(
      first.surface!.layout().pages.map(pageText)
    );
  }, 60_000);

  test('an editor with a review module opens its own session, and drops the waiting one', async () => {
    const parsed = posted(bytesOf(FILES[0]!));
    prepareParsedDocument(parsed);
    const reads = ooxmlPackageReadCount();
    const editor = await opened(parsed, { modules: [stubReviewModule()] });
    expect(ooxmlPackageReadCount()).toBe(reads);
    expect(editor.snapshot().parseError).toBeNull();
    expect(openedAheadWaiting(parsed.bytes)).toBe(false);
  }, 60_000);

  test('an editor with a content-lock policy opens its own session', async () => {
    const parsed = posted(bytesOf(FILES[0]!));
    prepareParsedDocument(parsed);
    const editor = await opened(parsed, {
      contentLockPolicy: { allowFormattingInLockedContent: true },
    });
    expect(editor.snapshot().parseError).toBeNull();
    expect(openedAheadWaiting(parsed.bytes)).toBe(false);
  }, 60_000);

  test('an editor handed plain bytes reads them, and leaves a prepared session alone', async () => {
    const parsed = posted(bytesOf(FILES[0]!));
    prepareParsedDocument(parsed);
    const reads = ooxmlPackageReadCount();
    await opened(parsed.bytes);
    expect(ooxmlPackageReadCount()).toBe(reads + 1);
    expect(openedAheadWaiting(parsed.bytes)).toBe(true);
  }, 60_000);

  test('a read paired with other bytes prepares nothing, and the editor reads the bytes', async () => {
    const parsed = { ...posted(bytesOf(FILES[0]!)), bytes: bytesOf(FILES[1]!) };
    expect(prepareParsedDocument(parsed)).toBe(false);
    const reads = ooxmlPackageReadCount();
    const fromParsed = await opened(parsed);
    expect(ooxmlPackageReadCount()).toBe(reads + 1);
    const plain = await opened(bytesOf(FILES[1]!));
    expect(fromParsed.surface!.layout().pages.map(pageText)).toEqual(
      plain.surface!.layout().pages.map(pageText)
    );
  }, 60_000);

  test('a refused file prepares nothing, and opens refused as before', () => {
    const parsed = posted(new TextEncoder().encode('not a zip'));
    expect(prepareParsedDocument(parsed)).toBe(false);
    expect(openedAheadWaiting(parsed.bytes)).toBe(false);
    const editor = createDocxEditor({ document: parsed });
    editors.push(editor);
    editor.attach(window.document.createElement('div'));
    expect(editor.snapshot().parseError).not.toBeNull();
  });

  test("the open scheduler's prepare steps take the waiting session, with no read", () => {
    const parsed = posted(bytesOf(FILES[0]!));
    prepareParsedDocument(parsed);
    holdParsedRead(parsed, parsed.bytes);
    const reads = ooxmlPackageReadCount();
    let step: ReturnType<typeof prepareOpenSteps> | void = prepareOpenSteps(
      parsed.bytes,
      {},
      () => undefined
    );
    while (typeof step === 'function') step = step();
    expect(ooxmlPackageReadCount()).toBe(reads);
    expect(openedAheadWaiting(parsed.bytes)).toBe(false);
    const taken = takePreparedOpen(parsed.bytes);
    expect(taken.openedSession?.ok).toBe(true);
    expect(taken.progressiveOpen).toBe(true);
  });

  test('only the most recent few wait', () => {
    const prepared = Array.from({ length: 6 }, () => posted(bytesOf(FILES[0]!)));
    for (const parsed of prepared) prepareParsedDocument(parsed);
    expect(prepared.map((parsed) => openedAheadWaiting(parsed.bytes))).toEqual([
      false,
      false,
      true,
      true,
      true,
      true,
    ]);
  });
});
