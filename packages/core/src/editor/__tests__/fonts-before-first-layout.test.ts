// A large document's fonts resolve before its first layout.
//
// A document past the open-yield threshold already mounts behind a painted frame. Its fonts
// used to be asked for only after that mount, so it was laid out on the fixed measurer and then
// laid out again, whole, when the shaped measurer arrived. What these tests pin down: the
// deferred mount waits for the fonts and comes up shaped, laid out once; a resolver slower than
// the wait does not hold the document back, and its answer re-measures the mounted surface in
// place; `exec` inside the wait mounts at once, as it does inside the yield window; and a small
// document keeps its synchronous mount.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { strToU8, zipSync } from 'fflate';
import { sha256FontBytes } from '../../layout/index.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { OPEN_PREPARE_WAIT_MS } from '../docx-editor-open-scheduler.ts';
import type { FontConfigurationFragment } from '../font-composition.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const fontBytes = new Uint8Array(
  readFileSync(new URL('../../layout/__tests__/fixtures/fonts/DejaVuSans.ttf', import.meta.url))
);
const dejaVu: FontConfigurationFragment = {
  sources: [
    {
      request: { family: 'DejaVu Sans', weight: 400, style: 'normal' },
      id: 'dejavu',
      bytes: fontBytes,
      hash: sha256FontBytes(fontBytes),
      faceIndex: 0,
    },
  ],
};

/** A document naming DejaVu Sans, padded past the yield threshold with a part layout never reads. */
function docx(text: string, padding = 0): Uint8Array {
  const lines: string[] = [];
  for (let length = 0, line = 1; length < padding; line += 1) {
    const next = `padding line ${line} with a little ordinary sentence text.\n`;
    lines.push(next);
    length += next.length;
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Default Extension="txt" ContentType="text/plain"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`
    ),
    ...(padding > 0 ? { 'word/padding.txt': strToU8(lines.join('')) } : {}),
  });
}

const LARGE = docx('a large document', 1024 * 1024);
const SMALL = docx('a small document');

const editors: DocxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

function open(
  bytes: Uint8Array,
  fonts: Parameters<typeof createDocxEditor>[0]['fonts']
): DocxEditorInstance {
  const editor = createDocxEditor({ document: bytes, fonts });
  editors.push(editor);
  editor.attach(document.createElement('div'));
  return editor;
}

async function until(check: () => boolean, timeoutMs = 5000): Promise<void> {
  const start = Date.now();
  while (!check()) {
    if (Date.now() - start > timeoutMs) throw new Error('condition not reached in time');
    await new Promise((resolve) => setTimeout(resolve, 2));
  }
}

const fullPasses = (editor: DocxEditorInstance) => editor.surface?.state().perf.fullPasses;

describe('a deferred open with fonts', () => {
  test('mounts on the shaped measurer and lays the document out once', async () => {
    const editor = open(LARGE, dejaVu);
    expect(editor.surface).toBeNull();
    expect(editor.snapshot().isOpening).toBe(true);

    await until(() => editor.surface !== null);
    // The first moment the surface exists, it is already shaped.
    expect(editor.fontMeasurement().measurer).toBe('shaped');
    expect(fullPasses(editor)).toBe(1);

    const mounted = editor.surface;
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(editor.surface).toBe(mounted);
    expect(fullPasses(editor)).toBe(1);
    expect(editor.fontMeasurement()).toMatchObject({ measurer: 'shaped', resolving: false });
  });

  test('a resolver slower than the wait does not hold the document back', async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => (release = resolve));
    const editor = open(LARGE, async () => {
      await gate;
      return dejaVu;
    });

    await until(() => editor.surface !== null, OPEN_PREPARE_WAIT_MS + 3000);
    expect(editor.fontMeasurement()).toMatchObject({ measurer: 'fixed', resolving: true });

    // Its answer still lands, on the surface already mounted.
    const mounted = editor.surface;
    release();
    await until(() => editor.fontMeasurement().measurer === 'shaped');
    expect(editor.surface).toBe(mounted);
    expect(editor.snapshot().isOpening).toBe(false);
  });

  test('a command inside the wait mounts at once, and the fonts follow in place', async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => (release = resolve));
    const editor = open(LARGE, async () => {
      await gate;
      return dejaVu;
    });
    await until(() => !editor.fontMeasurement().measurer || editor.fontMeasurement().resolving);

    expect(editor.exec({ type: 'insertText', text: 'x' }).ok).toBe(true);
    expect(editor.surface).not.toBeNull();
    const mounted = editor.surface;

    release();
    await until(() => editor.fontMeasurement().measurer === 'shaped');
    expect(editor.surface).toBe(mounted);
  });

  test('a small document keeps the synchronous mount', () => {
    const editor = open(SMALL, dejaVu);
    expect(editor.surface).not.toBeNull();
    expect(editor.snapshot().isOpening).toBe(false);
  });
});
