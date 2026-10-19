// A document opens on its first pages: the open ends (`isOpening`) once they are laid out, and
// the rest is laid out in slices while `isPaginating` holds. Every way through ends on the
// pages a document laid out in one task gets.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { BlockFragmentRecord, SemanticLayout } from '../../layout/semantic-records.ts';
import { createDocxEditor } from '../docx-editor.ts';
import { setOpenAfterPaintForTest } from '../docx-editor-open-scheduler.ts';
import { mountPaginatedSurface, type PaginatedSurfaceOptions } from '../paginated-surface.ts';
import { updateSurfaceMeasurement } from '../surface-measurement.ts';
import { setDeferredLayoutForTest } from '../surface-deferred-layout.ts';
import { continueProgressiveOpen, progressiveOpenPending } from '../surface-progressive-open.ts';
import type { OpeningSlicesOptions } from '../surface-progressive-open.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const paragraph = (text: string, extra = '') =>
  `<w:p><w:pPr>${extra}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** `count` body paragraphs with a heading every so often: a few dozen pages. */
function docx(count: number): Uint8Array {
  const parts: string[] = [];
  for (let index = 0; index < count; index += 1) {
    if (index % 30 === 0) parts.push(paragraph(`Heading ${index}`, '<w:keepNext/>'));
    parts.push(
      paragraph(
        `Body paragraph ${index}, long enough to wrap over a second line of the page at the ` +
          'default margins, so the document runs to many pages.'
      )
    );
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

/** Every fragment's identity, text and position, in page order. */
function signature(layout: SemanticLayout): string[] {
  const rows: string[] = [];
  const visit = (blocks: readonly BlockFragmentRecord[], page: number): void => {
    for (const block of blocks) {
      if (block.kind === 'table') continue;
      const text = block.lines.flatMap((line) => line.spans.map((span) => span.text)).join('');
      rows.push(`${page}|${block.paragraphId}|${text}|${block.box.y}|${block.lines.length}`);
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

const frame = () => new Promise((resolve) => requestAnimationFrame(() => resolve(undefined)));

function mount(
  bytes: Uint8Array,
  progressiveOpen?: OpeningSlicesOptions,
  container: HTMLElement = document.createElement('div')
) {
  const mounted = mountPaginatedSurface(container, bytes, {
    ...(progressiveOpen ? { progressiveOpen } : {}),
  } as PaginatedSurfaceOptions);
  if (!mounted.ok) throw new Error(mounted.reason);
  return mounted.surface;
}

/** A scroll container with stated geometry, as happy-dom reports none, around `container`. */
function scrollable(): { scroller: HTMLElement; container: HTMLElement; dispose(): void } {
  const scroller = document.createElement('div');
  scroller.className = 'docx-editor__scroll-container';
  const container = document.createElement('div');
  scroller.append(container);
  document.body.append(scroller);
  Object.defineProperty(scroller, 'clientHeight', { value: 600, configurable: true });
  Object.defineProperty(scroller, 'scrollHeight', { value: 1_000_000, configurable: true });
  let scrollTop = 0;
  Object.defineProperty(scroller, 'scrollTop', {
    get: () => scrollTop,
    set: (next: number) => {
      scrollTop = next;
    },
    configurable: true,
  });
  scroller.scrollTo = ((options: ScrollToOptions) => {
    scrollTop = options.top ?? 0;
  }) as HTMLElement['scrollTo'];
  return { scroller, container, dispose: () => scroller.remove() };
}

/** Queued discrete browser input, for as long as the returned restore has not run. */
function stubInputPending(): () => void {
  const descriptor = Object.getOwnPropertyDescriptor(window.navigator, 'scheduling');
  Object.defineProperty(window.navigator, 'scheduling', {
    configurable: true,
    value: { isInputPending: () => true },
  });
  return () => {
    if (descriptor) Object.defineProperty(window.navigator, 'scheduling', descriptor);
    else delete (window.navigator as Navigator & { scheduling?: unknown }).scheduling;
  };
}

function unslicedSignature(bytes: Uint8Array): string[] {
  const surface = mount(bytes);
  try {
    return signature(surface.layout());
  } finally {
    surface.destroy();
  }
}

// A pass defers at its first page past the view instead of after 40 ms of work, so a document
// this machine lays out in less still shows the slicing.
let restoreDeferral: () => void = () => {};
beforeAll(() => {
  restoreDeferral = setDeferredLayoutForTest(null, 0);
});
afterAll(() => restoreDeferral());

describe('a live opening', () => {
  test('mounts on its first pages and lays out the rest in many short steps', () => {
    const bytes = docx(400);
    const expected = unslicedSignature(bytes);
    const surface = mount(bytes, { live: true });
    try {
      // The mount laid out only the first pages.
      expect(progressiveOpenPending(surface)).toBe(true);
      const first = surface.layout().pages.length;
      let steps = 0;
      while (!continueProgressiveOpen(surface, 0)) steps += 1;
      expect(first).toBeLessThan(surface.layout().pages.length);
      // One step per pause point: the rest of the document was never one task.
      expect(steps).toBeGreaterThan(20);
      expect(signature(surface.layout())).toEqual(expected);
    } finally {
      surface.destroy();
    }
  }, 60_000);

  test('going straight from the first pages to the full layout ends on the same pages', () => {
    const bytes = docx(400);
    const surface = mount(bytes, { live: true, fullAfterFirst: true });
    try {
      let steps = 0;
      while (!continueProgressiveOpen(surface, 0)) steps += 1;
      expect(steps).toBeGreaterThan(20);
      expect(signature(surface.layout())).toEqual(unslicedSignature(bytes));
    } finally {
      surface.destroy();
    }
  }, 60_000);

  test('a relayout while paginating is one more pass in the slices, not a pass in its task', () => {
    const bytes = docx(400);
    const surface = mount(bytes, { live: true });
    try {
      for (let step = 0; step < 5; step += 1) continueProgressiveOpen(surface, 0);
      expect(progressiveOpenPending(surface)).toBe(true);
      const passesBefore = surface.state().perf.fullPasses;
      // What fonts landing does: a new measurement for the whole document.
      updateSurfaceMeasurement(surface, {});
      // Nothing was laid out in the call, and the opening is still under way.
      expect(surface.state().perf.fullPasses).toBe(passesBefore);
      expect(progressiveOpenPending(surface)).toBe(true);
      while (!continueProgressiveOpen(surface, 0)) {
        /* each step lays out a little more */
      }
      expect(signature(surface.layout())).toEqual(unslicedSignature(bytes));
    } finally {
      surface.destroy();
    }
  }, 60_000);

  test('a reveal of a paragraph not laid out yet lays out as far as it', () => {
    const bytes = docx(400);
    // The id the last heading gets, read from a mount laid out in one task.
    const reference = mount(bytes);
    const headings = reference
      .layout()
      .pages.flatMap((page) => page.fragments)
      .filter((fragment) => fragment.kind !== 'table');
    const last = (headings[headings.length - 1] as { paragraphId: string }).paragraphId;
    reference.destroy();
    const surface = mount(bytes, { live: true });
    try {
      const laidOut = () =>
        surface
          .layout()
          .pages.some((page) =>
            page.fragments.some(
              (fragment) => fragment.kind !== 'table' && fragment.paragraphId === last
            )
          );
      expect(laidOut()).toBe(false);
      surface.revealParagraph(last);
      expect(laidOut()).toBe(true);
      // The rest finishes as before, on the same pages.
      while (!continueProgressiveOpen(surface, 0)) {
        /* each step lays out a little more */
      }
      expect(signature(surface.layout())).toEqual(unslicedSignature(bytes));
    } finally {
      surface.destroy();
    }
  }, 60_000);

  test('a scroll to pages not laid out yet lays out as far as them', async () => {
    const bytes = docx(400);
    const reference = mount(bytes);
    const full = reference.layout();
    const expected = signature(full);
    const lastPage = full.pages[full.pages.length - 1]!;
    reference.destroy();
    const { scroller, container, dispose } = scrollable();
    const surface = mount(bytes, { live: true }, container);
    try {
      expect(progressiveOpenPending(surface)).toBe(true);
      expect(surface.layout().pages.length).toBeLessThan(full.pages.length);
      // Scroll to where the document's last page will stand (the surface is at 100%).
      scroller.scrollTop = lastPage.box.y;
      scroller.dispatchEvent(new Event('scroll'));
      // The repaint for a scroll is coalesced to a frame.
      await frame();
      await frame();
      const pages = surface.layout().pages;
      expect(pages.length).toBe(full.pages.length);
      const last = pages[pages.length - 1]!;
      expect(
        last.fragments.some((fragment) => fragment.kind !== 'table' && fragment.lines.length > 0)
      ).toBe(true);
      expect(signature(surface.layout())).toEqual(expected);
    } finally {
      surface.destroy();
      dispose();
    }
  }, 60_000);

  test('a step ends at its next pause point while input waits, and goes on without it', () => {
    const bytes = docx(400);
    const surface = mount(bytes, { live: true });
    const restore = stubInputPending();
    try {
      const before = surface.layout().pages.length;
      // A whole second of budget, given up at the first pause point.
      expect(continueProgressiveOpen(surface, 1000)).toBe(false);
      expect(progressiveOpenPending(surface)).toBe(true);
      expect(surface.layout().pages.length).toBeLessThanOrEqual(before + 1);
      restore();
      while (!continueProgressiveOpen(surface, 1000)) {
        /* each step lays out a little more */
      }
      expect(signature(surface.layout())).toEqual(unslicedSignature(bytes));
    } finally {
      restore();
      surface.destroy();
    }
  }, 60_000);
});

describe('an open behind a painted frame', () => {
  // Every document opens behind the frame here, so a document small enough to lay out
  // quickly shows the facade's part of the opening.
  let restoreEvery: () => void = () => {};
  beforeAll(() => {
    restoreEvery = setOpenAfterPaintForTest(true);
  });
  afterAll(() => restoreEvery());

  test('opens on its first pages, then fires paginated once', async () => {
    const bytes = docx(400);
    const editor = createDocxEditor({ document: bytes });
    let paginated = 0;
    editor.on('paginated', () => {
      paginated += 1;
    });
    const seen: { opening: boolean; paginating: boolean }[] = [];
    editor.on('selectionChange', (snapshot) => {
      seen.push({
        opening: snapshot.isOpening === true,
        paginating: snapshot.isPaginating === true,
      });
    });
    editor.attach(document.createElement('div'));
    // Behind a painted frame, though it is far under the large-document threshold.
    expect(editor.snapshot().isOpening).toBe(true);
    await until(() => paginated > 0);
    expect(seen.some((entry) => !entry.opening && entry.paginating)).toBe(true);
    expect(editor.snapshot().isPaginating).toBe(false);
    expect(signature(editor.surface!.layout())).toEqual(unslicedSignature(bytes));
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(paginated).toBe(1);
    editor.destroy();
  }, 60_000);

  test('moving the editor to another element while paginating still lays out every page', async () => {
    const bytes = docx(400);
    const editor = createDocxEditor({ document: bytes });
    let paginated = 0;
    editor.on('paginated', () => {
      paginated += 1;
    });
    const second = document.createElement('div');
    let moved = false;
    editor.on('selectionChange', (snapshot) => {
      if (moved || snapshot.isOpening || !snapshot.isPaginating) return;
      moved = true;
      editor.attach(second);
    });
    editor.attach(document.createElement('div'));
    await until(() => paginated > 0);
    expect(moved).toBe(true);
    expect(editor.snapshot().isPaginating).toBe(false);
    expect(signature(editor.surface!.layout())).toEqual(unslicedSignature(bytes));
    editor.destroy();
  }, 60_000);

  test('a document that mounts in one task still fires paginated', async () => {
    const restore = setOpenAfterPaintForTest(false);
    const bytes = docx(40);
    const editor = createDocxEditor({ document: bytes });
    let paginated = 0;
    editor.on('paginated', () => {
      paginated += 1;
    });
    editor.attach(document.createElement('div'));
    expect(editor.snapshot().isOpening).toBe(false);
    expect(editor.snapshot().isPaginating).toBe(false);
    expect(editor.surface).not.toBeNull();
    await until(() => paginated > 0);
    expect(paginated).toBe(1);
    editor.destroy();
    restore();
  }, 60_000);
});
