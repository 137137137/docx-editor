// The opening of a large document: a preview first, then the full layout in slices.
//
// What these tests pin down, against a host whose layouts are plain records: the first layout
// is a preview cut before the page its last block lands on; the full layout runs only in
// scheduled slices and is published once, at the end; a relayout drops the run in progress and
// starts it again from cold; `reach` grows the preview and finishes the full layout when the
// target lies too far in; a view near the end of the preview extends it; and `cancel` stops
// everything without publishing.

import { describe, expect, test } from 'bun:test';
import type { PageRecord, SemanticLayout } from '@docx-editor.dev/core/layout';
import type { LayoutSteps } from '../../layout/layout-steps.ts';
import {
  createOpeningLayout,
  createSliceScheduler,
  previewPagesOf,
  type OpeningLayoutHost,
} from '../surface-opening-layout.ts';

/** A page with body content, or a page that only carries notes. */
function page(index: number, body = true): PageRecord {
  return {
    id: `page-${index}`,
    index,
    box: { x: 0, y: index * 800, width: 600, height: 780 },
    contentBox: { x: 72, y: index * 800 + 72, width: 456, height: 636 },
    fragments: body ? [{ kind: 'paragraph', paragraphId: `p${index}` } as never] : [],
  } as PageRecord;
}

function layoutOf(pages: number, revision: number, notesPages = 0): SemanticLayout {
  const all: PageRecord[] = [];
  for (let index = 0; index < pages; index += 1) all.push(page(index));
  for (let index = 0; index < notesPages; index += 1) all.push(page(pages + index, false));
  return { revision, pages: all };
}

function* stepsOf(layout: () => SemanticLayout, pauses: number): LayoutSteps<SemanticLayout> {
  for (let step = 0; step < pauses; step += 1) yield;
  return layout();
}

/** A host whose body has `blocks` blocks, four to a page, and whose tasks run on demand. */
function harness(
  options: {
    blocks?: number;
    coverPages?: number;
    fullPauses?: number;
    /** Body pages a cut of `limit` blocks lays out, the cut's own page included. */
    previewPages?: (limit: number) => number;
  } = {}
) {
  const previewPages = options.previewPages ?? ((limit: number) => Math.ceil(limit / 4) + 1);
  const blocks = options.blocks ?? 400;
  let revision = 1;
  let clock = 0;
  const tasks: (() => void)[] = [];
  const published: { pages: number; complete: boolean }[] = [];
  const calls = { preview: [] as number[], full: 0, resetPreview: 0 };
  let coverPages = options.coverPages ?? 3;
  const host: OpeningLayoutHost = {
    // Every read of the clock advances it, so a slice ends after a bounded number of steps.
    now: () => (clock += 5),
    revision: () => revision,
    bodyBlockCount: () => blocks,
    previewSteps(limit) {
      calls.preview.push(limit);
      // The cut's last page and a notes-only page after it are not the document's.
      return stepsOf(() => layoutOf(previewPages(limit), revision, 1), 3);
    },
    resetPreview: () => {
      calls.resetPreview += 1;
    },
    fullSteps() {
      calls.full += 1;
      return stepsOf(() => layoutOf(blocks / 4, revision), options.fullPauses ?? 40);
    },
    covers: (layout) => layout.pages.length >= coverPages,
    publish: (layout, complete) => published.push({ pages: layout.pages.length, complete }),
    schedule(task) {
      tasks.push(task);
      return () => {
        const at = tasks.indexOf(task);
        if (at >= 0) tasks.splice(at, 1);
      };
    },
  };
  const runTasks = (limit = 1000): number => {
    let ran = 0;
    while (tasks.length > 0 && ran < limit) {
      tasks.shift()!();
      ran += 1;
    }
    return ran;
  };
  return {
    host,
    calls,
    published,
    tasks,
    runTasks,
    commit: () => {
      revision += 1;
    },
    setCoverPages: (pages: number) => {
      coverPages = pages;
    },
  };
}

describe('previewPagesOf', () => {
  test('keeps the pages before the last page with body content', () => {
    const cut = previewPagesOf(layoutOf(5, 1, 2));
    expect(cut.pages.map((entry) => entry.index)).toEqual([0, 1, 2, 3]);
  });

  test('keeps the content-control fragments on the kept pages only', () => {
    const layout: SemanticLayout = {
      ...layoutOf(4, 1),
      contentControls: [
        { id: 'early', fragments: [{ pageIndex: 0, box: page(0).box }] },
        { id: 'late', fragments: [{ pageIndex: 3, box: page(3).box }] },
      ] as never,
    };
    const cut = previewPagesOf(layout);
    expect(cut.contentControls?.map((control) => control.id)).toEqual(['early']);
  });
});

describe('createOpeningLayout', () => {
  test('starts on a preview and publishes the full layout once, from slices', () => {
    const { host, calls, published, runTasks } = harness();
    const opening = createOpeningLayout(host);
    const first = opening.start();
    // 12 blocks: three pages plus the cut's own page, of which the last is dropped.
    expect(first.pages).toHaveLength(3);
    expect(calls.preview).toEqual([12]);
    expect(calls.full).toBe(0);
    expect(opening.pending()).toBe(true);

    const slices = runTasks();
    expect(slices).toBeGreaterThan(1);
    expect(calls.full).toBe(1);
    expect(published).toEqual([{ pages: 100, complete: true }]);
    expect(opening.pending()).toBe(false);
  });

  test('grows the first preview until it covers the view', () => {
    const { host, calls } = harness({ coverPages: 10 });
    const opening = createOpeningLayout(host);
    expect(opening.start().pages.length).toBeGreaterThanOrEqual(10);
    expect(calls.preview).toEqual([12, 24, 48]);
  });

  test('a relayout drops the run in progress and starts the full layout again', () => {
    const { host, calls, published, runTasks, commit } = harness();
    const opening = createOpeningLayout(host);
    opening.start();
    runTasks(2);
    expect(calls.full).toBe(1);
    commit();
    const preview = opening.relayout();
    expect(preview.revision).toBe(2);
    runTasks();
    expect(calls.full).toBe(2);
    expect(published).toEqual([{ pages: 100, complete: true }]);
  });

  test('a relayout keeps the preview at the size it had', () => {
    const { host, calls, setCoverPages, commit } = harness();
    const opening = createOpeningLayout(host);
    opening.start();
    // A command placed the selection far in: the relayout does not chase it.
    setCoverPages(80);
    commit();
    expect(opening.relayout().pages).toHaveLength(3);
    expect(calls.preview).toEqual([12, 12]);
  });

  test('a run that ends on an older revision is not published', () => {
    const { host, calls, published, runTasks, commit } = harness({ fullPauses: 4 });
    const opening = createOpeningLayout(host);
    opening.start();
    runTasks(1);
    // A commit the opening was not told about: the run finishes but reads revision 1.
    commit();
    runTasks();
    expect(calls.full).toBe(2);
    expect(published.every((entry) => entry.complete)).toBe(true);
    expect(published).toHaveLength(1);
  });

  test('reach grows the preview while the target is in the first half of the body', () => {
    const { host, calls, published } = harness();
    const opening = createOpeningLayout(host);
    opening.start();
    let wanted = 20;
    opening.reach(() => (published.at(-1)?.pages ?? 0) >= wanted);
    expect(opening.pending()).toBe(true);
    expect(calls.full).toBe(0);
    expect(published.at(-1)).toEqual({ pages: 24, complete: false });

    // Past half the body, the rest of the full layout is laid out at once.
    wanted = 90;
    opening.reach(() => (published.at(-1)?.pages ?? 0) >= wanted);
    expect(opening.pending()).toBe(false);
    expect(published.at(-1)).toEqual({ pages: 100, complete: true });
  });

  test('a view near the end of the preview extends it in slices', () => {
    const { host, calls, published, runTasks, setCoverPages } = harness({ fullPauses: 400 });
    const opening = createOpeningLayout(host);
    const first = opening.start();
    runTasks(1);
    setCoverPages(6);
    opening.viewMoved(first);
    runTasks(3);
    expect(calls.preview).toContain(24);
    expect(published[0]).toEqual({ pages: 6, complete: false });
    expect(opening.pending()).toBe(true);
  });

  test('finish lays out the rest at once and resolves the waiters', async () => {
    const { host, published } = harness();
    const opening = createOpeningLayout(host);
    opening.start();
    const done = opening.whenComplete();
    opening.finish();
    await done;
    expect(published).toEqual([{ pages: 100, complete: true }]);
  });

  test('cancel stops the opening without publishing', async () => {
    const { host, published, runTasks, tasks } = harness();
    const opening = createOpeningLayout(host);
    opening.start();
    const done = opening.whenComplete();
    opening.cancel();
    await done;
    expect(tasks).toHaveLength(0);
    runTasks();
    expect(published).toEqual([]);
    expect(opening.pending()).toBe(false);
  });

  test('a body whose preview holds no page is laid out in full at once', () => {
    // Every block of the cut lands on its first page, which the preview drops.
    const { host, calls } = harness({ blocks: 80, previewPages: () => 1 });
    const opening = createOpeningLayout(host);
    expect(opening.start().pages).toHaveLength(20);
    expect(calls.full).toBe(1);
    expect(opening.pending()).toBe(false);
  });
});

describe('createSliceScheduler', () => {
  /** A visible page whose frames never come, as when it is hidden before its first frame. */
  const frameless = {
    requestAnimationFrame: () => 1,
    cancelAnimationFrame: () => undefined,
    document: { visibilityState: 'visible' },
  } as unknown as Window & typeof globalThis;
  const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

  test('runs a slice in a later task', async () => {
    let ran = 0;
    createSliceScheduler(null)(() => (ran += 1), false);
    expect(ran).toBe(0);
    await sleep(5);
    expect(ran).toBe(1);
  });

  test('runs the first slice without a frame when none comes', async () => {
    let ran = 0;
    createSliceScheduler(frameless)(() => (ran += 1), true);
    await sleep(300);
    expect(ran).toBe(1);
  });

  test('a cancelled slice never runs', async () => {
    let ran = 0;
    const schedule = createSliceScheduler(frameless);
    schedule(() => (ran += 1), true)();
    schedule(() => (ran += 1), false)();
    await sleep(300);
    expect(ran).toBe(0);
  });
});
