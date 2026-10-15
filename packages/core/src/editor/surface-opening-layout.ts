// A large document's first layout, spread over several tasks, with its first pages first.
//
// A long document with footnotes is laid out as one body pass after another while the
// footnote reserves settle: seconds of work, and it ran as one task, so the page painted
// nothing and answered nothing until the last page was placed. Here the open publishes a
// PREVIEW first. It lays out the first body blocks on a session of its own, and keeps only
// the pages that end before the last of those blocks, because a page that the cut reaches is
// not the document's page. The full layout then runs in short slices on the surface's own
// session, from cold, as an open without slices runs it, and replaces the preview when it
// ends. So the pages are the ones an open without slices lays out; only the moment they
// appear changes. Where the open lays out on the fixed measurer before the shaped one
// (`surface-fixed-first.ts`), each layout here is both passes. Every preview starts from a
// cold session, so that its passes start where the full layout's do.
//
// While the full layout runs:
//
//   - A relayout (a commit, an undo, a font or zoom change, an image that settled) is one more
//     pass over the session, as it is after an open without slices: the second pass of a
//     footnoted document starts its reserve search where the first one settled, so how many
//     passes a document has had decides its pages, and a pass cannot simply be dropped. Each
//     layout here is therefore the open's own and then one pass per relayout so far, and the
//     preview is laid out again at once the same way. A relayout that leaves the revision as
//     it was adds its pass to the run in progress; a commit starts the run again, since a
//     pass of an older revision is never finished and never published.
//   - A view that comes near the end of the preview extends it in slices, ahead of the
//     reader.
//   - `reach` lays out at once as far as a caller needs, for a reveal or a page jump: a
//     longer preview while that stays well inside the body, otherwise the rest of the full
//     layout.

import {
  createLayoutSession,
  storyBlocks,
  type LayoutSession,
  type RevisionAuthorFilter,
  type RevisionDisplayMode,
  type SemanticLayout,
  type SemanticSelection,
} from '@docx-editor.dev/core/layout';
import type { HeadlessDocumentView } from '@docx-editor.dev/core/store';
import { advanceLayoutSteps, drainLayoutSteps, type LayoutSteps } from '../layout/layout-steps.ts';
import { replaceLayoutSession } from '../layout/layout-session.ts';
import { layoutCoversView } from './surface-pages.ts';

/** The work a slice does before it gives the page a turn, in milliseconds. */
const SLICE_MS = 16;
/**
 * Body blocks the first preview lays out. It grows by doubling until it covers the view: a
 * page of a law review can hold two blocks above its footnotes, a brief's page twenty.
 */
const FIRST_PREVIEW_BLOCKS = 12;
/**
 * The share of the body a preview may cover. A target further in is reached by finishing the
 * full layout instead: a preview that long costs about as much as the rest of the full one.
 */
const PREVIEW_SHARE = 0.5;
/** How long the first slice waits for a frame before it runs without one. */
const AFTER_PAINT_FALLBACK_MS = 200;
/** Fewer body blocks than this lay out at once: a preview would not save anything. */
export const MIN_OPENING_BLOCKS = 64;

/** What the opening needs from the surface it lays out. */
export interface OpeningLayoutHost {
  now(): number;
  /** The model revision now. */
  revision(): number;
  /** How many blocks the body has. */
  bodyBlockCount(): number;
  /** Steps of the open's layout of the first `limit` body blocks, on the preview session. */
  previewSteps(limit: number): LayoutSteps<SemanticLayout>;
  /** Steps of one more pass over the preview session, as a relayout lays it out. */
  previewAgainSteps(limit: number): LayoutSteps<SemanticLayout>;
  /** Forget the preview session's state, because a pass on it stopped part way. */
  resetPreview(): void;
  /** Steps of the open's full layout of the current revision on the surface's session, from cold. */
  fullSteps(): LayoutSteps<SemanticLayout>;
  /** Steps of one more pass over the surface's session, as a relayout lays it out. */
  fullAgainSteps(): LayoutSteps<SemanticLayout>;
  /**
   * Whether `layout` reaches far enough for what the reader sees now, and, when `selection`,
   * holds the selection too.
   */
  covers(layout: SemanticLayout, selection: boolean): boolean;
  /** Show `layout`: a preview while `complete` is false, then the full layout. */
  publish(layout: SemanticLayout, complete: boolean): void;
  /** Run `task` in a later task, after the next paint when `afterPaint`. Answers a canceller. */
  schedule(task: () => void, afterPaint: boolean): () => void;
}

/** A surface's opening layout. */
export interface OpeningLayout {
  /**
   * The first layout to paint: a preview of the first pages, with the rest scheduled. A body
   * whose preview would hold no page is laid out in full at once, and the opening is complete.
   */
  start(): SemanticLayout;
  /** Whether pages after the preview are still being laid out. */
  pending(): boolean;
  /** A layout of the current revision for a relayout during the opening: a new preview. */
  relayout(): SemanticLayout;
  /** Lay out at once until `test` holds or the full layout is published. */
  reach(test: () => boolean): void;
  /** Extend the preview in slices when `layout`, the one on screen, no longer covers the view. */
  viewMoved(layout: SemanticLayout): void;
  /** Lay out the rest at once and publish the full layout. */
  finish(): void;
  /** Resolves once the full layout is published, or the opening is cancelled. */
  whenComplete(): Promise<void>;
  /** Stop: for teardown. Nothing more is published. */
  cancel(): void;
}

/**
 * The pages of a cut layout that a full layout repeats: those before the last page with body
 * content. That page lacks what follows the cut, and the pages after it hold only notes the
 * cut collected there. Content-control boundaries keep their fragments on the kept pages.
 */
export function previewPagesOf(layout: SemanticLayout): SemanticLayout {
  let frontier = layout.pages.length - 1;
  while (frontier >= 0 && layout.pages[frontier]!.fragments.length === 0) frontier -= 1;
  const keep = Math.max(0, frontier);
  if (keep === layout.pages.length) return layout;
  const contentControls = layout.contentControls
    ?.map((control) => ({
      ...control,
      fragments: control.fragments.filter((fragment) => fragment.pageIndex < keep),
    }))
    .filter((control) => control.fragments.length > 0);
  return {
    ...layout,
    pages: layout.pages.slice(0, keep),
    ...(contentControls ? { contentControls } : {}),
  };
}

/** Create the opening for a surface. Nothing runs until `start`. */
export function createOpeningLayout(host: OpeningLayoutHost): OpeningLayout {
  let total = 0;
  let limit = FIRST_PREVIEW_BLOCKS;
  let complete = false;
  let cancelled = false;
  /** The full layout in progress, and the revision it reads. Created by the first slice. */
  let full: { readonly steps: LayoutSteps<SemanticLayout>; readonly revision: number } | null =
    null;
  /** A longer preview in progress, laid out in slices. */
  let extension: {
    readonly steps: LayoutSteps<SemanticLayout>;
    readonly limit: number;
    readonly revision: number;
  } | null = null;
  let cancelTask: (() => void) | null = null;
  const waiters: (() => void)[] = [];
  /** Relayouts asked for during the opening: each is one more pass after the open's layout. */
  let relayouts = 0;

  /** The open's layout of a preview, then a pass for each relayout, the count read live. */
  function* previewRun(at: number): LayoutSteps<SemanticLayout> {
    let laid = yield* host.previewSteps(at);
    for (let pass = 0; pass < relayouts; pass += 1) laid = yield* host.previewAgainSteps(at);
    return laid;
  }

  /** The open's full layout, then a pass for each relayout, the count read live. */
  function* fullRun(): LayoutSteps<SemanticLayout> {
    let laid = yield* host.fullSteps();
    for (let pass = 0; pass < relayouts; pass += 1) laid = yield* host.fullAgainSteps();
    return laid;
  }

  const maxPreview = (): number => Math.floor(total * PREVIEW_SHARE);

  const settle = (): void => {
    for (const resolve of waiters.splice(0)) resolve();
  };

  /** Drop a longer preview that stopped part way; its session holds a pass mid-flow. */
  const dropExtension = (): void => {
    if (!extension) return;
    extension = null;
    host.resetPreview();
  };

  /**
   * The preview at the current limit. With `grow`, grown until it covers the view and the
   * selection, or may grow no more. A relayout does not grow it: the limit already covered
   * the view, and a selection placed far away by a command is a reveal's to reach.
   */
  const preview = (grow: boolean): SemanticLayout => {
    dropExtension();
    for (;;) {
      const laid = previewPagesOf(drainLayoutSteps(previewRun(limit)));
      if (!grow || host.covers(laid, true) || limit * 2 > maxPreview()) return laid;
      limit *= 2;
    }
  };

  const arm = (afterPaint = false): void => {
    if (cancelTask || complete || cancelled) return;
    cancelTask = host.schedule(slice, afterPaint);
  };

  const publishFull = (layout: SemanticLayout): void => {
    complete = true;
    full = null;
    dropExtension();
    host.resetPreview();
    host.publish(layout, true);
    settle();
  };

  function slice(): void {
    cancelTask = null;
    if (complete || cancelled) return;
    const deadline = host.now() + SLICE_MS;
    if (extension) {
      const run = extension;
      const step = advanceLayoutSteps(run.steps, deadline, host.now);
      if (step.done) {
        extension = null;
        if (run.revision === host.revision()) {
          limit = Math.max(limit, run.limit);
          const laid = previewPagesOf(step.value);
          host.publish(laid, false);
          // Still short of the view: the next slice extends it again.
          if (!host.covers(laid, false)) startExtension(limit * 2);
        }
      }
    }
    if (!extension && host.now() < deadline) {
      full ??= { steps: fullRun(), revision: host.revision() };
      const run = full;
      const step = advanceLayoutSteps(run.steps, deadline, host.now);
      if (step.done) {
        if (run.revision === host.revision()) publishFull(step.value);
        // A revision the run did not read: lay the current one out from cold.
        else full = null;
      }
    }
    arm();
  }

  function startExtension(next: number): void {
    if (extension || complete || next > maxPreview()) return;
    extension = { steps: previewRun(next), limit: next, revision: host.revision() };
    arm();
  }

  const finish = (): void => {
    if (complete || cancelled) return;
    cancelTask?.();
    cancelTask = null;
    const run =
      full && full.revision === host.revision()
        ? full
        : { steps: fullRun(), revision: host.revision() };
    publishFull(drainLayoutSteps(run.steps));
  };

  return {
    start() {
      total = host.bodyBlockCount();
      limit = Math.min(FIRST_PREVIEW_BLOCKS, maxPreview());
      const laid = limit > 0 ? preview(true) : null;
      if (!laid || laid.pages.length === 0) {
        // Nothing of the body can show before the cut: lay it all out now, as before.
        complete = true;
        return drainLayoutSteps(fullRun());
      }
      // After the preview paints: a slice queued now could run before the first frame.
      arm(true);
      return laid;
    },
    pending: () => !complete && !cancelled,
    relayout() {
      relayouts += 1;
      // A commit: the run in progress reads a revision that no longer exists, and the next
      // slice starts the current one from cold. Otherwise the run gains the pass.
      if (full && full.revision !== host.revision()) full = null;
      const laid = preview(false);
      arm();
      return laid;
    },
    reach(test) {
      while (!complete && !cancelled && !test()) {
        if (limit * 2 > maxPreview()) {
          finish();
          return;
        }
        limit *= 2;
        host.publish(preview(false), false);
      }
    },
    viewMoved(layout) {
      if (complete || cancelled || extension || host.covers(layout, false)) return;
      startExtension(limit * 2);
    },
    finish,
    whenComplete: () =>
      complete || cancelled
        ? Promise.resolve()
        : new Promise<void>((resolve) => waiters.push(resolve)),
    cancel() {
      cancelled = true;
      cancelTask?.();
      cancelTask = null;
      full = null;
      extension = null;
      settle();
    },
  };
}

/**
 * A task scheduler for slices: a message to a channel of our own, which runs as soon as the
 * page has handled what is waiting, without the minimum delay a nested timer gets. After a
 * paint when asked, so the preview is on screen before the first slice. A hidden page has no
 * frames, so there the slice is queued at once, and a page hidden before its frame came falls
 * back to a timer rather than stopping until it is shown again.
 */
export function createSliceScheduler(view: (Window & typeof globalThis) | null) {
  const queue = (task: () => void): (() => void) => {
    let live = true;
    const run = () => {
      if (live) task();
    };
    if (typeof MessageChannel === 'function') {
      const channel = new MessageChannel();
      channel.port1.onmessage = () => {
        channel.port1.close();
        run();
      };
      channel.port2.postMessage(null);
    } else {
      setTimeout(run, 0);
    }
    return () => {
      live = false;
    };
  };
  return (task: () => void, afterPaint: boolean): (() => void) => {
    const frame = view?.requestAnimationFrame?.bind(view);
    if (!afterPaint || !frame || view?.document.visibilityState === 'hidden') return queue(task);
    let cancelQueued: (() => void) | null = null;
    // Filled in below; `fire` may only run once both are set, but never assumes it.
    const waiting: { frame?: number; timer?: ReturnType<typeof setTimeout> } = {};
    const stopWaiting = (): void => {
      if (waiting.timer !== undefined) clearTimeout(waiting.timer);
      if (waiting.frame !== undefined) view?.cancelAnimationFrame?.(waiting.frame);
    };
    const fire = (): void => {
      if (cancelQueued) return;
      stopWaiting();
      cancelQueued = queue(task);
    };
    waiting.frame = frame(fire);
    waiting.timer = setTimeout(fire, AFTER_PAINT_FALLBACK_MS);
    return () => {
      stopWaiting();
      cancelQueued?.();
      cancelQueued = () => {};
    };
  };
}

/** The surface pieces an opening lays out with. */
export interface SurfaceOpeningInput {
  readonly container: HTMLElement;
  readonly session: HeadlessDocumentView;
  readonly layoutSession: LayoutSession;
  now(): number;
  displayMode(): RevisionDisplayMode;
  authorFilter(): RevisionAuthorFilter | undefined;
  scale(): number;
  selection(): SemanticSelection;
  /**
   * Steps of a layout of the current revision: on `previewSession` when given, cut to
   * `bodyBlockLimit` blocks, otherwise the surface's own layout on its own session.
   */
  steps(previewSession?: LayoutSession, bodyBlockLimit?: number): LayoutSteps<SemanticLayout>;
  /** Steps of one more pass of the current revision, as a relayout lays it out, on the same session. */
  stepsAgain(previewSession?: LayoutSession, bodyBlockLimit?: number): LayoutSteps<SemanticLayout>;
  publish(layout: SemanticLayout): void;
}

/**
 * The opening for a surface, or null when its body is too short to gain from one. The
 * preview has a layout session of its own; `previewSession()` answers it while the opening
 * is pending, for a read that has to agree with the pages on screen.
 */
export function createSurfaceOpening(
  input: SurfaceOpeningInput
): (OpeningLayout & { previewSession(): LayoutSession | null }) | null {
  const bodyBlockCount = (): number =>
    storyBlocks(input.session.part(), input.displayMode(), input.authorFilter()).length;
  if (bodyBlockCount() < MIN_OPENING_BLOCKS) return null;
  const previewSession = createLayoutSession();
  const opening = createOpeningLayout({
    now: input.now,
    revision: () => input.session.packageRevision(),
    bodyBlockCount,
    // Every preview from a cold session, so that it starts where the full layout does.
    previewSteps: (limit) => {
      replaceLayoutSession(previewSession, createLayoutSession());
      return input.steps(previewSession, limit);
    },
    previewAgainSteps: (limit) => input.stepsAgain(previewSession, limit),
    resetPreview: () => replaceLayoutSession(previewSession, createLayoutSession()),
    fullSteps: () => {
      // A cold session: the pass an open without slices runs. A dropped pass may have left
      // this one part way.
      replaceLayoutSession(input.layoutSession, createLayoutSession());
      return input.steps();
    },
    fullAgainSteps: () => input.stepsAgain(),
    covers: (layout, selection) =>
      layoutCoversView(
        input.container,
        layout,
        input.scale(),
        selection ? input.selection() : undefined
      ),
    publish: (layout) => input.publish(layout),
    schedule: createSliceScheduler(input.container.ownerDocument.defaultView),
  });
  return { ...opening, previewSession: () => (opening.pending() ? previewSession : null) };
}
