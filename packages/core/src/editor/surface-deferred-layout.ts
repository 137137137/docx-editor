// The pages after the screen, laid out after the commit that moved them.
//
// One edit near the top of a long document can move every later page: a paragraph that turns
// into a list item wraps one more line, and everything below shifts. Laying out hundreds of
// shifted pages inside the commit took most of a second. Here the commit runs the same pass
// as steps (`layout-steps.ts`) and stops once it has completed the pages the reader can see.
// It publishes those pages with the previous layout's pages after them, and finishes the same
// pass in short tasks, then publishes the real layout. Until then a page after the screen can
// show where its content stood before the edit.
//
// A newer edit drops the paused pass and restores the session to its state before that pass,
// then lays out from there; any other layout first finishes the paused pass. A pass that
// cannot be put back (several sections, footnotes) is asked to wrap up and finished instead of
// dropped: a footnote search then stops after its current round, and the next pass continues
// it. A document
// whose pass finishes within its budget, or that never reports pages, lays out as before.
//
// While a document is still being paginated after its opening (`surface-progressive-open.ts`),
// a pass runs in OPENING mode: it hands its later pages to the background as soon as the pages
// before the view are complete, however few pages the previous layout had, and the opening's
// own slices advance it (`advance`) instead of a background task of its own. So a relayout
// during the opening, an image that decoded or an edit, is one more pass over the session the
// opening laid out, and never a whole-document pass in one task.

import { drainLayoutSteps, type LayoutProgress, type LayoutSteps } from '../layout/layout-steps.ts';
import type { SemanticLayout } from '../layout/semantic-records.ts';

/** The foreground budget before a pass may hand its later pages to the background. */
let foregroundMs = 40;
/** Background slice length, in milliseconds: short enough to leave room for a frame. */
const BACKGROUND_SLICE_MS = 12;
/** Fewer pages than this left after the screen are cheaper to finish now. */
const MIN_DEFERRED_PAGES = 8;

let pageAfterViewForTest: number | null = null;

/** @internal Test hook: defer after `page`, with a foreground budget of `budgetMs`. */
export function setDeferredLayoutForTest(page: number | null, budgetMs = 40): () => void {
  const previous = [pageAfterViewForTest, foregroundMs] as const;
  pageAfterViewForTest = page;
  foregroundMs = budgetMs;
  return () => {
    [pageAfterViewForTest, foregroundMs] = previous;
  };
}

export interface DeferredLayoutHost {
  now(): number;
  /** The window the pages paint in: its frames, not the script's, decide "after paint". */
  view?(): Window;
  /** The first page index the commit may leave to the background, or undefined for none. */
  pageAfterView(): number | undefined;
  /** The layout on screen now. */
  current(): SemanticLayout | null;
  /** The model revision now; a finished pass for an older one is not published. */
  revision(): number;
  /** Publish a finished layout of the current revision. */
  publish(layout: SemanticLayout): void;
  /** Save the layout session's state before a pass; the answer puts it back. */
  snapshot(): () => void;
  /** Told when a pass starts or stops finishing in the background. */
  pendingChanged?(pending: boolean): void;
  /**
   * Whether the browser holds discrete input the page has not handled. A pass advanced by
   * the opening's slices ({@link DeferredLayout.advance}) pauses at its next step for it.
   */
  inputPending?(): boolean;
  /**
   * Told the milliseconds of main-thread work one pass took, its foreground part and every
   * background slice together, once the pass ends. A pass dropped for a newer edit is not a
   * pass of its own: its time is added to the pass that replaces it.
   */
  passTimed?(ms: number): void;
}

export interface DeferredLayout {
  /**
   * Lay out `steps` for `revision`: the finished layout, or an interim one. A pass with
   * `timed: false` reports no time through `passTimed`. With `opening`, the pass defers as
   * soon as the pages before the view are complete, and only {@link DeferredLayout.advance}
   * (or `finish`) moves it on; see the module comment.
   */
  run(
    steps: LayoutSteps<SemanticLayout>,
    revision: number,
    timed?: boolean,
    opening?: boolean
  ): SemanticLayout;
  /**
   * Advance a pass deferred in opening mode for about `budgetMs`. True once no pass is
   * deferred, which is also the answer when none was.
   */
  advance(budgetMs: number): boolean;
  /** Finish a deferred pass now. Its layout is published only if still current. */
  finish(): void;
  /**
   * Drop a deferred pass and put the session back as it was before the pass, so a newer
   * edit lays out from there instead of first paying for the rest of an outdated pass.
   */
  abandon(): void;
  /** Whether a pass is still finishing in the background. */
  pending(): boolean;
  /** Drop a deferred pass without finishing it: for teardown only. */
  cancel(): void;
}

function queueTask(run: () => void, view?: Window): void {
  if (typeof MessageChannel === 'function') {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.close();
      run();
    };
    channel.port2.postMessage(null);
    return;
  }
  (view ?? globalThis).setTimeout(run, 0);
}

/**
 * Run after the next frame paints, so the edit shows before any background work.
 *
 * The frame is the pages' window's: a frame requested from a hidden opener never comes.
 */
function afterPaint(run: () => void, view?: Window): void {
  const frames = view ?? globalThis;
  if (typeof frames.requestAnimationFrame === 'function') {
    frames.requestAnimationFrame(() => queueTask(run, view));
    return;
  }
  queueTask(run, view);
}

export function createDeferredLayout(host: DeferredLayoutHost): DeferredLayout {
  type Paused = {
    readonly steps: LayoutSteps<SemanticLayout>;
    readonly revision: number;
    readonly restore: () => void;
    /** Main-thread milliseconds the pass has taken so far. */
    spent: number;
    readonly timed: boolean;
    /** False when the pass advances state in place: it is finished, never put back. */
    readonly restorable: boolean;
    /** Asks a pass that is not restorable to end early; see {@link LayoutProgress.wrapUp}. */
    wrapUp: (() => void) | undefined;
    /** An opening pass: `advance` moves it, never a background task of its own. */
    readonly opening: boolean;
  };
  let paused: Paused | null = null;
  /** Time of dropped passes, owed to the next pass. */
  let carried = 0;
  const report = (timed: boolean, ms: number): void => {
    if (timed) host.passTimed?.(ms);
  };
  const setPaused = (next: Paused | null): void => {
    const changed = (paused === null) !== (next === null);
    paused = next;
    if (changed) host.pendingChanged?.(next !== null);
  };

  const settle = (layout: SemanticLayout, revision: number): void => {
    if (host.revision() === revision) host.publish(layout);
  };

  /**
   * Step the paused pass for `budgetMs`, or with `forInput` until input waits. True once it
   * ended (and was settled).
   */
  const stepPaused = (pass: Paused, budgetMs: number, forInput = false): boolean => {
    const began = host.now();
    const deadline = began + budgetMs;
    for (;;) {
      const next = pass.steps.next();
      if (!next.done && next.value?.wrapUp) pass.wrapUp = next.value.wrapUp;
      if (next.done) {
        setPaused(null);
        settle(next.value, pass.revision);
        report(pass.timed, pass.spent + host.now() - began);
        return true;
      }
      if (host.now() >= deadline || (forInput && host.inputPending?.() === true)) break;
    }
    pass.spent += host.now() - began;
    return false;
  };

  const drive = (): void => {
    const pass = paused;
    if (!pass || pass.opening) return;
    if (!stepPaused(pass, BACKGROUND_SLICE_MS)) queueTask(drive, host.view?.());
  };

  const finish = (): void => {
    const pass = paused;
    if (!pass) return;
    const began = host.now();
    setPaused(null);
    settle(drainLayoutSteps(pass.steps), pass.revision);
    report(pass.timed, pass.spent + host.now() - began);
  };

  return {
    run(steps, revision, timed = true, opening = false) {
      const restore = host.snapshot();
      const began = host.now();
      // Only the report takes the replaced passes' time; the budget below starts now, so
      // measuring never changes when a pass defers.
      const owed = timed ? carried : 0;
      if (timed) carried = 0;
      const deadline = began + foregroundMs;
      // An opening defers after the first page when no view says which pages show.
      const after = pageAfterViewForTest ?? host.pageAfterView() ?? (opening ? 1 : undefined);
      const previous = host.current();
      let progress: LayoutProgress | null = null;
      for (;;) {
        const next = steps.next();
        if (next.done) {
          report(timed, owed + host.now() - began);
          return next.value;
        }
        if (next.value) progress = next.value;
        if (
          after === undefined ||
          !progress?.finalize ||
          progress.pages.length <= after ||
          host.now() < deadline ||
          (!previous && !opening) ||
          (!opening &&
            previous !== null &&
            previous.pages.length - progress.pages.length < MIN_DEFERRED_PAGES)
        )
          continue;
        // The completed pages, then the previous layout's pages after them, until the pass
        // ends. Each keeps the index it is painted at. An opening's first pass has no
        // previous layout: its completed pages are all there is to show.
        const pages = previous
          ? [...progress.pages, ...previous.pages.slice(progress.pages.length)]
          : [...progress.pages];
        setPaused({
          steps,
          revision,
          restore,
          spent: owed + host.now() - began,
          timed,
          restorable: progress.restorable !== false,
          wrapUp: progress.wrapUp,
          opening,
        });
        if (!opening) afterPaint(drive, host.view?.());
        return progress.finalize(pages);
      }
    },
    finish,
    advance(budgetMs) {
      const pass = paused;
      return !pass || stepPaused(pass, budgetMs, true);
    },
    abandon() {
      const pass = paused;
      if (!pass) return;
      // Several sections or footnotes advance their session state in place, so the snapshot
      // taken before the pass cannot put it back: such a pass ends here instead, and its
      // layout reaches the screen only if no newer edit arrived meanwhile.
      if (!pass.restorable) {
        pass.wrapUp?.();
        return finish();
      }
      setPaused(null);
      pass.restore();
      if (pass.timed) carried += pass.spent;
    },
    pending: () => paused !== null,
    cancel() {
      setPaused(null);
    },
  };
}
