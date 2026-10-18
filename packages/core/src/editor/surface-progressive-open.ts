// The internal seam a large open uses to finish its first layout in slices.
//
// The surface mounts with a layout of only the first body blocks. Each step lays out more of
// the body, resuming from the previous pass, until the limit no longer cuts anything; then
// the normal full layout runs, and its pages after the first ones are laid out in further
// steps (`surface-deferred-layout.ts`, opening mode). The opening is pending until that full
// layout is published. Kept off the public surface contract, like `surface-measurement.ts`.
//
// WHAT A STEP RUNS NEVER DEPENDS ON THE CLOCK. The prefixes are sized by content and the full
// layout is the same pass however it is sliced, so a document paginates the same whether its
// slices are short or long, and whether a reveal ran some of them at once.

import { drainLayoutSteps, type LayoutSteps } from '../layout/layout-steps.ts';
import { showOpeningIndicator } from './surface-opening-indicator.ts';

const steps = new WeakMap<object, (budgetMs: number) => boolean>();
let stepCount = 0;

/** @internal How many opening steps have run, for tests that must see the slicing. */
export function progressiveOpenStepCount(): number {
  return stepCount;
}

/** Register `surface`'s opening step. */
export function registerProgressiveOpen(
  surface: object,
  step: (budgetMs: number) => boolean
): void {
  steps.set(surface, step);
}

/**
 * Advance `surface`'s opening layout for about `budgetMs`. True once the full layout is
 * published, and for a surface that opened without slicing.
 */
export function continueProgressiveOpen(surface: object, budgetMs: number): boolean {
  const step = steps.get(surface);
  if (!step) return true;
  stepCount += 1;
  const done = step(budgetMs);
  if (done) steps.delete(surface);
  return done;
}

/** Whether `surface` is still laying out its opening slices. */
export function progressiveOpenPending(surface: object): boolean {
  return steps.has(surface);
}

/**
 * The body text one prefix step adds, in characters, and its bounds in blocks.
 *
 * Sized by CONTENT, never by the time a pass took. Each prefix is its own footnote-reserve
 * problem and seeds the next one, so a schedule that followed the clock published a different
 * pagination on a slower or busier machine. Content is the same on every machine and tracks the
 * work closely enough: a footnoted body lays a step out in one task, so the step bounds the
 * longest pause, and every prefix also walks the blocks before it, so steps must not be tiny.
 */
const PREFIX_STEP_CHARACTERS = 150_000;
const MIN_PREFIX_STEP = 16;
const MAX_PREFIX_STEP = 4096;

/**
 * The body block limit after `limit`: at least {@link MIN_PREFIX_STEP} more blocks, then more
 * until the step holds {@link PREFIX_STEP_CHARACTERS} of text, at most {@link MAX_PREFIX_STEP}.
 * A function of the content alone.
 */
export function nextPrefixLimit(
  limit: number,
  total: number,
  blockAt: (index: number) => unknown
): number {
  let characters = 0;
  let next = limit;
  while (
    next < total &&
    next - limit < MAX_PREFIX_STEP &&
    (next - limit < MIN_PREFIX_STEP || characters < PREFIX_STEP_CHARACTERS)
  ) {
    characters += bodyBlockWeight(blockAt(next));
    next += 1;
  }
  return next;
}

const blockWeights = new WeakMap<object, number>();

/** A body block's text length, the measure a prefix step is sized by. */
export function bodyBlockWeight(block: unknown): number {
  if (typeof block !== 'object' || block === null) return 1;
  const cached = blockWeights.get(block);
  if (cached !== undefined) return cached;
  let weight = 1;
  const stack: unknown[] = [block];
  while (stack.length > 0) {
    const node = stack.pop() as { kind?: string; value?: string; children?: readonly unknown[] };
    if (node.kind === 'textValue') weight += node.value?.length ?? 0;
    else if (node.children) for (const child of node.children) stack.push(child);
  }
  blockWeights.set(block, weight);
  return weight;
}

/** What the slices need from the surface that is opening. */
export interface OpeningSlicesHost {
  now(): number;
  destroyed(): boolean;
  /** The body block limit of the opening, or null once a full layout ran. */
  limit(): number | null;
  setLimit(limit: number): void;
  bodyBlockCount(): number | undefined;
  /** A prefix layout to `limit` body blocks, as steps that may pause inside a long table. */
  prefixPass(limit: number): LayoutSteps<unknown>;
  /**
   * Show a finished prefix layout, so the pages it completed can be read and scrolled to while
   * the rest is laid out. Painting only: the next prefix resumes from the session either way.
   */
  publishPrefix?(layout: unknown): void;
  /**
   * Start the normal full layout, resumed from the last prefix; it ends the prefixes. It may
   * leave its later pages to {@link OpeningSlicesHost.advance}.
   */
  finish(): void;
  /**
   * Whether a pass left its later pages deferred: the mount's first pass, the full layout
   * `finish` starts, or a relayout during the opening.
   */
  deferred(): boolean;
  /** Advance the deferred pass for about `budgetMs`. True once it ended. */
  advance(budgetMs: number): boolean;
  /**
   * Whether a relayout arrived during the opening that nobody waited for (an image decoded,
   * fonts landed). It runs as one more pass once the prefixes are over.
   */
  relayoutPending?(): boolean;
  /** Run that relayout: one more pass over the session, its later pages deferred. */
  relayout?(): void;
  /** The body block at `index` of the last pass, for sizing the next step. */
  bodyBlock(index: number): unknown;
}

/** How one opening runs; see {@link startOpeningSlices}. */
export interface OpeningSlicesOptions {
  /**
   * The live editor's pages, shown as they are laid out. They are not muted and take input;
   * the indicator on the first page (`surface-opening-indicator.ts`) says the rest is still
   * being laid out. A preview fork leaves this off and shows its first pages muted under the
   * indicator until the rest is laid out.
   */
  readonly live?: boolean;
  /**
   * Go from the mounted prefix straight to the full layout, without the growing prefixes in
   * between: the order an open takes when its fit moves the scale right after the mount.
   */
  readonly fullAfterFirst?: boolean;
  /**
   * Lay the document out whole from the mount, in one pass from a new session whose pages
   * after the first ones follow in the slices, instead of growing prefixes. The pages are the
   * ones a document laid out in one task gets, and the ones the background one-pass check
   * (`surface-canonical-layout.ts`) would settle on. For a document the editor would
   * otherwise have mounted in one task.
   */
  readonly whole?: boolean;
}

/**
 * Lay a mounted surface's body out in growing prefixes (see {@link PREFIX_STEP_CHARACTERS}):
 * each extends the prefix and resumes from the previous one, and a prefix pass pauses between the
 * rows of a long table instead of laying it out in one task. Then the full layout, whose later
 * pages are laid out in further slices. A preview marks `container` while it runs, so the
 * loading overlay shows the document muted, and shows the loading indicator on the first page.
 *
 * Answers a function that finishes a paused prefix pass at once. Any other layout of the same
 * session must call it first, because the paused pass holds that session mid-flow.
 */
export function startOpeningSlices(
  surface: object,
  container: HTMLElement,
  host: OpeningSlicesHost,
  firstGrowth: number,
  options: OpeningSlicesOptions = {}
): () => void {
  let growth = firstGrowth;
  let fullNext = options.fullAfterFirst === true;
  let pass: LayoutSteps<unknown> | null = null;
  const finishPaused = (): void => {
    const paused = pass;
    pass = null;
    if (paused) drainLayoutSteps(paused);
  };
  const step = (budgetMs: number): boolean => {
    const deadline = host.now() + budgetMs;
    while (!host.destroyed()) {
      // A pass that deferred its later pages ends first: the next one resumes from it.
      if (host.deferred()) {
        if (!host.advance(Math.max(0, deadline - host.now()))) return false;
        if (host.now() >= deadline) return false;
        continue;
      }
      // The prefixes are over, by the full layout below or by any other full layout of the
      // session. A relayout nobody waited for is one more pass, in these slices.
      if (host.limit() === null) {
        if (!host.relayoutPending?.()) return true;
        host.relayout?.();
        if (host.now() >= deadline) return false;
        continue;
      }
      if (!pass) {
        const limit = host.limit()!;
        const total = host.bodyBlockCount();
        if (fullNext || (total !== undefined && limit >= total)) {
          fullNext = false;
          host.finish();
          continue;
        }
        const next =
          total === undefined
            ? limit + growth
            : nextPrefixLimit(limit, total, (index) => host.bodyBlock(index));
        host.setLimit(next);
        pass = host.prefixPass(next);
      }
      let result: IteratorResult<unknown, unknown>;
      do result = pass.next();
      while (!result.done && host.now() < deadline);
      if (!result.done) return false;
      pass = null;
      host.publishPrefix?.(result.value);
      // Before the first pass reports the block count, steps double; see the constants above.
      growth = Math.min(MAX_PREFIX_STEP, growth * 2);
      if (host.now() >= deadline) return false;
    }
    return true;
  };
  container.removeAttribute('data-docx-opening-preview');
  if (host.limit() !== null || host.deferred()) {
    // Live pages take input; a preview's show muted until they are all laid out.
    if (!options.live) container.setAttribute('data-docx-opening-preview', '');
    const indicator = showOpeningIndicator(surface, container);
    registerProgressiveOpen(surface, (budgetMs) => {
      const done = step(budgetMs);
      if (!done) indicator.place();
      else {
        container.removeAttribute('data-docx-opening-preview');
        indicator.remove();
      }
      return done;
    });
  }
  return finishPaused;
}
