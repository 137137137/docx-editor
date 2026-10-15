// The internal seam a large open uses to finish its first layout in slices.
//
// The surface mounts with a layout of only the first body blocks. Each step lays out more of
// the body, resuming from the previous pass, until the limit no longer cuts anything; the
// last step publishes the normal full layout. Kept off the public surface contract, like
// `surface-measurement.ts`.

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
  /** The normal full layout, resumed from the last prefix; it ends the opening. */
  finish(): void;
  /** The body block at `index` of the last pass, for sizing the next step. */
  bodyBlock(index: number): unknown;
}

/**
 * Lay a mounted surface's body out in growing prefixes (see {@link PREFIX_STEP_CHARACTERS}):
 * each extends the prefix and resumes from the previous one, and a prefix pass pauses between the
 * rows of a long table instead of laying it out in one task. Marks `container` while it runs,
 * so the loading overlay shows the document muted, and shows the loading indicator on the first
 * page.
 *
 * Answers a function that finishes a paused prefix pass at once. Any other layout of the same
 * session must call it first, because the paused pass holds that session mid-flow.
 */
export function startOpeningSlices(
  surface: object,
  container: HTMLElement,
  host: OpeningSlicesHost,
  firstGrowth: number
): () => void {
  let growth = firstGrowth;
  let pass: LayoutSteps<unknown> | null = null;
  const finishPaused = (): void => {
    const paused = pass;
    pass = null;
    if (paused) drainLayoutSteps(paused);
  };
  const step = (budgetMs: number): boolean => {
    const deadline = host.now() + budgetMs;
    while (host.limit() !== null && !host.destroyed()) {
      if (!pass) {
        const limit = host.limit()!;
        const total = host.bodyBlockCount();
        if (total !== undefined && limit >= total) {
          host.finish();
          return true;
        }
        const next =
          total === undefined
            ? limit + growth
            : nextPrefixLimit(limit, total, (index) => host.bodyBlock(index));
        host.setLimit(next);
        pass = host.prefixPass(next);
      }
      let done = false;
      do done = pass.next().done === true;
      while (!done && host.now() < deadline);
      if (!done) return false;
      pass = null;
      // Before the first pass reports the block count, steps double; see the constants above.
      growth = Math.min(MAX_PREFIX_STEP, growth * 2);
      if (host.now() >= deadline) return false;
    }
    return true;
  };
  container.removeAttribute('data-docx-opening-preview');
  if (host.limit() !== null) {
    container.setAttribute('data-docx-opening-preview', '');
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
