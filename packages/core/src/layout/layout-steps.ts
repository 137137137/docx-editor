// Layout work that can pause between blocks, as generators.
//
// A long document's first layout took seconds in one task: a footnote-heavy document lays
// its body out tens of times while the footnote reserves settle, and nothing in that loop
// gave the page a turn. The body pass and every layout function above it are generators, and
// each `yield` is a point where a caller may stop and continue in a later task. A synchronous
// caller drains them, so outside a sliced run nothing pauses and the answer is the same: the
// generators run the same statements in the same order as the plain functions they replaced.

/** Layout steps that end with a `T`. Every `yield` is a point where the caller may pause. */
export type LayoutSteps<T> = Generator<void, T, void>;

/** Run `steps` to the end and answer the result. */
export function drainLayoutSteps<T>(steps: LayoutSteps<T>): T {
  for (;;) {
    const next = steps.next();
    if (next.done) return next.value;
  }
}

/** A synchronous computation as steps that never pause. */
export function* layoutStepsOf<T>(run: () => T): LayoutSteps<T> {
  return run();
}

/**
 * Advance `steps` until it ends or `now()` reaches `deadline`.
 *
 * Answers `{ done: true, value }` when the work ended in this call. At least one step always
 * runs, so a run with an expired deadline still makes progress.
 */
export function advanceLayoutSteps<T>(
  steps: LayoutSteps<T>,
  deadline: number,
  now: () => number
): { readonly done: true; readonly value: T } | { readonly done: false } {
  for (;;) {
    const next = steps.next();
    if (next.done) return { done: true, value: next.value };
    if (now() >= deadline) return { done: false };
  }
}
