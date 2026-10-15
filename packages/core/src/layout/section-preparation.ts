/** Reconcile section preparation while keeping dependency validation with its owner. */

import { drainLayoutSteps, type LayoutSteps } from './layout-steps.ts';

export function prepareSectionBlocks<Node extends object, Prepared>(
  bodies: readonly Node[],
  previous: { readonly bodies: readonly Node[]; readonly prepared: readonly Prepared[] } | null,
  prepare: (body: Node) => Prepared
): Prepared[] {
  return drainLayoutSteps(prepareSectionBlocksSteps(bodies, previous, prepare));
}

/**
 * {@link prepareSectionBlocks} as steps that pause after each block it prepares, so a
 * long section's first preparation does not hold the page (`layout-steps.ts`).
 */
export function* prepareSectionBlocksSteps<Node extends object, Prepared>(
  bodies: readonly Node[],
  previous: { readonly bodies: readonly Node[]; readonly prepared: readonly Prepared[] } | null,
  prepare: (body: Node) => Prepared
): LayoutSteps<Prepared[]> {
  const prepared: Prepared[] = [];
  if (!previous) {
    for (const body of bodies) {
      prepared.push(prepare(body));
      yield;
    }
    return prepared;
  }
  // Match shifted tails after insert/delete without allocating a document-sized map.
  // Equal-length edits can reuse aligned entries directly. Reorders conservatively
  // prepare moved blocks again; node IDs alone cannot prove an entry is reusable.
  let suffix = 0;
  if (bodies.length !== previous.bodies.length) {
    while (
      suffix < Math.min(bodies.length, previous.bodies.length) &&
      bodies[bodies.length - 1 - suffix] === previous.bodies[previous.bodies.length - 1 - suffix]
    )
      suffix += 1;
  }
  for (let index = 0; index < bodies.length; index += 1) {
    const body = bodies[index]!;
    if (body === previous.bodies[index]) {
      prepared.push(previous.prepared[index]!);
    } else if (index >= bodies.length - suffix) {
      prepared.push(previous.prepared[index + previous.bodies.length - bodies.length]!);
    } else {
      prepared.push(prepare(body));
      yield;
    }
  }
  return prepared;
}

/** A text edit changes a paragraph root without changing its place in drawing order. */
export function sameSectionParagraphOrder(
  before: readonly { readonly kind: string; readonly id: string }[],
  after: readonly { readonly kind: string; readonly id: string }[]
): boolean {
  return (
    before.length === after.length &&
    before.every((node, index) => {
      const next = after[index]!;
      // A table's nested story order requires the exact subtree, not merely its ID.
      return (
        node === next ||
        (node.kind === 'paragraph' && next.kind === 'paragraph' && node.id === next.id)
      );
    })
  );
}

/** Column layout and wrap probes decide whether an authored paragraph frame participates. */
export function framePolicy(columns: number, disabled?: ReadonlySet<string>): string {
  if (columns !== 1) return 'columns';
  return disabled?.size ? JSON.stringify([...disabled].sort()) : 'body';
}
