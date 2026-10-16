// Resuming a body pass whose footnote reserves moved.
//
// The footnote reserve search lays the body out again each time it adopts a new reserve map,
// and a cold search adopts tens of them. Its settled prefix grows from the front: a round's
// map usually differs from the last only from some page on. The reserves are part of a pass's
// context key, so every round used to start from the first page. It does not have to.
//
// A pass reads the reserve of a page only while that page is the one it is filling
// (`contentHeight`), and pages are only ever added. So every paragraph placed before the pass
// opened a given page read reserves of earlier pages alone. When two passes share everything
// else (the context without its reserves, the producer, the parity rule, the blocks up to
// the resume point), the checkpoint taken before a paragraph on a page below the first
// reserve that moved describes the new pass exactly, and the pages completed by then are the
// new pass's own.
//
// Wrap exclusion zones are read per page the same way but are collected from the previous
// pass's pages, so a pass with any is never resumed here.
//
// The record is kept per checkpoint ARRAY: a pass stores a new array with every layout it
// publishes, and a session replaced wholesale (`replaceLayoutSession`) carries its array, so
// the record cannot outlive the checkpoints it describes or describe another pass's.

import type { ExclusionZone } from './drawing-exclusion.ts';
import type { FlowCheckpoint } from './layout-session.ts';
import { firstDifferingReserveSlot } from './note-reserves.ts';

/** What one pass read its reserves under. */
export interface PassReserveReads {
  /** The pass's context key with its reserve slice left out. */
  readonly context: string;
  readonly reserves: ReadonlyMap<number, number> | undefined;
  readonly pageIndexStart: number;
  readonly hasExclusionZones: boolean;
}

export function passReserveReads(
  context: string,
  reserves: ReadonlyMap<number, number> | undefined,
  pageIndexStart: number,
  options: {
    readonly drawingExclusionZonesByPage?: ReadonlyMap<number, readonly ExclusionZone[]>;
  }
): PassReserveReads {
  return {
    context,
    reserves,
    pageIndexStart,
    hasExclusionZones: (options.drawingExclusionZonesByPage?.size ?? 0) > 0,
  };
}

const records = new WeakMap<readonly FlowCheckpoint[], PassReserveReads>();
let suspended = false;

/** Lay every reserve round out from its first page, as before. @internal For tests. */
export function suspendReserveResumeForTest(): () => void {
  suspended = true;
  return () => {
    suspended = false;
  };
}

/** Record the reserves the pass that produced `checkpoints` laid its pages out under. */
export function recordPassReserves(
  checkpoints: readonly FlowCheckpoint[],
  reads: PassReserveReads
): void {
  if (reads.hasExclusionZones) {
    records.delete(checkpoints);
    return;
  }
  // A copy: the record must keep describing the pages even if a caller reuses its map.
  records.set(checkpoints, {
    ...reads,
    reserves: reads.reserves && reads.reserves.size > 0 ? new Map(reads.reserves) : undefined,
  });
}

/**
 * The block a pass may resume at when only its reserves differ from the pass that produced
 * `checkpoints`: the last checkpoint at or before `firstChanged` taken on a page below the
 * first reserve that moved. Zero when there is none, or the previous pass is not comparable.
 */
export function reserveResumeIndex(
  checkpoints: readonly FlowCheckpoint[],
  reads: PassReserveReads,
  firstChanged: number
): number {
  const record = records.get(checkpoints);
  if (suspended || reads.hasExclusionZones || !record || record.context !== reads.context) {
    return 0;
  }
  const slot = firstDifferingReserveSlot(
    record.reserves,
    record.pageIndexStart,
    reads.reserves,
    reads.pageIndexStart
  );
  let index = Math.min(firstChanged, checkpoints.length - 1);
  while (index > 0 && checkpoints[index]!.pageCount >= slot) index -= 1;
  return index;
}
