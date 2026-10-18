// Bytes opened ahead of their mount by the open scheduler's `prepare` task.
//
// Keyed by the exact bytes object, and taken once: a session is live editing state, so two
// mounts must never share one. An entry whose mount never comes dies with its bytes.
//
// The OPTIONAL parsed-document path (`@docx-editor.dev/core/parse`, `prepareParsedDocument`)
// joins here and nowhere else: a read held for bytes, or a session opened ahead from it,
// stands in for the package read. Bytes with neither — every ordinary source — take the same
// steps as before; the only addition on their path is a WeakMap miss.

import { openTreeSession, type OpenTreeSessionResult } from '@docx-editor.dev/core/binding';
import {
  openTreeSessionFromPackage,
  type OpenTreeSessionOptions,
} from '../binding/tree-session.ts';
import { readOoxmlPackageSteps, type OoxmlPackageResult } from '../store/package/ooxml-package.ts';
import { takeOpenedAhead } from './opened-ahead.ts';
import {
  isParsedDocument,
  parsedDocumentRead,
  type ParsedDocument,
} from '../store/package/parsed-document.ts';
import type { TreeDocxSessionView } from '../binding/tree-session-contract.ts';
import type { OoxmlElement, OoxmlNode } from '../store/package/ooxml-tree.ts';
import { nodeIndexSteps } from '../store/package/ooxml-edit.ts';
import { contentLockPolicyOf } from '../store/store/content-lock-policy.ts';
import {
  warmResolverGlyphFontFamilies,
  warmResolverGlyphFontTableRows,
} from './resolver-glyph-font-families.ts';
import { warmLayoutNodeDigests } from '../layout/layout-cache.ts';
import { setOpeningIndicatorHost, type OpeningIndicatorHost } from './surface-opening-indicator.ts';
import { PROGRESSIVE_OPEN_FIRST_BLOCKS } from './surface-progressive-open.ts';

/** About how long one font warm-up step runs, in milliseconds. */
const FONT_WARM_STEP_MS = 40;
/** Body blocks scanned per warm-up call, between deadline checks. */
const FONT_WARM_BATCH = 16;
/** Rows of one table scanned per warm-up call, between deadline checks. */
const FONT_WARM_ROWS = 32;

/** What font resolution reads from a session, whether prepared or mounted. */
export type FontSession = Pick<
  TreeDocxSessionView,
  'embeddedFonts' | 'currentPackage' | 'stylesRoot' | 'documentThemeFonts'
>;

const prepared = new WeakMap<Uint8Array, OpenTreeSessionResult>();
/** The package read a parsed document carries, by the bytes it opens; taken once, like a session. */
const parsedReads = new WeakMap<Uint8Array, OoxmlPackageResult>();

/**
 * Keep the read a {@link ParsedDocument} carries beside the bytes it opens, so their open skips
 * reading them. Any other source, and a parsed document whose read this engine will not use
 * (another wire version, a damaged stream, a read that does not name these bytes or exceeds
 * the reader's limits; see `parsedDocumentRead`), opens from its bytes as before — through
 * the bounded reader, so no source skips its limits.
 */
export function holdParsedRead(source: unknown, bytes: Uint8Array): void {
  if (!isParsedDocument(source)) return;
  const read = parsedDocumentRead(source as ParsedDocument);
  if (read) parsedReads.set(bytes, read);
}

/**
 * The open of `bytes` from what a parsed document left for them, claimed now: the session
 * opened ahead (`prepareParsedDocument`) when `options` are the ones it was opened with, else
 * a session opened from the held read. Both are removed at once — a session is live editing
 * state and owns its package — and `null` (no held read) means the bytes are read as usual.
 */
function takeParsedOpen(
  bytes: Uint8Array,
  options: () => PreparedOpenOptions
): (() => OpenTreeSessionResult) | null {
  const read = parsedReads.get(bytes);
  if (!read) return null;
  parsedReads.delete(bytes);
  const resolved = options();
  const ahead = takeOpenedAhead(bytes, resolved);
  return ahead ? () => ahead : () => openTreeSessionFromPackage(read, resolved);
}

/** The session options a prepared open must share with the mount it stands in for. */
export type PreparedOpenOptions = Pick<OpenTreeSessionOptions, 'reviewModel' | 'contentLockPolicy'>;

/**
 * The session options of one mount: `review` plus the host's lock policy. The prepared open of a
 * large document takes the same, or its session would ignore the policy the mount was given.
 */
export function sessionOpenOptions(
  review: Pick<OpenTreeSessionOptions, 'reviewModel'>,
  config: Pick<OpenTreeSessionOptions, 'contentLockPolicy'>
): PreparedOpenOptions {
  const contentLockPolicy = contentLockPolicyOf(config.contentLockPolicy);
  return contentLockPolicy ? { ...review, contentLockPolicy } : review;
}

/** Open `bytes` now with the session options the mount will pass, and return the result. */
export function prepareOpen(
  bytes: Uint8Array,
  options: PreparedOpenOptions
): OpenTreeSessionResult {
  const opened = openTreeSession(bytes, options);
  prepared.set(bytes, opened);
  return opened;
}

/**
 * {@link prepareOpen} in steps of about {@link FONT_WARM_STEP_MS}: the package is read in
 * short tasks (a long part reads in many), then the session opens and `then` decides the
 * steps after it. A large document parsed in one task for seconds.
 */
export function prepareOpenSteps(
  bytes: Uint8Array,
  options: PreparedOpenOptions,
  then: (opened: OpenTreeSessionResult) => ReturnType<Step>
): Step {
  // A parsed document's package is read already: the first step opens its session.
  const parsed = takeParsedOpen(bytes, () => options);
  if (parsed) {
    return () => {
      const opened = parsed();
      prepared.set(bytes, opened);
      return then(opened);
    };
  }
  const reader = readOoxmlPackageSteps(bytes);
  const step: Step = () => {
    const deadline = performance.now() + FONT_WARM_STEP_MS;
    for (;;) {
      const next = reader.next();
      if (next.done) {
        const opened = openTreeSessionFromPackage(next.value, options);
        prepared.set(bytes, opened);
        return then(opened);
      }
      if (performance.now() >= deadline) return step;
    }
  };
  return step;
}

/**
 * The mount options for a prepared open of `bytes`, or nothing. Removes the entry. A prepared
 * open is a large one, so its first layout also runs in slices after the mount. `options` is
 * asked only for a parsed document's open, so an ordinary mount computes nothing more.
 */
export function takePreparedOpen(
  bytes: Uint8Array,
  options: () => PreparedOpenOptions = () => ({})
): {
  readonly openedSession?: OpenTreeSessionResult;
  readonly progressiveOpen?: boolean;
} {
  const openedSession = prepared.get(bytes);
  if (openedSession === undefined) {
    // A parsed document opened without a prepare: its session opens from the held read.
    const parsed = takeParsedOpen(bytes, options);
    return parsed ? { openedSession: parsed() } : {};
  }
  prepared.delete(bytes);
  return { openedSession, progressiveOpen: true };
}

/**
 * The open scheduler's `prepare` hook: read `bytes` in steps, let `adopt` claim the opened
 * session for font work (null declines), then warm fonts and reads while it stays current.
 */
export function preparedOpenHook(
  sessionOptions: () => PreparedOpenOptions,
  adopt: (bytes: Uint8Array, session: TreeDocxSessionView) => (() => boolean) | null,
  startFonts: () => void | Promise<unknown>
): (bytes: Uint8Array) => Step {
  return (bytes) =>
    prepareOpenSteps(bytes, sessionOptions(), (opened) => {
      const current = opened.ok ? adopt(bytes, opened.session) : null;
      return opened.ok && current ? openSteps(opened.session, current, startFonts) : undefined;
    });
}

export { progressiveOpenPending as stillOpening } from './surface-progressive-open.ts';

/** The host's opening indicator, kept across the surfaces the facade mounts. */
export function openingIndicatorChrome(current: () => object | null) {
  let host: OpeningIndicatorHost | null = null;
  return {
    set(next: OpeningIndicatorHost | null): void {
      host = next;
      const surface = current();
      if (surface) setOpeningIndicatorHost(surface, next);
    },
    adopt: (surface: object): void => setOpeningIndicatorHost(surface, host),
  };
}

type Step = () => void | Promise<unknown> | Step;

/**
 * The prepare steps after the parse, each in a task of its own: scan `session`'s body for
 * font families about {@link FONT_WARM_STEP_MS} at a time, start the font work, digest the
 * body for layout keys the same way (a table row by row), then build the reads the mount
 * would otherwise build in its own task (the node index, the review items). The last step
 * returns the font work for the mount to wait on. A long document's scan, hashing, and reads
 * took seconds in one task. Stops when `current()` no longer holds, because a newer open
 * replaced this one.
 */
export function openSteps(
  session: TreeDocxSessionView,
  current: () => boolean,
  startFonts: () => void | Promise<unknown>
): Step {
  const body = session.part().root.children.find((child) => child.kind === 'body');
  const blocks = body?.kind === 'body' ? (body.children as readonly OoxmlElement[]) : [];
  // A table's rows come before the table, so its own digest only joins theirs. Only the blocks
  // the mount lays out: the later ones are hashed by the background passes that lay them out,
  // which pause between table rows, so the first page does not wait for the whole document.
  const digests: OoxmlNode[] = [];
  for (const block of blocks.slice(0, PROGRESSIVE_OPEN_FIRST_BLOCKS)) {
    if (block.kind === 'table') for (const row of block.children) digests.push(row);
    digests.push(block);
  }
  // Font work, in calls a deadline can fall between: plain blocks in batches, a table a few
  // rows at a time and then as a whole (its rows are cached by then).
  const fontCalls: (() => void)[] = [];
  for (let start = 0; start < blocks.length; ) {
    const block = blocks[start]!;
    if (block.kind === 'table') {
      for (let row = 0; row < block.children.length; row += FONT_WARM_ROWS) {
        const from = row;
        fontCalls.push(() => {
          warmResolverGlyphFontTableRows(session, block, from, from + FONT_WARM_ROWS);
        });
      }
      fontCalls.push(() => warmResolverGlyphFontFamilies(session, [block]));
      start += 1;
      continue;
    }
    let end = start;
    while (end < blocks.length && end - start < FONT_WARM_BATCH && blocks[end]!.kind !== 'table')
      end += 1;
    const batch = blocks.slice(start, end);
    fontCalls.push(() => warmResolverGlyphFontFamilies(session, batch));
    start = end;
  }
  let scanned = 0;
  let digested = 0;
  let fontsStarted = false;
  let fonts: void | Promise<unknown> = undefined;
  const indexStep = nodeIndexSteps(session.part().root);
  // Each read answers true once done; the node index takes several steps on a long part.
  const reads: (() => boolean)[] = [
    () => {
      const deadline = performance.now() + FONT_WARM_STEP_MS;
      return indexStep(() => performance.now() >= deadline);
    },
    () => {
      session.reviewItems();
      return true;
    },
  ];
  const step: Step = () => {
    if (!current()) return;
    const deadline = performance.now() + FONT_WARM_STEP_MS;
    if (!fontsStarted) {
      while (scanned < fontCalls.length && performance.now() < deadline) fontCalls[scanned++]!();
      if (scanned < fontCalls.length) return step;
      fontsStarted = true;
      fonts = startFonts();
      return step;
    }
    if (digested < digests.length) {
      while (digested < digests.length && performance.now() < deadline) {
        warmLayoutNodeDigests(digests.slice(digested, digested + FONT_WARM_BATCH));
        digested += FONT_WARM_BATCH;
      }
      return step;
    }
    const read = reads[0];
    if (!read) return fonts;
    if (read()) reads.shift();
    return step;
  };
  return step;
}
