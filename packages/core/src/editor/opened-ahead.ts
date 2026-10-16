// A parsed document opened ahead of the editor that will show it.
//
// After the package read (`parseDocument`, on a worker), an open still has main-thread work
// before it can resolve fonts: the session over the package (paragraph identity, the stores),
// and the families the font resolver is asked for, which scan every story with the style
// cascade. On a long law review that is a quarter of a second, and it ran after the host's
// "open" — between a reader's click and the first page. None of it needs the editor, its
// container or its fonts. So a host that knows a document is about to be opened can run it
// earlier, while the reader is still choosing, and the editor that opens the same parsed
// document then starts from that session: its own open is the fonts and the first layout.
//
// The session is the one the editor would have opened. It is opened from the same read with
// the options the editor opens with when no review module is registered, and nothing touches
// it before the editor takes it, so the editor sees the revision it would have seen. An editor
// with a review module opens its own session, as before. A session is taken once: a later open
// of the same bytes (a font remount, a re-attach) opens again from the read, as it always has.

import {
  openTreeSession,
  type OpenTreeSessionOptions,
  type OpenTreeSessionResult,
} from '../binding/tree-session.ts';
import {
  isParsedDocument,
  parsedDocumentRead,
  type ParsedDocument,
} from '../store/package/parsed-document.ts';
import { resolverGlyphFontFamilies } from './resolver-glyph-font-families.ts';

/**
 * How many documents can wait opened at once. A host opens one document at a time and may
 * prepare a few candidates; an older one still waiting is dropped rather than held.
 */
const MAX_WAITING = 4;

/** Sessions opened ahead, by the bytes of the document they were opened from, oldest first. */
const waiting = new Map<Uint8Array, OpenTreeSessionResult>();

/**
 * Open a parsed document's session now, ahead of the editor that will open it, and read what
 * its first open asks of it before the fonts: the families its font resolver is asked for.
 * An editor (`createDocxEditor`, an adapter's root) handed the same parsed document then
 * opens from that session instead of opening its own; nothing else about the open changes.
 *
 * Main-thread work, as much as the editor's own open would do at that point: call it when the
 * page has a moment, once a document is certain to be opened soon. Answers whether a session
 * is waiting for the document. A document the read refused, or one already waiting, does
 * nothing.
 *
 * @public
 */
export function prepareParsedDocument(parsed: ParsedDocument): boolean {
  if (!isParsedDocument(parsed)) return false;
  const { bytes } = parsed;
  if (waiting.has(bytes)) return true;
  const read = parsedDocumentRead(parsed);
  if (!read?.ok) return false;
  const result = openTreeSession(bytes, { read });
  if (!result.ok) return false;
  // The same reads the editor's open makes before it asks for fonts. Both keep their answer
  // with the session for its revision, so the open finds them done.
  result.session.documentFonts();
  resolverGlyphFontFamilies(result.session);
  waiting.set(bytes, result);
  while (waiting.size > MAX_WAITING) waiting.delete(waiting.keys().next().value!);
  return true;
}

/**
 * The session opened ahead for `bytes`, taken: answered once, and only to an open with the
 * options it was opened with — none but the read.
 */
export function takeOpenedAhead(
  bytes: Uint8Array,
  options: OpenTreeSessionOptions
): OpenTreeSessionResult | null {
  if (options.reviewModel !== undefined) return null;
  const result = waiting.get(bytes);
  if (!result) return null;
  waiting.delete(bytes);
  return result;
}

/** Test-observable: whether a session opened ahead still waits for `bytes`. */
export function openedAheadWaiting(bytes: Uint8Array): boolean {
  return waiting.has(bytes);
}
