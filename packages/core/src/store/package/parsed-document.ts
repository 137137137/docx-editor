// Reading a package on one thread and opening it on another.
//
// `openTreeSession` reads the package and builds the session in one call on the page's thread,
// and on a long document the read is most of it: the unzip and the XML of the body, notes,
// styles and numbering. This module splits the read off. `parseDocument` runs the SAME bounded
// reader (`readOoxmlPackage`, the same zip-bomb, part, relationship and XML limits) and writes
// its result — package or rejection — in the codec's wire form; `parsedDocumentRead` rebuilds
// it on the receiving thread, and the session opens from that with nothing left to read.
//
// Neither half touches a DOM, a window or a canvas: this lane compiles without the DOM lib.
//
// OPTIONAL. Nothing in an ordinary open comes here: an editor handed bytes reads them itself,
// in its own slices, exactly as it does without this module.
//
// Provenance. A read stands in for the bounded read of its bytes only when THIS realm issued
// it: `parseDocument` ran here, or `issueParsedDocument` accepted it as the reply to a request
// this page sent to a worker it registered (`createDocumentParser`, editor entry). Nothing
// about a value proves where it came from — a forged read can carry any length, fingerprint or
// hash of the real file — so no property of the value is consulted. Issuance is recorded by
// identity in a WeakSet no other realm can reach: a clone, a value from storage, a server or
// another origin, or a hand-built object is never in it, and opens from its BYTES, which always
// travel with it, through the ordinary bounded read.

import { readOoxmlPackage, type OoxmlPackageResult } from './ooxml-package.ts';
import {
  createValueGraphDecoder,
  encodeValueGraph,
  isEncodedValueGraph,
  UnencodableValueError,
  type EncodedValueGraph,
  type ValueGraphDecoder,
} from './parsed-document-codec.ts';

/**
 * A DOCX read ahead of its open, typically on a worker. Opt-in: see {@link parseDocument}.
 *
 * Reading a package — unzipping it and turning its XML parts into trees — is most of what an
 * open costs before layout, and none of it needs the page. {@link parseDocument} does that read
 * anywhere, a worker included, and answers this: plain data that crosses `postMessage`. Handed
 * to `createDocxEditor({ document })`, `load()` or an adapter's `document` prop, an open of a
 * parsed document this page issued skips the read and goes straight to fonts and the first
 * layout, on exactly the model a read of `bytes` in place would have produced. Any other
 * parsed document opens from its bytes.
 *
 * @public
 */
export interface ParsedDocument {
  /** Tells a parsed document from the other kinds of document source. */
  readonly kind: 'docx-editor.parsed-document';
  /** The wire format's version. A document of another version is opened from its bytes. */
  readonly version: number;
  /** The file that was read: what the editor keeps, saves from, and reads embedded fonts from. */
  readonly bytes: Uint8Array;
  /** The read itself. Opaque, and used only when this page issued the document. */
  readonly read: unknown;
}

const KIND = 'docx-editor.parsed-document';
/** The wire format. Bumped whenever the codec or the read's shape changes incompatibly. */
const VERSION = 3;

/** The message a page sends its parse worker, and the worker's answer to it. */
export const PARSE_REQUEST = 'docx-editor.parse-request';
export const PARSE_REPLY = 'docx-editor.parse-reply';

/** Parsed documents this realm issued: the only ones whose read is used. */
const issued = new WeakSet<ParsedDocument>();

/**
 * Read DOCX bytes into a {@link ParsedDocument}, on any thread.
 *
 * The read is the one an open runs, under the same limits, and a file it refuses is carried
 * as that refusal: the editor reports it exactly as it reports the same file opened from
 * bytes. An answer is used as a read only in the realm that issued it — here, or a page that
 * received it from its own worker through `createDocumentParser`; anywhere else it opens from
 * its bytes.
 *
 * @public
 */
export function parseDocument(bytes: Uint8Array): ParsedDocument {
  const result = readOoxmlPackage(bytes);
  let encoded: EncodedValueGraph | null = null;
  try {
    encoded = encodeValueGraph(result);
  } catch (error) {
    // A read the codec cannot carry exactly is not offered at all: the editor then reads the
    // bytes in place, as it would have without this module.
    if (!(error instanceof UnencodableValueError)) throw error;
  }
  const parsed: ParsedDocument = { kind: KIND, version: VERSION, bytes, read: encoded };
  issued.add(parsed);
  return parsed;
}

/**
 * The buffers a {@link ParsedDocument} can transfer in `postMessage` rather than copy: its
 * bytes and the read's. Transferring them detaches them on the sending side.
 *
 * @public
 */
export function parsedDocumentTransferables(parsed: ParsedDocument): ArrayBuffer[] {
  const buffers = new Set<ArrayBuffer>();
  const add = (view: ArrayBufferView): void => {
    // A shared buffer cannot be transferred; it crosses by reference anyway.
    if (view.buffer instanceof ArrayBuffer) buffers.add(view.buffer);
  };
  add(parsed.bytes);
  if (isEncodedValueGraph(parsed.read)) {
    add(parsed.read.ops);
    for (const array of parsed.read.bytes) add(array);
  }
  return [...buffers];
}

/** The scope a parse worker answers from: `self` in a dedicated worker. @public */
export interface ParseWorkerScope {
  addEventListener(type: 'message', listener: (event: { readonly data: unknown }) => void): void;
  postMessage(message: unknown, transfer: ArrayBuffer[]): void;
}

/**
 * Answer a page's parse requests from inside a dedicated worker: call it once with `self`.
 * The page side is `createDocumentParser` from `@docx-editor.dev/core/editor`.
 *
 * @public
 */
export function serveDocumentParsing(scope: ParseWorkerScope): void {
  scope.addEventListener('message', ({ data }) => {
    const request = data as { kind?: unknown; id?: unknown; bytes?: unknown } | null;
    if (request?.kind !== PARSE_REQUEST || typeof request.id !== 'number') return;
    if (!(request.bytes instanceof Uint8Array)) return;
    const parsed = parseDocument(request.bytes);
    scope.postMessage(
      { kind: PARSE_REPLY, id: request.id, parsed },
      parsedDocumentTransferables(parsed)
    );
  });
}

/** Whether a value is shaped as a {@link ParsedDocument}, by its tag and its bytes. */
export function isParsedDocument(value: unknown): value is ParsedDocument {
  if (typeof value !== 'object' || value === null) return false;
  const candidate = value as Partial<ParsedDocument>;
  return candidate.kind === KIND && candidate.bytes instanceof Uint8Array;
}

/**
 * Mark a parsed document received as the reply to this page's own request as issued here.
 * Only `createDocumentParser` calls it, and only for a reply its own worker sent to a request
 * still pending. Answers the document, or `null` for a value that is not one.
 */
export function issueParsedDocument(value: unknown): ParsedDocument | null {
  if (!isParsedDocument(value)) return null;
  issued.add(value);
  return value;
}

/** A parsed document's rebuild: in progress, or done with its read (or with none usable). */
interface Rebuild {
  readonly decoder: ValueGraphDecoder | null;
  done: boolean;
  read: OoxmlPackageResult | null;
}

/** Each parsed document's rebuild: started once, by whoever asks first, and shared. */
const rebuilds = new WeakMap<ParsedDocument, Rebuild>();

function rebuildOf(parsed: ParsedDocument): Rebuild {
  let rebuild = rebuilds.get(parsed);
  if (!rebuild) {
    const usable =
      issued.has(parsed) && parsed.version === VERSION && isEncodedValueGraph(parsed.read);
    rebuild = usable
      ? {
          decoder: createValueGraphDecoder(parsed.read as EncodedValueGraph),
          done: false,
          read: null,
        }
      : { decoder: null, done: true, read: null };
    rebuilds.set(parsed, rebuild);
  }
  return rebuild;
}

/** Run a rebuild until it finishes or `pause` asks it to stop; answers whether it finished. */
function advance(parsed: ParsedDocument, rebuild: Rebuild, pause?: () => boolean): boolean {
  if (rebuild.done) return true;
  if (!rebuild.decoder!.run(pause)) return false;
  const decoded = rebuild.decoder!.result();
  rebuild.read = decoded && isPackageResult(decoded.value) ? decoded.value : null;
  rebuild.done = true;
  releaseWireForm(parsed.read as EncodedValueGraph);
  return true;
}

/**
 * The read a {@link ParsedDocument} carries, rebuilt on this thread — or `null` when this
 * realm did not issue it, or it has none this engine can use (another wire version, a damaged
 * stream), and the bytes must be read in place instead.
 *
 * Rebuilding is one pass of object creation, a fraction of the read it replaces, and happens
 * once per parsed document. After it the wire form is released, so a host holding the parsed
 * document (as an adapter prop does) does not also hold a second copy of the read.
 */
export function parsedDocumentRead(parsed: ParsedDocument): OoxmlPackageResult | null {
  const rebuild = rebuildOf(parsed);
  advance(parsed, rebuild);
  return rebuild.read;
}

/** Rebuild a parsed document's read in tasks of about `sliceMs`, yielding between them. */
export async function rebuildParsedDocumentInSlices(
  parsed: ParsedDocument,
  sliceMs = 8,
  yieldToHost: () => Promise<void> = platformYield
): Promise<void> {
  const rebuild = rebuildOf(parsed);
  for (;;) {
    const deadline = now() + sliceMs;
    if (advance(parsed, rebuild, () => now() >= deadline)) return;
    await yieldToHost();
  }
}

/** The thread's clock. This lane has no DOM types, so the one global it reads is named here. */
const now = (): number =>
  (globalThis as { performance?: { now(): number } }).performance?.now() ?? Date.now();

/** Yield one task to the host, by the best means the platform has. */
function platformYield(): Promise<void> {
  const host = globalThis as {
    scheduler?: { yield?: () => Promise<void> };
    MessageChannel?: new () => {
      port1: { onmessage: (() => void) | null; close(): void };
      port2: { postMessage(message: null): void; close(): void };
    };
    setTimeout?: (callback: () => void, ms: number) => unknown;
  };
  if (typeof host.scheduler?.yield === 'function') return host.scheduler.yield();
  if (host.MessageChannel) {
    const channel = new host.MessageChannel();
    return new Promise((resolve) => {
      channel.port1.onmessage = () => {
        channel.port1.close();
        channel.port2.close();
        resolve();
      };
      channel.port2.postMessage(null);
    });
  }
  return new Promise((resolve) => host.setTimeout!(resolve, 0));
}

/** The members a read's result has: a decode that lost them is not used. */
function isPackageResult(value: unknown): value is OoxmlPackageResult {
  if (typeof value !== 'object' || value === null) return false;
  const result = value as { ok?: unknown; reason?: unknown; package?: unknown };
  if (result.ok === false) return typeof result.reason === 'string';
  if (result.ok !== true || typeof result.package !== 'object' || result.package === null) {
    return false;
  }
  const pkg = result.package as Record<string, unknown>;
  const contentTypes = pkg.contentTypes as Record<string, unknown> | undefined;
  return (
    pkg.parts instanceof Map &&
    pkg.partBytes instanceof Map &&
    pkg.relationships instanceof Map &&
    Array.isArray(pkg.externalTargets) &&
    typeof pkg.mainDocumentPart === 'string' &&
    typeof contentTypes === 'object' &&
    contentTypes !== null &&
    contentTypes.defaults instanceof Map &&
    contentTypes.overrides instanceof Map
  );
}

/** Drop the stream and tables a decoded read no longer needs; the byte arrays live on in it. */
function releaseWireForm(graph: EncodedValueGraph): void {
  try {
    const writable = graph as { -readonly [K in keyof EncodedValueGraph]: EncodedValueGraph[K] };
    writable.ops = new Int32Array(0);
    writable.strings = [];
    writable.shapes = [];
    writable.numbers = [];
    writable.bytes = [];
  } catch {
    // A frozen wire form stays as it is; the WeakMap answers for it all the same.
  }
}
