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
// in its own slices, exactly as it does without this module. Only a host that calls
// `parseDocument` (typically in a worker it started, the moment a file is chosen) and hands
// the answer to the editor takes this path.
//
// Trust model. The read is produced by `parseDocument` and nothing else: the reader's own
// bounded read, in a same-origin worker the host started itself. A parsed document is not a
// format to accept from a server, another origin or storage. Even so the editor does not take
// a read on faith: it opens one only when the wire version is this engine's, the stream
// decodes completely, the read names the file it travels with (its length and a fingerprint
// of the zip's central directory, which records every entry's CRC-32 and sizes), and the
// rebuilt package is within the reader's own limits (entries, total and ratio of expanded
// bytes, XML parts, part size, relationships, and every tree's depth and node count). Any
// value that fails one of those opens from its BYTES, which always travel with it, through
// the ordinary bounded read — so no value can carry the editor past the reader's limits.

import {
  DEFAULT_OOXML_PACKAGE_LIMITS,
  readOoxmlPackage,
  type OoxmlPackageResult,
} from './ooxml-package.ts';
import { DEFAULT_ZIP_LIMITS, ZIP_RATIO_EXEMPT_BYTES } from './zip.ts';
import { MAX_XML_DEPTH, XML_DEFAULT_MAX_ELEMENTS, XML_HARD_MAX_BYTES } from './xml-reader.ts';
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
 * to `createDocxEditor({ document })`, `load()` or an adapter's `document` prop, the open skips
 * the read and goes straight to fonts and the first layout, on exactly the model a read of
 * `bytes` in place would have produced.
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
  /**
   * The read itself. Opaque: only the engine's own {@link parseDocument} writes it, and a read
   * that fails the editor's checks (version, integrity, the reader's limits) is ignored and the
   * bytes are read in place instead.
   */
  readonly read: unknown;
}

const KIND = 'docx-editor.parsed-document';
/** The wire format. Bumped whenever the codec or the read's shape changes incompatibly. */
const VERSION = 2;

/** What the wire carries: the read, and the file it was read from. */
interface WireRead {
  readonly byteLength: number;
  readonly fingerprint: number;
  readonly result: OoxmlPackageResult;
}

/** Bytes at the end of a file the fingerprint covers: the zip's central directory, usually whole. */
const FINGERPRINT_TAIL = 128 * 1024;

/**
 * A fingerprint of a file, cheap at any size: FNV-1a over its length and its last
 * {@link FINGERPRINT_TAIL} bytes. A zip ends with its central directory, which records every
 * entry's name, CRC-32 and sizes, so a read paired with any other file — a different document,
 * an edited copy, a truncated one — is told apart without hashing the whole file.
 */
function fingerprintOf(bytes: Uint8Array): number {
  let hash = 0x811c9dc5;
  const mix = (byte: number): void => {
    hash = Math.imul(hash ^ byte, 0x01000193);
  };
  for (let length = bytes.byteLength, i = 0; i < 4; i += 1, length = Math.floor(length / 256)) {
    mix(length & 0xff);
  }
  for (let index = Math.max(0, bytes.byteLength - FINGERPRINT_TAIL); index < bytes.byteLength; ) {
    mix(bytes[index++]!);
  }
  return hash >>> 0;
}

/**
 * Read DOCX bytes into a {@link ParsedDocument}, on any thread.
 *
 * An OPTIONAL path for reading a file before any editor exists — in a worker the host starts
 * itself, the moment the file is chosen — so the editor that later opens it skips the read.
 * An editor handed plain bytes never comes here and opens exactly as it otherwise would.
 *
 * The read is the one an open runs, under the same limits, and a file it refuses is carried
 * as that refusal: the editor reports it exactly as it reports the same file opened from
 * bytes. The editor checks a read before it uses one (wire version, integrity against the
 * bytes it travels with, the reader's limits) and opens the bytes in place when it fails, so
 * only answers of this function are ever opened as reads. Post the answer with
 * {@link parsedDocumentTransferables} to move its buffers rather than copy them.
 *
 * @public
 */
export function parseDocument(bytes: Uint8Array): ParsedDocument {
  const result = readOoxmlPackage(bytes);
  const wire: WireRead = {
    byteLength: bytes.byteLength,
    fingerprint: fingerprintOf(bytes),
    result,
  };
  let encoded: EncodedValueGraph | null = null;
  try {
    encoded = encodeValueGraph(wire);
  } catch (error) {
    // A read the codec cannot carry exactly is not offered at all: the editor then reads the
    // bytes in place, as it would have without this module.
    if (!(error instanceof UnencodableValueError)) throw error;
  }
  return { kind: KIND, version: VERSION, bytes, read: encoded };
}

/**
 * The buffers a {@link ParsedDocument} can transfer in `postMessage` rather than copy: its
 * bytes and the read's. Transferring them detaches them on the sending side, so pass this
 * only where the sender is done with them — a worker that just parsed.
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

/** Whether a value is a {@link ParsedDocument}, by its tag and its bytes. */
export function isParsedDocument(value: unknown): value is ParsedDocument {
  if (typeof value !== 'object' || value === null) return false;
  const candidate = value as Partial<ParsedDocument>;
  return candidate.kind === KIND && candidate.bytes instanceof Uint8Array;
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
    const usable = parsed.version === VERSION && isEncodedValueGraph(parsed.read);
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
  rebuild.read = decoded ? usableRead(decoded.value, parsed.bytes) : null;
  rebuild.done = true;
  releaseWireForm(parsed.read as EncodedValueGraph);
  return true;
}

/**
 * The read a {@link ParsedDocument} carries, rebuilt on this thread — or `null` when it has
 * none this engine can use (another wire version, a damaged stream), and the bytes must be
 * read in place instead.
 *
 * Rebuilding is one pass of object creation, a fraction of the read it replaces, and happens
 * once per parsed document; a rebuild a host started in slices is finished here in one go.
 * After it the wire form is released, so a host holding the parsed document (as an adapter
 * prop does) does not also hold a second copy of the read.
 */
export function parsedDocumentRead(parsed: ParsedDocument): OoxmlPackageResult | null {
  const rebuild = rebuildOf(parsed);
  advance(parsed, rebuild);
  return rebuild.read;
}

/**
 * Rebuild a received {@link ParsedDocument}'s read now — in the message handler that received
 * it, say — so that the open which follows costs nothing more. Answers the document, or `null`
 * for a value that is not one.
 *
 * Optional: an editor handed a parsed document rebuilds it on open if this has not.
 *
 * @public
 */
export function receiveParsedDocument(value: unknown): ParsedDocument | null {
  if (!isParsedDocument(value)) return null;
  parsedDocumentRead(value);
  return value;
}

/** How {@link receiveParsedDocumentInSlices} spreads the rebuild. */
export interface ReceiveInSlicesOptions {
  /** The longest a slice runs before it yields, in milliseconds. Default 8. */
  readonly sliceMs?: number;
  /** How a slice yields to the thread. Default: the platform's `scheduler.yield`, else a
   *  message-channel task, else `setTimeout(0)`. */
  readonly yieldToHost?: () => Promise<void>;
}

/**
 * {@link receiveParsedDocument}, a slice at a time: the rebuild runs in short tasks so the
 * page never stalls on it, and the promise settles when it is done. Opening the document
 * before then is fine — the open finishes what is left of the rebuild at once.
 *
 * @public
 */
export async function receiveParsedDocumentInSlices(
  value: unknown,
  options: ReceiveInSlicesOptions = {}
): Promise<ParsedDocument | null> {
  if (!isParsedDocument(value)) return null;
  const sliceMs = options.sliceMs ?? 8;
  const yieldToHost = options.yieldToHost ?? platformYield;
  const rebuild = rebuildOf(value);
  for (;;) {
    const deadline = now() + sliceMs;
    if (advance(value, rebuild, () => now() >= deadline)) return value;
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

/**
 * The read a decoded wire value carries, when it may stand in for a read of `bytes`: it names
 * those bytes, has a read's members, and is within the reader's limits. Otherwise `null`, and
 * the bytes are read in place.
 */
function usableRead(value: unknown, bytes: Uint8Array): OoxmlPackageResult | null {
  if (typeof value !== 'object' || value === null) return null;
  const wire = value as Partial<WireRead>;
  if (wire.byteLength !== bytes.byteLength || wire.fingerprint !== fingerprintOf(bytes)) {
    return null;
  }
  const result = wire.result;
  return isPackageResult(result) && withinReaderLimits(result, bytes) ? result : null;
}

/**
 * Whether a rebuilt package is one the reader could have produced from `bytes` under the
 * limits an open applies (`readOoxmlPackage` with none configured): entry count, expanded
 * bytes and their ratio to the file, XML parts and each one's source size, relationships, and
 * each tree's depth and node count. A refusal carries nothing to bound.
 */
function withinReaderLimits(result: OoxmlPackageResult, bytes: Uint8Array): boolean {
  if (!result.ok) return true;
  const pkg = result.package;
  if (pkg.partBytes.size > DEFAULT_ZIP_LIMITS.maxEntries) return false;
  let expanded = 0;
  for (const [name, data] of pkg.partBytes) {
    if (typeof name !== 'string' || !(data instanceof Uint8Array)) return false;
    expanded += data.byteLength;
  }
  if (expanded > DEFAULT_ZIP_LIMITS.maxTotalBytes) return false;
  const maxRatio = DEFAULT_ZIP_LIMITS.maxRatio ?? 200;
  if (expanded > ZIP_RATIO_EXEMPT_BYTES && expanded / Math.max(1, bytes.byteLength) > maxRatio) {
    return false;
  }
  if (pkg.parts.size > DEFAULT_OOXML_PACKAGE_LIMITS.maxXmlParts) return false;
  for (const [name, part] of pkg.parts) {
    const source = pkg.partBytes.get(name);
    if (!source || source.byteLength > XML_HARD_MAX_BYTES) return false;
    if (typeof part !== 'object' || part === null || !treeWithinLimits(part.root)) return false;
  }
  let relationships = pkg.externalTargets.length;
  for (const records of pkg.relationships.values()) {
    if (!Array.isArray(records)) return false;
    relationships += records.length;
  }
  return relationships <= DEFAULT_OOXML_PACKAGE_LIMITS.maxRelationships;
}

/**
 * Whether a part's tree is no deeper than the XML reader admits and holds no more nodes than
 * it could produce (elements, plus at most one text node beside each). Walked with its own
 * stack, so a forged tree cannot exhaust the call stack; a cycle exceeds the count and stops.
 */
function treeWithinLimits(root: unknown): boolean {
  const maxNodes = 2 * XML_DEFAULT_MAX_ELEMENTS + 1;
  const stack: { readonly node: unknown; readonly depth: number }[] = [{ node: root, depth: 1 }];
  let nodes = 0;
  while (stack.length > 0) {
    const { node, depth } = stack.pop()!;
    if (typeof node !== 'object' || node === null) return false;
    nodes += 1;
    if (nodes > maxNodes || depth > MAX_XML_DEPTH + 1) return false;
    const children = (node as { children?: unknown }).children;
    if (children === undefined) continue;
    if (!Array.isArray(children)) return false;
    for (const child of children) stack.push({ node: child, depth: depth + 1 });
  }
  return true;
}

/** The members a read's result has, checked before it stands in for a read of the bytes. */
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
