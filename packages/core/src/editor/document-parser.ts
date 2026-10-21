// The page side of a parse worker. OPTIONAL: nothing here runs unless a host creates one.
//
// A read skips the bounded read of its bytes only when this page issued it (see the header of
// `parsed-document.ts`). Here that means: the reply came from the dedicated Worker this page
// constructed and registered, and it answers a request this parser sent and has not given up
// on. A reply to an aborted or superseded request, a message from anywhere else, and a parsed
// document that reaches the editor by any other route are never issued, and open from bytes.

import {
  PARSE_REPLY,
  PARSE_REQUEST,
  issueParsedDocument,
  rebuildParsedDocumentInSlices,
  type ParsedDocument,
} from '../store/package/parsed-document.ts';

/** @public */
export interface DocumentParseOptions {
  /** Abort to give up on this file: the promise settles `null` and a late reply is dropped. */
  readonly signal?: AbortSignal;
}

/** @public */
export interface DocumentParser {
  /**
   * Read `bytes` on the worker. Settles with the parsed document, its read already rebuilt on
   * this page in short tasks, or `null` when the request was aborted, the parser disposed, or
   * the worker failed. `bytes` are copied, never detached, so the host can still open them.
   */
  parse(bytes: Uint8Array, options?: DocumentParseOptions): Promise<ParsedDocument | null>;
  /** Stop listening to the worker and settle every pending request `null`. */
  dispose(): void;
}

/**
 * Parse documents on a dedicated Worker this page constructed, whose script calls
 * `serveDocumentParsing(self)` from `@docx-editor.dev/core/parse`. A parsed document it answers
 * opens without a package read when handed to an editor on this page.
 *
 * @public
 */
export function createDocumentParser(worker: Worker): DocumentParser {
  if (typeof Worker !== 'function' || !(worker instanceof Worker)) {
    throw new TypeError('createDocumentParser takes a dedicated Worker this page constructed');
  }
  let nextId = 0;
  const pending = new Map<number, (reply: unknown) => void>();
  const onMessage = (event: MessageEvent): void => {
    const reply = event.data as { kind?: unknown; id?: unknown; parsed?: unknown } | null;
    if (reply?.kind !== PARSE_REPLY || typeof reply.id !== 'number') return;
    const settle = pending.get(reply.id);
    pending.delete(reply.id);
    settle?.(reply.parsed);
  };
  const failAll = (): void => {
    const settles = [...pending.values()];
    pending.clear();
    for (const settle of settles) settle(null);
  };
  worker.addEventListener('message', onMessage);
  worker.addEventListener('error', failAll);

  return {
    async parse(bytes, { signal } = {}) {
      if (signal?.aborted) return null;
      const id = nextId++;
      const reply = await new Promise<unknown>((resolve) => {
        const abort = (): void => {
          if (pending.delete(id)) resolve(null);
        };
        pending.set(id, (value) => {
          signal?.removeEventListener('abort', abort);
          resolve(value);
        });
        signal?.addEventListener('abort', abort, { once: true });
        const copy = bytes.slice();
        worker.postMessage({ kind: PARSE_REQUEST, id, bytes: copy }, [copy.buffer]);
      });
      const parsed = issueParsedDocument(reply);
      if (!parsed) return null;
      await rebuildParsedDocumentInSlices(parsed);
      return signal?.aborted ? null : parsed;
    },
    dispose() {
      worker.removeEventListener('message', onMessage);
      worker.removeEventListener('error', failAll);
      failAll();
    },
  };
}
