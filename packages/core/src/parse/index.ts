/**
 * `@docx-editor.dev/core/parse` — read a DOCX ahead of its open, on a worker.
 *
 * OPTIONAL, for prefetching. Reading a package is most of what an open costs before layout.
 * This entry does that read before any editor exists and where it cannot hold up the page — a
 * dedicated Web Worker — and the page takes the answer through `createDocumentParser` from
 * `@docx-editor.dev/core/editor`; an editor then opens it without reading the file again. An
 * editor handed plain bytes never comes here.
 *
 * Provenance: a parsed document skips the read only on the page that issued it — the reply its
 * own worker sent to a request it is still waiting on. Any other value (a clone, one from a
 * server, another origin or storage) opens from its bytes through the ordinary bounded read.
 *
 * @example In the worker
 * ```ts
 * import { serveDocumentParsing } from '@docx-editor.dev/core/parse';
 *
 * serveDocumentParsing(self);
 * ```
 *
 * Nothing here touches a DOM, a window or a canvas.
 *
 * @packageDocumentation
 * @public
 */

export {
  parseDocument,
  parsedDocumentTransferables,
  serveDocumentParsing,
  type ParsedDocument,
  type ParseWorkerScope,
} from '../store/package/parsed-document.ts';
