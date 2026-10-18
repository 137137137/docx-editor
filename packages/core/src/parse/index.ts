/**
 * `@docx-editor.dev/core/parse` — read a DOCX ahead of its open, on any thread.
 *
 * OPTIONAL, for prefetching. Reading a package is most of what an open costs before layout.
 * This entry does that read before any editor exists and where it cannot hold up the page — a
 * Web Worker — and hands back a `ParsedDocument` that crosses `postMessage`; the editor opens
 * from it without reading the file again. An editor handed plain bytes never comes here.
 *
 * Trust model: a parsed document comes from `parseDocument` in a same-origin worker the host
 * started itself, never from a server, another origin or storage. The editor checks it all the
 * same — wire version, the bytes it names, the reader's limits — and opens anything that fails
 * from its bytes, through the ordinary bounded read.
 *
 * @example In a worker
 * ```ts
 * import { parseDocument, parsedDocumentTransferables } from '@docx-editor.dev/core/parse';
 *
 * self.onmessage = (event) => {
 *   const parsed = parseDocument(new Uint8Array(event.data));
 *   self.postMessage(parsed, parsedDocumentTransferables(parsed));
 * };
 * ```
 *
 * @example On the page
 * ```ts
 * import { receiveParsedDocumentInSlices } from '@docx-editor.dev/core/parse';
 *
 * worker.onmessage = async (event) => {
 *   // Rebuilds the read in short tasks; `receiveParsedDocument` does it in one.
 *   const parsed = await receiveParsedDocumentInSlices(event.data);
 *   if (parsed) editor.load(parsed); // or <DocxEditor.Root document={parsed} />
 * };
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
  receiveParsedDocument,
  receiveParsedDocumentInSlices,
  type ParsedDocument,
  type ReceiveInSlicesOptions,
} from '../store/package/parsed-document.ts';
