---
'@docx-editor.dev/core': minor
---

Parse a document on a worker before its editor opens. In the worker, `serveDocumentParsing(self)` from `@docx-editor.dev/core/parse` answers requests with the open's own bounded package read. On the page, `createDocumentParser(worker)` from `@docx-editor.dev/core/editor` sends the worker a copy of the bytes and settles with a `ParsedDocument` once its read is rebuilt in short tasks, or with `null` if the request was aborted. `createDocxEditor({ document })`, `load()` and an adapter's `document` prop accept the parsed document: the open skips the read and goes straight to fonts and the first layout, with the same pages and the same `save()` as the bytes. The path is opt-in: an editor handed bytes opens exactly as before. A read is used only on the page that issued it, meaning the reply its own worker sent to a request still pending; any other parsed document opens from its bytes through the ordinary bounded read.
