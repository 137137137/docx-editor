---
'@docx-editor.dev/core': minor
'@docx-editor.dev/react': minor
'@docx-editor.dev/vue': minor
---

`openAfterPaint: 'every'` (a `createDocxEditor` option and a `DocxEditor.Root` prop) opens every document the way a large one opens: behind a painted frame, with its fonts resolved before its first layout, on its first pages while the rest is laid out. The default, `'large'`, keeps a small document's synchronous mount. The pages and line breaks are the same either way: a small document deferred this way still runs the pass at the zoom the editor held before its fit.
