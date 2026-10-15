---
'@docx-editor.dev/core': patch
---

A relayout while a large document is still being laid out (an image that finished decoding, fonts that arrived, a zoom) is now one more pass after the opening's layout, as it is after an open without slices, instead of starting the layout again: a footnoted document's pages depend on how many passes it has had, so they are now the ones an open without slices produces.
