---
'@docx-editor.dev/core': minor
'@docx-editor.dev/react': minor
'@docx-editor.dev/vue': minor
---

A document that opens behind a painted frame now shows its first pages as soon as they are laid out. `snapshot().isOpening` turns false then, and the rest of the document is laid out in short tasks while the new `snapshot().isPaginating` holds; the new `paginated` event fires once when every page is laid out (and just after the mount for a document laid out in one task). The pages can be read, edited, searched and scrolled meanwhile; a reveal of something not laid out yet, or a scroll to pages not laid out yet, lays out as far as them. The tasks give way to input: none starts while a key or a click waits, and they resume once it has settled. The fitted zoom (`zoomMode: 'auto'`) is settled before the first layout. A relayout during the opening, such as an image that decoded or fonts that landed, is one more pass in the same slices. When every page is laid out the pages are the ones the open laid out before. An external document refresh still shows its replacement muted, without input, until every page is laid out and the refresh reports its result.
