---
'@docx-editor.dev/core': patch
---

Laying out a long document does each paragraph's fixed work once instead of again on every line or every pass: the drawing offsets and anchored atoms of a paragraph, a table cell paragraph's resolved inputs and break properties, and the cache key's digest are kept per immutable node; the key's inputs are compared before anything is hashed; the SHA-256 that digests them in a browser no longer copies or allocates per call; and the canvas measurer's cache moves an entry it hits instead of deleting and re-inserting it. The pages are the same.
