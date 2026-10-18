---
'@docx-editor.dev/core': patch
---

Laying out a long document keys each paragraph with less work: the key's memo compares its inputs before anything is hashed, the key text is digested in a browser without a fresh UTF-8 buffer per call, and the canvas measurer's cache moves an entry it hits instead of deleting and re-inserting it. The pages are the same.
