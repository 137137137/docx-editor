---
'@docx-editor.dev/core': patch
---

A document opened with its fonts already resolved breaks its pages where it did when its fonts resolved after it opened. Its first layout runs on the fallback measurer before the shaped one, without painting it, because a document with footnotes settles its footnote space from where the previous layout left it. A small document again mounts at the zoom the editor holds and fits after, as before.
