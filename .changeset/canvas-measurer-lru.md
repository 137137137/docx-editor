---
'@docx-editor.dev/core': patch
---

The canvas measurer's width and line-metrics caches move an entry they hit to most recently used instead of deleting and re-inserting it. What is cached and what is evicted are unchanged.
