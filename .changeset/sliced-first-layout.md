---
'@docx-editor.dev/core': minor
---

A large document now opens on its first pages, which you can read and edit at once, and lays out the rest in short tasks, so the page stays responsive while it opens. The pages and line breaks are the same as before. `snapshot().isPaginating` is true until every page is laid out, and the `paginated` editor event fires when it turns false. On the paginated surface, `paginating()`, `whenPaginated()` and `finishPagination()` report and complete the remaining layout, and a reveal of something not yet laid out lays out as far as it needs first.
