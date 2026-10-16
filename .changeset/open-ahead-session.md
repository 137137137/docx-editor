---
'@docx-editor.dev/core': minor
---

`prepareParsedDocument(parsed)` (`@docx-editor.dev/core/editor`) opens a parsed document's session ahead of its editor and reads the font families its open asks for. An editor handed the same parsed document opens from that session, so the work between `load()` and the font resolution is already done: on a long document, a quarter of a second of the page's thread moves to before the open. The session is the one the editor would have opened, from the same read with the same options, and the pages and `save()` are unchanged. A session is taken once; an editor with a review module opens its own.
