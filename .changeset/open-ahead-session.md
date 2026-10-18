---
'@docx-editor.dev/core': minor
---

`prepareParsedDocument(parsed)` (`@docx-editor.dev/core/editor`) opens a parsed document's session ahead of its editor and reads the font families its open asks for. An editor handed the same parsed document opens from that session, so the work between `load()` and the font resolution is already done. Opt-in: the session is the one the editor would have opened, from the same checked read with the same options, and the pages and `save()` are unchanged. A session is taken once; an editor with a review module or a content-lock policy opens its own, and an editor handed bytes never uses it.
