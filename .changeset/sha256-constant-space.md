---
'@docx-editor.dev/core': patch
---

Hash complete blocks directly from the caller's byte view and allocate only the 64- or 128-byte padding tail. Preserve digest encoding, native hashing, and caller bytes while avoiding a payload-sized copy in the pure SHA-256 path.
