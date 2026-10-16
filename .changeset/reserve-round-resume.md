---
'@docx-editor.dev/core': patch
---

The footnote reserve search no longer lays the whole body out again in each round. A round whose reserves first differ from the last round's on some page resumes from the checkpoint before that page, carrying the earlier pages over: a pass reads a page's reserve only while it fills that page, so those pages are the ones a full round lays out. A long footnoted document finishes its first layout in about half the time, with the same pages.
