---
'@docx-editor.dev/core': patch
---

A drawing wrapped top and bottom now clears its band when it is set behind the text, as Word does: the paragraphs after it flow below it instead of under it. A band offset below its paragraph or line is placed where the drawing is, so the anchor paragraph's own text above it keeps its line.
