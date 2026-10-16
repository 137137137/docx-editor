---
'@docx-editor.dev/core': patch
---

`findMatches` builds each paragraph's searchable text once per revision instead of once per query. A host asking many queries of one document (placing a list of findings, say) paid a walk of every story and a rebuild of every paragraph's text for each one: about 33 ms a query on a 120-page law review, now under 1 ms. The answers are the same.
