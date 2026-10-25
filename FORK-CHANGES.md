# Changes in this fork

This fork of [docx-editor](https://github.com/eigenpal/docx-editor) is licensed under the Apache License 2.0, like the original. As section 4(b) of that license requires, every file changed from the original release is listed here with what changed.

## Modified by Counselcite from @docx-editor.dev 2.31.0

A section's note number format when the section authors note properties without one:

- `packages/core/src/store/package/note-properties.ts`: `sectionFormat` numbers such a section in the default format, as Word does, instead of the document-wide format from settings.xml.
- Tests: `packages/core/src/store/package/__tests__/note-properties-format.test.ts` (new).
- `packages/core/src/layout/document-notes-input.ts`, `packages/core/src/layout/field-noteref.ts`: the last section reads its own body-level `w:sectPr` for its note properties, as `surface-note-state.ts` already did, instead of the document defaults.
- Tests: `packages/core/src/layout/__tests__/final-section-note-format.test.ts` (new).

A document search that builds each paragraph's text once per revision:

- `packages/core/src/binding/document-search.ts`: a paragraph's searchable text is kept per part and paragraph node, so a second query of the same revision does not build it again.
- `packages/core/src/binding/document-search-frames.ts`: `searchStoryParagraphs`, a story's paragraphs kept per story root.
- `.changeset/search-text-once.md`, `.collaboration/changes/search-text-once.json` (new).
- Tests: `packages/core/src/binding/__tests__/document-search.test.ts`.

A `wrapTopAndBottom` band anchored on a later line of its paragraph:

- `packages/core/src/layout/top-and-bottom-clearance.ts`: `topAndBottomBandAnchorY` measures a paragraph-relative offset from the paragraph's top, not from the line that holds the anchor.
- `packages/core/src/layout/drawing-exclusion.ts`: passes the paragraph's top.
- `.changeset/top-and-bottom-behind-text.md`, `.collaboration/changes/top-and-bottom-behind-text.json` (new).
- Tests: `packages/core/src/layout/__tests__/top-and-bottom-band.test.ts` (new).

A paragraph's fixed layout work done once:

- `packages/core/src/layout/layout-cache.ts`: the paragraph key memo compares its raw inputs before any digest is taken.
- `packages/core/src/layout/canvas-measurer.ts`: the width and metrics caches keep their recency order in a linked list beside the map.
- `packages/core/src/store/package/sha256.ts`: `sha256Text` encodes into a reused buffer on the pure hasher.
- `.changeset/layout-work-once.md`, `.collaboration/changes/layout-work-once.json` (new).
- Tests: `packages/core/src/store/__tests__/sha256.test.ts`.

A document parsed off the page's thread (optional), and its session opened ahead of the editor:

- `packages/core/src/store/package/parsed-document.ts` and `parsed-document-codec.ts` (new): `parseDocument`, the wire form of a read, and the checks a received read must pass (version, integrity, the bytes it travels with, the reader's limits) before it is used; anything else opens from its bytes.
- `packages/core/src/parse/index.ts`, `packages/core/src/parse/tsconfig.json` (new): the DOM-free `@docx-editor.dev/core/parse` entry.
- `packages/core/src/editor/opened-ahead.ts` (new): `prepareParsedDocument`.
- `packages/core/src/editor/docx-editor-prepared-open.ts`, `docx-editor-support.ts`, `docx-editor.ts`, `index.ts`: an open holding a parsed read uses it; ordinary sources open as before.
- `packages/core/src/contracts/editor.ts`: `DocumentSource` takes a `ParsedDocument`.
- `packages/core/package.json`, `packages/core/tsup.config.ts`, `packages/core/tsconfig.json`: the `./parse` subpath.
- `packages/core/src/__tests__/core-lane-graph.ts`, `core-lane-graph.test.ts`, `scripts/check-lane-boundaries.mjs`: the `parse` lane.
- `.oxlintrc.json`: the line cap of `contracts/editor.ts`.
- `docs/api/docx-editor-core/parse.api.md` (new), `contracts-editor.api.md`, `editor.api.md`, `index.api.md`; `docs/site/content/guides/loading-and-saving.mdx`; `packages/core/README.md`.
- `.changeset/parse-off-thread.md`, `.changeset/open-ahead-session.md`, `.collaboration/changes/parse-off-thread.json`, `.collaboration/changes/open-ahead-session.json` (new).
- Tests: `packages/core/src/store/__tests__/parsed-document.test.ts`, `same-graph.ts`, `packages/core/src/editor/__tests__/parsed-document-open.test.ts`, `opened-ahead.test.ts` (new).

A document opened on its first pages, the rest laid out in slices:

- `packages/core/src/layout/note-pagination.ts`: the footnote reserve scan and the note attach pause between pages; synchronous callers drain them.
- `packages/core/src/editor/surface-deferred-layout.ts`: an opening mode that defers a pass after the first pages and is advanced only by the opening's slices.
- `packages/core/src/editor/surface-progressive-open.ts`, `surface-runtime-options.ts`: the mount paints the first pages; later prefixes and the full layout run in slices; a relayout during the opening is one more pass.
- `packages/core/src/editor/paginated-surface.ts`: pagination state, deferred flushes while paginating, reveals that lay out to their target.
- `packages/core/src/editor/docx-editor-open-scheduler.ts`, `docx-editor.ts`, `docx-editor-types.ts`, `docx-editor-equality.ts`, `loading-snapshot.ts`, `editor-events.ts`: `isOpening` ends at the mount; `snapshot().isPaginating`; the `paginated` event; the `openAfterPaint` option.
- `packages/core/src/editor/docx-editor-zoom.ts`, `zoom-controller.ts`: the fitted zoom settled before the first layout (`prefit`, `fitFor`).
- `packages/core/src/contracts/editor.ts`, `editor-events.ts`: `EditorSnapshot.isPaginating`, `EditorEvents.paginated`.
- `packages/react/src/editor/DocxEditorRoot.tsx`, `DocxEditorLoading.tsx`, `packages/vue/src/editor/DocxEditorRoot.ts`, `useDocxEditorRoot.ts`, `DocxEditorLoading.ts`: the `openAfterPaint` prop; the first-page indicator kept while paginating.
- `scripts/check-max-lines-caps.mjs`, `scripts/feature-parity-baseline.json`.
- `docs/api/docx-editor-core/contracts-editor.api.md`, `editor.api.md`, `index.api.md`, `docs/api/docx-editor-react/index.api.md`, `docs/api/docx-editor-vue/index.api.md`; `docs/site/content/react/composition.mdx`, `hooks.mdx`, `docs/site/content/vue/composition.mdx`, `composables.mdx`.
- `.changeset/page-one-first.md`, `.collaboration/changes/page-one-first.json` (new).
- Tests: `packages/core/src/editor/__tests__/page-one-first.test.ts` (new); `progressive-open.test.ts`, `opening-indicator.test.ts`, `opened-ahead.test.ts`, `parsed-document-open.test.ts`.
