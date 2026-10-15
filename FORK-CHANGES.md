# Changes in this fork

This fork of [docx-editor](https://github.com/eigenpal/docx-editor) is licensed under the Apache License 2.0, like the original. As section 4(b) of that license requires, every file changed from the original release is listed here with what changed.

## Modified by Counselcite from @docx-editor.dev 2.21.0

Fonts and fit before the first layout:

- `packages/core/src/editor/docx-editor.ts`: a large document is opened before its mount, its fonts are resolved and its fitted zoom is settled before the first layout, and the mount waits for the fonts for a bounded time; a font resolution that finishes before the mount no longer re-measures the previous document; font helpers moved to `docx-editor-open-fonts.ts`.
- `packages/core/src/editor/docx-editor-open-fonts.ts` (new): the families a resolver is asked for, the first page's width, and the font error report, read from the parsed document.
- `packages/core/src/editor/docx-editor-open-scheduler.ts`: a scheduled open can carry preparation that the mount waits for, at most `OPEN_PREPARE_WAIT_MS`.
- `packages/core/src/editor/docx-editor-zoom.ts`: `prefit` settles a fit mode's scale from a page width before anything is laid out.
- `packages/core/src/editor/zoom-controller.ts`: `fitFor` answers the fit for a page width without applying it.
- `packages/core/src/editor/live-font-resolution.ts`: `adopt` records a resolution started before the mount as the load's initial one.
- `packages/core/src/editor/paginated-surface.ts`: the surface can mount a session the facade already opened.
- `packages/core/src/editor/index.ts`: exports `fitZoom` and `FitZoomInput`.
- `packages/core/src/editor/zoom-fit.ts`: documents the bounds of `FitZoomInput`.
- `packages/core/src/contracts/editor.ts`: documents the font wait on `load` and `isOpening`.
- `packages/core/src/layout/layout-shaping.ts`: `createLayoutShaping` copies every face before it awaits and digests the copies asynchronously.
- `packages/core/src/layout/font-resource.ts`: an asynchronous twin of `prepareFontResourceDefinition`.
- `packages/core/src/store/package/sha256.ts`: `sha256BytesAsync`, which uses `crypto.subtle` where there is no synchronous native hasher.
- `docs/api/docx-editor-core/editor.api.md`: the API snapshot for the new exports.
- `.changeset/open-once-at-fit.md` (new): the release note.
- `.collaboration/changes/open-once-at-fit.json` (new): the collaboration compatibility decision (no impact).
- Tests: `packages/core/src/editor/__tests__/fonts-before-first-layout.test.ts` and `open-scheduler-prepare.test.ts` (new); `font-configuration.test.ts`, `live-font-resolution.test.ts`, `zoom-controller.test.ts`, `zoom-lane.test.ts` in the same folder, and `packages/core/src/store/__tests__/sha256.test.ts`.

The first layout of a large document, in slices:

- `packages/core/src/layout/layout-steps.ts` (new): layout steps as generators that pause between blocks, with helpers to drain or advance them.
- `packages/core/src/layout/semantic-layout.ts`: the body pass, its drawing-exclusion and frame-fallback passes, the section preparation and the table flow pause between blocks and rows (`layoutSemanticDocumentSteps`); `layoutSemanticDocument` drains the same steps. A `bodyBlockLimit` lays out only the first body blocks.
- `packages/core/src/layout/semantic-layout-options.ts`: the `bodyBlockLimit` option.
- `packages/core/src/layout/body-block-limit.ts` (new): the body prefix and its clipped sections for `bodyBlockLimit`.
- `packages/core/src/layout/document-layout-coordinator.ts`: `layoutDocumentViewSteps`, and `bodyBlockLimit` passed through.
- `packages/core/src/layout/column-balance-layout.ts`: the column-balancing driver runs its passes as steps.
- `packages/core/src/layout/multi-section-layout.ts`: the multi-section driver runs its sections as steps.
- `packages/core/src/layout/note-pagination.ts`: the footnote reserve search, the reserve computation, note attachment and endnote placement run as steps that pause between passes and pages.
- `packages/core/src/layout/section-preparation.ts`: section preparation pauses after each block it prepares.
- `packages/core/src/layout/table-flow-pagination.ts`: table pagination pauses between body rows.
- `packages/core/src/export/export-session.ts`: names the new coordinator option.
- `packages/core/src/editor/surface-opening-layout.ts` (new): a large document opens on a preview of its first pages and lays out the rest in slices; relayouts, reveals and scrolling during that time.
- `packages/core/src/editor/paginated-surface.ts`: mounts through the opening when asked, routes relayouts, reveals and scrolling through it during the opening, and reports it (`paginating`, `whenPaginated`, `finishPagination`).
- `packages/core/src/editor/paginated-surface-contract.ts`: documents `paginating`, `whenPaginated` and `finishPagination`.
- `packages/core/src/editor/surface-pages.ts`: `layoutCoversView`, whether a layout reaches past the view.
- `packages/core/src/editor/docx-editor.ts`: a document large enough to open behind a painted frame opens progressively; `snapshot().isPaginating` and the `paginated` event; two per-layout indexes moved to `docx-editor-layout-index.ts`.
- `packages/core/src/editor/docx-editor-layout-index.ts` (new): the per-layout note-scope and paragraph-order indexes, moved from `docx-editor.ts`.
- `packages/core/src/editor/docx-editor-equality.ts` and `packages/core/src/editor/loading-snapshot.ts`: the `isPaginating` snapshot field.
- `packages/core/src/contracts/editor.ts`: `EditorSnapshot.isPaginating` and the `load` documentation; `HyperlinkInfo` and `PageSetup` moved to `editor-page-reads.ts`.
- `packages/core/src/contracts/editor-page-reads.ts` (new): `HyperlinkInfo` and `PageSetup`, moved from `editor.ts` and re-exported there.
- `packages/core/src/contracts/editor-events.ts`: the `paginated` event.
- `docs/api/docx-editor-core/contracts-editor.api.md`, `editor.api.md`, `index.api.md`, `layout.api.md`: the API snapshots for the new members.
- `docs/site/content/react/composition.mdx`, `docs/site/content/react/hooks.mdx`, `docs/site/content/vue/composables.mdx`: `isPaginating` and the `paginated` event.
- `.changeset/sliced-first-layout.md` (new): the release note.
- `.collaboration/changes/sliced-first-layout.json` (new): the collaboration compatibility decision (no impact).
- Tests: `packages/core/src/layout/__tests__/layout-steps.test.ts`, `packages/core/src/editor/__tests__/surface-opening-layout.test.ts` and `progressive-open.test.ts` (new).

The first layout of a document opened with its fonts, as the open without them laid it out:

- `packages/core/src/editor/surface-fixed-first.ts` (new): when a layout runs the fixed measurer's pass before the shaped one, and the fixed measurer an unshaped mount resolves.
- `packages/core/src/editor/paginated-surface.ts`: a mount whose fonts were resolved first lays out on the fixed measurer and then the shaped one on the same session, painting only the second; so does an opening that started on the fixed measurer when the fonts arrive part way.
- `packages/core/src/editor/surface-opening-layout.ts`: every preview starts from a cold session when the layouts run both passes.
- `packages/core/src/editor/docx-editor.ts`: tells the surface when the fonts were resolved before its mount; only a deferred open settles its fit before the mount.
- `scripts/check-eslint-max-lines-globs.mjs`: the declared length of `paginated-surface.ts`.
- `.changeset/open-once-at-fit.md`: the release note, corrected; `.changeset/fixed-pass-first.md` and `.collaboration/changes/fixed-pass-first.json` (new): the release note and the collaboration compatibility decision (no impact).
- Tests: `progressive-open.test.ts`, `fonts-before-first-layout.test.ts` and `zoom-controller.test.ts` in `packages/core/src/editor/__tests__/`.

Every document opened the way a large one is, when the host asks:

- `packages/core/src/editor/docx-editor-types.ts`: the `openAfterPaint` option.
- `packages/core/src/editor/docx-editor-open-scheduler.ts`: an open yields for every document when asked; `isLarge` says whether it would by default.
- `packages/core/src/editor/docx-editor.ts`: passes the option to the scheduler, and tells the surface the zoom a small document held before its fit.
- `packages/core/src/editor/surface-fixed-first.ts` and `packages/core/src/editor/paginated-surface.ts`: the pass at that held zoom runs first, unpainted.
- `packages/react/src/editor/DocxEditorRoot.tsx`, `packages/vue/src/editor/DocxEditorRoot.ts`, `packages/vue/src/editor/useDocxEditorRoot.ts`: the `openAfterPaint` prop.
- `docs/site/content/react/composition.mdx`, `docs/site/content/vue/composition.mdx`: the prop documented.
- `docs/api/docx-editor-core/editor.api.md`, `index.api.md`, `docs/api/docx-editor-react/index.api.md`, `docs/api/docx-editor-vue/index.api.md`: the API snapshots.
- `scripts/feature-parity-baseline.json`: the prop, added to both adapters' roots.
- `.changeset/open-after-paint-every.md` and `.collaboration/changes/open-after-paint-every.json` (new).
- Tests: `fonts-before-first-layout.test.ts` and `progressive-open.test.ts` in `packages/core/src/editor/__tests__/`.

A relayout during a sliced opening, as after an unsliced one:

- `packages/core/src/editor/surface-opening-layout.ts`: a relayout during the opening is one more pass after the open's layout, added to the run in progress (or, after a commit, to the run started again); every preview starts from a cold session.
- `packages/core/src/editor/paginated-surface.ts`: the open's layout runs on the measurer, scale and furniture of the mount, and a relayout's pass on the current ones.
- `packages/core/src/editor/surface-fixed-first.ts`: the fixed pass runs first only for an open whose fonts were resolved before its mount.
- `scripts/check-eslint-max-lines-globs.mjs`: the declared length of `paginated-surface.ts`.
- `.changeset/opening-relayout-pass.md` and `.collaboration/changes/opening-relayout-pass.json` (new): the release note and the collaboration compatibility decision (no impact).
- Tests: `surface-opening-layout.test.ts` and `progressive-open.test.ts` in `packages/core/src/editor/__tests__/`.

A document parsed off the page's thread, and opened from that read:

- `packages/core/src/store/package/parsed-document.ts` (new): `parseDocument`, which runs the open's bounded package read on any thread and answers a `ParsedDocument` that crosses `postMessage`; `parsedDocumentTransferables`; `receiveParsedDocument` and `receiveParsedDocumentInSlices`, which rebuild the read on the receiving thread in one task or in short ones; `parsedDocumentRead`, the rebuilt read an open uses.
- `packages/core/src/store/package/parsed-document-codec.ts` (new): the wire form of a read: one integer stream with tables of distinct strings and object shapes, rebuilt with the key order, sharing and frozen state the read had, resumably.
- `packages/core/src/parse/index.ts` and `packages/core/src/parse/tsconfig.json` (new): the `@docx-editor.dev/core/parse` entry, DOM-free.
- `packages/core/package.json`, `packages/core/tsup.config.ts`, `packages/core/tsconfig.json`: the `./parse` subpath.
- `packages/core/src/contracts/editor.ts`: `DocumentSource` takes a `ParsedDocument`.
- `packages/core/src/binding/tree-session.ts`: `openTreeSession` takes the package already read from its bytes (`read`).
- `packages/core/src/editor/docx-editor-support.ts`: `createDocumentSources`, which resolves a source to its bytes and opens those bytes from the read a parsed document carried; `normalizeSource` and `unloadableSourceReason` know a parsed document.
- `packages/core/src/editor/docx-editor.ts`: sources resolve and open through `createDocumentSources`.
- `packages/core/src/__tests__/core-lane-graph.ts` and `core-lane-graph.test.ts`, `scripts/check-lane-boundaries.mjs`: the `parse` lane, neutral, importing the store.
- `docs/api/docx-editor-core/parse.api.md` (new), `contracts-editor.api.md`, `index.api.md`: the API snapshots.
- `docs/site/content/guides/loading-and-saving.mdx`: parsing in a worker.
- `.changeset/parse-off-thread.md` and `.collaboration/changes/parse-off-thread.json` (new): the release note and the collaboration compatibility decision (no impact).
- Tests: `packages/core/src/store/__tests__/parsed-document.test.ts`, `same-graph.ts` and `packages/core/src/editor/__tests__/parsed-document-open.test.ts` (new).
