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
