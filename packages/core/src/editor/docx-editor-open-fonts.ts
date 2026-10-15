// What the facade reads from a parsed document before its first layout.
//
// A document's first layout should be its only one: on the measurer its fonts make, and at
// the scale its zoom mode settles on. Both answers come from the parsed tree alone — the
// families the document names and the size of its first page — so they are read from the
// session before the surface mounts rather than from the surface after it. Kept out of
// `docx-editor.ts`, which sits against the max-lines gate, together with the font error
// report the same lane raises.

import type { TreeDocxSessionView } from '@docx-editor.dev/core/binding';
import {
  FontResolutionError,
  HarfBuzzShapingError,
  fontRequestKey,
  type LayoutShapingOptions,
} from '@docx-editor.dev/core/layout';
import type { EditorFontError, FontConfiguration } from '@docx-editor.dev/core/contracts/editor';
import {
  DEFAULT_SECTION_PROPERTIES,
  enumerateDocumentSections,
  geometryOfSection,
} from '../layout/index.ts';
import { layoutProjectionOf, type ReviewDisplayMode } from '../layout/revision-projection.ts';
import { fontResolverFamilies } from './font-availability.ts';
import { MAX_RESOLVER_FAMILIES } from './font-composition.ts';
import { resolverGlyphFontFamilies } from './resolver-glyph-font-families.ts';
import { toEditorFontError, warnFontFailureOnce } from './font-configuration.ts';
import { composeFontOrigins, defineFontResolver } from './font-resolver.ts';
import { normalizeFontResolverResult, type FontConfigurationBase } from './font-composition.ts';
import type { DocxEditorConfig } from './docx-editor-types.ts';

/**
 * The families a font resolver is asked for on behalf of a document: the family at the caret
 * when it is not the configured default, then every family the document declares, bounded,
 * with the symbol faces its glyphs need reserved a place.
 */
export function documentFontRequestFamilies(
  session: TreeDocxSessionView,
  selected: string | undefined,
  defaultFamily: string | undefined
): readonly string[] {
  return fontResolverFamilies(
    [
      ...new Set([
        ...(selected && selected !== defaultFamily ? [selected] : []),
        ...session.documentFonts(),
      ]),
    ],
    resolverGlyphFontFamilies(session),
    MAX_RESOLVER_FAMILIES
  );
}

/**
 * The first page's width in CSS pixels at 100%, as layout will give it: the first section's
 * page size, read with the display mode the surface lays out with, because that mode decides
 * which section a paragraph belongs to.
 *
 * Only a prediction. The fit that runs after the mount still reads the laid-out page, and if
 * the two ever disagree that fit moves the scale exactly as it did before.
 */
export function firstPageWidthPx(
  session: TreeDocxSessionView,
  displayMode: ReviewDisplayMode
): number {
  const [first] = enumerateDocumentSections(session.part(), layoutProjectionOf(displayMode));
  return geometryOfSection(first?.properties ?? DEFAULT_SECTION_PROPERTIES).width * (96 / 72);
}

/**
 * Report one font failure to the host's `onFontError` and the editor's `error` event, whose
 * listeners `listening` says exist.
 */
export function reportEditorFontError(
  error: EditorFontError,
  onFontError: ((error: EditorFontError) => void) | undefined,
  listening: boolean,
  emit: (error: EditorFontError) => void
): void {
  // A shaper that never loaded is not a per-face degradation: NOTHING measures correctly,
  // and the fault is the host's build rather than the document's content. Said out loud
  // exactly once, and only when nobody is listening — a host that reports it in its own
  // UI does not need library noise it cannot switch off. Before this was fixed the same
  // misconfiguration failed the BUILD, so a green build with a silent console would be a
  // strictly worse trade (#282).
  const shaperNeverLoaded =
    error.code === 'wasmUnavailable' ||
    (error.cause instanceof HarfBuzzShapingError && error.cause.code === 'unsupportedRuntime');
  if (shaperNeverLoaded && !onFontError && !listening) {
    warnFontFailureOnce(error);
  }
  // A host handler that throws must not abort font resolution — reporting a dropped
  // face would then cost the whole shaped measurer, and the catch's own report would
  // throw again as an unhandled rejection.
  try {
    onFontError?.(error);
  } catch {
    /* the host's reporting problem, not the document's */
  }
  emit(error);
}

/**
 * The requests in `composed` that `shaping` refused, each reported as it is found.
 *
 * Per-face degradation: an embedded face the validator refused still resolves — to a typed
 * error. Probing here (map lookups, no shaping work) is what turns a silent fixed-measurer
 * fallback into a diagnosable one.
 */
export function refusedFontRequests(
  shaping: LayoutShapingOptions,
  composed: FontConfiguration,
  report: (error: EditorFontError) => void
): Set<string> {
  const refused = new Set<string>();
  for (const source of composed.sources ?? []) {
    const resolved = shaping.fonts.resolve(source.request);
    if (resolved instanceof FontResolutionError) {
      report(toEditorFontError(resolved));
      refused.add(fontRequestKey(source.request));
    }
  }
  return refused;
}

/**
 * The faces of `fonts` that `shaping` admitted, as the validated copies it owns, for the paint
 * side to register. Never the raw file view: what paints must be what was measured.
 */
export function admittedFontSources(shaping: LayoutShapingOptions, fonts: FontConfiguration) {
  return (fonts.sources ?? [])
    .map((source) => {
      const resolved = shaping.fonts.resolve(source.request);
      return resolved instanceof FontResolutionError || resolved.id !== source.id
        ? null
        : {
            request: source.request,
            id: resolved.id,
            bytes: resolved.bytes,
            hash: resolved.hash,
            faceIndex: resolved.faceIndex,
          };
    })
    .filter((source): source is NonNullable<typeof source> => source !== null);
}

/**
 * The configuration's answer for one document: `raw` as the host gave it, `explicit` with a
 * resolver's answer normalized.
 *
 * On-demand configurations resolve per load rather than at construction, which is the whole
 * point of the function form: the document is parsed by now, so the resolver is told what
 * the file actually asks for and can skip everything else. A later pass of the same load
 * composes with the earlier answer, so only new families cost anything.
 */
export async function resolveConfiguredFonts(
  configured: DocxEditorConfig['fonts'],
  previous: FontConfigurationBase | undefined,
  request: { readonly families: readonly string[]; readonly defaultFamily: string },
  report: (error: EditorFontError) => void
): Promise<{
  raw: FontConfigurationBase | undefined;
  explicit: FontConfigurationBase | undefined;
}> {
  if (typeof configured !== 'function') return { raw: configured, explicit: configured };
  const raw = previous
    ? await composeFontOrigins(
        [previous, defineFontResolver((next) => configured(next))],
        request,
        { onOriginFailure: ({ cause }) => report(toEditorFontError(cause)) }
      )
    : await configured(request);
  return { raw, explicit: normalizeFontResolverResult(raw) };
}
