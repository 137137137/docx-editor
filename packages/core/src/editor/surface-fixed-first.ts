// The fixed measurer first, when the fonts were resolved before the mount.
//
// An open whose fonts resolve after its mount lays the document out twice: on the fixed
// measurer, then on the shaped one, on the same layout session. The second pass is not a cold
// one. The footnote reserve search starts from the reserves the first pass settled on
// (`notePageBottomReserves`), and a document with more than one settled answer keeps the one
// nearest that start; laid out cold on the shaped measurer, such a document breaks its pages
// elsewhere. So an open that already has its fonts still runs the fixed pass first, on a
// measurer resolved exactly as an unshaped mount resolves it (no font alias: nothing is
// registered before the first mount of a document), and lays the shaped pass over that
// session, so the pages are the ones the two-pass open produces. The fixed pass is never
// painted, and during a sliced opening it runs in the slices with the rest.

import { resolveDefaultSurfaceMeasurer, type TextMeasurer } from '@docx-editor.dev/core/layout';
import { tryCreateBrowserCanvasContext } from './browser-canvas-context.ts';

/** A measurer and the producer identity its layouts are cached under. */
export interface LayoutMeasurement {
  readonly measurer: TextMeasurer;
  readonly producer: string;
}

/** When a surface's layouts run the fixed pass first, and on which measurer. */
export interface FixedFirst {
  /** Whether a layout now runs the fixed pass before the one it publishes. */
  runs(): boolean;
  /** The fixed measurer at `scale`, as a mount with no measurer resolves it. */
  measurement(scale: number): LayoutMeasurement;
}

export interface FixedFirstInput {
  /** The facade resolved the fonts before this mount, so it mounted on the shaped measurer. */
  readonly fontsFirst: boolean;
  /** The surface mounted on its fixed measurer; a measurer that arrives during its opening
   *  makes the opening's layouts the two passes too. */
  readonly mountedFixed: boolean;
  /** Whether the surface measures with a host (shaped) measurer now. */
  hostMeasured(): boolean;
  readonly document: Document;
}

export function createFixedFirst(input: FixedFirstInput): FixedFirst {
  let resolved: { readonly scale: number; readonly measurement: LayoutMeasurement } | null = null;
  return {
    runs: () => (input.fontsFirst || input.mountedFixed) && input.hostMeasured(),
    measurement(scale) {
      if (resolved?.scale !== scale) {
        const fixed = resolveDefaultSurfaceMeasurer(scale, {
          context: tryCreateBrowserCanvasContext(input.document),
        });
        resolved = {
          scale,
          measurement: {
            measurer: fixed.measurer,
            producer: `${fixed.producer ?? 'fixed-measurer'}@scale:${scale}`,
          },
        };
      }
      return resolved.measurement;
    },
  };
}
