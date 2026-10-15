// The body prefix a preview of the first pages lays out (`SemanticLayoutOptions.bodyBlockLimit`).
//
// Only the flow is cut. Sections are enumerated over the whole body first and then clipped,
// so a prefix that ends inside the first of several sections still lays out under that
// section's own properties rather than the properties of the body's last `w:sectPr`.

import type { OoxmlElement } from '@docx-editor.dev/core/store';
import type { DocumentSection } from './section-properties.ts';

/** The body blocks and sections to lay out, cut to the first `limit` blocks. */
export function limitBodyBlocks(
  blocks: readonly OoxmlElement[],
  sections: readonly DocumentSection[],
  limit: number | undefined
): { readonly blocks: readonly OoxmlElement[]; readonly sections: readonly DocumentSection[] } {
  if (limit === undefined || limit >= blocks.length) return { blocks, sections };
  const cut = Math.max(0, Math.floor(limit));
  const clipped: DocumentSection[] = [];
  for (const section of sections) {
    if (section.blockStart >= cut && clipped.length > 0) break;
    clipped.push(
      section.blockEndExclusive > cut ? { ...section, blockEndExclusive: cut } : section
    );
  }
  return { blocks: blocks.slice(0, cut), sections: clipped };
}
