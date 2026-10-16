// A table cell paragraph's resolved layout and break inputs, read once per context.
//
// A table is placed more than once in one layout (its rows are measured before they are
// placed), and again in every footnote-reserve round that reaches it. The inputs are a
// function of the paragraph node, which is immutable, and of the width, style table, list
// item and table style it is resolved against, so they are kept per node for the few contexts
// a paragraph meets. Nothing reads them but to lay the paragraph out; nothing writes them.

import type { OoxmlElement, OoxmlProperty } from '@docx-editor.dev/core/store';
import type { ResolvedListItem } from './list-resolve.ts';
import {
  prepareParagraphBreakInputs,
  type ParagraphBreakDependencies,
} from './paragraph-break-request.ts';
import type { ResolvedTabStops } from './paragraph-tabs.ts';
import {
  resolveParagraphLayoutInputs,
  type ParagraphLayoutInputs,
  type StyleCascadeTable,
  type TableCellStyleFormatting,
} from './style-cascade.ts';

interface CellInputsEntry {
  readonly width: number;
  readonly styleCascade: StyleCascadeTable | undefined;
  readonly listItem: ResolvedListItem | undefined;
  readonly tableCellStyle: TableCellStyleFormatting | undefined;
  readonly inputs: ParagraphLayoutInputs;
}

/** Contexts kept per paragraph: a cell's own width, and the odd probe at another. */
const MAX_CONTEXTS = 4;

const byParagraph = new WeakMap<OoxmlElement, CellInputsEntry[]>();

/** `resolveParagraphLayoutInputs` for a paragraph in a table cell. */
export function cellParagraphLayoutInputs(
  paragraph: OoxmlElement,
  width: number,
  styleCascade: StyleCascadeTable | undefined,
  listItem: ResolvedListItem | undefined,
  tableCellStyle: TableCellStyleFormatting | undefined
): ParagraphLayoutInputs {
  let entries = byParagraph.get(paragraph);
  const known = entries?.find(
    (entry) =>
      entry.width === width &&
      entry.styleCascade === styleCascade &&
      entry.listItem === listItem &&
      entry.tableCellStyle === tableCellStyle
  );
  if (known) return known.inputs;
  const inputs = resolveParagraphLayoutInputs(
    paragraph,
    width,
    styleCascade,
    listItem,
    tableCellStyle,
    true
  );
  if (!entries) {
    entries = [];
    byParagraph.set(paragraph, entries);
  }
  entries.push({ width, styleCascade, listItem, tableCellStyle, inputs });
  if (entries.length > MAX_CONTEXTS) entries.shift();
  return inputs;
}

interface CellBreakEntry {
  readonly defaultTabStopPt: number | undefined;
  readonly dependencies: ParagraphBreakDependencies;
  readonly tabStops: ResolvedTabStops;
  readonly properties: readonly OoxmlProperty[];
}

const breakInputsByLayoutInputs = new WeakMap<ParagraphLayoutInputs, CellBreakEntry[]>();

/**
 * `prepareParagraphBreakInputs` for inputs from {@link cellParagraphLayoutInputs}. The
 * property list comes back frozen and the same for the same context, so the break key reads
 * its token once (`propertiesToken`).
 */
export function cellParagraphBreakInputs(
  inputs: ParagraphLayoutInputs,
  defaultTabStopPt: number | undefined,
  dependencies: ParagraphBreakDependencies
): { readonly tabStops: ResolvedTabStops; readonly properties: readonly OoxmlProperty[] } {
  let entries = breakInputsByLayoutInputs.get(inputs);
  const known = entries?.find(
    (entry) =>
      entry.defaultTabStopPt === defaultTabStopPt &&
      entry.dependencies.listToken === dependencies.listToken &&
      entry.dependencies.hostedListToken === dependencies.hostedListToken &&
      entry.dependencies.refToken === dependencies.refToken
  );
  if (known) return known;
  const prepared = prepareParagraphBreakInputs(inputs, defaultTabStopPt, dependencies);
  const entry: CellBreakEntry = {
    defaultTabStopPt,
    dependencies: { ...dependencies },
    tabStops: prepared.tabStops,
    properties: Object.freeze([...prepared.properties]),
  };
  if (!entries) {
    entries = [];
    breakInputsByLayoutInputs.set(inputs, entries);
  }
  entries.push(entry);
  if (entries.length > MAX_CONTEXTS) entries.shift();
  return entry;
}
