import { errors, type Page } from '@playwright/test';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const PORT = 5275;
const FIXTURE = 'synthetic-long-edit.docx';
/**
 * The tough companion fixture: ~170+ pages of NUMBERED clauses with ~310 dense tracked
 * replacements — the paths (list cascade, tracked-run structure, review cards) the plain
 * fixture never exercises.
 */
export const EDIT_BROWSER_TRACKED_FIXTURE = 'synthetic-tracked-numbered.docx';
/**
 * The stress fixture: ~1,000 pages, ~1,060 tracked changes. Deterministic but
 * uncommitted — regenerate with `bun scripts/create-synthetic-tracked-fixture.mjs --huge`.
 */
export const EDIT_BROWSER_HUGE_FIXTURE = 'synthetic-huge-tracked.docx';
/**
 * The pinned 521-page reported reproduction (SHA-256 in
 * `e2e/fixtures/typing-perf-521pp.manifest.json`) — 12,820 paragraphs, 81 sections, tables,
 * drawings, notes and TOCs. The synthetic fixtures above are sized for deterministic gates;
 * this is the shape where per-keystroke costs that scale with the document actually show,
 * and its rows lead the typing-latency table.
 */
export const EDIT_BROWSER_PINNED_HUGE_FIXTURE = 'typing-perf-521pp.docx';
export const REVIEW_RAIL_ENABLED = process.env.EDIT_BROWSER_BENCH_REVIEW_RAIL !== '0';

export function editBrowserBenchUrl(fixture: string = FIXTURE): string {
  return `http://localhost:${PORT}/?perfE2e=1&fixture=${fixture}&reviewRail=${
    REVIEW_RAIL_ENABLED ? '1' : '0'
  }`;
}

export const EDIT_BROWSER_BENCH_URL = editBrowserBenchUrl();

export interface EnginePerf {
  readonly layoutMs: number;
  readonly paintMs: number;
  readonly selectionMs: number;
  readonly placed: number;
  readonly total: number;
  readonly reusedPages: number;
  readonly fullPasses: number;
  readonly staleDiscards: number;
  readonly cancelledRuns: number;
}

export interface BrowserSample {
  readonly inputTaskMs: number;
  readonly frameMs: number;
  readonly eventDurationMs: number | null;
  readonly eventDelayMs: number | null;
  readonly engine: EnginePerf;
  readonly domNodes: number;
  readonly materializedPages: number;
  readonly selectionSpans: number;
}

export interface TimingSummary {
  readonly medianMs: number;
  readonly p95Ms: number;
  readonly minMs: number;
  readonly maxMs: number;
}

export function positiveInteger(value: string | undefined, fallback: number): number {
  if (value === undefined) return fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1)
    throw new Error(`expected positive integer: ${value}`);
  return parsed;
}

export function nonNegativeNumber(value: string | undefined, fallback: number): number {
  if (value === undefined) return fallback;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0)
    throw new Error(`expected non-negative number: ${value}`);
  return parsed;
}

export function summarize(values: readonly number[]): TimingSummary {
  const sorted = [...values].sort((a, b) => a - b);
  return {
    medianMs: sorted[Math.floor(sorted.length / 2)]!,
    p95Ms: sorted[Math.max(0, Math.ceil(sorted.length * 0.95) - 1)]!,
    minMs: sorted[0]!,
    maxMs: sorted[sorted.length - 1]!,
  };
}

export function summarizeOptional(values: readonly (number | null)[]): TimingSummary | null {
  const present = values.filter((value): value is number => value !== null);
  return present.length > 0 ? summarize(present) : null;
}

export async function twoFrames(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
      })
  );
}

export async function heapBytes(page: Page): Promise<number | null> {
  return page.evaluate(() => {
    const collect = (globalThis as typeof globalThis & { gc?: () => void }).gc;
    collect?.();
    const memory = (
      performance as Performance & {
        memory?: { readonly usedJSHeapSize: number };
      }
    ).memory;
    return memory?.usedJSHeapSize ?? null;
  });
}

export function percentChange(next: number, before: number): number {
  return before === 0 ? 0 : ((next - before) / before) * 100;
}

export async function delay(ms: number): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, ms));
}

/** How long an open may take to finish its complete layout before the harness gives up. */
const LAYOUT_TIMEOUT_MS = 60_000;

export async function loadHarness(
  page: Page,
  installMeasurementProbe: (page: Page) => Promise<void>,
  measurementProbe = true,
  fixture: string = FIXTURE
): Promise<void> {
  await page.goto(editBrowserBenchUrl(fixture), { waitUntil: 'domcontentloaded' });
  await page.waitForFunction(() => !!window.__DOCX_EDITOR_E2E__?.ready());
  await page.waitForFunction(() => window.__DOCX_EDITOR_E2E__?.fontMeasurer() === 'shaped');
  // A materialized first page can precede the remaining pagination during progressive open.
  // Keep the existing preparation budget, and start measurement after the complete layout.
  try {
    await page.waitForFunction(
      () => {
        const materialized = document.querySelector('.docx-page[data-materialized="true"]');
        if (!materialized) return false;
        const bounds = materialized.getBoundingClientRect();
        return (
          getComputedStyle(materialized).visibility === 'visible' &&
          bounds.width > 0 &&
          bounds.height > 0 &&
          !document.querySelector('[data-docx-opening-preview], [data-docx-layout-pending]')
        );
      },
      undefined,
      { timeout: LAYOUT_TIMEOUT_MS }
    );
  } catch (error) {
    if (!(error instanceof errors.TimeoutError)) throw error;
    // Say what was still pending, so a timeout is not mistaken for a slow benchmark.
    // A blocked page must not hold the report until the test timeout.
    const state = await Promise.race([
      page.evaluate(() => ({
        pages: document.querySelectorAll('.docx-page').length,
        materializedPages: document.querySelectorAll('.docx-page[data-materialized="true"]').length,
        openingPreview: !!document.querySelector('[data-docx-opening-preview]'),
        layoutPending: !!document.querySelector('[data-docx-layout-pending]'),
      })),
      delay(5_000).then(() => null),
    ]).catch(() => null);
    throw new Error(
      `${fixture} did not finish layout in ${LAYOUT_TIMEOUT_MS / 1000} s: ${JSON.stringify(state)}`,
      {
        cause: error,
      }
    );
  }
  if (measurementProbe) await installMeasurementProbe(page);
}

export const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const EDIT_BROWSER_FIXTURE = FIXTURE;
