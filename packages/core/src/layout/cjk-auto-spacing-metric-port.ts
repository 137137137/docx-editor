import type { ResolvedRunStyle } from './run-style.ts';
import type { TextMeasurer } from './text-measurer.ts';

type AuthoredGapMetric = (style: ResolvedRunStyle) => number | undefined;
const metrics = new WeakMap<TextMeasurer, AuthoredGapMetric>();

// Session-local derived metrics. This port is not part of the public measurer contract.
export function withAuthoredCjkGapMetric<T extends TextMeasurer>(
  measurer: T,
  metric: AuthoredGapMetric
): T {
  metrics.set(measurer, metric);
  return measurer;
}

export function authoredCjkGapMetric(
  measurer: TextMeasurer | undefined,
  style: ResolvedRunStyle
): number | undefined {
  return measurer ? metrics.get(measurer)?.(style) : undefined;
}
