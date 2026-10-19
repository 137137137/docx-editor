// Pagination slices give way to input: none starts while the browser holds a key or a click,
// a slice that meets input ends at its next pause point, and the slices resume once the input
// has settled rather than in the next task.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { createOpenScheduler, PAGINATE_INPUT_SETTLE_MS } from '../docx-editor-open-scheduler.ts';

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function until(check: () => boolean, timeoutMs = 5000): Promise<void> {
  const start = Date.now();
  while (!check()) {
    if (Date.now() - start > timeoutMs) throw new Error('condition not reached in time');
    await wait(1);
  }
}

describe('pagination under input', () => {
  test('no slice starts while input waits, and the slices resume once it has settled', async () => {
    let input = true;
    let slices = 0;
    const events: string[] = [];
    const scheduler = createOpenScheduler({
      mount: () => events.push('mount'),
      scheduled: () => events.push('scheduled'),
      continueOpen: () => {
        slices += 1;
        return slices >= 3;
      },
      paginated: () => events.push('paginated'),
      inputPending: () => input,
    });
    scheduler.schedule(new Uint8Array(1));
    await until(() => events.includes('mount'));
    expect(scheduler.isScheduled()).toBe(false);
    // Several settle periods pass with input waiting: not one slice runs.
    await wait(PAGINATE_INPUT_SETTLE_MS * 3);
    expect(slices).toBe(0);
    expect(events).not.toContain('paginated');

    input = false;
    await until(() => events.includes('paginated'));
    expect(slices).toBe(3);
    expect(events).toEqual(['scheduled', 'mount', 'scheduled', 'paginated']);
  });

  test('a slice that stopped for input resumes after the settle period, not in the next task', async () => {
    let input = false;
    let slices = 0;
    let paginated = 0;
    const scheduler = createOpenScheduler({
      mount: () => undefined,
      scheduled: () => undefined,
      continueOpen: () => {
        slices += 1;
        // The first slice meets input part-way and gives the rest of its budget up.
        input = slices === 1;
        return slices >= 2;
      },
      paginated: () => {
        paginated += 1;
      },
      inputPending: () => input,
    });
    scheduler.schedule(new Uint8Array(1));
    await until(() => slices === 1);
    // The next tasks come and go; the second slice waits for the input to settle.
    await wait(5);
    expect(slices).toBe(1);
    expect(paginated).toBe(0);

    input = false;
    await until(() => paginated === 1);
    expect(slices).toBe(2);
  });
});
