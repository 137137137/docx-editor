// An open that prepares before it mounts.
//
// The facade prepares a deferred open by resolving the document's fonts, so the first layout is
// the shaped one. What these tests pin down is the scheduler's side of that: the mount waits for
// the preparation and runs exactly once, whichever of the preparation, its time limit, a
// `flush` or a `cancel` comes first; a failed preparation still mounts; a preparation with
// nothing to wait for mounts at once; and on the yielding path the preparation itself starts
// only after the painted frame, so the loading screen is up before the parse.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { createOpenScheduler, OPEN_PREPARE_WAIT_MS } from '../docx-editor-open-scheduler.ts';

const BYTES = new Uint8Array([1, 2, 3]);

function harness(prepareWaitMs = 40) {
  const mounts: Uint8Array[] = [];
  let scheduledCalls = 0;
  const scheduler = createOpenScheduler({
    mount: (bytes) => mounts.push(bytes),
    scheduled: () => {
      scheduledCalls += 1;
    },
    prepareWaitMs,
  });
  return { scheduler, mounts, scheduledCalls: () => scheduledCalls };
}

function deferred(): { promise: Promise<void>; resolve: () => void; reject: () => void } {
  let resolve!: () => void;
  let reject!: () => void;
  const promise = new Promise<void>((res, rej) => {
    resolve = res;
    reject = () => rej(new Error('preparation failed'));
  });
  return { promise, resolve, reject };
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

describe('an open with preparation', () => {
  test('waits for the preparation, then mounts once', async () => {
    const { scheduler, mounts, scheduledCalls } = harness(1000);
    const work = deferred();
    scheduler.schedule(BYTES, { yieldFrame: false, prepare: () => work.promise });

    expect(mounts).toHaveLength(0);
    expect(scheduler.isScheduled()).toBe(true);
    expect(scheduledCalls()).toBe(1);

    work.resolve();
    await sleep(0);
    expect(mounts).toEqual([BYTES]);
    expect(scheduler.isScheduled()).toBe(false);
  });

  test('mounts without the preparation once the wait runs out, and never twice', async () => {
    const { scheduler, mounts } = harness(30);
    const work = deferred();
    scheduler.schedule(BYTES, { yieldFrame: false, prepare: () => work.promise });

    await sleep(60);
    expect(mounts).toHaveLength(1);
    expect(scheduler.isScheduled()).toBe(false);

    work.resolve();
    await sleep(0);
    expect(mounts).toHaveLength(1);
  });

  test('a failed preparation still mounts', async () => {
    const { scheduler, mounts } = harness(1000);
    const work = deferred();
    scheduler.schedule(BYTES, { yieldFrame: false, prepare: () => work.promise });
    work.reject();
    await sleep(0);
    expect(mounts).toHaveLength(1);
  });

  test('nothing to prepare mounts before `schedule` returns, and reports no window', () => {
    const { scheduler, mounts, scheduledCalls } = harness();
    scheduler.schedule(BYTES, { yieldFrame: false, prepare: () => null });
    expect(mounts).toEqual([BYTES]);
    expect(scheduler.isScheduled()).toBe(false);
    expect(scheduledCalls()).toBe(0);
  });

  test('`flush` mounts now without waiting, and the late preparation adds nothing', async () => {
    const { scheduler, mounts } = harness(1000);
    const work = deferred();
    scheduler.schedule(BYTES, { yieldFrame: false, prepare: () => work.promise });

    scheduler.flush();
    expect(mounts).toHaveLength(1);
    expect(scheduler.isScheduled()).toBe(false);

    work.resolve();
    await sleep(0);
    expect(mounts).toHaveLength(1);
  });

  test('`cancel` hands the bytes back and nothing mounts', async () => {
    const { scheduler, mounts } = harness(20);
    const work = deferred();
    scheduler.schedule(BYTES, { yieldFrame: false, prepare: () => work.promise });

    expect(scheduler.cancel()).toBe(BYTES);
    work.resolve();
    await sleep(40);
    expect(mounts).toHaveLength(0);
  });

  test('on the yielding path the preparation starts after the painted frame', async () => {
    const { scheduler, mounts } = harness(1000);
    let prepared = 0;
    scheduler.schedule(BYTES, {
      prepare: () => {
        prepared += 1;
        return null;
      },
    });
    expect(prepared).toBe(0);
    expect(scheduler.isScheduled()).toBe(true);

    for (let i = 0; i < 50 && mounts.length === 0; i++) await sleep(10);
    expect(prepared).toBe(1);
    expect(mounts).toHaveLength(1);
  });

  test('waits long enough for fonts from memory, and not so long a stalled origin hides the page', () => {
    expect(OPEN_PREPARE_WAIT_MS).toBeGreaterThanOrEqual(500);
    expect(OPEN_PREPARE_WAIT_MS).toBeLessThanOrEqual(2000);
  });
});
