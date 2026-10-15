import { expect, test } from 'bun:test';
import { createLiveFontResolution } from '../live-font-resolution.ts';

const tick = () => new Promise((resolve) => setTimeout(resolve, 0));

test('static configurations initialize without scanning document font needs on load or edits', async () => {
  let scans = 0;
  let resolutions = 0;
  const queue = createLiveFontResolution(
    () => ({
      generation: 1,
      dynamic: false,
      families: () => {
        scans++;
        return ['Calibri'];
      },
    }),
    async () => {
      resolutions++;
    }
  );
  queue.schedule(true);
  await tick();
  queue.schedule();
  await tick();
  expect(resolutions).toBe(1);
  expect(scans).toBe(0);
});

test('a resolution started before the mount is the initial one, not a second', async () => {
  // The facade resolves a deferred document's fonts before its surface mounts. The mount then
  // schedules the initial pass as always, and that must wait for the adopted work rather than
  // start the same resolution again.
  const requests: (readonly string[])[] = [];
  let families = ['Calibri', 'Cambria'];
  const queue = createLiveFontResolution(
    () => ({ generation: 7, dynamic: true, families: () => families }),
    async (asked) => {
      requests.push(asked);
    }
  );
  let finish!: () => void;
  queue.adopt(7, ['Calibri', 'Cambria'], new Promise<void>((resolve) => (finish = resolve)));

  queue.schedule(true);
  await tick();
  expect(requests).toEqual([]);

  // Settling re-reads the document: only a family the adopted pass did not ask for is asked.
  families = ['Calibri', 'Cambria', 'Courier New'];
  finish();
  await tick();
  expect(requests).toEqual([['Courier New']]);
});

test('an adopted resolution that fails still frees the queue', async () => {
  const requests: (readonly string[])[] = [];
  let families = ['Calibri'];
  const queue = createLiveFontResolution(
    () => ({ generation: 1, dynamic: true, families: () => families }),
    async (asked) => {
      requests.push(asked);
    }
  );
  queue.adopt(1, ['Calibri'], Promise.reject(new Error('origin down')));
  families = ['Calibri', 'Arial'];
  await tick();
  queue.schedule();
  await tick();
  expect(requests).toEqual([['Arial']]);
});
