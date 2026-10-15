// The open-yield scheduler: one painted frame between "open this document" and the
// blocking mount, so a host's loading screen can actually show.
//
// Opening a document parses, lays out and paints in one synchronous pass — seconds of
// blocked main thread on a long file — and a mount that runs in the same task as
// `load()`/`attach()` blocks the very frame that would have painted the host's loading
// screen. Documents past {@link OPEN_PAINT_YIELD_CONTENT_BYTES} of CONTENT therefore
// schedule the mount behind one painted frame, and `snapshot().isOpening` is true for
// exactly that window. Small documents keep the synchronous path: they need no loading
// flash, and every existing synchronous caller stays as it was. So does every
// environment without `requestAnimationFrame` — headless and server hosts mount
// synchronously by definition.
//
// A scheduled open can also carry PREPARATION: work the mount should not start without,
// which today is resolving the document's fonts so that its first layout is already the
// shaped one. The mount waits for it, but only for {@link OPEN_PREPARE_WAIT_MS}: a slow
// font origin costs the reader one extra layout when it lands, never a document that
// does not open. A small document with nothing to prepare still mounts synchronously.

/**
 * The UNCOMPRESSED size past which an open earns a painted loading frame first.
 *
 * Mount cost tracks the content, and the zipped size lies about it: a 200-page
 * tracked-changes document is ~1.6 MB of XML but zips under 90 KiB, while WordprocessingML
 * routinely compresses 10–20×. So the threshold reads the true entry sizes from the ZIP
 * central directory (see {@link zipContentExceeds} — a bounded scan, no decompression).
 * Documents under this mount well inside a frame or two on current hardware; over it,
 * the synchronous mount visibly freezes the page the document was opened from.
 */
const OPEN_PAINT_YIELD_CONTENT_BYTES = 512 * 1024;

/**
 * Whether the ZIP's entries sum past `limit` uncompressed — WITHOUT inflating anything.
 * The central directory records every entry's uncompressed size; walking it costs
 * microseconds on any real document.
 *
 * Attacker-controlled input rules (the bytes come straight from a file): the entry
 * count is a bounded uint16, the walk advances monotonically and bounds-checks every
 * read, the sizes are only summed and compared — never fed to an allocation — and the
 * sum exits early at the limit, so a forged huge size simply means "defer", which the
 * real parser then judges. Anything malformed answers `false` and the open proceeds on
 * the synchronous path, where the parser reports the actual error.
 */
function zipContentExceeds(bytes: Uint8Array, limit: number): boolean {
  const length = bytes.byteLength;
  if (length < 22) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, length);
  // End-of-central-directory: signature PK\x05\x06, at most 64 KiB of trailing comment.
  const scanFloor = Math.max(0, length - (64 * 1024 + 22));
  let eocd = -1;
  for (let at = length - 22; at >= scanFloor; at -= 1) {
    if (view.getUint32(at, true) === 0x06054b50) {
      eocd = at;
      break;
    }
  }
  if (eocd < 0) return false;
  const entryCount = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  let total = 0;
  for (let entry = 0; entry < entryCount; entry += 1) {
    if (offset + 46 > length) return false;
    if (view.getUint32(offset, true) !== 0x02014b50) return false;
    // A ZIP64 sentinel (0xFFFFFFFF) sums as huge and exits below: correct — it IS huge.
    total += view.getUint32(offset + 24, true);
    if (total >= limit) return true;
    offset +=
      46 +
      view.getUint16(offset + 28, true) +
      view.getUint16(offset + 30, true) +
      view.getUint16(offset + 32, true);
  }
  return false;
}

/**
 * How long a mount waits for its preparation before it goes ahead without it.
 *
 * Long enough for fonts served from the host's own origin or memory, which settle in tens of
 * milliseconds. Short enough that a stalled origin still leaves a document on screen within
 * the time a reader waits for one, laid out on the fixed measurer and laid out again when the
 * fonts arrive, which is what every open did before.
 */
export const OPEN_PREPARE_WAIT_MS = 1500;

/** How one open is scheduled. */
export interface OpenScheduleOptions {
  /** Paint one frame before the work starts, so a loading screen can show. Defaults to true. */
  readonly yieldFrame?: boolean;
  /**
   * Work to finish before the mount, started after the yielded frame, or at once when there is
   * none. Returns null when there is nothing to wait for, and the mount then runs immediately.
   * A rejected promise counts as finished: preparation is an optimization and its failures
   * are reported by the work itself.
   */
  readonly prepare?: (bytes: Uint8Array) => Promise<unknown> | null;
}

/** What the facade hands the scheduler; both close over facade-owned state. */
export interface OpenSchedulerHooks {
  /** The real synchronous mount (`mountBytes`). */
  readonly mount: (bytes: Uint8Array) => void;
  /** Called once when a mount is scheduled — the facade bumps and emits here. */
  readonly scheduled: () => void;
  /** How long a mount waits for its preparation; {@link OPEN_PREPARE_WAIT_MS} by default. */
  readonly prepareWaitMs?: number;
}

/** The deferred-mount window, owned by the facade. See the module comment. */
export interface OpenScheduler {
  /** Whether this open is worth (and able to get) a painted frame before the mount. */
  shouldYield(bytes: Uint8Array): boolean;
  /**
   * Schedule the mount behind one painted frame and any preparation, and report the state
   * move. A mount that turns out to need neither runs before this returns.
   */
  schedule(bytes: Uint8Array, options?: OpenScheduleOptions): void;
  /** Cancel a scheduled open and hand its bytes back to the caller to re-route. */
  cancel(): Uint8Array | null;
  /**
   * Run a scheduled open NOW, without waiting for its frame or its preparation. For callers
   * that need the document synchronously — a `save()` or `exec` issued inside the yield
   * window must see the document that was just loaded, not a "no document is loaded"
   * refusal the next frame would disprove.
   */
  flush(): void;
  /** Whether a mount is currently waiting on its frame — `snapshot().isOpening`. */
  isScheduled(): boolean;
}

/** Headless and server hosts have no frame to yield, and mount without one. */
const canYieldFrame = (): boolean =>
  typeof requestAnimationFrame === 'function' && typeof cancelAnimationFrame === 'function';

export function createOpenScheduler(hooks: OpenSchedulerHooks): OpenScheduler {
  let scheduled: { readonly bytes: Uint8Array; cancel(): void } | null = null;

  const cancel = (): Uint8Array | null => {
    if (!scheduled) return null;
    const { bytes } = scheduled;
    scheduled.cancel();
    scheduled = null;
    return bytes;
  };

  return {
    // The zipped size is a shortcut, not the measure: a file already past the limit
    // zipped cannot be under it unpacked in any way that mounts fast.
    shouldYield: (bytes) =>
      canYieldFrame() &&
      (bytes.byteLength >= OPEN_PAINT_YIELD_CONTENT_BYTES ||
        zipContentExceeds(bytes, OPEN_PAINT_YIELD_CONTENT_BYTES)),

    // `requestAnimationFrame` fires BEFORE the pending paint, so the heavy work goes
    // into a task queued from inside it — the first slot guaranteed to run after the
    // loading screen is on screen. Until then the previous document (if any) stays
    // mounted under the host's overlay. A HIDDEN tab never fires rAF at all, so a
    // plain-timer fallback mounts anyway: a document opened in the background must be
    // there when the tab is next looked at, not still waiting for a frame.
    schedule(bytes, options = {}) {
      const { yieldFrame = true, prepare } = options;
      let live = true;
      let raf: number | null = null;
      const timers = new Set<ReturnType<typeof setTimeout>>();
      const later = (callback: () => void, ms: number): ReturnType<typeof setTimeout> => {
        const timer = setTimeout(() => {
          timers.delete(timer);
          callback();
        }, ms);
        timers.add(timer);
        return timer;
      };
      const stop = (): void => {
        live = false;
        if (raf !== null) cancelAnimationFrame(raf);
        raf = null;
        for (const timer of timers) clearTimeout(timer);
        timers.clear();
      };
      const entry = { bytes, cancel: stop };
      // Reached by the preparation settling, by its time limit, or by the frame when there is
      // no preparation, whichever comes first; `live` makes the others no-ops.
      const run = (): void => {
        if (!live) return;
        stop();
        if (scheduled === entry) scheduled = null;
        hooks.mount(bytes);
      };
      const begin = (): void => {
        if (!live) return;
        const pending = prepare?.(bytes) ?? null;
        if (!pending) {
          run();
          return;
        }
        later(run, hooks.prepareWaitMs ?? OPEN_PREPARE_WAIT_MS);
        pending.then(run, run);
      };
      scheduled = entry;
      if (yieldFrame && canYieldFrame()) {
        const fallback = later(() => {
          if (raf !== null) cancelAnimationFrame(raf);
          raf = null;
          begin();
        }, 250);
        raf = requestAnimationFrame(() => {
          raf = null;
          clearTimeout(fallback);
          timers.delete(fallback);
          later(begin, 0);
        });
      } else {
        begin();
      }
      if (scheduled === entry) hooks.scheduled();
    },

    cancel,

    flush() {
      const bytes = cancel();
      if (bytes) hooks.mount(bytes);
    },

    isScheduled: () => scheduled !== null,
  };
}
