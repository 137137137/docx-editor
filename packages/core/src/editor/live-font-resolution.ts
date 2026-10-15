/** Serialize font updates within a document. A new document can supersede a pending fetch. */
export function createLiveFontResolution(
  read: () => { generation: number; families: () => readonly string[]; dynamic: boolean } | null,
  resolve: (families: readonly string[]) => Promise<void>
): {
  schedule(initial?: boolean): void;
  /**
   * Record a document's first resolution, started before its surface mounted, as this
   * generation's initial one. The mount's own `schedule(true)` then waits for it instead of
   * starting a second, and later passes ask only for families it did not.
   */
  adopt(generation: number, families: readonly string[], running: Promise<void>): void;
} {
  let current: { generation: number; requested: Set<string>; running: boolean } | undefined;
  let queued = false;
  const pump = (): void => {
    queued = false;
    const state = read();
    if (!state) return;
    const initial = current?.generation !== state.generation;
    if (initial) current = { generation: state.generation, requested: new Set(), running: false };
    const active = current!;
    if (active.running || (!initial && !state.dynamic)) return;
    const families = state.dynamic
      ? state.families().filter((family) => !active.requested.has(family.toLowerCase()))
      : [];
    if (!initial && families.length === 0) return;
    for (const family of families) active.requested.add(family.toLowerCase());
    active.running = true;
    // Failed families stay attempted until the next document load. A failing provider
    // must not retry on every keystroke. resolve owns error reporting and degradation.
    void resolve(families).finally(() => {
      active.running = false;
      if (current === active) schedule();
    });
  };
  const schedule = (initial = false): void => {
    if (initial) {
      pump();
      return;
    }
    if (queued) return;
    queued = true;
    queueMicrotask(pump);
  };
  const adopt = (generation: number, families: readonly string[], running: Promise<void>): void => {
    const active = {
      generation,
      requested: new Set(families.map((family) => family.toLowerCase())),
      running: true,
    };
    current = active;
    const settle = (): void => {
      active.running = false;
      if (current === active) schedule();
    };
    void running.then(settle, settle);
  };
  return { schedule, adopt };
}
