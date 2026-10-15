import type { DocumentChange, EditorError, EditorSnapshot } from './editor.ts';
import type { HistoryDiagnostic } from './editor-scope.ts';

/**
 * What `editor.on(...)` can be subscribed to, and what each handler receives.
 *
 * These are PUSH notifications and are not interchangeable with reading `snapshot()`: a snapshot
 * read cannot observe an event that was never emitted, which is why adapter behaviour is asserted
 * against these rather than against the snapshot.
 */
export interface EditorEvents {
  /** A document mutation committed, with the ids it touched. */
  change: (change: DocumentChange) => void;
  /** The selection or its derived formatting moved. */
  selectionChange: (snapshot: EditorSnapshot) => void;
  error: (error: EditorError) => void;
  historyDiagnostic: (diagnostic: HistoryDiagnostic) => void;
  /**
   * Every page of the mounted document is laid out. A large document mounts on a preview of
   * its first pages and fires this when the rest is laid out; any other document fires it
   * just after it mounts. See {@link EditorSnapshot.isPaginating}.
   */
  paginated: () => void;
}
