import { useCallback, useRef } from 'react';

/** Undo/redo for immutable documents held by a parent. */
export function useHistory<T>(value: T, onChange: (v: T) => void, limit = 100) {
  const past = useRef<T[]>([]);
  const future = useRef<T[]>([]);
  const cur = useRef(value);
  cur.current = value;

  const commit = useCallback(
    (next: T) => {
      past.current.push(cur.current);
      if (past.current.length > limit) past.current.shift();
      future.current = [];
      onChange(next);
    },
    [onChange, limit],
  );
  const undo = useCallback(() => {
    const p = past.current.pop();
    if (p === undefined) return;
    future.current.push(cur.current);
    onChange(p);
  }, [onChange]);
  const redo = useCallback(() => {
    const f = future.current.pop();
    if (f === undefined) return;
    past.current.push(cur.current);
    onChange(f);
  }, [onChange]);
  return { commit, undo, redo, canUndo: () => past.current.length > 0, canRedo: () => future.current.length > 0 };
}
