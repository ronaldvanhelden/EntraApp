import type { ReactNode } from 'react';
import { useSearchParams } from 'react-router-dom';

// Click-to-sort for the list tables. A header cycles ascending → descending →
// off (back to the order the list loaded in). The choice lives in the URL as
// ?sort=key or ?sort=-key, so it survives going into a row and coming back.
// Sorting is done on the rows that are loaded — a list that shows only its
// first page sorts that page.

export type SortValue = string | number | boolean | null | undefined;
export type SortAccessors<T> = Record<string, (row: T) => SortValue>;

export interface SortState {
  key: string | null;
  desc: boolean;
  toggle: (key: string) => void;
}

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
const isEmpty = (v: SortValue) => v === null || v === undefined || v === '';

function compare(a: SortValue, b: SortValue): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return collator.compare(String(a), String(b));
}

export function useSort<T>(rows: T[], accessors: SortAccessors<T>) {
  const [params, setParams] = useSearchParams();
  const raw = params.get('sort') ?? '';
  const desc = raw.startsWith('-');
  const name = desc ? raw.slice(1) : raw;
  const key = name in accessors ? name : null;

  const toggle = (next: string) =>
    setParams(
      (prev) => {
        const p = new URLSearchParams(prev);
        if (key !== next) p.set('sort', next);
        else if (!desc) p.set('sort', `-${next}`);
        else p.delete('sort');
        return p;
      },
      { replace: true },
    );

  let sorted = rows;
  if (key) {
    const get = accessors[key];
    // Rows without a value stay at the bottom in either direction.
    sorted = [...rows].sort((x, y) => {
      const a = get(x);
      const b = get(y);
      if (isEmpty(a) || isEmpty(b)) return Number(isEmpty(a)) - Number(isEmpty(b));
      return desc ? compare(b, a) : compare(a, b);
    });
  }
  return { sorted, sort: { key, desc, toggle } as SortState };
}

export function SortTh({
  id,
  sort,
  className,
  children,
}: {
  id: string;
  sort: SortState;
  className?: string;
  children: ReactNode;
}) {
  const active = sort.key === id;
  return (
    <th
      className={className}
      aria-sort={active ? (sort.desc ? 'descending' : 'ascending') : undefined}
    >
      <button
        type="button"
        className={`th-sort${active ? ' active' : ''}`}
        title={
          !active ? 'Sort ascending' : sort.desc ? 'Back to the default order' : 'Sort descending'
        }
        onClick={() => sort.toggle(id)}
      >
        {children}
        <span className="th-sort-arrow" aria-hidden="true">
          {active ? (sort.desc ? '▼' : '▲') : '↕'}
        </span>
      </button>
    </th>
  );
}
