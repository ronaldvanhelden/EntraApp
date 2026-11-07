import { useMemo, useState } from 'react';
import {
  TAG_PREFIX,
  addTag,
  hasTag,
  isOurTag,
  removeTag,
  sortTags,
  tagLabel,
  toStoredTag,
} from '../lib/tags';
import { TagChip } from './TagChip';

export interface TagSuggestion {
  tag: string;
  count: number;
}

interface Props {
  tags: string[] | undefined;
  // Persists the full replacement collection — Graph PATCHes tags wholesale.
  onSave: (next: string[]) => Promise<void>;
  disabled?: boolean;
  // Tags already in use elsewhere in the tenant (debble: namespace only),
  // most used first. Offered as one-click chips while adding.
  suggestions?: TagSuggestion[];
}

const MAX_SUGGESTIONS = 8;

export function TagEditor({ tags, onSave, disabled, suggestions }: Props) {
  const [adding, setAdding] = useState(false);
  const [value, setValue] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const commit = async (next: string[]) => {
    setBusy(true);
    setError(null);
    try {
      await onSave(next);
      setValue('');
      setAdding(false);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const add = (input = value) => {
    if (!toStoredTag(input)) {
      setAdding(false);
      setValue('');
      return;
    }
    void commit(addTag(tags, input));
  };

  // Suggestions not already on this app, narrowed by what's been typed.
  // Prefix matches rank above substring matches; ties keep usage order.
  const offered = useMemo(() => {
    if (!suggestions?.length) return [];
    const q = tagLabel(toStoredTag(value)).toLowerCase();
    const scored = suggestions
      .filter((s) => isOurTag(s.tag) && !hasTag(tags, s.tag))
      .map((s) => {
        const label = tagLabel(s.tag).toLowerCase();
        const rank = !q ? 0 : label.startsWith(q) ? 0 : label.includes(q) ? 1 : -1;
        return { ...s, rank, exact: label === q };
      })
      .filter((s) => s.rank >= 0 && !s.exact);
    scored.sort((a, b) => a.rank - b.rank || b.count - a.count);
    return scored.slice(0, MAX_SUGGESTIONS);
  }, [suggestions, tags, value]);

  const current = sortTags(tags);
  const locked = disabled || busy;

  return (
    <div className="stack" style={{ gap: 6 }}>
      <div className="row wrap" style={{ gap: 6 }}>
        {current.map((t) => (
          <TagChip
            key={t}
            tag={t}
            disabled={locked}
            // Foreign tags are read-only: a PATCH carries them through, but the
            // chip offers no ×.
            onRemove={isOurTag(t) ? () => commit(removeTag(tags, t)) : undefined}
          />
        ))}
        {current.length === 0 && !adding && <span className="muted">No tags</span>}
        {adding ? (
          <span className="row" style={{ gap: 4 }}>
            <span className="mono muted" style={{ fontSize: 12 }}>
              {TAG_PREFIX}
            </span>
            <input
              autoFocus
              value={value}
              disabled={locked}
              placeholder="Production"
              style={{ width: 150 }}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') add();
                if (e.key === 'Escape') {
                  setAdding(false);
                  setValue('');
                }
              }}
            />
            <button disabled={locked} onClick={() => add()}>
              {busy ? 'Saving…' : 'Add'}
            </button>
            <button
              className="ghost"
              disabled={locked}
              onClick={() => {
                setAdding(false);
                setValue('');
              }}
            >
              Cancel
            </button>
          </span>
        ) : (
          <button className="ghost" disabled={locked} onClick={() => setAdding(true)}>
            + Add tag
          </button>
        )}
      </div>
      {adding && offered.length > 0 && (
        <div className="row wrap tag-suggestions" style={{ gap: 6 }}>
          <span className="muted" style={{ fontSize: 12 }}>
            In use:
          </span>
          {offered.map((s) => (
            <button
              key={s.tag}
              type="button"
              className="ghost tag-suggestion"
              disabled={locked}
              title={`${s.tag} — used on ${s.count} ${s.count === 1 ? 'app' : 'apps'}`}
              onClick={() => add(s.tag)}
            >
              {tagLabel(s.tag)} <span className="muted">({s.count})</span>
            </button>
          ))}
        </div>
      )}
      {error && (
        <p className="error" style={{ margin: 0 }}>
          {error}
        </p>
      )}
    </div>
  );
}
