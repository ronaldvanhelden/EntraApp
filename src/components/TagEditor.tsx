import { useState } from 'react';
import {
  TAG_PREFIX,
  addTag,
  isOurTag,
  removeTag,
  sortTags,
  toStoredTag,
} from '../lib/tags';
import { TagChip } from './TagChip';

interface Props {
  tags: string[] | undefined;
  // Persists the full replacement collection — Graph PATCHes tags wholesale.
  onSave: (next: string[]) => Promise<void>;
  disabled?: boolean;
}

export function TagEditor({ tags, onSave, disabled }: Props) {
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

  const add = () => {
    if (!toStoredTag(value)) {
      setAdding(false);
      setValue('');
      return;
    }
    commit(addTag(tags, value));
  };

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
        {current.length === 0 && !adding && (
          <span className="muted">No tags</span>
        )}
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
            <button disabled={locked} onClick={add}>
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
          <button
            className="ghost"
            disabled={locked}
            onClick={() => setAdding(true)}
          >
            + Add tag
          </button>
        )}
      </div>
      {error && (
        <p className="error" style={{ margin: 0 }}>
          {error}
        </p>
      )}
    </div>
  );
}
