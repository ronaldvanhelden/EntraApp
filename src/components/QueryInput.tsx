import { useMemo, useRef, useState } from 'react';
import {
  QUERY_HELP,
  queryProblem,
  suggest,
  type Entity,
  type Suggestion,
} from '../lib/searchQuery';

// A search box that understands the query language in lib/searchQuery.ts and
// completes it while typing: field names first, then the values of fields
// that have a fixed set. Nothing is picked until the arrow keys are used, so
// Enter and plain typing behave like an ordinary input; Tab takes the first
// suggestion.

interface Props {
  value: string;
  onChange: (value: string) => void;
  // Which kinds of object this box searches: decides the fields on offer.
  entities: Entity[];
  placeholder?: string;
  // Goes on the wrapper, which takes the input's place in the layout.
  className?: string;
  autoFocus?: boolean;
  ariaLabel?: string;
}

export function QueryInput({
  value,
  onChange,
  entities,
  placeholder,
  className,
  autoFocus,
  ariaLabel,
}: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [focused, setFocused] = useState(false);
  const [caret, setCaret] = useState(value.length);
  const [active, setActive] = useState(-1);
  // Escape hides the list until the text changes again.
  const [dismissed, setDismissed] = useState(false);

  const items = useMemo(
    () => (focused && !dismissed ? suggest(value, caret, entities) : []),
    [focused, dismissed, value, caret, entities],
  );
  const problem = useMemo(() => queryProblem(value, entities), [value, entities]);
  const index = Math.min(active, items.length - 1);

  const accept = (s: Suggestion) => {
    const next = value.slice(0, s.from) + s.insert + value.slice(s.to);
    const pos = s.from + s.insert.length;
    onChange(next);
    setCaret(pos);
    setActive(-1);
    // After React has put the new text in the input.
    requestAnimationFrame(() => {
      inputRef.current?.focus();
      inputRef.current?.setSelectionRange(pos, pos);
    });
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (items.length === 0) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      setActive(
        e.key === 'ArrowDown'
          ? (index + 1) % items.length
          : index <= 0
            ? items.length - 1
            : index - 1,
      );
    } else if (e.key === 'Tab' && !e.shiftKey) {
      e.preventDefault();
      accept(items[Math.max(index, 0)]);
    } else if (e.key === 'Enter' && index >= 0) {
      e.preventDefault();
      accept(items[index]);
    } else if (e.key === 'Escape') {
      e.stopPropagation();
      setDismissed(true);
    }
  };

  const syncCaret = () => setCaret(inputRef.current?.selectionStart ?? value.length);
  const open = items.length > 0;

  return (
    <div className={`qinput${className ? ` ${className}` : ''}`}>
      <input
        ref={inputRef}
        type="search"
        role="combobox"
        aria-expanded={open}
        aria-controls="qinput-list"
        aria-autocomplete="list"
        aria-activedescendant={open && index >= 0 ? `qinput-opt-${index}` : undefined}
        aria-label={ariaLabel}
        aria-invalid={problem ? true : undefined}
        title={`Fields: ${QUERY_HELP}`}
        placeholder={placeholder}
        autoFocus={autoFocus}
        autoComplete="off"
        spellCheck={false}
        value={value}
        onChange={(e) => {
          onChange(e.target.value);
          setCaret(e.target.selectionStart ?? e.target.value.length);
          setActive(-1);
          setDismissed(false);
        }}
        onKeyDown={onKeyDown}
        onKeyUp={syncCaret}
        onClick={syncCaret}
        onFocus={() => {
          setFocused(true);
          syncCaret();
        }}
        onBlur={() => setFocused(false)}
      />
      {open && (
        <div
          className="qinput-list"
          id="qinput-list"
          role="listbox"
          // Keeps focus in the input while a suggestion is clicked.
          onMouseDown={(e) => e.preventDefault()}
        >
          {items.map((s, i) => (
            <div
              key={s.insert}
              id={`qinput-opt-${i}`}
              role="option"
              aria-selected={i === index}
              className={`qinput-option${i === index ? ' active' : ''}`}
              onClick={() => accept(s)}
            >
              <span className="mono">{s.label}</span>
              <span className="muted">{s.detail}</span>
            </div>
          ))}
          <div className="qinput-foot">
            <span>
              <kbd>Tab</kbd> to complete
            </span>
            <span>{QUERY_HELP}</span>
          </div>
        </div>
      )}
      {!open && focused && problem && <div className="qinput-problem">{problem}</div>}
    </div>
  );
}
