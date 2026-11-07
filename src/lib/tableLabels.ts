// On phones, `.table` rows are rendered as stacked cards (see the mobile
// breakpoint in index.css). Each cell then shows its column name via
// `td::before { content: attr(data-label) }`. Rather than threading a label
// through every <td> by hand, this watcher copies the <th> text onto the
// matching cells of every `.table` whenever the DOM changes.

function labelTable(table: HTMLTableElement) {
  const headRow = table.tHead?.rows[0];
  if (!headRow) return;
  const heads: string[] = [];
  for (const th of Array.from(headRow.cells)) {
    const text = th.textContent?.trim() ?? '';
    for (let i = 0; i < (th.colSpan || 1); i++) heads.push(text);
  }
  for (const body of Array.from(table.tBodies)) {
    for (const row of Array.from(body.rows)) {
      let col = 0;
      for (const cell of Array.from(row.cells)) {
        const span = cell.colSpan || 1;
        // Cells spanning several columns are empty states / notices: no label.
        const label = span === 1 ? (heads[col] ?? '') : '';
        if (cell.getAttribute('data-label') !== label) cell.setAttribute('data-label', label);
        col += span;
      }
    }
  }
}

export function installTableLabels(root: HTMLElement = document.body): () => void {
  let frame = 0;
  const run = () => {
    frame = 0;
    root.querySelectorAll<HTMLTableElement>('table.table').forEach(labelTable);
  };
  const schedule = () => {
    if (!frame) frame = requestAnimationFrame(run);
  };
  // Only childList/characterData are observed, so our own attribute writes
  // never retrigger the observer.
  const observer = new MutationObserver(schedule);
  observer.observe(root, { childList: true, subtree: true, characterData: true });
  run();
  return () => {
    observer.disconnect();
    if (frame) cancelAnimationFrame(frame);
  };
}
