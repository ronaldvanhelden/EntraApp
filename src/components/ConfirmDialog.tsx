import { useCallback, useRef, useState, type ReactNode } from 'react';
import { Modal } from './Modal';

export interface ConfirmOptions {
  title: string;
  /** Body of the dialog. A string is rendered as a paragraph. */
  message?: ReactNode;
  /** Label of the destructive button. Defaults to "Remove". */
  confirmLabel?: string;
  cancelLabel?: string;
}

interface Pending extends ConfirmOptions {
  resolve: (ok: boolean) => void;
}

// Small destructive-action confirmation built on Modal. Rendered by useConfirm;
// not meant to be used directly.
function ConfirmDialog({
  title,
  message,
  confirmLabel = 'Remove',
  cancelLabel = 'Cancel',
  onResolve,
}: ConfirmOptions & { onResolve: (ok: boolean) => void }) {
  return (
    <Modal
      title={title}
      className="modal-sm"
      onClose={() => onResolve(false)}
      footer={
        <>
          <button type="button" onClick={() => onResolve(false)}>
            {cancelLabel}
          </button>
          <button type="button" className="danger" autoFocus onClick={() => onResolve(true)}>
            {confirmLabel}
          </button>
        </>
      }
    >
      {typeof message === 'string' ? <p style={{ margin: 0, fontSize: 14 }}>{message}</p> : message}
    </Modal>
  );
}

// Promise-based confirmation. Usage:
//
//   const { confirm, confirmDialog } = useConfirm();
//   if (!(await confirm({ title: 'Remove X?', message: '...' }))) return;
//   ...
//   return <>{...}{confirmDialog}</>;
//
// The returned element must be rendered somewhere in the component tree.
export function useConfirm() {
  const [pending, setPending] = useState<Pending | null>(null);
  const pendingRef = useRef<Pending | null>(null);

  const confirm = useCallback((opts: ConfirmOptions) => {
    // If a dialog is somehow already open, treat the new request as cancelled.
    pendingRef.current?.resolve(false);
    return new Promise<boolean>((resolve) => {
      const p: Pending = { ...opts, resolve };
      pendingRef.current = p;
      setPending(p);
    });
  }, []);

  const confirmDialog = pending ? (
    <ConfirmDialog
      {...pending}
      onResolve={(ok) => {
        pendingRef.current = null;
        setPending(null);
        pending.resolve(ok);
      }}
    />
  ) : null;

  return { confirm, confirmDialog };
}
