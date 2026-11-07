import { useEffect, useState } from 'react';
import { useGraphToken } from '../auth/useGraphToken';
import { getMe } from '../graph/me';
import {
  createSchemaExtension,
  deleteSchemaExtension,
  listApplicationOwnerIds,
  listSchemaExtensionsOwnedBy,
  NO_BOOL_INT_TARGETS,
  SCHEMA_PROPERTY_TYPES,
  SCHEMA_TARGET_TYPES,
  updateSchemaExtension,
  type SchemaExtension,
  type SchemaExtensionProperty,
  type SchemaExtensionStatus,
  type SchemaPropertyType,
} from '../graph/schemaExtensions';
import type { Application } from '../graph/types';
import { CopyButton } from './CopyButton';
import { useConfirm } from './ConfirmDialog';
import { Modal } from './Modal';

// Schema extensions owned by this app registration: list, create, extend
// (additive only), move through the lifecycle, delete while in development.

interface Props {
  app: Application;
}

type Load<T> = { kind: 'loading' } | { kind: 'ok'; data: T } | { kind: 'error'; message: string };

const STATUS_BADGE: Record<SchemaExtensionStatus, string> = {
  InDevelopment: 'pending',
  Available: 'granted',
  Deprecated: 'expired',
};
const STATUS_LABEL: Record<SchemaExtensionStatus, string> = {
  InDevelopment: 'In development',
  Available: 'Available',
  Deprecated: 'Deprecated',
};

export function SchemaExtensionsManager({ app }: Props) {
  const token = useGraphToken();
  const { confirm, confirmDialog } = useConfirm();
  const [extensions, setExtensions] = useState<Load<SchemaExtension[]>>({ kind: 'loading' });
  // null = couldn't determine (owners unreadable); we then let Graph decide.
  const [isOwner, setIsOwner] = useState<boolean | null | 'loading'>('loading');
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [modal, setModal] = useState<'closed' | 'new' | SchemaExtension>('closed');

  useEffect(() => {
    let cancelled = false;
    setExtensions({ kind: 'loading' });
    setIsOwner('loading');
    listSchemaExtensionsOwnedBy(token, app.appId)
      .then((list) => {
        if (!cancelled)
          setExtensions({ kind: 'ok', data: [...list].sort((a, b) => a.id.localeCompare(b.id)) });
      })
      .catch((e: unknown) => {
        if (!cancelled)
          setExtensions({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
      });
    Promise.all([getMe(token), listApplicationOwnerIds(token, app.id)])
      .then(([me, owners]) => {
        if (!cancelled) setIsOwner(owners.some((o) => o.toLowerCase() === me.id.toLowerCase()));
      })
      .catch(() => {
        if (!cancelled) setIsOwner(null);
      });
    return () => {
      cancelled = true;
    };
  }, [token, app.id, app.appId]);

  const canEdit = isOwner !== false && isOwner !== 'loading';
  const list = extensions.kind === 'ok' ? extensions.data : [];
  const replace = (ext: SchemaExtension) =>
    setExtensions((prev) =>
      prev.kind === 'ok'
        ? {
            kind: 'ok',
            data: prev.data.some((e) => e.id === ext.id)
              ? prev.data.map((e) => (e.id === ext.id ? ext : e))
              : [...prev.data, ext].sort((a, b) => a.id.localeCompare(b.id)),
          }
        : prev,
    );

  const run = async (key: string, fn: () => Promise<void>) => {
    setBusy(key);
    setError(null);
    try {
      await fn();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  const setStatus = async (ext: SchemaExtension, status: SchemaExtensionStatus) => {
    const ok = await confirm(
      status === 'Available'
        ? {
            title: `Make ${ext.id} available?`,
            message: (
              <p style={{ margin: 0, fontSize: 14 }}>
                Any app in any tenant can then store data with this schema.{' '}
                <strong>It can't be moved back to development or deleted afterwards</strong> — only
                extended or deprecated.
              </p>
            ),
            confirmLabel: 'Make available',
          }
        : {
            title: `Deprecate ${ext.id}?`,
            message: (
              <p style={{ margin: 0, fontSize: 14 }}>
                <strong>The definition can no longer be read, extended or deleted by anyone</strong>
                , including this app. Data already stored with it stays readable and editable on the
                resources.
              </p>
            ),
            confirmLabel: 'Deprecate',
          },
    );
    if (!ok) return;
    await run(`status:${ext.id}`, async () => {
      await updateSchemaExtension(token, ext.id, { owner: app.appId, status });
      replace({ ...ext, status });
    });
  };

  const remove = async (ext: SchemaExtension) => {
    const ok = await confirm({
      title: `Delete ${ext.id}?`,
      message: (
        <p style={{ margin: 0, fontSize: 14 }}>
          Removes the definition. Data already written with it on users, groups etc. becomes
          inaccessible through this schema.
        </p>
      ),
      confirmLabel: 'Delete',
    });
    if (!ok) return;
    await run(`delete:${ext.id}`, async () => {
      await deleteSchemaExtension(token, ext.id);
      setExtensions((prev) =>
        prev.kind === 'ok' ? { kind: 'ok', data: prev.data.filter((e) => e.id !== ext.id) } : prev,
      );
    });
  };

  return (
    <>
      <div className="card">
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ margin: 0 }}>Schema extensions</h3>
            <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>
              Typed custom properties on users, groups, devices and more, owned by this app (
              <span className="mono">owner = {app.appId}</span>).
            </div>
          </div>
          <button
            className="primary"
            style={{ whiteSpace: 'nowrap' }}
            disabled={!canEdit || !!busy}
            onClick={() => setModal('new')}
          >
            + New extension
          </button>
        </div>
        {isOwner === false && (
          <p className="muted" style={{ fontSize: 12, marginBottom: 0 }}>
            You're not an owner of this app registration. Graph only lets owners of the owning app
            create or change its schema extensions — admin roles don't count — so everything here is
            read-only for you.
          </p>
        )}
        {isOwner === null && (
          <p className="muted" style={{ fontSize: 12, marginBottom: 0 }}>
            Couldn't read this app's owners. Changes are allowed here, but Graph will refuse them
            unless you're an owner of this app registration.
          </p>
        )}
      </div>

      {error && <div className="card error">{error}</div>}

      {extensions.kind === 'loading' && <span className="spinner" />}
      {extensions.kind === 'error' && <div className="card error">{extensions.message}</div>}
      {extensions.kind === 'ok' && list.length === 0 && (
        <div className="card empty">This app doesn't own any schema extensions.</div>
      )}

      {list.map((ext) => {
        const inDev = ext.status === 'InDevelopment';
        const available = ext.status === 'Available';
        return (
          <div className="card" key={ext.id}>
            <div className="row wrap" style={{ justifyContent: 'space-between', gap: 8 }}>
              <div style={{ minWidth: 0 }}>
                <div className="row" style={{ gap: 8 }}>
                  <span className="mono" style={{ fontSize: 14, fontWeight: 600 }}>
                    {ext.id}
                  </span>
                  <CopyButton value={ext.id} label="extension id" />
                  <span className={`badge ${STATUS_BADGE[ext.status]}`}>
                    {STATUS_LABEL[ext.status]}
                  </span>
                </div>
                {ext.description && (
                  <div className="muted" style={{ fontSize: 13, marginTop: 4 }}>
                    {ext.description}
                  </div>
                )}
              </div>
              {canEdit && ext.status !== 'Deprecated' && (
                <div className="row wrap" style={{ gap: 6 }}>
                  <button disabled={!!busy} onClick={() => setModal(ext)}>
                    Extend
                  </button>
                  {inDev && (
                    <button
                      className="primary"
                      disabled={!!busy}
                      onClick={() => void setStatus(ext, 'Available')}
                    >
                      Make available
                    </button>
                  )}
                  {available && (
                    <button
                      className="danger"
                      disabled={!!busy}
                      onClick={() => void setStatus(ext, 'Deprecated')}
                    >
                      Deprecate
                    </button>
                  )}
                  {inDev && (
                    <button className="danger" disabled={!!busy} onClick={() => void remove(ext)}>
                      Delete
                    </button>
                  )}
                </div>
              )}
            </div>

            <div className="kv" style={{ marginTop: 12 }}>
              <div className="k">Applies to</div>
              <div className="row wrap" style={{ gap: 4 }}>
                {ext.targetTypes.map((t) => (
                  <span key={t} className="badge">
                    {t}
                  </span>
                ))}
              </div>
              <div className="k">Use in Graph</div>
              <div className="mono muted" style={{ fontSize: 12 }}>
                $select=id,{ext.id}
              </div>
            </div>

            <table className="table" style={{ marginTop: 12 }}>
              <thead>
                <tr>
                  <th>Property</th>
                  <th>Type</th>
                </tr>
              </thead>
              <tbody>
                {ext.properties.map((p) => (
                  <tr key={p.name} style={{ cursor: 'default' }}>
                    <td className="mono">{p.name}</td>
                    <td>{p.type}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      })}

      {extensions.kind === 'ok' && (
        <p className="muted" style={{ fontSize: 12 }}>
          Deprecated extensions can't be read back from Graph, so they don't appear here.
        </p>
      )}

      {modal !== 'closed' && (
        <ExtensionModal
          existing={modal === 'new' ? undefined : modal}
          onClose={() => setModal('closed')}
          onSubmit={async (input) => {
            if (modal === 'new') {
              const created = await createSchemaExtension(token, {
                ...input,
                owner: app.appId,
              });
              replace(created);
            } else {
              await updateSchemaExtension(token, modal.id, {
                owner: app.appId,
                description: input.description,
                targetTypes: input.targetTypes,
                properties: input.properties,
              });
              replace({ ...modal, ...input, id: modal.id });
            }
            setModal('closed');
          }}
        />
      )}
      {confirmDialog}
    </>
  );
}

const ID_RE = /^[A-Za-z][A-Za-z0-9_]*$/;
const PROP_RE = /^[A-Za-z][A-Za-z0-9_]*$/;

interface DraftProperty {
  key: number;
  name: string;
  type: SchemaPropertyType;
}

function ExtensionModal({
  existing,
  onClose,
  onSubmit,
}: {
  existing?: SchemaExtension;
  onClose: () => void;
  onSubmit: (input: {
    id: string;
    description?: string;
    targetTypes: string[];
    properties: SchemaExtensionProperty[];
  }) => Promise<void>;
}) {
  const [id, setId] = useState('');
  const [description, setDescription] = useState(existing?.description ?? '');
  const [targets, setTargets] = useState<string[]>(existing?.targetTypes ?? ['user']);
  const [drafts, setDrafts] = useState<DraftProperty[]>(
    existing ? [] : [{ key: 1, name: '', type: 'String' }],
  );
  const [nextKey, setNextKey] = useState(2);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const lockedTargets = new Set((existing?.targetTypes ?? []).map((t) => t.toLowerCase()));
  const existingProps = existing?.properties ?? [];
  const allNames = [...existingProps.map((p) => p.name), ...drafts.map((d) => d.name.trim())]
    .filter(Boolean)
    .map((n) => n.toLowerCase());
  const dupNames = new Set(allNames.filter((n, i) => allNames.indexOf(n) !== i));
  const outlookTargets = targets.filter((t) => NO_BOOL_INT_TARGETS.includes(t));
  const badTypeForTarget =
    outlookTargets.length > 0 &&
    [...existingProps, ...drafts].some((p) => p.type === 'Boolean' || p.type === 'Integer');

  const draftsValid = drafts.every((d) => PROP_RE.test(d.name.trim()));
  const idValid = !!existing || ID_RE.test(id.trim());
  const totalProps = existingProps.length + drafts.length;
  const valid =
    idValid &&
    draftsValid &&
    dupNames.size === 0 &&
    targets.length > 0 &&
    totalProps > 0 &&
    !badTypeForTarget &&
    (!existing ||
      drafts.length > 0 ||
      description.trim() !== (existing.description ?? '') ||
      targets.length !== existing.targetTypes.length);

  const toggleTarget = (t: string) =>
    setTargets((prev) => (prev.includes(t) ? prev.filter((x) => x !== t) : [...prev, t]));

  const submit = async () => {
    if (!valid) return;
    setSubmitting(true);
    setError(null);
    try {
      await onSubmit({
        id: existing?.id ?? id.trim(),
        description: description.trim() || undefined,
        targetTypes: targets,
        properties: [
          ...existingProps,
          ...drafts.map((d) => ({ name: d.name.trim(), type: d.type })),
        ],
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setSubmitting(false);
    }
  };

  return (
    <Modal
      title={existing ? `Extend ${existing.id}` : 'New schema extension'}
      onClose={() => !submitting && onClose()}
      footer={
        <>
          <button disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          <button className="primary" disabled={submitting || !valid} onClick={submit}>
            {submitting ? 'Saving…' : existing ? 'Save' : 'Create'}
          </button>
        </>
      }
    >
      {existing ? (
        <p className="muted" style={{ fontSize: 12, marginTop: 0 }}>
          Changes are additive: you can add properties and target types and change the description.
          Existing properties and target types can't be removed or renamed.
        </p>
      ) : (
        <label className="field">
          <span>Id</span>
          <input
            autoFocus
            className="mono"
            value={id}
            onChange={(e) => setId(e.target.value)}
            placeholder="contoso_employeeData"
          />
          <span className="muted" style={{ fontSize: 11, marginTop: 4, display: 'block' }}>
            Use <span className="mono">{'{verifieddomain}_{name}'}</span> (domains under .com, .net,
            .org, .edu or .gov only) or just a name — Graph then assigns{' '}
            <span className="mono">ext&lt;8 chars&gt;_name</span>. Can't be changed later.
          </span>
          {id.trim() && !ID_RE.test(id.trim()) && (
            <span className="error" style={{ fontSize: 12 }}>
              Start with a letter; letters, digits and underscores only.
            </span>
          )}
        </label>
      )}

      <label className="field">
        <span>Description</span>
        <input value={description} onChange={(e) => setDescription(e.target.value)} />
      </label>

      <div className="field">
        <span
          style={{ display: 'block', fontSize: 12, color: 'var(--text-muted)', marginBottom: 4 }}
        >
          Applies to
        </span>
        <div className="row wrap" style={{ gap: 4 }}>
          {SCHEMA_TARGET_TYPES.map((t) => {
            const locked = lockedTargets.has(t.toLowerCase());
            const on = targets.some((x) => x.toLowerCase() === t.toLowerCase());
            return (
              <button
                key={t}
                type="button"
                className={on ? 'primary' : 'ghost'}
                style={{ minHeight: 28, padding: '2px 10px', fontSize: 12 }}
                disabled={locked}
                title={locked ? 'Already part of the definition — cannot be removed' : undefined}
                onClick={() => toggleTarget(t)}
              >
                {t}
              </button>
            );
          })}
        </div>
      </div>

      <div className="field">
        <span
          style={{ display: 'block', fontSize: 12, color: 'var(--text-muted)', marginBottom: 4 }}
        >
          Properties
        </span>
        {existingProps.map((p) => (
          <div key={p.name} className="row" style={{ gap: 6, marginBottom: 6 }}>
            <input className="mono" value={p.name} disabled style={{ flex: 2 }} />
            <input value={p.type} disabled style={{ flex: 1 }} />
            <span style={{ width: 32 }} />
          </div>
        ))}
        {drafts.map((d) => (
          <div key={d.key} className="row" style={{ gap: 6, marginBottom: 6 }}>
            <input
              className="mono"
              style={{ flex: 2 }}
              placeholder="propertyName"
              value={d.name}
              onChange={(e) =>
                setDrafts((prev) =>
                  prev.map((x) => (x.key === d.key ? { ...x, name: e.target.value } : x)),
                )
              }
            />
            <select
              style={{ flex: 1 }}
              value={d.type}
              onChange={(e) =>
                setDrafts((prev) =>
                  prev.map((x) =>
                    x.key === d.key ? { ...x, type: e.target.value as SchemaPropertyType } : x,
                  ),
                )
              }
            >
              {SCHEMA_PROPERTY_TYPES.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
            <button
              type="button"
              className="ghost"
              style={{ width: 32, padding: 0 }}
              aria-label="Remove property"
              onClick={() => setDrafts((prev) => prev.filter((x) => x.key !== d.key))}
            >
              ✕
            </button>
          </div>
        ))}
        <button
          type="button"
          className="ghost"
          onClick={() => {
            setDrafts((prev) => [...prev, { key: nextKey, name: '', type: 'String' }]);
            setNextKey((k) => k + 1);
          }}
        >
          + Add property
        </button>
        <div className="muted" style={{ fontSize: 11, marginTop: 4 }}>
          String and Binary hold up to 256 characters/bytes; DateTime is stored as UTC.
        </div>
        {drafts.some((d) => d.name.trim() && !PROP_RE.test(d.name.trim())) && (
          <div className="error" style={{ fontSize: 12 }}>
            Property names start with a letter; letters, digits and underscores only.
          </div>
        )}
        {dupNames.size > 0 && (
          <div className="error" style={{ fontSize: 12 }}>
            Duplicate property name: {[...dupNames].join(', ')}
          </div>
        )}
        {badTypeForTarget && (
          <div className="error" style={{ fontSize: 12 }}>
            Boolean and Integer aren't supported on {outlookTargets.join(', ')}.
          </div>
        )}
      </div>
      {error && <p className="error">{error}</p>}
    </Modal>
  );
}
