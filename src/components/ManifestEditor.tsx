import { useEffect, useMemo, useRef, useState } from 'react';
// edcore.main = the editor with all its contributions (find, folding,
// suggest, hover…) but none of the bundled languages; JSON is added below.
// This keeps the lazily loaded chunk far smaller than all of monaco-editor.
import * as monaco from 'monaco-editor/esm/vs/editor/edcore.main';
import 'monaco-editor/esm/vs/language/json/monaco.contribution';
import EditorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker';
import JsonWorker from 'monaco-editor/esm/vs/language/json/json.worker?worker';
import { useGraphToken } from '../auth/useGraphToken';
import {
  getApplication,
  getApplicationManifest,
  updateApplication,
  type UpdateApplicationPatch,
} from '../graph/applications';
import type { Application } from '../graph/types';
import { APPLICATION_SCHEMA, READ_ONLY_PROPERTIES } from '../lib/applicationSchema';
import { useConfirm } from './ConfirmDialog';

(self as unknown as { MonacoEnvironment: monaco.Environment }).MonacoEnvironment = {
  getWorker: (_id, label) => (label === 'json' ? new JsonWorker() : new EditorWorker()),
};

const SCHEMA_URI = 'https://entra.2cloud.nl/schemas/graph-application.json';

monaco.languages.json.jsonDefaults.setDiagnosticsOptions({
  validate: true,
  allowComments: false,
  enableSchemaRequest: false,
  schemas: [
    { uri: SCHEMA_URI, fileMatch: ['graph-application-*.json'], schema: APPLICATION_SCHEMA },
  ],
});

type Manifest = Record<string, unknown>;

// Top-level properties that differ between the loaded and the edited JSON.
// Graph PATCH replaces each top-level property wholesale (collections too),
// so the property is the unit of change.
function diffManifest(original: Manifest, edited: Manifest) {
  const changed: string[] = [];
  const removed: string[] = [];
  const readOnly: string[] = [];
  for (const key of Object.keys(edited)) {
    if (JSON.stringify(original[key]) === JSON.stringify(edited[key])) continue;
    if (READ_ONLY_PROPERTIES.has(key)) readOnly.push(key);
    else changed.push(key);
  }
  for (const key of Object.keys(original)) if (!(key in edited)) removed.push(key);
  return { changed, removed, readOnly };
}

function stripOData(m: Manifest): Manifest {
  return Object.fromEntries(Object.entries(m).filter(([k]) => !k.startsWith('@odata.')));
}

// The full application object as Graph v1.0 returns it, in a JSON editor with
// schema validation, hover docs and completion. Save PATCHes only the
// top-level properties that changed.
export default function ManifestEditor({
  app,
  onAppChange,
}: {
  app: Application;
  onAppChange: (next: Application) => void;
}) {
  const token = useGraphToken();
  const { confirm, confirmDialog } = useConfirm();
  const host = useRef<HTMLDivElement>(null);
  const editor = useRef<monaco.editor.IStandaloneCodeEditor | null>(null);
  const [original, setOriginal] = useState<Manifest | null>(null);
  const [text, setText] = useState('');
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [markers, setMarkers] = useState(0);
  const [saving, setSaving] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);
  const [reloadKey, setReloadKey] = useState(0);

  useEffect(() => {
    let cancelled = false;
    setLoadError(null);
    getApplicationManifest(token, app.id)
      .then((m) => {
        if (cancelled) return;
        const clean = stripOData(m);
        setOriginal(clean);
        setText(JSON.stringify(clean, null, 2));
      })
      .catch((e: unknown) => {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, [token, app.id, reloadKey]);

  // One editor per mount; the model is swapped when a fresh copy loads.
  useEffect(() => {
    if (!host.current) return;
    const ed = monaco.editor.create(host.current, {
      theme: 'vs-dark',
      automaticLayout: true,
      minimap: { enabled: false },
      fontSize: 13,
      tabSize: 2,
      scrollBeyondLastLine: false,
      formatOnPaste: true,
      fixedOverflowWidgets: true,
    });
    editor.current = ed;
    const sub = monaco.editor.onDidChangeMarkers(() => {
      const model = ed.getModel();
      if (!model) return;
      setMarkers(
        monaco.editor
          .getModelMarkers({ resource: model.uri })
          .filter((m) => m.severity === monaco.MarkerSeverity.Error).length,
      );
    });
    return () => {
      sub.dispose();
      ed.getModel()?.dispose();
      ed.dispose();
      editor.current = null;
    };
  }, []);

  useEffect(() => {
    const ed = editor.current;
    if (!ed || original === null) return;
    const uri = monaco.Uri.parse(`inmemory://manifest/graph-application-${app.id}.json`);
    monaco.editor.getModel(uri)?.dispose();
    const model = monaco.editor.createModel(JSON.stringify(original, null, 2), 'json', uri);
    const sub = model.onDidChangeContent(() => setText(model.getValue()));
    ed.setModel(model);
    return () => sub.dispose();
  }, [original, app.id]);

  const parsed = useMemo((): { value?: Manifest; error?: string } => {
    try {
      const v = JSON.parse(text);
      if (!v || typeof v !== 'object' || Array.isArray(v))
        return { error: 'Root must be an object.' };
      return { value: v as Manifest };
    } catch (e: unknown) {
      return { error: e instanceof Error ? e.message : String(e) };
    }
  }, [text]);

  const diff = original && parsed.value ? diffManifest(original, parsed.value) : null;
  const dirty =
    !!diff && (diff.changed.length > 0 || diff.removed.length > 0 || diff.readOnly.length > 0);
  const blocked = !!parsed.error || !!diff?.removed.length || !!diff?.readOnly.length;

  const save = async () => {
    if (!diff || !parsed.value || blocked || diff.changed.length === 0) return;
    const ok = await confirm({
      title: 'Save manifest changes',
      message: (
        <>
          <p style={{ marginTop: 0 }}>
            These top-level properties will be replaced on the app registration:
          </p>
          <p className="mono">{diff.changed.join(', ')}</p>
          <p className="muted" style={{ marginBottom: 0 }}>
            Each property is sent whole. Collections such as redirectUris or appRoles replace the
            existing list.
          </p>
        </>
      ),
      confirmLabel: 'Save',
    });
    if (!ok) return;
    const body = Object.fromEntries(diff.changed.map((k) => [k, parsed.value![k]]));
    setSaving(true);
    setSaveError(null);
    setNotice(null);
    try {
      await updateApplication(token, app.id, body as UpdateApplicationPatch);
      const fresh = await getApplication(token, app.id);
      onAppChange({ ...app, ...fresh });
      setNotice(
        `Saved ${diff.changed.length} ${diff.changed.length === 1 ? 'property' : 'properties'}.`,
      );
      setReloadKey((k) => k + 1);
    } catch (e: unknown) {
      setSaveError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  };

  const discard = () => {
    if (original) editor.current?.getModel()?.setValue(JSON.stringify(original, null, 2));
  };

  const format = () => editor.current?.getAction('editor.action.formatDocument')?.run();

  return (
    <div className="card">
      <div
        className="row wrap"
        style={{ justifyContent: 'space-between', gap: 8, marginBottom: 8 }}
      >
        <div>
          <h3 style={{ margin: 0 }}>Manifest</h3>
          <div className="muted" style={{ fontSize: 12 }}>
            The app registration as Microsoft Graph v1.0 returns it (not the legacy Azure AD Graph
            manifest from the portal). Ctrl+Space for completion, hover for docs.
          </div>
        </div>
        <div className="row" style={{ gap: 6 }}>
          <button className="ghost" onClick={format} disabled={!original}>
            Format
          </button>
          <button
            className="ghost"
            onClick={() => setReloadKey((k) => k + 1)}
            disabled={saving}
            title="Load the current version from Graph (discards edits)"
          >
            Reload
          </button>
          <button onClick={discard} disabled={!dirty || saving}>
            Discard
          </button>
          <button
            className="primary"
            onClick={save}
            disabled={!dirty || blocked || saving || markers > 0 || !diff?.changed.length}
          >
            {saving ? 'Saving…' : diff?.changed.length ? `Save (${diff.changed.length})` : 'Save'}
          </button>
        </div>
      </div>

      {loadError && <p className="error">{loadError}</p>}
      {saveError && <p className="error">{saveError}</p>}
      {notice && !dirty && <p style={{ color: 'var(--success)', margin: '4px 0' }}>{notice}</p>}
      {parsed.error && original && <p className="error">JSON: {parsed.error}</p>}
      {!parsed.error && markers > 0 && (
        <p className="error">
          {markers} {markers === 1 ? 'error' : 'errors'} — see the squiggles in the editor.
        </p>
      )}
      {diff?.readOnly.length ? (
        <p className="error">
          Read-only properties changed: <span className="mono">{diff.readOnly.join(', ')}</span>.
          Revert them to save. Client secrets are managed in the Credentials tab.
        </p>
      ) : null}
      {diff?.removed.length ? (
        <p className="error">
          Properties removed: <span className="mono">{diff.removed.join(', ')}</span>. Graph can't
          delete a property; set it to null or an empty value instead.
        </p>
      ) : null}
      {diff && diff.changed.length > 0 && !blocked && (
        <p className="muted" style={{ fontSize: 12, margin: '4px 0' }}>
          Changed: <span className="mono">{diff.changed.join(', ')}</span>
        </p>
      )}

      <div
        ref={host}
        className="manifest-editor"
        style={{ visibility: original ? 'visible' : 'hidden' }}
      />
      {!original && !loadError && <span className="spinner" />}
      {confirmDialog}
    </div>
  );
}
