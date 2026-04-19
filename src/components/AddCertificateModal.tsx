import { useState } from 'react';
import { Modal } from './Modal';
import {
  exportPem,
  exportPfx,
  generateSelfSignedCertificate,
  parseDerCertificate,
  parsePemCertificate,
  parsePfxCertificate,
  triggerDownload,
  type GeneratedCertBundle,
  type ParsedCertificate,
} from '../lib/cert';

type Mode = 'upload' | 'generate';
type UploadFormat = 'pem' | 'pfx';
type OutFormat = 'pfx' | 'pem';

export interface CertSubmitPayload {
  displayName: string;
  keyBase64: string;
  thumbprintBase64: string;
  // Optional — Graph extracts start/end from the certificate DER when
  // omitted. We pass them only for display in the optimistic local update.
  startDateTime?: string;
  endDateTime?: string;
}

interface Props {
  onClose: () => void;
  onSubmit: (payload: CertSubmitPayload) => Promise<void>;
}

export function AddCertificateModal({ onClose, onSubmit }: Props) {
  const [mode, setMode] = useState<Mode>('generate');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [generated, setGenerated] = useState<GeneratedCertBundle | null>(null);
  const [downloadFormat, setDownloadFormat] = useState<OutFormat>('pfx');
  const [downloadPassword, setDownloadPassword] = useState('');
  const [downloaded, setDownloaded] = useState(false);

  return (
    <Modal
      title={generated ? 'Certificate ready' : 'New certificate'}
      onClose={() => !submitting && onClose()}
    >
      {!generated && (
        <div className="tabs" style={{ marginBottom: 16 }}>
          <button
            className={mode === 'generate' ? 'active' : ''}
            onClick={() => setMode('generate')}
            disabled={submitting}
          >
            Generate new
          </button>
          <button
            className={mode === 'upload' ? 'active' : ''}
            onClick={() => setMode('upload')}
            disabled={submitting}
          >
            Upload existing
          </button>
        </div>
      )}

      {!generated && mode === 'generate' && (
        <GenerateTab
          submitting={submitting}
          setSubmitting={setSubmitting}
          setError={setError}
          onGenerated={setGenerated}
          onRegistered={async (parsed, displayName) => {
            const payload: CertSubmitPayload = {
              displayName,
              keyBase64: parsed.certDerB64,
              thumbprintBase64: parsed.thumbprintB64,
              startDateTime: parsed.notBefore.toISOString(),
              endDateTime: parsed.notAfter.toISOString(),
            };
            await onSubmit(payload);
          }}
        />
      )}

      {!generated && mode === 'upload' && (
        <UploadTab
          submitting={submitting}
          setSubmitting={setSubmitting}
          setError={setError}
          onSubmit={onSubmit}
          onDone={onClose}
        />
      )}

      {generated && (
        <>
          <p className="error">
            This is the only time the private key will be available. Download
            it now — it cannot be retrieved later.
          </p>
          <div className="kv">
            <Kv label="Subject" value={generated.subjectCN} />
            <Kv label="Thumbprint" value={generated.thumbprintHex} mono />
            <Kv
              label="Valid from"
              value={generated.notBefore.toLocaleDateString()}
            />
            <Kv
              label="Valid until"
              value={generated.notAfter.toLocaleDateString()}
            />
          </div>

          <h4 style={{ marginTop: 20, marginBottom: 8 }}>Download</h4>
          <div className="row" style={{ gap: 16, marginBottom: 12 }}>
            <label className="row" style={{ gap: 6 }}>
              <input
                type="radio"
                style={{ width: 'auto' }}
                checked={downloadFormat === 'pfx'}
                onChange={() => setDownloadFormat('pfx')}
              />
              PFX (.pfx)
            </label>
            <label className="row" style={{ gap: 6 }}>
              <input
                type="radio"
                style={{ width: 'auto' }}
                checked={downloadFormat === 'pem'}
                onChange={() => setDownloadFormat('pem')}
              />
              PEM (.pem)
            </label>
          </div>
          {downloadFormat === 'pfx' && (
            <label className="field">
              <span>PFX password</span>
              <input
                type="password"
                value={downloadPassword}
                onChange={(e) => setDownloadPassword(e.target.value)}
                placeholder="Required by most tools"
                autoFocus
              />
            </label>
          )}
          <div className="row" style={{ gap: 8, marginTop: 12 }}>
            <button
              className="primary"
              onClick={() => {
                try {
                  if (downloadFormat === 'pfx') {
                    if (!downloadPassword) {
                      setError('Set a password for the PFX.');
                      return;
                    }
                    const bytes = exportPfx(generated, downloadPassword);
                    triggerDownload(
                      filename(generated.subjectCN, 'pfx'),
                      bytes,
                      'application/x-pkcs12',
                    );
                  } else {
                    const text = exportPem(generated);
                    triggerDownload(
                      filename(generated.subjectCN, 'pem'),
                      text,
                      'application/x-pem-file',
                    );
                  }
                  setDownloaded(true);
                  setError(null);
                } catch (e: unknown) {
                  setError(e instanceof Error ? e.message : String(e));
                }
              }}
            >
              Download {downloadFormat.toUpperCase()}
            </button>
            <button onClick={onClose}>
              {downloaded ? 'Done' : 'Close'}
            </button>
          </div>
          {downloaded && (
            <p className="muted" style={{ fontSize: 12, marginTop: 8 }}>
              Downloaded. You can now close this dialog.
            </p>
          )}
        </>
      )}

      {error && <p className="error" style={{ marginTop: 12 }}>{error}</p>}
    </Modal>
  );
}

function Kv({
  label,
  value,
  mono,
}: {
  label: string;
  value: string;
  mono?: boolean;
}) {
  return (
    <>
      <div className="k">{label}</div>
      <div className={mono ? 'mono' : undefined} style={{ wordBreak: 'break-all' }}>
        {value || '—'}
      </div>
    </>
  );
}

function filename(subject: string, ext: string): string {
  const safe = (subject || 'entraapp-cert').replace(/[^a-z0-9_-]+/gi, '_');
  return `${safe}.${ext}`;
}

function GenerateTab({
  submitting,
  setSubmitting,
  setError,
  onGenerated,
  onRegistered,
}: {
  submitting: boolean;
  setSubmitting: (b: boolean) => void;
  setError: (s: string | null) => void;
  onGenerated: (b: GeneratedCertBundle) => void;
  onRegistered: (parsed: ParsedCertificate, displayName: string) => Promise<void>;
}) {
  const [commonName, setCommonName] = useState('');
  const [months, setMonths] = useState(12);

  const generateAndRegister = async () => {
    const cn = commonName.trim() || 'entraapp-generated';
    setSubmitting(true);
    setError(null);
    try {
      const bundle = await generateSelfSignedCertificate({
        commonName: cn,
        months,
      });
      // Transition to the "ready" screen up front so the user can download
      // the private key even if the Graph registration step fails — nothing
      // is more frustrating than a generated-but-invisible private key.
      onGenerated(bundle);
      try {
        await onRegistered(bundle, cn);
      } catch (e: unknown) {
        setError(
          `Generated successfully, but Graph registration failed: ${
            e instanceof Error ? e.message : String(e)
          }. You can still download the private key below.`,
        );
      }
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <>
      <label className="field">
        <span>Subject (common name)</span>
        <input
          autoFocus
          value={commonName}
          onChange={(e) => setCommonName(e.target.value)}
          placeholder="e.g. my-app-cert"
        />
      </label>
      <label className="field">
        <span>Validity</span>
        <select
          value={months}
          onChange={(e) => setMonths(Number(e.target.value))}
        >
          <option value={3}>3 months</option>
          <option value={6}>6 months</option>
          <option value={12}>12 months</option>
          <option value={24}>24 months</option>
        </select>
      </label>
      <p className="muted" style={{ fontSize: 12 }}>
        A new 2048-bit RSA keypair will be generated in your browser. The
        public certificate will be registered with Entra; the private key
        stays in memory until you download it.
      </p>
      <div className="row" style={{ gap: 8, marginTop: 12 }}>
        <button
          className="primary"
          disabled={submitting}
          onClick={generateAndRegister}
        >
          {submitting ? 'Generating…' : 'Generate & register'}
        </button>
      </div>
    </>
  );
}

function UploadTab({
  submitting,
  setSubmitting,
  setError,
  onSubmit,
  onDone,
}: {
  submitting: boolean;
  setSubmitting: (b: boolean) => void;
  setError: (s: string | null) => void;
  onSubmit: (payload: CertSubmitPayload) => Promise<void>;
  onDone: () => void;
}) {
  const [format, setFormat] = useState<UploadFormat>('pem');
  const [file, setFile] = useState<File | null>(null);
  const [password, setPassword] = useState('');
  const [displayName, setDisplayName] = useState('');

  const submit = async () => {
    if (!file) {
      setError('Pick a file first.');
      return;
    }
    setSubmitting(true);
    setError(null);
    try {
      const buf = await file.arrayBuffer();
      let parsed: ParsedCertificate;
      if (format === 'pfx') {
        parsed = parsePfxCertificate(buf, password);
      } else {
        const text = new TextDecoder().decode(buf);
        if (/-----BEGIN CERTIFICATE-----/.test(text)) {
          parsed = parsePemCertificate(text);
        } else {
          parsed = parseDerCertificate(buf);
        }
      }
      await onSubmit({
        displayName: displayName.trim() || parsed.subjectCN || file.name,
        keyBase64: parsed.certDerB64,
        thumbprintBase64: parsed.thumbprintB64,
        startDateTime: parsed.notBefore.toISOString(),
        endDateTime: parsed.notAfter.toISOString(),
      });
      onDone();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <>
      <div className="row" style={{ gap: 16, marginBottom: 12 }}>
        <label className="row" style={{ gap: 6 }}>
          <input
            type="radio"
            style={{ width: 'auto' }}
            checked={format === 'pem'}
            onChange={() => setFormat('pem')}
          />
          PEM / DER (.pem, .cer, .crt)
        </label>
        <label className="row" style={{ gap: 6 }}>
          <input
            type="radio"
            style={{ width: 'auto' }}
            checked={format === 'pfx'}
            onChange={() => setFormat('pfx')}
          />
          PFX (.pfx, .p12)
        </label>
      </div>

      <label className="field">
        <span>Certificate file</span>
        <input
          type="file"
          accept={
            format === 'pfx'
              ? '.pfx,.p12,application/x-pkcs12'
              : '.pem,.cer,.crt,.der,application/x-x509-ca-cert,application/x-pem-file'
          }
          onChange={(e) => setFile(e.target.files?.[0] ?? null)}
        />
      </label>

      {format === 'pfx' && (
        <label className="field">
          <span>PFX password (if any)</span>
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="Leave blank if unencrypted"
          />
        </label>
      )}

      <label className="field">
        <span>Description (optional)</span>
        <input
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="Defaults to subject CN"
        />
      </label>

      <p className="muted" style={{ fontSize: 12 }}>
        Only the public certificate is registered with Entra. For PFX uploads,
        the private key is extracted in your browser to read the cert and then
        discarded.
      </p>

      <div className="row" style={{ gap: 8, marginTop: 12 }}>
        <button
          className="primary"
          disabled={submitting || !file}
          onClick={submit}
        >
          {submitting ? 'Uploading…' : 'Upload & register'}
        </button>
      </div>
    </>
  );
}
