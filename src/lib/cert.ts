import forge from 'node-forge';

// Shapes returned/consumed by UI. All byte fields are base64-encoded.
export interface ParsedCertificate {
  // Base64(DER) of the X.509 certificate. This is what Graph wants in
  // keyCredentials.key.
  certDerB64: string;
  // Base64(SHA-1(DER)) — Graph's customKeyIdentifier.
  thumbprintB64: string;
  // Hex SHA-1 thumbprint for display.
  thumbprintHex: string;
  subjectCN: string;
  issuerCN: string;
  notBefore: Date;
  notAfter: Date;
}

export interface GeneratedCertBundle extends ParsedCertificate {
  // Private key material in PEM form (kept in memory only to produce exports).
  privateKeyPem: string;
  // Certificate in PEM form.
  certPem: string;
}

function binaryStringToBase64(bin: string): string {
  return forge.util.encode64(bin);
}

function derToThumbprint(derBinary: string): { b64: string; hex: string } {
  const md = forge.md.sha1.create();
  md.update(derBinary);
  const hex = md.digest().toHex().toUpperCase();
  const binary = forge.util.hexToBytes(hex);
  return { b64: binaryStringToBase64(binary), hex };
}

function certToParsed(cert: forge.pki.Certificate): ParsedCertificate {
  const asn1 = forge.pki.certificateToAsn1(cert);
  const derBinary = forge.asn1.toDer(asn1).getBytes();
  const { b64: thumbprintB64, hex: thumbprintHex } = derToThumbprint(derBinary);
  const cn = (attrs: forge.pki.CertificateField[]) =>
    (attrs.find((a) => a.shortName === 'CN' || a.name === 'commonName')
      ?.value as string | undefined) ?? '';
  return {
    certDerB64: binaryStringToBase64(derBinary),
    thumbprintB64,
    thumbprintHex,
    subjectCN: cn(cert.subject.attributes),
    issuerCN: cn(cert.issuer.attributes),
    notBefore: cert.validity.notBefore,
    notAfter: cert.validity.notAfter,
  };
}

// Parse a PEM-encoded X.509 certificate (public only). Returns the first
// certificate block encountered.
export function parsePemCertificate(pemText: string): ParsedCertificate {
  const msgs = forge.pem.decode(pemText);
  const certBlock = msgs.find((m) => /CERTIFICATE/i.test(m.type));
  if (!certBlock) throw new Error('No CERTIFICATE block found in PEM.');
  const cert = forge.pki.certificateFromPem(
    forge.pem.encode({
      type: 'CERTIFICATE',
      body: certBlock.body,
    }),
  );
  return certToParsed(cert);
}

// Parse a DER-encoded certificate (e.g. .cer file that's raw DER, not PEM).
export function parseDerCertificate(bytes: ArrayBuffer): ParsedCertificate {
  const binary = forge.util.createBuffer(new Uint8Array(bytes)).getBytes();
  const asn1 = forge.asn1.fromDer(binary);
  const cert = forge.pki.certificateFromAsn1(asn1);
  return certToParsed(cert);
}

// Parse a PKCS#12 / PFX bundle. We only care about the certificate (public
// portion) — the private key stays on the user's disk.
export function parsePfxCertificate(
  bytes: ArrayBuffer,
  password: string,
): ParsedCertificate {
  const binary = forge.util.createBuffer(new Uint8Array(bytes)).getBytes();
  const asn1 = forge.asn1.fromDer(binary);
  const p12 = forge.pkcs12.pkcs12FromAsn1(asn1, false, password);
  const bags = p12.getBags({ bagType: forge.pki.oids.certBag });
  const certBag = bags[forge.pki.oids.certBag]?.[0];
  const cert = certBag?.cert;
  if (!cert) throw new Error('No certificate found in PFX.');
  return certToParsed(cert);
}

// Generate an RSA-2048 keypair and a self-signed X.509 certificate valid for
// the requested number of months. Common name defaults to the display name.
export async function generateSelfSignedCertificate(options: {
  commonName: string;
  months: number;
}): Promise<GeneratedCertBundle> {
  const { commonName, months } = options;
  const keys = await new Promise<forge.pki.rsa.KeyPair>((resolve, reject) => {
    forge.pki.rsa.generateKeyPair(
      { bits: 2048, workers: -1 },
      (err, kp) => (err ? reject(err) : resolve(kp)),
    );
  });
  const cert = forge.pki.createCertificate();
  cert.publicKey = keys.publicKey;
  cert.serialNumber = randomSerial();
  cert.validity.notBefore = new Date();
  cert.validity.notAfter = new Date();
  cert.validity.notAfter.setMonth(
    cert.validity.notAfter.getMonth() + months,
  );
  const attrs: forge.pki.CertificateField[] = [
    { name: 'commonName', value: commonName },
  ];
  cert.setSubject(attrs);
  cert.setIssuer(attrs);
  cert.setExtensions([
    { name: 'basicConstraints', cA: false },
    {
      name: 'keyUsage',
      digitalSignature: true,
      keyEncipherment: true,
    },
    {
      name: 'extKeyUsage',
      clientAuth: true,
    },
    { name: 'subjectKeyIdentifier' },
  ]);
  cert.sign(keys.privateKey, forge.md.sha256.create());

  const parsed = certToParsed(cert);
  return {
    ...parsed,
    privateKeyPem: forge.pki.privateKeyToPem(keys.privateKey),
    certPem: forge.pki.certificateToPem(cert),
  };
}

function randomSerial(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  // Ensure first byte is positive (avoid negative ASN.1 integer).
  bytes[0] &= 0x7f;
  return Array.from(bytes)
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

// Build a PKCS#12 / PFX file as raw bytes.
export function exportPfx(
  bundle: GeneratedCertBundle,
  password: string,
): Uint8Array {
  const privateKey = forge.pki.privateKeyFromPem(bundle.privateKeyPem);
  const cert = forge.pki.certificateFromPem(bundle.certPem);
  const p12 = forge.pkcs12.toPkcs12Asn1(privateKey, [cert], password, {
    // 3DES is the widely interoperable default on Windows (OpenSSL legacy).
    algorithm: '3des',
    friendlyName: bundle.subjectCN || 'entraapp-cert',
  });
  const derBinary = forge.asn1.toDer(p12).getBytes();
  return binaryStringToBytes(derBinary);
}

// Build a single PEM file containing the cert and the PKCS#8 private key.
// Safe to open in text editors, and accepted by most tooling (openssl, curl,
// Python's ssl module, etc).
export function exportPem(bundle: GeneratedCertBundle): string {
  return `${bundle.certPem}\n${bundle.privateKeyPem}`;
}

function binaryStringToBytes(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

// Trigger a browser download for a byte array or text blob.
export function triggerDownload(
  filename: string,
  data: Uint8Array | string,
  mimeType: string,
): void {
  // Wrap Uint8Array in a fresh ArrayBuffer slice so TS's stricter BlobPart
  // typing (which rejects SharedArrayBuffer-backed views) is satisfied.
  const blob =
    typeof data === 'string'
      ? new Blob([data], { type: mimeType })
      : new Blob([data.slice().buffer], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
