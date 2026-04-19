import { GRAPH_BASE } from '../auth/config';

export class GraphError extends Error {
  status: number;
  code?: string;
  constructor(status: number, message: string, code?: string) {
    super(message);
    this.status = status;
    this.code = code;
  }
}

export interface GraphRequestInit {
  method?: string;
  body?: unknown;
  headers?: Record<string, string>;
  query?: Record<string, string | number | boolean | undefined>;
  // Opt in to Microsoft Graph advanced query ($search, $count, advanced
  // $filter operators, $orderby combined with $filter). Sets
  // ConsistencyLevel: eventual and $count=true on the request.
  advanced?: boolean;
}

function buildUrl(
  path: string,
  query?: GraphRequestInit['query'],
  advanced?: boolean,
): string {
  const url = new URL(path.startsWith('http') ? path : `${GRAPH_BASE}${path}`);
  if (query) {
    for (const [k, v] of Object.entries(query)) {
      if (v === undefined) continue;
      url.searchParams.set(k, String(v));
    }
  }
  if (advanced && !url.searchParams.has('$count')) {
    url.searchParams.set('$count', 'true');
  }
  return url.toString();
}

export async function graph<T>(
  getToken: () => Promise<string>,
  path: string,
  init: GraphRequestInit = {},
): Promise<T> {
  const token = await getToken();
  const res = await fetch(buildUrl(path, init.query, init.advanced), {
    method: init.method ?? 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(init.advanced ? { ConsistencyLevel: 'eventual' } : {}),
      ...init.headers,
    },
    body: init.body !== undefined ? JSON.stringify(init.body) : undefined,
  });

  if (res.status === 204) return undefined as T;

  const text = await res.text();
  const parsed = text ? safeJson(text) : null;

  if (!res.ok) {
    const err = (parsed as { error?: { message?: string; code?: string } })
      ?.error;
    throw new GraphError(
      res.status,
      err?.message ?? `HTTP ${res.status}`,
      err?.code,
    );
  }
  return parsed as T;
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

export interface Paged<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export async function graphAll<T>(
  getToken: () => Promise<string>,
  path: string,
  init: GraphRequestInit = {},
  maxPages = 10,
): Promise<T[]> {
  const out: T[] = [];
  let next: string | undefined;
  let pages = 0;
  let first = true;

  while (pages < maxPages) {
    const pagePath = next ?? path;
    // Subsequent page requests must preserve the ConsistencyLevel header
    // (and any caller-supplied headers); the @odata.nextLink URL already
    // encodes the original query params.
    const page = await graph<Paged<T>>(
      getToken,
      pagePath,
      first
        ? init
        : {
            method: 'GET',
            advanced: init.advanced,
            headers: init.headers,
          },
    );
    out.push(...page.value);
    next = page['@odata.nextLink'];
    first = false;
    pages++;
    if (!next) break;
  }
  return out;
}
