import { toStoredTag } from './tags';

// A small query language for the search boxes, loosely after SharePoint's:
//
//   visitor                 free text, as before
//   upn:visitor             the field starts with / has a word starting with
//   upn=admin@contoso.com   the field is exactly
//   -upn:visitor            leave out what matches (also -visitor, -type=guest)
//   created>=2024-06        dates: > >= < <= = with YYYY, YYYY-MM or YYYY-MM-DD
//   name:"conference room"  quotes keep spaces together
//
// Terms are AND-ed. Each is translated to Microsoft Graph: $search for the
// properties Graph indexes for it, $filter for the rest. Graph has no
// "contains" filter and treats empty values inconsistently under "not", so an
// exclusion on a property that may be empty is applied to the loaded rows
// instead of being sent along.

export type Entity = 'user' | 'group' | 'app' | 'sp';

type Kind = 'text' | 'id' | 'bool' | 'date' | 'enum' | 'tags';

interface Field {
  name: string;
  aliases?: string[];
  prop: string;
  label: string;
  kind: Kind;
  // Graph can $search this property (word-prefix matching).
  search?: boolean;
  // Never empty, so "not" filters are safe to send to Graph.
  required?: boolean;
  // enum: accepted value → the $filter it stands for.
  values?: Record<string, string>;
}

const NAME: Field = {
  name: 'name',
  aliases: ['displayname'],
  prop: 'displayName',
  label: 'Display name',
  kind: 'text',
  search: true,
  required: true,
};
const ID: Field = { name: 'id', prop: 'id', label: 'Object ID', kind: 'id', required: true };
const APP_ID: Field = {
  name: 'appid',
  aliases: ['clientid'],
  prop: 'appId',
  label: 'Application (client) ID',
  kind: 'id',
  required: true,
};
const CREATED: Field = {
  name: 'created',
  prop: 'createdDateTime',
  label: 'Created on',
  kind: 'date',
};
const MAIL: Field = {
  name: 'mail',
  aliases: ['email'],
  prop: 'mail',
  label: 'Mail address',
  kind: 'text',
  search: true,
};
const ENABLED: Field = {
  name: 'enabled',
  prop: 'accountEnabled',
  label: 'Can sign in',
  kind: 'bool',
};
const UNIFIED = "groupTypes/any(t:t eq 'Unified')";

const FIELDS: Record<Entity, Field[]> = {
  user: [
    NAME,
    {
      name: 'upn',
      aliases: ['userprincipalname'],
      prop: 'userPrincipalName',
      label: 'User principal name',
      kind: 'text',
      search: true,
      required: true,
    },
    MAIL,
    {
      name: 'firstname',
      aliases: ['givenname'],
      prop: 'givenName',
      label: 'First name',
      kind: 'text',
    },
    { name: 'lastname', aliases: ['surname'], prop: 'surname', label: 'Last name', kind: 'text' },
    {
      name: 'company',
      aliases: ['companyname'],
      prop: 'companyName',
      label: 'Company',
      kind: 'text',
    },
    { name: 'department', prop: 'department', label: 'Department', kind: 'text' },
    { name: 'jobtitle', aliases: ['title'], prop: 'jobTitle', label: 'Job title', kind: 'text' },
    {
      name: 'type',
      prop: 'userType',
      label: 'Member or guest',
      kind: 'enum',
      values: { member: "userType eq 'Member'", guest: "userType eq 'Guest'" },
    },
    ENABLED,
    CREATED,
    ID,
  ],
  group: [
    NAME,
    MAIL,
    {
      name: 'description',
      prop: 'description',
      label: 'Description',
      kind: 'text',
      search: true,
    },
    {
      name: 'nickname',
      aliases: ['alias', 'mailnickname'],
      prop: 'mailNickname',
      label: 'Mail nickname',
      kind: 'text',
    },
    {
      name: 'type',
      prop: 'groupTypes',
      label: 'Kind of group',
      kind: 'enum',
      values: {
        security: `securityEnabled eq true and not(${UNIFIED})`,
        m365: UNIFIED,
        distribution: `mailEnabled eq true and securityEnabled eq false and not(${UNIFIED})`,
      },
    },
    CREATED,
    ID,
  ],
  app: [
    NAME,
    APP_ID,
    { name: 'uniquename', prop: 'uniqueName', label: 'Unique name', kind: 'text' },
    { name: 'tag', prop: 'tags', label: 'Tag', kind: 'tags' },
    {
      name: 'audience',
      prop: 'signInAudience',
      label: 'Sign-in audience',
      kind: 'enum',
      values: {
        myorg: "signInAudience eq 'AzureADMyOrg'",
        multiorg: "signInAudience eq 'AzureADMultipleOrgs'",
        personal:
          "signInAudience in ('AzureADandPersonalMicrosoftAccount','PersonalMicrosoftAccount')",
      },
    },
    CREATED,
    ID,
  ],
  sp: [
    NAME,
    APP_ID,
    {
      name: 'type',
      prop: 'servicePrincipalType',
      label: 'Kind of enterprise app',
      kind: 'enum',
      values: {
        application: "servicePrincipalType eq 'Application'",
        managedidentity: "servicePrincipalType eq 'ManagedIdentity'",
        legacy: "servicePrincipalType eq 'Legacy'",
      },
    },
    ENABLED,
    ID,
  ],
};

const ENTITIES = Object.keys(FIELDS) as Entity[];
const findField = (entity: Entity, name: string) =>
  FIELDS[entity].find((f) => f.name === name || f.aliases?.includes(name));
const isFieldName = (name: string) => ENTITIES.some((e) => findField(e, name));

// ------------------------------------------------------------------ Parsing

type Op = ':' | '=' | '>' | '>=' | '<' | '<=';

interface Term {
  neg: boolean;
  // Lower-cased field name as typed; undefined for free text.
  field?: string;
  op?: Op;
  value: string;
  start: number;
  end: number;
}

const unquote = (s: string) => s.replace(/"/g, '');

function parse(q: string): Term[] {
  const terms: Term[] = [];
  // Runs of non-space characters; a quoted stretch may contain spaces.
  for (const m of q.matchAll(/(?:[^\s"]+|"[^"]*"?)+/g)) {
    const raw = m[0];
    const neg = raw.startsWith('-') && raw.length > 1;
    const rest = neg ? raw.slice(1) : raw;
    const f = /^([A-Za-z]\w*)(>=|<=|:|=|>|<)(.*)$/.exec(rest);
    const base = { neg, start: m.index, end: m.index + raw.length };
    if (f && isFieldName(f[1].toLowerCase())) {
      terms.push({ ...base, field: f[1].toLowerCase(), op: f[2] as Op, value: unquote(f[3]) });
    } else {
      terms.push({ ...base, value: unquote(rest) });
    }
  }
  return terms;
}

// ------------------------------------------------------- Graph translation

const odata = (v: string) => v.replace(/'/g, "''");
const searchClause = (prop: string, v: string) =>
  `"${prop}:${v.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
const GUID = /^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
const TRUE = ['true', 'yes', '1'];
const FALSE = ['false', 'no', '0'];

const iso = (d: Date) => d.toISOString().replace(/\.\d{3}Z$/, 'Z');

// "2024", "2024-06" or "2024-06-03" → the first instant of that period and
// the first instant after it (UTC).
function period(v: string): [Date, Date] | null {
  const m = /^(\d{4})(?:-(\d{1,2}))?(?:-(\d{1,2}))?$/.exec(v);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = m[2] ? Number(m[2]) - 1 : 0;
  const d = m[3] ? Number(m[3]) : 1;
  const start = new Date(Date.UTC(y, mo, d));
  if (Number.isNaN(start.getTime())) return null;
  const end = m[3]
    ? new Date(Date.UTC(y, mo, d + 1))
    : m[2]
      ? new Date(Date.UTC(y, mo + 1, 1))
      : new Date(Date.UTC(y + 1, 0, 1));
  return [start, end];
}

function dateFilter(prop: string, op: Op, v: string): string | null {
  const p = period(v);
  if (!p) return null;
  const [start, end] = p;
  // Graph offers ge and le here, not gt and lt.
  const before = (d: Date) => iso(new Date(d.getTime() - 1000));
  switch (op) {
    case '>':
      return `${prop} ge ${iso(end)}`;
    case '>=':
      return `${prop} ge ${iso(start)}`;
    case '<':
      return `${prop} le ${before(start)}`;
    case '<=':
      return `${prop} le ${before(end)}`;
    default:
      return `${prop} ge ${iso(start)} and ${prop} le ${before(end)}`;
  }
}

type Row = Record<string, unknown>;
const text = (row: Row, prop: string) => String(row[prop] ?? '').toLowerCase();

export interface GraphQuery {
  // Query parameters to merge with $select. `filter` is AND-ed in (a tab's
  // own narrowing); `ordered` asks for A–Z when there is no $search, which
  // can't be combined with $orderby.
  params: (filter?: string, ordered?: boolean) => Record<string, string>;
  // Applies the exclusions that couldn't be sent to Graph.
  keep: (row: object) => boolean;
  // The free-text part, for ranking the hits.
  text: string;
}

// null: plain text without any syntax — search as before.
// 'none': the query uses a field this kind of object doesn't have (or a value
// that makes no sense), so nothing of this kind can match.
export function buildQuery(q: string, entity: Entity): GraphQuery | null | 'none' {
  const terms = parse(q).filter((t) => t.value !== '');
  if (!terms.some((t) => t.field || t.neg)) return null;

  const search: string[] = [];
  const freeText: string[] = [];
  const filters: string[] = [];
  const drop: ((row: Row) => boolean)[] = [];
  const not = (expr: string) => `not(${expr})`;

  for (const t of terms) {
    const v = t.value;
    const lower = v.toLowerCase();
    if (!t.field) {
      if (!t.neg) {
        freeText.push(v);
      } else {
        filters.push(not(`startswith(displayName,'${odata(v)}')`));
        drop.push((r) => text(r, 'displayName').includes(lower));
      }
      continue;
    }
    const f = findField(entity, t.field);
    if (!f) return 'none';
    const op = t.op ?? ':';
    const exact = op === '=';
    if (f.kind !== 'date' && op !== ':' && op !== '=') return 'none';

    if (f.kind === 'text') {
      const expr = exact ? `${f.prop} eq '${odata(v)}'` : `startswith(${f.prop},'${odata(v)}')`;
      if (!t.neg) {
        if (!exact && f.search) search.push(searchClause(f.prop, v));
        else filters.push(expr);
      } else {
        if (f.required) filters.push(not(expr));
        drop.push((r) => (exact ? text(r, f.prop) === lower : text(r, f.prop).includes(lower)));
      }
    } else if (f.kind === 'id') {
      // Object ids only compare whole; an appId can be searched by its start.
      if (f.prop === 'id' && !GUID.test(v)) return 'none';
      const expr =
        GUID.test(v) || exact
          ? `${f.prop} eq '${odata(v)}'`
          : `startswith(${f.prop},'${odata(v)}')`;
      filters.push(t.neg ? not(expr) : expr);
    } else if (f.kind === 'bool') {
      if (![...TRUE, ...FALSE].includes(lower)) return 'none';
      filters.push(`${f.prop} eq ${TRUE.includes(lower) !== t.neg}`);
    } else if (f.kind === 'date') {
      const expr = dateFilter(f.prop, op, v);
      if (!expr) return 'none';
      filters.push(t.neg ? not(expr) : expr);
    } else if (f.kind === 'enum') {
      const expr = f.values?.[lower];
      if (!expr) return 'none';
      filters.push(t.neg ? not(expr) : `(${expr})`);
    } else {
      // Tags: ours are stored with a prefix, so match with and without it.
      const stored = toStoredTag(v).toLowerCase();
      const hit = (tag: string) =>
        exact ? tag === lower || tag === stored : tag.startsWith(lower) || tag.startsWith(stored);
      if (!t.neg) {
        const one = (x: string) =>
          exact ? `tags/any(t:t eq '${odata(x)}')` : `tags/any(t:startswith(t,'${odata(x)}'))`;
        filters.push(`(${one(v)} or ${one(toStoredTag(v))})`);
      } else {
        drop.push((r) =>
          (Array.isArray(r[f.prop]) ? (r[f.prop] as unknown[]) : []).some((tag) =>
            hit(String(tag).toLowerCase()),
          ),
        );
      }
    }
  }

  // Free text searches the name; as the only search clause it may also look
  // at the other indexed properties (Graph can't group OR inside AND here).
  if (freeText.length === 1 && search.length === 0) {
    const props = FIELDS[entity].filter((f) => f.search).map((f) => f.prop);
    search.push(props.map((p) => searchClause(p, freeText[0])).join(' OR '));
  } else {
    for (const v of freeText) search.push(searchClause('displayName', v));
  }

  return {
    text: freeText.join(' '),
    keep: (row) => !drop.some((d) => d(row as Row)),
    params: (filter, ordered) => {
      const all = filter ? [...filters, `(${filter})`] : filters;
      return {
        ...(search.length ? { $search: search.join(' AND ') } : {}),
        ...(all.length ? { $filter: all.join(' and ') } : {}),
        ...(!search.length && ordered ? { $orderby: 'displayName' } : {}),
      };
    },
  };
}

// ------------------------------------------------------------- Completion

export interface Suggestion {
  // What the list shows, and what explains it.
  label: string;
  detail: string;
  // Replaces text[from, to).
  insert: string;
  from: number;
  to: number;
}

const OPERATORS = ': starts with · = exactly · - in front excludes';

// What could be typed at the caret: field names while a word is being typed,
// values once a field with a fixed set of them is chosen.
export function suggest(text: string, caret: number, entities: Entity[]): Suggestion[] {
  const chunk = [...text.matchAll(/(?:[^\s"]+|"[^"]*"?)+/g)].find(
    (m) => m.index <= caret && caret <= m.index + m[0].length,
  );
  if (!chunk) return [];
  const from = chunk.index;
  const to = chunk.index + chunk[0].length;
  const typed = text.slice(from, caret);
  const neg = typed.startsWith('-') ? '-' : '';
  const rest = typed.slice(neg.length);
  const f = /^([A-Za-z]\w*)(>=|<=|:|=|>|<)(.*)$/.exec(rest);

  if (f && isFieldName(f[1].toLowerCase())) {
    const name = f[1].toLowerCase();
    const prefix = unquote(f[3]).toLowerCase();
    const fields = entities.flatMap((e) => findField(e, name) ?? []);
    const values = new Map<string, string>();
    for (const fd of fields) {
      if (fd.kind === 'bool') {
        values.set('true', fd.label);
        values.set('false', `Not: ${fd.label.toLowerCase()}`);
      }
      for (const v of Object.keys(fd.values ?? {})) values.set(v, fd.label);
    }
    return [...values]
      .filter(([v]) => v.startsWith(prefix) && v !== prefix)
      .map(([v, detail]) => ({
        label: v,
        detail,
        insert: `${neg}${f[1]}${f[2]}${v} `,
        from,
        to,
      }));
  }

  if (!/^[A-Za-z]\w*$/.test(rest)) return [];
  const prefix = rest.toLowerCase();
  const seen = new Map<string, Suggestion>();
  for (const e of entities) {
    for (const fd of FIELDS[e]) {
      const hit = [fd.name, ...(fd.aliases ?? [])].some((n) => n.startsWith(prefix));
      if (!hit || seen.has(fd.name)) continue;
      const op = fd.kind === 'date' ? '>=' : fd.kind === 'bool' || fd.kind === 'enum' ? '=' : ':';
      seen.set(fd.name, {
        label: `${neg}${fd.name}${op}`,
        detail: fd.label,
        insert: `${neg}${fd.name}${op}`,
        from,
        to,
      });
    }
  }
  return [...seen.values()].slice(0, 8);
}

// Something to tell the user about the query as typed, if anything: a field
// none of these lists know, or a value that can't be right.
export function queryProblem(q: string, entities: Entity[]): string | null {
  for (const t of parse(q)) {
    // Looks like field syntax, but with a name nothing knows.
    const stray = !t.field && /^([A-Za-z]\w*)[:=](?!\/\/)./.exec(t.value);
    if (stray) return `“${stray[1]}” isn't a field here.`;
    if (!t.field || t.value === '') continue;
    const fields = entities.flatMap((e) => findField(e, t.field!) ?? []);
    if (fields.length === 0) return `“${t.field}” isn't a field here.`;
    const f = fields[0];
    const v = t.value.toLowerCase();
    if (f.kind === 'date' && !period(t.value))
      return `${t.field} takes a date: 2024, 2024-06 or 2024-06-03.`;
    if (f.kind !== 'date' && t.op !== ':' && t.op !== '=')
      return `${t.field} only works with : and =.`;
    if (f.kind === 'bool' && ![...TRUE, ...FALSE].includes(v))
      return `${t.field} is true or false.`;
    if (f.kind === 'enum' && !fields.some((x) => x.values?.[v]))
      return `${t.field} is one of: ${[...new Set(fields.flatMap((x) => Object.keys(x.values ?? {})))].join(', ')}.`;
    if (f.prop === 'id' && !GUID.test(t.value)) return 'id takes a complete object ID.';
  }
  return null;
}

export const QUERY_HELP = OPERATORS;
