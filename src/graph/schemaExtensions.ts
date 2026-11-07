import { graph, graphAll } from './client';

type TokenFn = () => Promise<string>;

// Graph schema extensions: strongly typed custom properties on users, groups,
// devices, … Each definition is owned by one app registration (its appId in
// `owner`). In delegated calls like ours, Graph only lets the signed-in user
// create or change a definition when they are an owner of that app
// registration, and the request must name the owner explicitly.

export type SchemaExtensionStatus = 'InDevelopment' | 'Available' | 'Deprecated';
export type SchemaPropertyType = 'Binary' | 'Boolean' | 'DateTime' | 'Integer' | 'String';

export interface SchemaExtensionProperty {
  name: string;
  type: SchemaPropertyType;
}

export interface SchemaExtension {
  id: string;
  description?: string | null;
  owner: string;
  status: SchemaExtensionStatus;
  targetTypes: string[];
  properties: SchemaExtensionProperty[];
}

export const SCHEMA_TARGET_TYPES = [
  'user',
  'group',
  'administrativeUnit',
  'device',
  'organization',
  'contact',
  'event',
  'message',
  'post',
  'todoTask',
  'todoTaskList',
] as const;

export const SCHEMA_PROPERTY_TYPES: SchemaPropertyType[] = [
  'String',
  'Integer',
  'Boolean',
  'DateTime',
  'Binary',
];

// Per the docs, Boolean and Integer properties aren't supported on these
// Outlook-backed resources.
export const NO_BOOL_INT_TARGETS = ['contact', 'message', 'event', 'post'];

export function listSchemaExtensionsOwnedBy(token: TokenFn, appId: string) {
  return graphAll<SchemaExtension>(token, '/schemaExtensions', {
    query: { $filter: `owner eq '${appId.replace(/'/g, "''")}'` },
  });
}

export interface CreateSchemaExtensionInput {
  // Either "{verifiedDomainPrefix}_{name}" or just a name, in which case
  // Graph assigns "ext{8 random chars}_{name}".
  id: string;
  description?: string;
  owner: string;
  targetTypes: string[];
  properties: SchemaExtensionProperty[];
}

export function createSchemaExtension(token: TokenFn, input: CreateSchemaExtensionInput) {
  return graph<SchemaExtension>(token, '/schemaExtensions', { method: 'POST', body: input });
}

// Changes must be additive: properties and target types can be added, never
// removed. `owner` has to be repeated on every update in delegated calls.
export function updateSchemaExtension(
  token: TokenFn,
  id: string,
  patch: {
    owner: string;
    description?: string;
    status?: SchemaExtensionStatus;
    targetTypes?: string[];
    properties?: SchemaExtensionProperty[];
  },
) {
  return graph<void>(token, `/schemaExtensions/${encodeURIComponent(id)}`, {
    method: 'PATCH',
    body: patch,
  });
}

// Only possible while the extension is InDevelopment.
export function deleteSchemaExtension(token: TokenFn, id: string) {
  return graph<void>(token, `/schemaExtensions/${encodeURIComponent(id)}`, { method: 'DELETE' });
}

// Object ids of the app registration's owners (users and service principals).
export async function listApplicationOwnerIds(token: TokenFn, applicationObjectId: string) {
  const rows = await graphAll<{ id: string }>(
    token,
    `/applications/${applicationObjectId}/owners`,
    {
      query: { $select: 'id' },
    },
  );
  return rows.map((r) => r.id);
}
