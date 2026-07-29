// Tags live on the directory object (`application.tags`), not in this browser,
// so they're shared with everyone who manages the tenant. Entra also mirrors
// them onto the associated service principal.
//
// The collection is shared with Entra itself and with other tooling — portal
// apps carry things like `HideApp` and
// `WindowsAzureActiveDirectoryIntegratedApp`. So every tag this app manages is
// namespaced with `debble:`. Tags without that prefix are shown but never
// edited or removed here, and a PATCH always carries them through untouched.
//
// "debble:Favorite" is an ordinary tag with a reserved name — that keeps
// favourites and tags one mechanism instead of two.

export const TAG_PREFIX = 'debble:';

export const FAVORITE_LABEL = 'Favorite';
export const FAVORITE_TAG = `${TAG_PREFIX}${FAVORITE_LABEL}`;

function eq(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

// Is this tag one of ours (i.e. safe to edit)?
export function isOurTag(tag: string): boolean {
  return tag.toLowerCase().startsWith(TAG_PREFIX.toLowerCase());
}

// The part a user sees and types: "debble:Production" → "Production".
export function tagLabel(tag: string): string {
  return isOurTag(tag) ? tag.slice(TAG_PREFIX.length) : tag;
}

// Turn user input into the stored form, tolerating a hand-typed prefix.
export function toStoredTag(input: string): string {
  const normalized = input.trim().replace(/\s+/g, ' ');
  if (!normalized) return '';
  return isOurTag(normalized) ? normalized : `${TAG_PREFIX}${normalized}`;
}

export function ourTags(tags: string[] | undefined): string[] {
  return (tags ?? []).filter(isOurTag);
}

export function foreignTags(tags: string[] | undefined): string[] {
  return (tags ?? []).filter((t) => !isOurTag(t));
}

export function hasTag(tags: string[] | undefined, tag: string): boolean {
  return (tags ?? []).some((t) => eq(t, tag));
}

export function isFavorite(tags: string[] | undefined): boolean {
  return hasTag(tags, FAVORITE_TAG);
}

// Case-insensitive add that keeps the existing casing when the tag is already
// present, so we never create "debble:prod" next to "debble:Prod".
export function addTag(tags: string[] | undefined, input: string): string[] {
  const current = tags ?? [];
  const stored = toStoredTag(input);
  if (!stored || hasTag(current, stored)) return [...current];
  return [...current, stored];
}

// Only ever drops one of our own tags — foreign entries survive by construction.
export function removeTag(tags: string[] | undefined, tag: string): string[] {
  if (!isOurTag(tag)) return [...(tags ?? [])];
  return (tags ?? []).filter((t) => !eq(t, tag));
}

export function withFavorite(
  tags: string[] | undefined,
  favorite: boolean,
): string[] {
  return favorite ? addTag(tags, FAVORITE_TAG) : removeTag(tags, FAVORITE_TAG);
}

// Display order: Favorite, then our tags alphabetically, then foreign tags.
export function sortTags(tags: string[] | undefined): string[] {
  const byLabel = (a: string, b: string) =>
    tagLabel(a).localeCompare(tagLabel(b));
  const ours = ourTags(tags).sort((a, b) => {
    if (eq(a, FAVORITE_TAG)) return -1;
    if (eq(b, FAVORITE_TAG)) return 1;
    return byLabel(a, b);
  });
  return [...ours, ...foreignTags(tags).sort(byLabel)];
}
