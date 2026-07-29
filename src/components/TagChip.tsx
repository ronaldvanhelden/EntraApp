import {
  FAVORITE_LABEL,
  FAVORITE_TAG,
  TAG_PREFIX,
  hasTag,
  isOurTag,
  tagLabel,
} from '../lib/tags';

interface Props {
  // Stored form, e.g. "debble:Production".
  tag: string;
  // Renders an × that removes the tag. Ignored for tags we don't manage.
  onRemove?: () => void;
  // Clicking the chip filters by this tag. Ignored for tags we don't manage.
  onClick?: () => void;
  disabled?: boolean;
  title?: string;
}

// A single directory tag. Ours (prefixed `debble:`) render as editable chips
// with the prefix stripped; tags written by Entra or other tooling render muted
// and inert so they can't be filtered on or deleted by accident.
export function TagChip({ tag, onRemove, onClick, disabled, title }: Props) {
  const ours = isOurTag(tag);
  const favorite = hasTag([tag], FAVORITE_TAG);
  const label = tagLabel(tag);
  const body = (
    <>
      {favorite && <span aria-hidden>★</span>}
      {favorite ? FAVORITE_LABEL : label}
    </>
  );

  if (!ours) {
    return (
      <span
        className="tag-chip foreign"
        title={`${tag} — set outside EntraApp (no ${TAG_PREFIX} prefix), read-only here`}
      >
        <span className="tag-chip-main">{label}</span>
      </span>
    );
  }

  return (
    <span
      className={`tag-chip${favorite ? ' favorite' : ''}`}
      title={title ?? tag}
    >
      {onClick ? (
        <button
          type="button"
          className="tag-chip-main"
          disabled={disabled}
          onClick={(e) => {
            e.stopPropagation();
            onClick();
          }}
        >
          {body}
        </button>
      ) : (
        <span className="tag-chip-main">{body}</span>
      )}
      {onRemove && (
        <button
          type="button"
          className="tag-chip-remove"
          aria-label={`Remove tag ${label}`}
          title={`Remove tag ${label}`}
          disabled={disabled}
          onClick={(e) => {
            e.stopPropagation();
            onRemove();
          }}
        >
          ×
        </button>
      )}
    </span>
  );
}

// Star toggle used in list rows and on the detail header. Writes the reserved
// debble:Favorite tag on the application.
export function FavoriteStar({
  favorite,
  onToggle,
  disabled,
}: {
  favorite: boolean;
  onToggle: () => void;
  disabled?: boolean;
}) {
  return (
    <button
      type="button"
      className={`favorite-star${favorite ? ' on' : ''}`}
      aria-pressed={favorite}
      aria-label={favorite ? 'Remove from favorites' : 'Add to favorites'}
      title={
        favorite
          ? `Remove the ${FAVORITE_TAG} tag`
          : `Tag as ${FAVORITE_TAG} (stored on the app registration)`
      }
      disabled={disabled}
      onClick={(e) => {
        e.stopPropagation();
        onToggle();
      }}
    >
      {favorite ? '★' : '☆'}
    </button>
  );
}
