// fileTypeIcon.ts — the ONE home for mapping a VBWD Space node to an fe-core
// Icon name (Finder-parity slice: per-file-type icons). Every view that
// renders a node icon (SpaceHome's list/grid rows today; any future one)
// reads it through this pure helper rather than re-deriving the mapping
// (DRY) — a template `v-if` chain per view would drift.
//
// Every icon name this module can return MUST already exist in fe-core's
// `ICON_PATHS` registry (`vbwd-fe-core/src/components/ui/icons.ts`) — an
// unmapped name silently falls back to a neutral dot there, which
// `tests/office-file-type-icon.spec.ts` guards against via `hasIcon`.

export type FileTypeIconInput = {
  kind: 'folder' | 'document';
  name: string;
  mime_type?: string | null;
};

const FOLDER_ICON = 'layers';
const DEFAULT_DOCUMENT_ICON = 'document';

/** MIME-type prefixes (checked with `startsWith`) mapped to an icon. */
const MIME_PREFIX_ICON_MAP: ReadonlyArray<readonly [prefix: string, icon: string]> = [
  ['image/', 'image'],
  ['audio/', 'bell'],
  ['video/', 'layout'],
];

/** Exact MIME types that do not fit a prefix rule above. */
const MIME_EXACT_ICON_MAP: Readonly<Record<string, string>> = {
  'application/pdf': 'invoice',
  'application/zip': 'bag',
  'application/x-7z-compressed': 'bag',
  'application/x-rar-compressed': 'bag',
  'application/x-tar': 'bag',
  'application/gzip': 'bag',
  'text/csv': 'bar-chart',
  'application/vnd.ms-excel': 'bar-chart',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'bar-chart',
};

/** Extension fallback — used when the server-sniffed MIME type is generic
 * (`application/octet-stream`) or absent, e.g. a browser File before upload. */
const EXTENSION_ICON_MAP: Readonly<Record<string, string>> = {
  pdf: 'invoice',
  // Image extensions matter even though `image/*` is covered by the MIME prefix
  // rule: the server SNIFFS the stored mime from the bytes, so a file whose
  // content is not recognisably an image (or is generic octet-stream) arrives
  // here as text/plain and would otherwise fall back to the generic document
  // icon despite an unambiguous .png name.
  png: 'image',
  jpg: 'image',
  jpeg: 'image',
  gif: 'image',
  webp: 'image',
  svg: 'image',
  bmp: 'image',
  zip: 'bag',
  rar: 'bag',
  '7z': 'bag',
  tar: 'bag',
  gz: 'bag',
  csv: 'bar-chart',
  xlsx: 'bar-chart',
  xls: 'bar-chart',
  mp3: 'bell',
  wav: 'bell',
  ogg: 'bell',
  flac: 'bell',
  m4a: 'bell',
  mp4: 'layout',
  mov: 'layout',
  webm: 'layout',
  avi: 'layout',
  mkv: 'layout',
  js: 'git-branch',
  ts: 'git-branch',
  jsx: 'git-branch',
  tsx: 'git-branch',
  py: 'git-branch',
  json: 'git-branch',
  html: 'git-branch',
  css: 'git-branch',
  vue: 'git-branch',
  sh: 'git-branch',
};

/** Resolve the fe-core Icon name for one VBWD Space node. */
export function fileTypeIcon(node: FileTypeIconInput): string {
  if (node.kind === 'folder') return FOLDER_ICON;
  return (
    iconForMimeType(node.mime_type) ?? iconForExtension(node.name) ?? DEFAULT_DOCUMENT_ICON
  );
}

function iconForMimeType(mimeType: string | null | undefined): string | undefined {
  if (!mimeType) return undefined;
  if (MIME_EXACT_ICON_MAP[mimeType]) return MIME_EXACT_ICON_MAP[mimeType];
  const prefixMatch = MIME_PREFIX_ICON_MAP.find(([prefix]) => mimeType.startsWith(prefix));
  return prefixMatch?.[1];
}

function iconForExtension(name: string): string | undefined {
  const dotIndex = name.lastIndexOf('.');
  if (dotIndex <= 0 || dotIndex === name.length - 1) return undefined;
  const extension = name.slice(dotIndex + 1).toLowerCase();
  return EXTENSION_ICON_MAP[extension];
}
