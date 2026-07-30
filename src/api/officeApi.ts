// fe-user client for the VBWD Office backend (S147-1 vault + S147-2 sharing).
//
// Consumes the session-scoped vault + share-CRUD endpoints, all `@require_auth`
// + `@require_user_permission('office.use')` and owner-resolved from the
// backend session — this client never sends an owner id.
//
//   GET    /api/v1/office/nodes?parent_id=            — list children (folders + documents)
//   POST   /api/v1/office/folders                     — create folder
//   POST   /api/v1/office/documents                   — multipart upload
//   GET    /api/v1/office/documents/<id>               — metadata
//   GET    /api/v1/office/documents/<id>/content        — stream (download)
//   GET    /api/v1/office/documents/<id>/versions       — history
//   POST   /api/v1/office/documents/<id>/restore        — restore a version
//   PATCH  /api/v1/office/nodes/<id>                    — rename / move
//   DELETE /api/v1/office/nodes/<id>                    — trash (?purge=true to purge)
//   GET    /api/v1/office/usage                         — {bytes_used, bytes_quota}
//   GET    /api/v1/office/nodes/<id>/shares             — list shares on a node
//   POST   /api/v1/office/nodes/<id>/shares             — create a share (token shown ONCE)
//   PATCH  /api/v1/office/shares/<id>                    — update permission/expiry/password
//   DELETE /api/v1/office/shares/<id>                    — revoke
//   GET    /api/v1/office/shared-with-me                — named shares addressed to me
//   GET    /api/v1/office/shares/<id>/access-log         — the C6 audit trail
//
// `officePublicShareApi` (below, separate export) talks to the ANONYMOUS
// `/api/v1/office/public/<token>*` routes and deliberately never sends the
// session Authorization header it borrows from `authHeaders()` — a stranger
// with just the link must be able to use every one of those calls.

const API = '/api/v1/office';

function authHeaders(): Record<string, string> {
  const token = localStorage.getItem('auth_token') || '';
  return token ? { Authorization: `Bearer ${token}` } : {};
}

async function get<T>(url: string, params?: Record<string, string>): Promise<T> {
  const queryString = params ? '?' + new URLSearchParams(params).toString() : '';
  const response = await fetch(url + queryString, { headers: authHeaders() });
  if (!response.ok) throw new Error(`GET ${url} failed: ${response.status}`);
  return response.json();
}

async function sendJson<T>(url: string, method: string, body?: unknown): Promise<T> {
  const response = await fetch(url, {
    method,
    headers: { ...authHeaders(), 'Content-Type': 'application/json' },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  if (!response.ok) throw new Error(`${method} ${url} failed: ${response.status}`);
  return response.json();
}

/** Prefer the server-provided attachment name from Content-Disposition. */
function filenameFromDisposition(response: Response, fallback: string): string {
  const disposition = response.headers.get('Content-Disposition') || '';
  const match = /filename\*?=(?:UTF-8'')?"?([^";]+)"?/i.exec(disposition);
  return match ? decodeURIComponent(match[1]) : fallback;
}

/**
 * Fetch an authed download as a Blob. Downloads are `@require_auth` with a
 * Bearer session token a plain `<a href>` navigation can't carry (→ 401), so
 * we fetch with the auth header and hand back a Blob the caller turns into an
 * object-URL download.
 */
async function downloadBlob(
  url: string,
  fallbackName: string,
): Promise<{ blob: Blob; filename: string }> {
  const response = await fetch(url, { headers: authHeaders() });
  if (!response.ok) throw new Error(`Download failed: ${response.status}`);
  return { blob: await response.blob(), filename: filenameFromDisposition(response, fallbackName) };
}

export type OfficeNodeKind = 'folder' | 'document';
export type OfficeDocType = 'file' | 'text' | 'sheet';

export interface OfficeNode {
  id: string;
  parent_id: string | null;
  kind: OfficeNodeKind;
  name: string;
  trashed_at: string | null;
  created_at: string;
  updated_at: string;
  doc_type?: OfficeDocType | null;
  mime_type?: string | null;
  size_bytes?: number | null;
}

/** The REAL wire shape of one `GET /nodes` (and `/shared-with-me`) list
 * entry: `NodeSummary.to_dict()` (routes.py) nests document facts under a
 * `document` sub-object rather than flattening them. `SpaceHome.vue` /
 * `OfficePreviewPane.vue` read `node.mime_type` / `node.size_bytes` FLAT
 * (`OfficeNode`, above) — left un-normalised, every document in a listing
 * silently shows "0 B" and never previews inline. `normalizeListedNode`
 * is the one place that bridges the two shapes. */
interface OfficeNodeListItemWire extends OfficeNode {
  document?: {
    id: string;
    doc_type: OfficeDocType;
    mime_type: string;
    size_bytes: number;
  };
}

function normalizeListedNode(wireNode: OfficeNodeListItemWire): OfficeNode {
  if (!wireNode.document) return wireNode;
  return {
    ...wireNode,
    doc_type: wireNode.document.doc_type,
    mime_type: wireNode.document.mime_type,
    size_bytes: wireNode.document.size_bytes,
  };
}

export interface OfficeVersion {
  id: string;
  version_no: number;
  size_bytes: number;
  sha256: string;
  created_by_user_id: string;
  created_at: string;
}

export interface OfficeUsage {
  bytes_used: number;
  bytes_quota: number;
}

// ---------------------------------------------------------------------------
// S147-2 — sharing
// ---------------------------------------------------------------------------

export type OfficeSharePermission = 'view' | 'comment' | 'edit';

export interface OfficeShare {
  id: string;
  node_id: string;
  created_by_user_id: string;
  permission: OfficeSharePermission;
  subject_user_id: string | null;
  allow_anonymous: boolean;
  has_password: boolean;
  expires_at: string | null;
  revoked_at: string | null;
  last_used_at: string | null;
  created_at: string;
  updated_at: string;
}

/** The create-share response — the ONLY moment the plaintext token exists
 * on the wire. It is never fetchable again after this response. */
export interface OfficeShareWithToken extends OfficeShare {
  token: string;
}

export interface OfficeShareAccessLogEntry {
  id: string;
  share_id: string;
  action: string;
  ip_hash: string | null;
  occurred_at: string;
}

export interface OfficeSharedWithMeEntry extends OfficeNode {
  share: OfficeShare;
}

export interface CreateShareInput {
  permission: OfficeSharePermission;
  recipient_email?: string;
  allow_anonymous?: boolean;
  password?: string;
  expires_at?: string | null;
}

export interface UpdateShareInput {
  permission?: OfficeSharePermission;
  expires_at?: string | null;
  password?: string | null;
  allow_anonymous?: boolean;
}

/**
 * Upload progress callback fired as the browser reports bytes sent. Wired
 * through XMLHttpRequest (fetch has no cross-browser upload-progress event).
 */
export type UploadProgressHandler = (fractionComplete: number) => void;

/**
 * POST a multipart upload with real progress reporting via XMLHttpRequest —
 * `fetch()` has no cross-browser upload-progress event, which the drag-drop
 * progress indicator needs.
 */
function uploadMultipart<T>(
  url: string,
  formData: FormData,
  onProgress?: UploadProgressHandler,
): Promise<T> {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open('POST', url, true);
    for (const [header, value] of Object.entries(authHeaders())) {
      request.setRequestHeader(header, value);
    }
    request.upload.onprogress = (event) => {
      if (onProgress && event.lengthComputable) {
        onProgress(event.loaded / event.total);
      }
    };
    request.onload = () => {
      if (request.status >= 200 && request.status < 300) {
        try {
          resolve(JSON.parse(request.responseText) as T);
        } catch {
          reject(new Error('Upload response was not valid JSON'));
        }
      } else {
        reject(new Error(`Upload failed: ${request.status}`));
      }
    };
    request.onerror = () => reject(new Error('Upload failed: network error'));
    request.send(formData);
  });
}

export const officeApi = {
  listNodes(parentId: string | null): Promise<OfficeNode[]> {
    return get<{ items: OfficeNodeListItemWire[] }>(
      `${API}/nodes`,
      parentId ? { parent_id: parentId } : {},
    ).then((r) => r.items.map(normalizeListedNode));
  },

  createFolder(name: string, parentId: string | null): Promise<OfficeNode> {
    return sendJson(`${API}/folders`, 'POST', { name, parent_id: parentId });
  },

  uploadDocument(
    file: File,
    parentId: string | null,
    onProgress?: UploadProgressHandler,
  ): Promise<OfficeNode> {
    const formData = new FormData();
    formData.append('file', file, file.name);
    if (parentId) formData.append('parent_id', parentId);
    return uploadMultipart(`${API}/documents`, formData, onProgress);
  },

  getDocument(documentId: string): Promise<OfficeNode> {
    return get(`${API}/documents/${documentId}`);
  },

  listVersions(documentId: string): Promise<OfficeVersion[]> {
    return get<{ items: OfficeVersion[] }>(`${API}/documents/${documentId}/versions`).then(
      (r) => r.items,
    );
  },

  /** `versionNo` — the backend's restore contract is `{version_no: <int>}`
   * (`routes.py` rejects anything else with 400); the version's OWN id is
   * never sent, only its ordinal. */
  restoreVersion(documentId: string, versionNo: number): Promise<OfficeVersion> {
    return sendJson(`${API}/documents/${documentId}/restore`, 'POST', { version_no: versionNo });
  },

  /** Browser-download / preview-stream URL (session-auth). */
  contentUrl(documentId: string): string {
    return `${API}/documents/${documentId}/content`;
  },

  /** Fetch a document's content as an auth-carrying Blob (see `downloadBlob`). */
  downloadDocument(documentId: string, fallbackName: string): Promise<{ blob: Blob; filename: string }> {
    return downloadBlob(`${API}/documents/${documentId}/content`, fallbackName);
  },

  renameNode(nodeId: string, name: string): Promise<OfficeNode> {
    return sendJson(`${API}/nodes/${nodeId}`, 'PATCH', { name });
  },

  moveNode(nodeId: string, parentId: string | null): Promise<OfficeNode> {
    return sendJson(`${API}/nodes/${nodeId}`, 'PATCH', { parent_id: parentId });
  },

  async trashNode(nodeId: string, purge = false): Promise<void> {
    const query = purge ? '?purge=true' : '';
    const response = await fetch(`${API}/nodes/${nodeId}${query}`, {
      method: 'DELETE',
      headers: authHeaders(),
    });
    if (!response.ok) throw new Error(`DELETE ${API}/nodes/${nodeId} failed: ${response.status}`);
  },

  getUsage(): Promise<OfficeUsage> {
    return get(`${API}/usage`);
  },

  // -------------------------------------------------------------------
  // S147-2 — owner-side share CRUD
  // -------------------------------------------------------------------

  listShares(nodeId: string): Promise<OfficeShare[]> {
    return get<{ items: OfficeShare[] }>(`${API}/nodes/${nodeId}/shares`).then((r) => r.items);
  },

  createShare(nodeId: string, input: CreateShareInput): Promise<OfficeShareWithToken> {
    return sendJson(`${API}/nodes/${nodeId}/shares`, 'POST', input);
  },

  updateShare(shareId: string, input: UpdateShareInput): Promise<OfficeShare> {
    return sendJson(`${API}/shares/${shareId}`, 'PATCH', input);
  },

  async revokeShare(shareId: string): Promise<void> {
    const response = await fetch(`${API}/shares/${shareId}`, {
      method: 'DELETE',
      headers: authHeaders(),
    });
    if (!response.ok) throw new Error(`DELETE ${API}/shares/${shareId} failed: ${response.status}`);
  },

  listSharedWithMe(): Promise<OfficeSharedWithMeEntry[]> {
    return get<{ items: OfficeSharedWithMeEntry[] }>(`${API}/shared-with-me`).then((r) => r.items);
  },

  shareAccessLog(shareId: string): Promise<OfficeShareAccessLogEntry[]> {
    return get<{ items: OfficeShareAccessLogEntry[] }>(`${API}/shares/${shareId}/access-log`).then(
      (r) => r.items,
    );
  },

  /** Open an entry from "Shared with me" — session-only (no token); the
   * caller's own login is what AccessResolver checks on the backend. */
  getSharedDocument(nodeId: string): Promise<OfficeNode> {
    return get(`${API}/shared-with-me/${nodeId}`);
  },

  downloadSharedDocument(nodeId: string, fallbackName: string): Promise<{ blob: Blob; filename: string }> {
    return downloadBlob(`${API}/shared-with-me/${nodeId}/content`, fallbackName);
  },
};

// ---------------------------------------------------------------------------
// S147-3 — VBWD Docs (the editor, edit lease, AI helper). A Doc is an
// office_node of doc_type='text'; every route here is session-auth
// (`@require_auth`), reached only as the owner or a NAMED share — never an
// anonymous share token (see the backend's `OfficeDocEditorService`
// docstring). Content is a structured JSON node tree (ProseMirror/Tiptap-
// shaped) — never HTML (epic D7).
// ---------------------------------------------------------------------------

/** A ProseMirror/Tiptap-shaped node — deliberately untyped beyond the shape
 * Tiptap itself produces (`editor.getJSON()`); the backend is the schema's
 * one source of truth (`doc_content.py`). */
// A Tiptap/ProseMirror document is a genuinely open-ended node tree, so the
// value type is unknown rather than `any` — callers must narrow instead of
// silently inheriting an escape hatch.
export type OfficeDocContentModel = Record<string, unknown>;

export type OfficeDocAccess = 'owner' | 'edit' | 'comment' | 'view';

export interface OfficeDocLeaseState {
  held: boolean;
  holder_user_id: string | null;
  is_self: boolean;
  granted: boolean;
  expires_at: string | null;
}

export interface OfficeDocView {
  id: string;
  parent_id: string | null;
  kind: OfficeNodeKind;
  name: string;
  trashed_at: string | null;
  created_at: string;
  updated_at: string;
  document: {
    id: string;
    node_id: string;
    doc_type: OfficeDocType;
    mime_type: string;
    size_bytes: number;
    ai_enabled: boolean;
    current_version_id: string | null;
    created_at: string;
    updated_at: string;
  };
  content: OfficeDocContentModel;
  version_no: number;
  access: OfficeDocAccess;
  lease: OfficeDocLeaseState;
}

/** The `409` conflict bodies ``PUT /docs/<id>`` can return — a stale
 * ``base_version_no`` or a lease held by someone else. Thrown as one typed
 * error so the editor can branch on `.kind` without re-parsing JSON. */
export class OfficeDocConflictError extends Error {
  readonly kind: 'stale_version' | 'locked';
  readonly currentVersionNo?: number;
  readonly holderUserId?: string;
  readonly expiresAt?: string;

  constructor(body: {
    error: string;
    current_version_no?: number;
    holder_user_id?: string;
    expires_at?: string;
  }) {
    super(body.error === 'locked' ? 'Document is locked by another editor' : 'A newer version already exists');
    this.name = 'OfficeDocConflictError';
    this.kind = body.error === 'locked' ? 'locked' : 'stale_version';
    this.currentVersionNo = body.current_version_no;
    this.holderUserId = body.holder_user_id;
    this.expiresAt = body.expires_at;
  }
}

/** The `403` the AI route returns when the capability id is refused, the
 * caller lacks edit access, or `ai_enabled` is off for this document. */
export class OfficeAiForbiddenError extends Error {
  constructor(message = 'The AI helper is not available for this document') {
    super(message);
    this.name = 'OfficeAiForbiddenError';
  }
}

/** The `429` the AI route returns once the caller's monthly budget is used
 * up — no provider call was made for this request. */
export class OfficeAiBudgetExceededError extends Error {
  constructor(message = 'Monthly AI helper budget exhausted') {
    super(message);
    this.name = 'OfficeAiBudgetExceededError';
  }
}

export type OfficeAiCapability =
  | 'continue_writing'
  | 'rewrite_shorter'
  | 'rewrite_longer'
  | 'rewrite_formal'
  | 'rewrite_plain'
  | 'summarize'
  | 'fix_grammar'
  | 'translate'
  | 'outline';

export interface RunAiCapabilityInput {
  capability: OfficeAiCapability;
  selectionText?: string;
  contextBefore?: string;
  contextAfter?: string;
  targetLanguage?: string;
}

export interface OfficeAiResult {
  capability: OfficeAiCapability;
  connection_slug: string;
  proposed_text: string;
}

// ---------------------------------------------------------------------------
// S147-3 toolbar bundle — image assets, "insert table from file", and
// Print/PDF/DOCX/Markdown export for VBWD Docs.
// ---------------------------------------------------------------------------

/** ``POST /docs/<id>/assets`` — the ``src`` is an ``office-asset:<node_id>``
 * REFERENCE, never a fetchable URL by itself (see ``fetchAssetBlob`` below). */
export interface OfficeDocAssetUploadResult {
  node_id: string;
  src: string;
}

export interface OfficeDocTableImportResult {
  rows: string[][];
  row_count: number;
  column_count: number;
}

export type OfficeDocExportFormat = 'pdf' | 'docx' | 'md';

async function docPutOrThrowConflict(
  url: string,
  body: unknown,
): Promise<OfficeVersion> {
  const response = await fetch(url, {
    method: 'PUT',
    headers: { ...authHeaders(), 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (response.status === 409) {
    throw new OfficeDocConflictError(await response.json());
  }
  if (!response.ok) throw new Error(`PUT ${url} failed: ${response.status}`);
  return response.json();
}

export const officeDocApi = {
  createDoc(name: string, parentId: string | null): Promise<OfficeDocView> {
    return sendJson(`${API}/docs`, 'POST', { name, parent_id: parentId });
  },

  getDoc(nodeId: string): Promise<OfficeDocView> {
    return get(`${API}/docs/${nodeId}`);
  },

  /** Save (autosave or explicit). Carries `baseVersionNo` — the never-
   * trust-the-client guard; a stale value rejects with `OfficeDocConflictError`
   * rather than silently overwriting a concurrent change. */
  saveDoc(
    nodeId: string,
    content: OfficeDocContentModel,
    baseVersionNo: number,
    aiEnabled?: boolean,
  ): Promise<OfficeVersion> {
    const body: Record<string, unknown> = { content, base_version_no: baseVersionNo };
    if (aiEnabled !== undefined) body.ai_enabled = aiEnabled;
    return docPutOrThrowConflict(`${API}/docs/${nodeId}`, body);
  },

  acquireLease(nodeId: string): Promise<OfficeDocLeaseState> {
    return sendJson(`${API}/docs/${nodeId}/lease`, 'POST');
  },

  async releaseLease(nodeId: string): Promise<void> {
    await fetch(`${API}/docs/${nodeId}/lease`, { method: 'DELETE', headers: authHeaders() });
  },

  /** Poll target for the lease banner — deliberately polling, not SSE (a
   * long-lived stream per open document starves gunicorn workers). */
  presence(nodeId: string): Promise<OfficeDocLeaseState> {
    return get(`${API}/docs/${nodeId}/presence`);
  },

  async runAiCapability(nodeId: string, input: RunAiCapabilityInput): Promise<OfficeAiResult> {
    const response = await fetch(`${API}/docs/${nodeId}/ai`, {
      method: 'POST',
      headers: { ...authHeaders(), 'Content-Type': 'application/json' },
      body: JSON.stringify({
        capability: input.capability,
        selection_text: input.selectionText || '',
        context_before: input.contextBefore || '',
        context_after: input.contextAfter || '',
        target_language: input.targetLanguage,
      }),
    });
    if (response.status === 403) {
      const body = (await response.json().catch(() => ({}))) as { error?: string };
      throw new OfficeAiForbiddenError(body.error);
    }
    if (response.status === 429) {
      const body = (await response.json().catch(() => ({}))) as { error?: string };
      throw new OfficeAiBudgetExceededError(body.error);
    }
    if (!response.ok) {
      const body = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(body.error || `AI request failed: ${response.status}`);
    }
    return response.json();
  },

  // -------------------------------------------------------------------
  // S147-3 toolbar bundle — insert image, insert table from file, export.
  // -------------------------------------------------------------------

  /** Upload an image into the Doc's hidden ``_assets`` folder — returns
   * the ``office-asset:`` src the editor writes into an image node. */
  uploadAsset(
    nodeId: string,
    file: File,
    onProgress?: UploadProgressHandler,
  ): Promise<OfficeDocAssetUploadResult> {
    const formData = new FormData();
    formData.append('file', file, file.name);
    return uploadMultipart(`${API}/docs/${nodeId}/assets`, formData, onProgress);
  },

  /** Fetch an inserted image's bytes as an auth-carrying Blob — a bare
   * ``<img src="office-asset:...">`` cannot load (it is not a URL scheme a
   * browser understands), so every rendered image goes through this. */
  async fetchAssetBlob(nodeId: string, assetNodeId: string): Promise<Blob> {
    const response = await fetch(`${API}/docs/${nodeId}/assets/${assetNodeId}`, {
      headers: authHeaders(),
    });
    if (!response.ok) throw new Error(`Fetching asset failed: ${response.status}`);
    return response.blob();
  },

  /** Recognise a table from an uploaded ``.csv``/``.xlsx`` — the caller
   * inserts the returned rows as a real Tiptap table node. ``format`` is
   * inferred from the filename when omitted (mirrors the backend). */
  importTableFromFile(
    nodeId: string,
    file: File,
    format?: 'csv' | 'xlsx',
  ): Promise<OfficeDocTableImportResult> {
    const formData = new FormData();
    formData.append('file', file, file.name);
    if (format) formData.append('format', format);
    return uploadMultipart(`${API}/docs/${nodeId}/table-import`, formData);
  },

  /** Print/PDF/DOCX/Markdown export — the file bytes, ready for the caller
   * to turn into a download (mirrors ``officeSheetApi.exportWorkbook``). */
  async exportDoc(
    nodeId: string,
    format: OfficeDocExportFormat,
    fallbackName: string,
  ): Promise<{ blob: Blob; filename: string }> {
    const response = await fetch(`${API}/docs/${nodeId}/export?format=${format}`, {
      method: 'POST',
      headers: authHeaders(),
    });
    if (!response.ok) {
      const body = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(body.error || `Export failed: ${response.status}`);
    }
    return { blob: await response.blob(), filename: filenameFromDisposition(response, fallbackName) };
  },
};

// ---------------------------------------------------------------------------
// S147-4 — VBWD Spreadsheets. A Sheet is an office_node of doc_type='sheet';
// the edit lease is reused VERBATIM from Docs — `officeDocApi.acquireLease`/
// `releaseLease`/`presence` above work unchanged against a Sheet's node id
// (the backend widened `OfficeDocEditorService`'s lease methods to accept
// doc_type in {text, sheet}) — one home per behaviour, no separate lease
// client here.
// ---------------------------------------------------------------------------

/** A single cell's stored/cached value on the wire — a bare JSON scalar for
 * number/text/boolean/blank, or the `{t, v}` wrapper for an error/date
 * (`sheet_content.py`'s `encode_cell_value`/`decode_cell_value`). */
export type OfficeSheetCellValue =
  | number
  | string
  | boolean
  | null
  | { t: 'error' | 'date'; v: string };

export interface OfficeSheetCellModel {
  f?: string;
  v?: OfficeSheetCellValue;
}

/** The DISPLAY-only per-cell formatting slot (S147-4 drag/toolbar slice) —
 * narrowed to exactly the fields `sheet_content.normalize_cell_style`
 * accepts on the backend; sending anything else is rejected with a 400. A
 * style update REPLACES the stored style for that cell wholesale (the
 * backend does not merge), so a caller that wants to flip one flag must
 * spread the cell's current style first. */
export type OfficeSheetNumberFormat = 'general' | 'number' | 'currency' | 'percent' | 'date';
export type OfficeSheetTextAlignment = 'left' | 'center' | 'right';

export interface OfficeSheetCellStyle {
  format?: OfficeSheetNumberFormat;
  decimals?: number;
  bold?: boolean;
  italic?: boolean;
  align?: OfficeSheetTextAlignment;
}

export interface OfficeSheetTabModel {
  name: string;
  cells: Record<string, OfficeSheetCellModel>;
  /** Present only once at least one cell on this sheet has been styled —
   * mirrors the backend's `apply_presentation`, which omits an empty map
   * rather than shipping `{}` on every sheet. */
  styles?: Record<string, OfficeSheetCellStyle>;
  /** `"A1:C1"`-style range strings — present only once at least one merge
   * exists on this sheet. */
  merges?: string[];
}

export interface OfficeSheetWorkbookModel {
  sheets: OfficeSheetTabModel[];
  active_sheet: string;
}

export interface OfficeSheetView {
  id: string;
  parent_id: string | null;
  kind: OfficeNodeKind;
  name: string;
  trashed_at: string | null;
  created_at: string;
  updated_at: string;
  document: {
    id: string;
    node_id: string;
    doc_type: OfficeDocType;
    mime_type: string;
    size_bytes: number;
    ai_enabled: boolean;
    current_version_id: string | null;
    created_at: string;
    updated_at: string;
  };
  workbook: OfficeSheetWorkbookModel;
  version_no: number;
  access: OfficeDocAccess;
  lease: OfficeDocLeaseState;
}

/** One edit sent to `PUT /sheets/<id>/cells`. Exactly one of `formula` /
 * `value` / `clear` / `fill_from` / `style` (with `address`), or `merge` /
 * `unmerge` (sheet-scoped, no `address`) — the backend rejects any other
 * combination with a 400 (`OfficeSheetEditorService._apply_change`).
 * `style: null` clears a cell's formatting. */
export interface OfficeSheetCellChange {
  sheet: string;
  address?: string;
  formula?: string;
  value?: number | string | boolean | null;
  clear?: boolean;
  /** The fill-handle drag (S147-4): copy `fill_from`'s formula/value (and
   * style) into `address`, translating relative references server-side —
   * see `OfficeSheetEditorService._apply_fill`. */
  fill_from?: string;
  style?: OfficeSheetCellStyle | null;
  /** `"A1:C1"`-style range text — replaces any merge it overlaps. */
  merge?: string;
  /** Any cell address inside the merge to remove (not necessarily its
   * top-left corner). */
  unmerge?: string;
}

export interface OfficeSheetSaveResult {
  version_no: number;
  changes: Record<string, OfficeSheetCellValue>;
}

export interface OfficeSheetImportResult {
  version_no: number;
  unmapped_formulas: { sheet: string; address: string; formula: string }[];
}

export type OfficeSheetExportFormat = 'csv' | 'xlsx' | 'pdf';

async function sheetPutOrThrowConflict(
  url: string,
  body: unknown,
): Promise<OfficeSheetSaveResult> {
  const response = await fetch(url, {
    method: 'PUT',
    headers: { ...authHeaders(), 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (response.status === 409) {
    throw new OfficeDocConflictError(await response.json());
  }
  if (!response.ok) throw new Error(`PUT ${url} failed: ${response.status}`);
  return response.json();
}

export const officeSheetApi = {
  createSheet(name: string, parentId: string | null): Promise<OfficeSheetView> {
    return sendJson(`${API}/sheets`, 'POST', { name, parent_id: parentId });
  },

  getSheet(nodeId: string): Promise<OfficeSheetView> {
    return get(`${API}/sheets/${nodeId}`);
  },

  /** `baseVersionNo` is the never-trust-the-client guard — a stale value
   * rejects with `OfficeDocConflictError` (the SAME class Docs throws;
   * one shape, one place the caller branches on `.kind`). */
  saveCells(
    nodeId: string,
    changes: OfficeSheetCellChange[],
    baseVersionNo: number,
  ): Promise<OfficeSheetSaveResult> {
    return sheetPutOrThrowConflict(`${API}/sheets/${nodeId}/cells`, {
      changes,
      base_version_no: baseVersionNo,
    });
  },

  recalc(nodeId: string): Promise<OfficeSheetSaveResult> {
    return sendJson(`${API}/sheets/${nodeId}/recalc`, 'POST');
  },

  async exportWorkbook(
    nodeId: string,
    format: OfficeSheetExportFormat,
    fallbackName: string,
  ): Promise<{ blob: Blob; filename: string }> {
    const response = await fetch(`${API}/sheets/${nodeId}/export?format=${format}`, {
      method: 'POST',
      headers: authHeaders(),
    });
    if (!response.ok) throw new Error(`Export failed: ${response.status}`);
    return { blob: await response.blob(), filename: filenameFromDisposition(response, fallbackName) };
  },

  async importWorkbook(nodeId: string, file: File): Promise<OfficeSheetImportResult> {
    const formData = new FormData();
    formData.append('file', file, file.name);
    return uploadMultipart(`${API}/sheets/${nodeId}/import`, formData);
  },
};

// ---------------------------------------------------------------------------
// S147-2 — the PUBLIC (no session) client for `/s/:token`. Every function
// here deliberately omits `authHeaders()` — a stranger with only the link
// must be able to call all of them; an optional bearer token, when a
// logged-in visitor happens to have one, is layered on top explicitly by the
// caller via `extraHeaders` rather than assumed (mirrors the backend's
// `@optional_auth`: presence is opt-in, never silently borrowed).
// ---------------------------------------------------------------------------

const PUBLIC_API = '/api/v1/office/public';

/** The 404 the backend returns for an invalid/revoked/expired/disallowed
 * token — deliberately indistinguishable from "never existed" (C2). */
export class OfficePublicShareUnavailableError extends Error {
  constructor() {
    super('This link is no longer available.');
    this.name = 'OfficePublicShareUnavailableError';
  }
}

/** The 401 the backend returns when a password-protected share's content is
 * requested without (or with an expired) unlock grant. */
export class OfficePublicPasswordRequiredError extends Error {
  constructor() {
    super('A password is required to open this link.');
    this.name = 'OfficePublicPasswordRequiredError';
  }
}

export interface OfficePublicShareDocument {
  mime_type: string;
  size_bytes: number;
}

export interface OfficePublicShareMetadata {
  name: string;
  kind: OfficeNodeKind;
  permission: OfficeSharePermission | 'owner';
  requires_password: boolean;
  document?: OfficePublicShareDocument;
}

function shareGrantHeaders(grantToken?: string | null): Record<string, string> {
  return grantToken ? { 'X-Share-Grant': grantToken } : {};
}

export const officePublicShareApi = {
  async getMetadata(token: string): Promise<OfficePublicShareMetadata> {
    const response = await fetch(`${PUBLIC_API}/${encodeURIComponent(token)}`);
    if (response.status === 404) throw new OfficePublicShareUnavailableError();
    if (!response.ok) throw new Error(`GET public share failed: ${response.status}`);
    return response.json();
  },

  /** Exchange a password for a short-lived grant token (C3) — never a
   * session; the caller re-sends it as `X-Share-Grant` on content calls. */
  async unlock(token: string, password: string): Promise<string> {
    const response = await fetch(`${PUBLIC_API}/${encodeURIComponent(token)}/unlock`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ password }),
    });
    if (response.status === 404) throw new OfficePublicShareUnavailableError();
    if (!response.ok) {
      const body = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(body.error || `Unlock failed: ${response.status}`);
    }
    const body = (await response.json()) as { grant_token: string };
    return body.grant_token;
  },

  /** Fetch the shared document's bytes as a Blob (view/comment/edit alike). */
  async download(
    token: string,
    fallbackName: string,
    grantToken?: string | null,
  ): Promise<{ blob: Blob; filename: string }> {
    const response = await fetch(`${PUBLIC_API}/${encodeURIComponent(token)}/content`, {
      headers: shareGrantHeaders(grantToken),
    });
    if (response.status === 401) throw new OfficePublicPasswordRequiredError();
    if (response.status === 404) throw new OfficePublicShareUnavailableError();
    if (!response.ok) throw new Error(`Download failed: ${response.status}`);
    return { blob: await response.blob(), filename: filenameFromDisposition(response, fallbackName) };
  },

  /** Replace the shared document's content — only ever reachable when the
   * resolved permission is `edit` (the backend 403s otherwise). */
  async replaceContent(token: string, file: File, grantToken?: string | null): Promise<void> {
    const response = await fetch(`${PUBLIC_API}/${encodeURIComponent(token)}/content`, {
      method: 'PUT',
      headers: shareGrantHeaders(grantToken),
      body: file,
    });
    if (response.status === 401) throw new OfficePublicPasswordRequiredError();
    if (response.status === 404) throw new OfficePublicShareUnavailableError();
    if (!response.ok) {
      const body = (await response.json().catch(() => ({}))) as { error?: string };
      throw new Error(body.error || `Edit failed: ${response.status}`);
    }
  },
};
