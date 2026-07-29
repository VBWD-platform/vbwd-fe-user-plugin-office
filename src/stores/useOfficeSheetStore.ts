import { defineStore } from 'pinia';
import { ref } from 'vue';
import {
  officeDocApi,
  officeSheetApi,
  OfficeDocConflictError,
  type OfficeDocLeaseState,
  type OfficeSheetCellChange,
  type OfficeSheetCellModel,
  type OfficeSheetCellStyle,
  type OfficeSheetCellValue,
  type OfficeSheetExportFormat,
  type OfficeSheetImportResult,
  type OfficeSheetNumberFormat,
  type OfficeSheetTabModel,
  type OfficeSheetTextAlignment,
  type OfficeSheetView,
  type OfficeSheetWorkbookModel,
} from '../api/officeApi';

const PRESENCE_POLL_INTERVAL_MS = 5000;
const LEASE_HEARTBEAT_INTERVAL_MS = 30000;

// Batches keystrokes into ONE `PUT .../cells` call rather than one per
// character — the server still only recalculates + returns the dirty
// sub-graph either way (requirement #5); this just avoids a request storm
// while typing. The edited cell itself is applied to the LOCAL model
// immediately (optimistic client-side echo, per the sprint's "Recalculation"
// section) so typing feels instant even before the round trip completes.
const SAVE_DEBOUNCE_MS = 400;

/** Access levels that may save cells or hold the edit lease — mirrors the
 * backend's `EDIT_CAPABLE_ACCESS` (owner or an edit-permission share). */
const EDIT_CAPABLE_ACCESS = new Set(['owner', 'edit']);

/** Mirrors the backend's `sheet_content.MAXIMUM_STYLE_DECIMAL_PLACES` — the
 * decimals +/- toolbar buttons must clamp to the same ceiling the backend
 * would otherwise reject with a 400. */
const MAXIMUM_STYLE_DECIMAL_PLACES = 10;
const DEFAULT_STYLE_DECIMALS = 2;

export type OfficeSheetConflictKind = 'stale_version' | 'locked' | null;

function pendingKey(sheet: string, address: string): string {
  return `${sheet}!${address}`;
}

/** Parse whatever the user typed into the formula bar / a cell into a
 * `PUT .../cells` change entry — `=`-prefixed is a formula, empty clears the
 * cell, a bare number parses numerically, anything else is text. Mirrors the
 * backend's own `=`-prefix convention (`parser.parse_formula`). */
export function changeFromRawInput(
  sheet: string,
  address: string,
  rawInput: string,
): OfficeSheetCellChange {
  if (rawInput.startsWith('=')) return { sheet, address, formula: rawInput };
  if (rawInput.trim() === '') return { sheet, address, clear: true };
  const asNumber = Number(rawInput);
  if (!Number.isNaN(asNumber) && rawInput.trim() !== '') {
    return { sheet, address, value: asNumber };
  }
  return { sheet, address, value: rawInput };
}

export const useOfficeSheetStore = defineStore('officeSheet', () => {
  const nodeId = ref<string | null>(null);
  const name = ref('');
  const workbook = ref<OfficeSheetWorkbookModel | null>(null);
  const versionNo = ref(0);
  const access = ref<OfficeSheetView['access'] | null>(null);
  const lease = ref<OfficeDocLeaseState | null>(null);
  const activeSheetName = ref('');

  const loading = ref(false);
  const loadError = ref<string | null>(null);
  const saving = ref(false);
  const dirty = ref(false);
  const conflict = ref<OfficeSheetConflictKind>(null);
  const importReport = ref<OfficeSheetImportResult['unmapped_formulas']>([]);

  const pendingChanges = new Map<string, OfficeSheetCellChange>();
  let saveTimer: ReturnType<typeof setTimeout> | null = null;
  let presenceTimer: ReturnType<typeof setInterval> | null = null;
  let heartbeatTimer: ReturnType<typeof setInterval> | null = null;

  function canEdit(): boolean {
    return !!access.value && EDIT_CAPABLE_ACCESS.has(access.value);
  }

  function getSheetTab(sheetName: string): OfficeSheetTabModel | undefined {
    return workbook.value?.sheets.find((tab) => tab.name === sheetName);
  }

  function getCell(sheetName: string, address: string): OfficeSheetCellModel {
    return getSheetTab(sheetName)?.cells[address] ?? {};
  }

  function getCellStyle(sheetName: string, address: string): OfficeSheetCellStyle {
    return getSheetTab(sheetName)?.styles?.[address] ?? {};
  }

  function getMerges(sheetName: string): string[] {
    return getSheetTab(sheetName)?.merges ?? [];
  }

  function applyView(view: OfficeSheetView): void {
    nodeId.value = view.id;
    name.value = view.name;
    workbook.value = view.workbook;
    versionNo.value = view.version_no;
    access.value = view.access;
    lease.value = view.lease;
    activeSheetName.value = view.workbook.active_sheet || view.workbook.sheets[0]?.name || '';
    conflict.value = null;
  }

  async function load(id: string): Promise<void> {
    loading.value = true;
    loadError.value = null;
    try {
      const view = await officeSheetApi.getSheet(id);
      applyView(view);
      dirty.value = false;
      if (canEdit()) await acquireLease();
      startPresencePolling();
      startLeaseHeartbeat();
    } catch (caught) {
      loadError.value = (caught as Error).message;
    } finally {
      loading.value = false;
    }
  }

  async function createSheet(sheetName: string, parentId: string | null): Promise<OfficeSheetView> {
    const view = await officeSheetApi.createSheet(sheetName, parentId);
    applyView(view);
    startPresencePolling();
    startLeaseHeartbeat();
    return view;
  }

  /** Optimistic local echo + queue the change for the debounced save — the
   * SERVER remains the recalculation authority (requirement #5); this is
   * purely so the edited cell shows the user's own keystrokes instantly. */
  function setCellFromRawInput(sheetName: string, address: string, rawInput: string): void {
    if (!workbook.value || !canEdit()) return;
    const change = changeFromRawInput(sheetName, address, rawInput);
    applyLocalEcho(sheetName, address, change);
    pendingChanges.set(pendingKey(sheetName, address), change);
    dirty.value = true;
    scheduleSave();
  }

  function applyLocalEcho(sheetName: string, address: string, change: OfficeSheetCellChange): void {
    const tab = getSheetTab(sheetName);
    if (!tab) return;
    if (change.clear) {
      delete tab.cells[address];
      return;
    }
    const nextCell: OfficeSheetCellModel = {};
    if (change.formula !== undefined) nextCell.f = change.formula;
    if (change.value !== undefined) nextCell.v = change.value as OfficeSheetCellValue;
    tab.cells[address] = nextCell;
  }

  /** Adds a new, empty tab — both locally (instant UI feedback) and on the
   * server, via a no-op "clear A1" change on the new sheet name (the
   * backend's `_apply_change` registers ANY referenced sheet name into the
   * workbook, even when the resulting cell stays blank — so this needs no
   * dedicated "create sheet" endpoint). Renaming/removing a tab is NOT built
   * in this slice (a deliberate MVP boundary, not an oversight). */
  function addSheetTab(sheetName: string): void {
    if (!workbook.value || !canEdit()) return;
    if (getSheetTab(sheetName)) return;
    workbook.value.sheets.push({ name: sheetName, cells: {} });
    activeSheetName.value = sheetName;
    const change: OfficeSheetCellChange = { sheet: sheetName, address: 'A1', clear: true };
    pendingChanges.set(pendingKey(sheetName, 'A1'), change);
    dirty.value = true;
    scheduleSave();
  }

  // ------------------------------------------------------------------
  // Toolbar / fill-handle presentation edits (S147-4 drag/toolbar slice).
  //
  // Cell VALUES are echoed locally for instant feedback (`applyLocalEcho`,
  // above); presentation edits are queued the same way through
  // `pendingChanges` + `scheduleSave`, but merge/unmerge/fill are NOT
  // echoed locally — the backend's overlap-resolution (merge) and formula
  // translation (fill) are non-trivial rules that already exist in exactly
  // one place (`OfficeSheetEditorService`); duplicating them here would be
  // a second, driftable copy for no real benefit, since a toolbar click is
  // an infrequent, discrete action, not a hot per-keystroke path. Instead
  // `flushSave` re-fetches the workbook after any batch that touched
  // presentation (see `changeTouchesPresentation`/`refreshPresentation`
  // below) — the server stays the single source of truth. Style toggles ARE
  // echoed locally (simple object replace, no overlap rule) for instant
  // bold/italic/format feedback.
  // ------------------------------------------------------------------

  function applyStyleUpdater(
    sheetName: string,
    addresses: string[],
    updater: (current: OfficeSheetCellStyle) => OfficeSheetCellStyle,
  ): void {
    if (!workbook.value || !canEdit()) return;
    const tab = getSheetTab(sheetName);
    if (!tab) return;
    if (!tab.styles) tab.styles = {};
    for (const address of addresses) {
      const nextStyle = updater(tab.styles[address] ?? {});
      const hasAnyField = Object.keys(nextStyle).length > 0;
      if (hasAnyField) {
        tab.styles[address] = nextStyle;
      } else {
        delete tab.styles[address];
      }
      pendingChanges.set(pendingKey(sheetName, address), {
        sheet: sheetName,
        address,
        style: hasAnyField ? nextStyle : null,
      });
    }
    dirty.value = true;
    scheduleSave();
  }

  function toggleBold(sheetName: string, addresses: string[]): void {
    const allBoldAlready = addresses.length > 0 && addresses.every(
      (address) => !!getCellStyle(sheetName, address).bold,
    );
    applyStyleUpdater(sheetName, addresses, (current) => ({ ...current, bold: !allBoldAlready }));
  }

  function toggleItalic(sheetName: string, addresses: string[]): void {
    const allItalicAlready = addresses.length > 0 && addresses.every(
      (address) => !!getCellStyle(sheetName, address).italic,
    );
    applyStyleUpdater(sheetName, addresses, (current) => ({ ...current, italic: !allItalicAlready }));
  }

  function setAlignment(sheetName: string, addresses: string[], align: OfficeSheetTextAlignment): void {
    applyStyleUpdater(sheetName, addresses, (current) => ({ ...current, align }));
  }

  /** Switching to number/currency seeds a default decimal count (matching
   * Excel/Sheets' own first-click behaviour); switching to general/percent/
   * date drops any previously-set decimals rather than carrying a stale
   * value across an unrelated format. */
  function setNumberFormat(sheetName: string, addresses: string[], format: OfficeSheetNumberFormat): void {
    applyStyleUpdater(sheetName, addresses, (current) => {
      const next: OfficeSheetCellStyle = { ...current, format };
      if (format === 'number' || format === 'currency') {
        next.decimals = current.decimals ?? DEFAULT_STYLE_DECIMALS;
      } else {
        delete next.decimals;
      }
      return next;
    });
  }

  function adjustDecimals(sheetName: string, addresses: string[], delta: number): void {
    applyStyleUpdater(sheetName, addresses, (current) => {
      const baseline = current.decimals ?? DEFAULT_STYLE_DECIMALS;
      const nextDecimals = Math.min(
        MAXIMUM_STYLE_DECIMAL_PLACES,
        Math.max(0, baseline + delta),
      );
      return { ...current, decimals: nextDecimals };
    });
  }

  function mergeRange(sheetName: string, rangeText: string): void {
    if (!nodeId.value || !canEdit()) return;
    pendingChanges.set(pendingKey(sheetName, `merge:${rangeText}`), { sheet: sheetName, merge: rangeText });
    dirty.value = true;
    scheduleSave();
  }

  function unmergeAt(sheetName: string, addressText: string): void {
    if (!nodeId.value || !canEdit()) return;
    pendingChanges.set(pendingKey(sheetName, `unmerge:${addressText}`), {
      sheet: sheetName,
      unmerge: addressText,
    });
    dirty.value = true;
    scheduleSave();
  }

  /** The fill-handle drag: copy ONE source cell's formula/value/style into
   * every target address, via the backend's `fill_from` operation (never
   * re-derived client-side — see the section docstring above). */
  function fillCells(sheetName: string, sourceAddress: string, targetAddresses: string[]): void {
    if (!nodeId.value || !canEdit()) return;
    for (const targetAddress of targetAddresses) {
      pendingChanges.set(pendingKey(sheetName, targetAddress), {
        sheet: sheetName,
        address: targetAddress,
        fill_from: sourceAddress,
      });
    }
    dirty.value = true;
    scheduleSave();
  }

  function changeTouchesPresentation(change: OfficeSheetCellChange): boolean {
    return (
      change.style !== undefined ||
      change.merge !== undefined ||
      change.unmerge !== undefined ||
      change.fill_from !== undefined
    );
  }

  /** Re-fetch the workbook (cells + styles + merges) after a save batch that
   * touched presentation — cheaper than a full `load()` (no lease
   * re-acquire, no polling restart) and the simplest correct way to pick up
   * server-computed state (merge overlap resolution, fill-handle style
   * copy) this store deliberately does not replicate locally. */
  async function refreshPresentation(): Promise<void> {
    if (!nodeId.value) return;
    const view = await officeSheetApi.getSheet(nodeId.value);
    workbook.value = view.workbook;
    versionNo.value = view.version_no;
  }

  function scheduleSave(): void {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(flushSave, SAVE_DEBOUNCE_MS);
  }

  async function flushSave(): Promise<boolean> {
    if (!nodeId.value || pendingChanges.size === 0) return true;
    const changes = Array.from(pendingChanges.values());
    const touchesPresentation = changes.some(changeTouchesPresentation);
    pendingChanges.clear();
    saving.value = true;
    try {
      const result = await officeSheetApi.saveCells(nodeId.value, changes, versionNo.value);
      applySaveResult(result);
      if (touchesPresentation) await refreshPresentation();
      dirty.value = pendingChanges.size > 0;
      conflict.value = null;
      return true;
    } catch (caught) {
      if (caught instanceof OfficeDocConflictError) {
        conflict.value = caught.kind;
        if (caught.kind === 'locked' && lease.value) {
          lease.value = {
            ...lease.value,
            held: true,
            is_self: false,
            granted: false,
            holder_user_id: caught.holderUserId ?? lease.value.holder_user_id,
            expires_at: caught.expiresAt ?? lease.value.expires_at,
          };
        }
        return false;
      }
      throw caught;
    } finally {
      saving.value = false;
    }
  }

  /** Merge the server's recalculated deltas (`"Sheet1!B1": value`) back into
   * the local model — ONLY the changed cells arrive (requirement #5), so
   * everything else in the grid is left exactly as the user sees it. */
  function applySaveResult(result: { version_no: number; changes: Record<string, OfficeSheetCellValue> }): void {
    versionNo.value = result.version_no;
    for (const [qualifiedAddress, value] of Object.entries(result.changes)) {
      const separatorIndex = qualifiedAddress.indexOf('!');
      if (separatorIndex < 0) continue;
      const sheetName = qualifiedAddress.slice(0, separatorIndex);
      const address = qualifiedAddress.slice(separatorIndex + 1);
      const tab = getSheetTab(sheetName);
      if (!tab) continue;
      const existing = tab.cells[address] ?? {};
      tab.cells[address] = { ...existing, v: value };
    }
  }

  async function recalcAll(): Promise<void> {
    if (!nodeId.value) return;
    const result = await officeSheetApi.recalc(nodeId.value);
    applySaveResult(result);
  }

  async function exportWorkbook(format: OfficeSheetExportFormat): Promise<{ blob: Blob; filename: string }> {
    if (!nodeId.value) throw new Error('No sheet is open');
    return officeSheetApi.exportWorkbook(nodeId.value, format, `${name.value}.${format}`);
  }

  async function importWorkbook(file: File): Promise<void> {
    if (!nodeId.value) return;
    const result = await officeSheetApi.importWorkbook(nodeId.value, file);
    importReport.value = result.unmapped_formulas;
    await load(nodeId.value);
  }

  function clearImportReport(): void {
    importReport.value = [];
  }

  async function acquireLease(): Promise<void> {
    if (!nodeId.value) return;
    try {
      lease.value = await officeDocApi.acquireLease(nodeId.value);
      if (lease.value.granted) conflict.value = null;
    } catch {
      // A failed heartbeat should not crash the grid — the next poll retries.
    }
  }

  async function releaseLease(): Promise<void> {
    if (!nodeId.value) return;
    try {
      await officeDocApi.releaseLease(nodeId.value);
    } catch {
      // Best-effort — the lease also self-expires.
    }
  }

  async function refreshPresence(): Promise<void> {
    if (!nodeId.value) return;
    try {
      lease.value = await officeDocApi.presence(nodeId.value);
    } catch {
      // A missed poll is not a product failure — the next tick retries.
    }
  }

  function startPresencePolling(): void {
    stopPresencePolling();
    presenceTimer = setInterval(refreshPresence, PRESENCE_POLL_INTERVAL_MS);
  }

  function stopPresencePolling(): void {
    if (presenceTimer) clearInterval(presenceTimer);
    presenceTimer = null;
  }

  function startLeaseHeartbeat(): void {
    stopLeaseHeartbeat();
    heartbeatTimer = setInterval(() => {
      if (canEdit() && lease.value?.is_self) acquireLease();
    }, LEASE_HEARTBEAT_INTERVAL_MS);
  }

  function stopLeaseHeartbeat(): void {
    if (heartbeatTimer) clearInterval(heartbeatTimer);
    heartbeatTimer = null;
  }

  function stopPolling(): void {
    stopPresencePolling();
    stopLeaseHeartbeat();
  }

  async function takeover(): Promise<boolean> {
    await acquireLease();
    return !!lease.value?.granted;
  }

  function reset(): void {
    stopPolling();
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = null;
    pendingChanges.clear();
    nodeId.value = null;
    name.value = '';
    workbook.value = null;
    versionNo.value = 0;
    access.value = null;
    lease.value = null;
    activeSheetName.value = '';
    loading.value = false;
    loadError.value = null;
    saving.value = false;
    dirty.value = false;
    conflict.value = null;
    importReport.value = [];
  }

  return {
    nodeId,
    name,
    workbook,
    versionNo,
    access,
    lease,
    activeSheetName,
    loading,
    loadError,
    saving,
    dirty,
    conflict,
    importReport,
    canEdit,
    getSheetTab,
    getCell,
    getCellStyle,
    getMerges,
    load,
    createSheet,
    setCellFromRawInput,
    addSheetTab,
    toggleBold,
    toggleItalic,
    setAlignment,
    setNumberFormat,
    adjustDecimals,
    mergeRange,
    unmergeAt,
    fillCells,
    flushSave,
    recalcAll,
    exportWorkbook,
    importWorkbook,
    clearImportReport,
    acquireLease,
    releaseLease,
    refreshPresence,
    startPresencePolling,
    stopPolling,
    takeover,
    reset,
  };
});
