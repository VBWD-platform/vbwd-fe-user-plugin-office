<template>
  <div
    class="office-sheet-editor"
    data-testid="office-sheet-editor"
  >
    <header class="office-sheet-header">
      <router-link
        to="/dashboard/office"
        class="office-sheet-back"
        data-testid="office-sheet-back"
      >
        {{ $t('office.sheet.backToFiles') }}
      </router-link>
      <h2 class="office-sheet-name">
        {{ store.name }}
      </h2>
      <div class="office-sheet-header-actions">
        <span
          class="office-sheet-save-status"
          data-testid="office-sheet-save-status"
        >{{ $t(saveStatusLabel) }}</span>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          data-testid="office-sheet-recalc-btn"
          :disabled="!store.canEdit()"
          @click="onRecalc"
        >
          {{ $t('office.sheet.recalc') }}
        </button>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          data-testid="office-sheet-export-csv-btn"
          @click="onExport('csv')"
        >
          {{ $t('office.sheet.exportCsv') }}
        </button>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          data-testid="office-sheet-export-xlsx-btn"
          @click="onExport('xlsx')"
        >
          {{ $t('office.sheet.exportXlsx') }}
        </button>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          data-testid="office-sheet-import-btn"
          :disabled="!store.canEdit()"
          @click="triggerImportInput"
        >
          {{ $t('office.sheet.import') }}
        </button>
        <input
          ref="importInputRef"
          type="file"
          accept=".csv,.xlsx"
          class="office-sheet-hidden-input"
          data-testid="office-sheet-import-input"
          @change="onImportInputChange"
        >
      </div>
    </header>

    <div
      v-if="store.loading"
      class="office-sheet-loading"
      data-testid="office-sheet-loading"
    >
      {{ $t('office.loading') }}
    </div>

    <div
      v-else-if="store.loadError"
      class="office-sheet-error"
      data-testid="office-sheet-load-error"
    >
      {{ store.loadError }}
    </div>

    <template v-else>
      <div
        v-if="importReportEntries.length > 0"
        class="office-sheet-banner office-sheet-banner--warning"
        data-testid="office-sheet-import-report"
      >
        <div>
          <strong>{{ $t('office.sheet.importReportTitle') }}</strong>
          <p>{{ $t('office.sheet.importReportHint') }}</p>
          <ul>
            <li
              v-for="entry in importReportEntries"
              :key="`${entry.sheet}!${entry.address}`"
            >
              {{ entry.sheet }}!{{ entry.address }}: {{ entry.formula }}
            </li>
          </ul>
        </div>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--sm"
          data-testid="office-sheet-import-report-dismiss"
          @click="store.clearImportReport()"
        >
          {{ $t('office.versions.close') }}
        </button>
      </div>

      <div
        v-if="!store.canEdit()"
        class="office-sheet-banner office-sheet-banner--info"
        data-testid="office-sheet-view-only-banner"
      >
        {{ $t('office.sheet.viewOnlyBanner') }}
      </div>

      <div
        v-if="showLeaseBanner"
        class="office-sheet-banner office-sheet-banner--warning"
        data-testid="office-sheet-lease-banner"
      >
        <span>{{ $t('office.sheet.lease.otherEditing') }}</span>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--sm vbwd-btn--primary"
          data-testid="office-sheet-lease-takeover"
          @click="onTakeover"
        >
          {{ $t('office.sheet.lease.takeOver') }}
        </button>
      </div>

      <div
        v-if="store.conflict === 'stale_version'"
        class="office-sheet-banner office-sheet-banner--warning"
        data-testid="office-sheet-conflict-banner"
      >
        <span>{{ $t('office.sheet.conflict.staleVersion') }}</span>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--sm vbwd-btn--primary"
          data-testid="office-sheet-reload"
          @click="onReload"
        >
          {{ $t('office.sheet.reload') }}
        </button>
      </div>

      <div
        class="office-sheet-formula-bar"
        data-testid="office-sheet-formula-bar"
      >
        <span
          class="office-sheet-formula-bar-address"
          data-testid="office-sheet-formula-bar-address"
        >{{ activeAddressLabel }}</span>
        <input
          v-model="formulaBarValue"
          type="text"
          class="vbwd-input office-sheet-formula-bar-input"
          data-testid="office-sheet-formula-bar-input"
          :placeholder="$t('office.sheet.formulaBarPlaceholder')"
          :disabled="!activeCell || !store.canEdit()"
          @keydown.enter="commitFormulaBar"
        >
      </div>

      <div
        class="office-sheet-tabs"
        data-testid="office-sheet-tabs"
      >
        <button
          v-for="tab in tabs"
          :key="tab.name"
          type="button"
          class="office-sheet-tab"
          :class="{ 'is-active': tab.name === store.activeSheetName }"
          data-testid="office-sheet-tab"
          @click="onSelectTab(tab.name)"
        >
          {{ tab.name }}
        </button>
        <button
          v-if="store.canEdit()"
          type="button"
          class="office-sheet-tab office-sheet-tab--add"
          data-testid="office-sheet-add-sheet-btn"
          @click="onAddSheet"
        >
          {{ $t('office.sheet.addSheet') }}
        </button>
      </div>

      <div
        ref="scrollRef"
        class="office-sheet-grid"
        data-testid="office-sheet-grid"
        @scroll="onScroll"
      >
        <div
          class="office-sheet-col-header-row"
          :style="{ width: totalGridWidthPx }"
        >
          <div class="office-sheet-corner-header" />
          <div
            v-for="column in totalColumns"
            :key="column"
            class="office-sheet-col-header"
            data-testid="office-sheet-col-header"
            :style="{ width: `${columnWidth}px` }"
          >
            {{ columnLetters(column) }}
          </div>
        </div>

        <div
          class="office-sheet-body-spacer"
          :style="{ height: `${totalRows * rowHeight}px`, width: totalGridWidthPx }"
        >
          <div
            class="office-sheet-visible-rows"
            :style="{ transform: `translateY(${(firstVisibleRow - 1) * rowHeight}px)` }"
          >
            <div
              v-for="row in visibleRows"
              :key="row"
              class="office-sheet-row"
              :style="{ height: `${rowHeight}px` }"
            >
              <div
                class="office-sheet-row-header"
                data-testid="office-sheet-row-header"
                :style="{ width: `${rowHeaderWidth}px` }"
              >
                {{ row }}
              </div>
              <div
                v-for="column in totalColumns"
                :key="column"
                class="office-sheet-cell"
                data-testid="office-sheet-cell"
                :class="{ 'is-selected': isSelected(row, column) }"
                :style="{ width: `${columnWidth}px` }"
                @click="onSelectCell(row, column)"
                @dblclick="onStartInlineEdit(row, column)"
              >
                <input
                  v-if="isEditingInline(row, column)"
                  ref="inlineInputRef"
                  v-model="editingText"
                  type="text"
                  class="office-sheet-cell-input"
                  data-testid="office-sheet-cell-input"
                  @keydown.enter="commitInlineEdit"
                  @blur="commitInlineEdit"
                >
                <span
                  v-else
                  class="office-sheet-cell-display"
                >{{ displayValue(row, column) }}</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref } from 'vue';
import { useRoute } from 'vue-router';
import { useOfficeSheetStore } from '../stores/useOfficeSheetStore';
import { columnIndexToLetters, formatA1 } from '../utils/sheetAddress';
import type { OfficeSheetCellModel, OfficeSheetExportFormat } from '../api/officeApi';

// A virtualised, scrollable canvas of a fixed size — NOT the ceiling
// (`sheet_max_rows`/`sheet_max_columns`, up to 100,000 x 1,000 server-side).
// Rendering the full ceiling as DOM rows would defeat the point of
// virtualisation; a fresh spreadsheet application shows a bounded default
// canvas too (Excel/Sheets simply grow it lazily). Only ROWS are virtualised
// here — 26 columns is cheap to render in full and is the deliberately
// narrower MVP scope (documented, not an oversight).
const TOTAL_ROWS = 500;
const TOTAL_COLUMNS = 26;
const ROW_HEIGHT_PX = 24;
const COLUMN_WIDTH_PX = 100;
const ROW_HEADER_WIDTH_PX = 40;
const VIRTUALIZATION_BUFFER_ROWS = 4;
const DEFAULT_VIEWPORT_HEIGHT_PX = 420;

const route = useRoute();
const store = useOfficeSheetStore();

const totalRows = TOTAL_ROWS;
const totalColumns = TOTAL_COLUMNS;
const rowHeight = ROW_HEIGHT_PX;
const columnWidth = COLUMN_WIDTH_PX;
const rowHeaderWidth = ROW_HEADER_WIDTH_PX;

const nodeId = computed(() => String(route.params.id || ''));

const scrollRef = ref<HTMLElement | null>(null);
const scrollTop = ref(0);
const viewportHeight = ref(DEFAULT_VIEWPORT_HEIGHT_PX);

const activeCell = ref<{ row: number; column: number } | null>(null);
const formulaBarValue = ref('');
const editingCell = ref<{ row: number; column: number } | null>(null);
const editingText = ref('');
const inlineInputRef = ref<HTMLInputElement[] | HTMLInputElement | null>(null);
const importInputRef = ref<HTMLInputElement | null>(null);

const totalGridWidthPx = computed(() => `${rowHeaderWidth + totalColumns * columnWidth}px`);

const tabs = computed(() => store.workbook?.sheets ?? []);

const importReportEntries = computed(() => store.importReport);

const showLeaseBanner = computed(
  () => store.canEdit() && !!store.lease?.held && !store.lease?.is_self,
);

const saveStatusLabel = computed(() => {
  if (store.saving) return 'office.sheet.saving';
  if (store.conflict || store.dirty) return 'office.sheet.unsavedChanges';
  return 'office.sheet.saved';
});

const firstVisibleRow = computed(() => {
  const rawStart = Math.floor(scrollTop.value / rowHeight) + 1 - VIRTUALIZATION_BUFFER_ROWS;
  return Math.max(1, rawStart);
});

const lastVisibleRow = computed(() => {
  const visibleCount = Math.ceil(viewportHeight.value / rowHeight) + VIRTUALIZATION_BUFFER_ROWS * 2;
  return Math.min(totalRows, firstVisibleRow.value + visibleCount);
});

const visibleRows = computed(() => {
  const rows: number[] = [];
  for (let row = firstVisibleRow.value; row <= lastVisibleRow.value; row += 1) rows.push(row);
  return rows;
});

const activeAddressLabel = computed(() =>
  activeCell.value ? formatA1(activeCell.value.column, activeCell.value.row) : '',
);

function columnLetters(column: number): string {
  return columnIndexToLetters(column);
}

function isSelected(row: number, column: number): boolean {
  return activeCell.value?.row === row && activeCell.value?.column === column;
}

function isEditingInline(row: number, column: number): boolean {
  return editingCell.value?.row === row && editingCell.value?.column === column;
}

function cellModelAt(row: number, column: number): OfficeSheetCellModel {
  return store.getCell(store.activeSheetName, formatA1(column, row));
}

/** The value the user would see typed back in — a formula's own source, or
 * the literal value re-rendered as editable text (never the FORMATTED
 * display text, which may not round-trip, e.g. a thousands separator). */
function rawTextFor(row: number, column: number): string {
  const cell = cellModelAt(row, column);
  if (cell.f) return cell.f;
  return literalToRawText(cell.v);
}

function literalToRawText(value: OfficeSheetCellModel['v']): string {
  if (value === undefined || value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'object') return value.v;
  return String(value);
}

/** Number/date/error formatting for DISPLAY only — editing always goes
 * through the raw source (see `rawTextFor`). */
function displayValue(row: number, column: number): string {
  const cell = cellModelAt(row, column);
  const value = cell.v;
  if (value === undefined || value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') return Number.isInteger(value) ? String(value) : value.toLocaleString();
  if (typeof value === 'object') {
    if (value.t === 'date') return new Date(value.v).toLocaleDateString();
    return value.v; // error code, e.g. "#NAME?"
  }
  return value;
}

function onSelectCell(row: number, column: number): void {
  if (editingCell.value && !isEditingInline(row, column)) commitInlineEdit();
  activeCell.value = { row, column };
  formulaBarValue.value = rawTextFor(row, column);
}

function onStartInlineEdit(row: number, column: number): void {
  if (!store.canEdit()) return;
  activeCell.value = { row, column };
  editingCell.value = { row, column };
  editingText.value = rawTextFor(row, column);
  nextTick(() => {
    const input = Array.isArray(inlineInputRef.value) ? inlineInputRef.value[0] : inlineInputRef.value;
    input?.focus();
  });
}

function commitInlineEdit(): void {
  const cell = editingCell.value;
  editingCell.value = null;
  if (!cell) return;
  store.setCellFromRawInput(store.activeSheetName, formatA1(cell.column, cell.row), editingText.value);
  if (activeCell.value?.row === cell.row && activeCell.value?.column === cell.column) {
    formulaBarValue.value = rawTextFor(cell.row, cell.column);
  }
}

function commitFormulaBar(): void {
  if (!activeCell.value) return;
  const { row, column } = activeCell.value;
  store.setCellFromRawInput(store.activeSheetName, formatA1(column, row), formulaBarValue.value);
}

function onSelectTab(sheetName: string): void {
  store.activeSheetName = sheetName;
  activeCell.value = null;
  formulaBarValue.value = '';
}

function onAddSheet(): void {
  // eslint-disable-next-line no-alert
  const sheetName = window.prompt('Sheet name', `Sheet${tabs.value.length + 1}`);
  if (!sheetName) return;
  store.addSheetTab(sheetName);
}

async function onRecalc(): Promise<void> {
  await store.recalcAll();
}

async function onExport(format: OfficeSheetExportFormat): Promise<void> {
  const { blob, filename } = await store.exportWorkbook(format);
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function triggerImportInput(): void {
  importInputRef.value?.click();
}

async function onImportInputChange(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (file) await store.importWorkbook(file);
}

async function onTakeover(): Promise<void> {
  await store.takeover();
}

async function onReload(): Promise<void> {
  await store.load(nodeId.value);
}

function onScroll(): void {
  if (scrollRef.value) scrollTop.value = scrollRef.value.scrollTop;
}

onMounted(async () => {
  await store.load(nodeId.value);
  if (scrollRef.value) viewportHeight.value = scrollRef.value.clientHeight || DEFAULT_VIEWPORT_HEIGHT_PX;
});

onBeforeUnmount(() => {
  store.flushSave();
  store.releaseLease();
  store.stopPolling();
  store.reset();
});
</script>

<style scoped>
.office-sheet-editor {
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 480px;
  background: var(--vbwd-color-surface, #fff);
  border-radius: var(--vbwd-card-radius, 8px);
  overflow: hidden;
}
.office-sheet-header {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 16px;
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
  flex-wrap: wrap;
}
.office-sheet-back {
  color: var(--vbwd-color-primary, #3498db);
  font-size: 0.85rem;
}
.office-sheet-name {
  flex: 1;
  margin: 0;
  font-size: 1rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-sheet-header-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}
.office-sheet-save-status {
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-sheet-hidden-input {
  display: none;
}
.office-sheet-loading,
.office-sheet-error {
  padding: 40px;
  text-align: center;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-sheet-banner {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 8px 16px;
  font-size: 0.85rem;
}
.office-sheet-banner--info {
  background: rgba(52, 152, 219, 0.08);
  color: var(--vbwd-color-primary, #3498db);
}
.office-sheet-banner--warning {
  background: rgba(243, 156, 18, 0.12);
  color: #b9770e;
}
.office-sheet-formula-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 16px;
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
}
.office-sheet-formula-bar-address {
  min-width: 48px;
  font-family: monospace;
  font-size: 0.8rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-sheet-formula-bar-input {
  flex: 1;
}
.office-sheet-tabs {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 6px 16px;
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
  overflow-x: auto;
}
.office-sheet-tab {
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: 4px;
  background: transparent;
  padding: 4px 10px;
  font-size: 0.8rem;
  cursor: pointer;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-sheet-tab.is-active {
  background: var(--vbwd-color-primary, #3498db);
  color: #fff;
  border-color: var(--vbwd-color-primary, #3498db);
}
.office-sheet-tab--add {
  color: var(--vbwd-color-primary, #3498db);
  border-style: dashed;
}
.office-sheet-grid {
  flex: 1;
  overflow: auto;
  position: relative;
}
.office-sheet-col-header-row {
  display: flex;
  position: sticky;
  top: 0;
  z-index: 2;
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-sheet-corner-header {
  width: 40px;
  flex-shrink: 0;
  border-right: 1px solid var(--vbwd-color-border, #e9ecef);
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
}
.office-sheet-col-header {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 0.75rem;
  font-weight: 600;
  color: var(--vbwd-color-text-secondary, #666);
  border-right: 1px solid var(--vbwd-color-border, #e9ecef);
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
  height: 24px;
}
.office-sheet-body-spacer {
  position: relative;
}
.office-sheet-visible-rows {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
}
.office-sheet-row {
  display: flex;
}
.office-sheet-row-header {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 0.7rem;
  color: var(--vbwd-color-text-secondary, #666);
  background: var(--vbwd-bg-hover, #f5f6f7);
  border-right: 1px solid var(--vbwd-color-border, #e9ecef);
  border-bottom: 1px solid var(--vbwd-border-light, #eee);
}
.office-sheet-cell {
  flex-shrink: 0;
  height: 100%;
  display: flex;
  align-items: center;
  padding: 0 6px;
  font-size: 0.8rem;
  border-right: 1px solid var(--vbwd-border-light, #eee);
  border-bottom: 1px solid var(--vbwd-border-light, #eee);
  cursor: cell;
  overflow: hidden;
}
.office-sheet-cell.is-selected {
  outline: 2px solid var(--vbwd-color-primary, #3498db);
  outline-offset: -2px;
}
.office-sheet-cell-display {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  width: 100%;
}
.office-sheet-cell-input {
  width: 100%;
  height: 100%;
  border: none;
  outline: none;
  font-size: 0.8rem;
  padding: 0;
  background: #fff;
}
</style>
