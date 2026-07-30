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
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          data-testid="office-sheet-ai-toggle-panel"
          @click="showAiSidebar = !showAiSidebar"
        >
          {{ $t('office.sheet.ai.title') }}
        </button>
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

      <div class="office-sheet-workspace">
        <div class="office-sheet-main-column">
          <div
            class="office-sheet-formula-bar"
            data-testid="office-sheet-formula-bar"
          >
            <span
              class="office-sheet-formula-bar-address"
              data-testid="office-sheet-formula-bar-address"
            >{{ activeAddressLabel }}</span>
            <input
              ref="formulaBarInputRef"
              v-model="formulaBarValue"
              type="text"
              class="vbwd-input office-sheet-formula-bar-input"
              data-testid="office-sheet-formula-bar-input"
              :placeholder="$t('office.sheet.formulaBarPlaceholder')"
              :disabled="!activeCell || !store.canEdit()"
              @focus="formulaBarFocused = true"
              @blur="formulaBarFocused = false"
              @keydown.enter="commitFormulaBar"
              @keydown.escape="cancelFormulaBarEdit"
            >
          </div>

          <div
            class="office-sheet-toolbar"
            data-testid="office-sheet-toolbar"
          >
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              :class="{ 'is-active': isBoldActive }"
              data-testid="office-sheet-tb-bold"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onToggleBold"
            >
              <strong>B</strong>
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              :class="{ 'is-active': isItalicActive }"
              data-testid="office-sheet-tb-italic"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onToggleItalic"
            >
              <em>I</em>
            </button>

            <span class="office-sheet-toolbar-separator" />

            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-align-left"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetAlign('left')"
            >
              {{ $t('office.sheet.toolbar.alignLeft') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-align-center"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetAlign('center')"
            >
              {{ $t('office.sheet.toolbar.alignCenter') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-align-right"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetAlign('right')"
            >
              {{ $t('office.sheet.toolbar.alignRight') }}
            </button>

            <span class="office-sheet-toolbar-separator" />

            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-format-general"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetFormat('general')"
            >
              {{ $t('office.sheet.toolbar.formatGeneral') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-format-number"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetFormat('number')"
            >
              {{ $t('office.sheet.toolbar.formatNumber') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-format-currency"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetFormat('currency')"
            >
              {{ $t('office.sheet.toolbar.formatCurrency') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-format-percent"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetFormat('percent')"
            >
              {{ $t('office.sheet.toolbar.formatPercent') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-format-date"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onSetFormat('date')"
            >
              {{ $t('office.sheet.toolbar.formatDate') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-decimals-dec"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onAdjustDecimals(-1)"
            >
              .0
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-decimals-inc"
              :disabled="!store.canEdit() || selectionAddresses.length === 0"
              @click="onAdjustDecimals(1)"
            >
              .00
            </button>

            <span class="office-sheet-toolbar-separator" />

            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-merge"
              :disabled="!store.canEdit() || selectionAddresses.length < 2"
              @click="onMerge"
            >
              {{ $t('office.sheet.toolbar.merge') }}
            </button>
            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-unmerge"
              :disabled="!store.canEdit() || !activeCell"
              @click="onUnmerge"
            >
              {{ $t('office.sheet.toolbar.unmerge') }}
            </button>

            <span class="office-sheet-toolbar-separator" />

            <button
              type="button"
              class="office-sheet-toolbar-btn"
              data-testid="office-sheet-tb-print"
              @click="onPrint"
            >
              {{ $t('office.sheet.toolbar.print') }}
            </button>

            <div class="office-sheet-export-menu">
              <button
                type="button"
                class="office-sheet-toolbar-btn"
                data-testid="office-sheet-tb-export"
                @click="exportMenuOpen = !exportMenuOpen"
              >
                {{ $t('office.sheet.toolbar.export') }}
              </button>
              <div
                v-if="exportMenuOpen"
                class="office-sheet-export-menu-list"
                data-testid="office-sheet-tb-export-menu"
              >
                <button
                  type="button"
                  class="office-sheet-export-menu-item"
                  data-testid="office-sheet-tb-export-csv"
                  @click="onExportFormat('csv')"
                >
                  {{ $t('office.sheet.exportCsv') }}
                </button>
                <button
                  type="button"
                  class="office-sheet-export-menu-item"
                  data-testid="office-sheet-tb-export-xlsx"
                  @click="onExportFormat('xlsx')"
                >
                  {{ $t('office.sheet.exportXlsx') }}
                </button>
                <button
                  type="button"
                  class="office-sheet-export-menu-item"
                  data-testid="office-sheet-tb-export-pdf"
                  @click="onExportFormat('pdf')"
                >
                  {{ $t('office.sheet.toolbar.exportPdf') }}
                </button>
              </div>
            </div>
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
              @click="showAddSheetDialog = true"
            >
              {{ $t('office.sheet.addSheet') }}
            </button>
          </div>

          <div
            ref="scrollRef"
            class="office-sheet-grid"
            data-testid="office-sheet-grid"
            tabindex="0"
            @scroll="onScroll"
            @keydown="onGridKeydown"
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
                    :class="{ 'is-selected': isSelected(row, column), 'is-merged-away': isMergedAway(row, column) }"
                    :style="{ width: `${columnWidth}px`, ...cellDisplayStyle(row, column) }"
                    @mousedown="onCellMouseDown(row, column, $event)"
                    @mouseenter="onCellMouseEnter(row, column)"
                    @dblclick="onStartInlineEdit(row, column)"
                  >
                    <input
                      v-if="isEditingInline(row, column)"
                      ref="inlineInputRef"
                      v-model="editingText"
                      type="text"
                      class="office-sheet-cell-input"
                      data-testid="office-sheet-cell-input"
                      @keydown.enter.stop="commitInlineEditAndMove('down')"
                      @keydown.tab.stop.prevent="commitInlineEditAndMove('right')"
                      @keydown.escape.stop="cancelInlineEdit"
                      @blur="commitInlineEdit"
                    >
                    <span
                      v-else-if="!isMergedAway(row, column)"
                      class="office-sheet-cell-display"
                    >{{ displayValue(row, column) }}</span>
                  </div>
                </div>
              </div>

              <div
                v-for="overlay in mergedRangeOverlays"
                :key="`merge-${overlay.range.startRow}-${overlay.range.startColumn}`"
                class="office-sheet-merge-overlay"
                :style="overlay.style"
              >
                {{ overlay.text }}
              </div>

              <div
                v-if="selectionOverlayStyle"
                class="office-sheet-selection"
                data-testid="office-sheet-selection"
                :style="selectionOverlayStyle"
              />

              <div
                v-if="formulaRangeHighlightStyle"
                class="office-sheet-range-highlight"
                data-testid="office-sheet-range-highlight"
                :style="formulaRangeHighlightStyle"
              />

              <div
                v-if="fillPreviewOverlayStyle"
                class="office-sheet-fill-preview"
                :style="fillPreviewOverlayStyle"
              />

              <div
                v-if="fillHandleStyle && store.canEdit()"
                class="office-sheet-fill-handle"
                data-testid="office-sheet-fill-handle"
                :style="fillHandleStyle"
                @mousedown="onFillHandleMouseDown"
              />
            </div>
          </div>
        </div>

        <OfficeSheetAiSidebar
          v-if="showAiSidebar"
          :ai-enabled="store.aiEnabled"
          :can-toggle="store.access === 'owner'"
          :can-propose-formula="store.canEdit()"
          :has-active-cell="!!activeCell"
          :running="store.aiRunning"
          :error="store.aiError"
          :proposal="store.aiProposal"
          @toggle-enabled="onToggleAi"
          @run-capability="onRunAiCapability"
          @accept="onAcceptAiProposal"
          @discard="store.clearAiProposal()"
        />
      </div>
    </template>

    <OfficeInputDialog
      v-if="showAddSheetDialog"
      :title="$t('office.sheet.addSheet')"
      :placeholder="$t('office.sheet.addSheetPrompt')"
      :confirm-label="$t('office.sheet.addSheetConfirm')"
      :cancel-label="$t('office.sheet.addSheetCancel')"
      :initial-value="defaultNewSheetName"
      testid-prefix="office-sheet-add-sheet"
      @confirm="onAddSheetDialogConfirm"
      @cancel="showAddSheetDialog = false"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref } from 'vue';
import { useRoute } from 'vue-router';
import { useOfficeSheetStore } from '../stores/useOfficeSheetStore';
import {
  addressesInRange,
  columnIndexToLetters,
  findContainingRange,
  formatA1,
  normalizeRange,
  parseRangeText,
  rangeReferenceFor,
  type SheetCellAddress,
  type SheetCellRange,
} from '../utils/sheetAddress';
import {
  applyRangeInsertion,
  beginRangeInsertion,
  isFormulaInput,
  type FormulaRangeDragState,
} from '../utils/formulaRangeInsertion';
import { formatNumericCellValue } from '../utils/sheetCellFormat';
import OfficeInputDialog from '../components/OfficeInputDialog.vue';
import OfficeSheetAiSidebar from '../components/OfficeSheetAiSidebar.vue';
import type {
  OfficeSheetAiCapability,
  OfficeSheetCellModel,
  OfficeSheetExportFormat,
  OfficeSheetNumberFormat,
  OfficeSheetTextAlignment,
} from '../api/officeApi';

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
const FILL_HANDLE_SIZE_PX = 8;

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

// -- Selection / active cell -------------------------------------------------
const activeCell = ref<SheetCellAddress | null>(null);
const selectionAnchor = ref<SheetCellAddress | null>(null);
const formulaBarValue = ref('');
const formulaBarFocused = ref(false);
const formulaBarInputRef = ref<HTMLInputElement | null>(null);
const editingCell = ref<SheetCellAddress | null>(null);
const editingText = ref('');
const inlineInputRef = ref<HTMLInputElement[] | HTMLInputElement | null>(null);
const importInputRef = ref<HTMLInputElement | null>(null);
const exportMenuOpen = ref(false);
const showAddSheetDialog = ref(false);
const showAiSidebar = ref(false);

// -- Drag interaction state (S147-4 headline feature) ------------------------
// One small state machine covers all three drag gestures the grid supports —
// a plain selection drag, a formula-range PICK drag (dragging while editing
// a "="-prefixed formula inserts/replaces the referenced range at the
// cursor), and a fill-handle drag. Exactly one is active at a time.
type DragMode = 'none' | 'select' | 'formula-range' | 'fill';
const dragMode = ref<DragMode>('none');
const formulaRangeAnchorCell = ref<SheetCellAddress | null>(null);
const formulaRangeState = ref<FormulaRangeDragState | null>(null);
const formulaRangeTarget = ref<'formula-bar' | 'inline'>('formula-bar');
const formulaRangeHighlight = ref<SheetCellRange | null>(null);
const fillSourceBounds = ref<SheetCellRange | null>(null);
const fillPreviewRange = ref<SheetCellRange | null>(null);

const totalGridWidthPx = computed(() => `${rowHeaderWidth + totalColumns * columnWidth}px`);

const tabs = computed(() => store.workbook?.sheets ?? []);

const defaultNewSheetName = computed(() => `Sheet${tabs.value.length + 1}`);

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

/** The normalised rectangle of the current selection — a single cell when no
 * drag/shift-arrow extension has happened, otherwise the anchor..focus box.
 * Computed from indices only (never per-cell), per the perf requirement. */
const selectionBounds = computed<SheetCellRange | null>(() => {
  if (!activeCell.value) return null;
  const anchor = selectionAnchor.value ?? activeCell.value;
  return normalizeRange(anchor, activeCell.value);
});

const selectionAddresses = computed<string[]>(() =>
  selectionBounds.value ? addressesInRange(selectionBounds.value) : [],
);

/** The A1(:A1) range text the AI helper reads — a bare address for a single
 * cell, `"A1:C3"` for a drag-selected block. Reused verbatim as the `range`
 * the backend bounds/serialises (S147-3.5) — never re-derived server-side. */
const selectionRangeText = computed<string | null>(() =>
  selectionBounds.value
    ? rangeReferenceFor(
      { column: selectionBounds.value.startColumn, row: selectionBounds.value.startRow },
      { column: selectionBounds.value.endColumn, row: selectionBounds.value.endRow },
    )
    : null,
);

const isBoldActive = computed(
  () => selectionAddresses.value.length > 0
    && selectionAddresses.value.every(
      (address) => !!store.getCellStyle(store.activeSheetName, address).bold,
    ),
);

const isItalicActive = computed(
  () => selectionAddresses.value.length > 0
    && selectionAddresses.value.every(
      (address) => !!store.getCellStyle(store.activeSheetName, address).italic,
    ),
);

/** The formula-range PICK context: the formula bar or the inline cell
 * editor currently holds `"="`-prefixed text AND has focus — the one
 * condition under which a click/drag on the grid inserts a reference
 * instead of changing the selection (requirement #1, the headline
 * feature). */
const isFormulaRangeDragContext = computed(() => {
  if (formulaBarFocused.value) return isFormulaInput(formulaBarValue.value);
  if (editingCell.value) return isFormulaInput(editingText.value);
  return false;
});

function columnLetters(column: number): string {
  return columnIndexToLetters(column);
}

function isSelected(row: number, column: number): boolean {
  return activeCell.value?.row === row && activeCell.value?.column === column;
}

function isEditingInline(row: number, column: number): boolean {
  return editingCell.value?.row === row && editingCell.value?.column === column;
}

function currentInlineInputElement(): HTMLInputElement | null {
  return Array.isArray(inlineInputRef.value) ? inlineInputRef.value[0] ?? null : inlineInputRef.value;
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
 * through the raw source (see `rawTextFor`). Number formatting (currency /
 * percent / decimals / date) NEVER changes the stored value, only what is
 * rendered — the engine keeps computing on the real number. */
function displayValue(row: number, column: number): string {
  const cell = cellModelAt(row, column);
  const value = cell.v;
  if (value === undefined || value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') {
    return formatNumericCellValue(value, store.getCellStyle(store.activeSheetName, formatA1(column, row)));
  }
  if (typeof value === 'object') {
    if (value.t === 'date') return new Date(value.v).toLocaleDateString();
    return value.v; // error code, e.g. "#NAME?"
  }
  return value;
}

function cellDisplayStyle(row: number, column: number): Record<string, string> {
  const style = store.getCellStyle(store.activeSheetName, formatA1(column, row));
  const declarations: Record<string, string> = {};
  if (style.bold) declarations.fontWeight = 'bold';
  if (style.italic) declarations.fontStyle = 'italic';
  if (style.align) declarations.textAlign = style.align;
  return declarations;
}

// ---------------------------------------------------------------------------
// Merged cells — a covered (non-top-left) cell renders blank and redirects
// selection/edit to its merge's top-left corner; the visible "one big cell"
// look comes from an absolute overlay drawn on top (see
// `mergedRangeOverlays`), the same technique the selection rectangle uses,
// rather than restructuring the per-row flex layout.
// ---------------------------------------------------------------------------

function containingMerge(row: number, column: number): SheetCellRange | null {
  return findContainingRange(store.getMerges(store.activeSheetName), row, column);
}

function isMergedAway(row: number, column: number): boolean {
  const merge = containingMerge(row, column);
  return !!merge && !(merge.startRow === row && merge.startColumn === column);
}

function resolveMergeTopLeft(row: number, column: number): SheetCellAddress {
  const merge = containingMerge(row, column);
  return merge ? { row: merge.startRow, column: merge.startColumn } : { row, column };
}

const mergedRangeOverlays = computed(() =>
  store.getMerges(store.activeSheetName)
    .map((rangeText) => parseRangeText(rangeText))
    .filter((range): range is SheetCellRange => !!range)
    .map((range) => ({
      range,
      text: displayValue(range.startRow, range.startColumn),
      style: {
        top: `${(range.startRow - 1) * rowHeight}px`,
        left: `${rowHeaderWidth + (range.startColumn - 1) * columnWidth}px`,
        width: `${(range.endColumn - range.startColumn + 1) * columnWidth}px`,
        height: `${(range.endRow - range.startRow + 1) * rowHeight}px`,
        ...cellDisplayStyle(range.startRow, range.startColumn),
      },
    })),
);

// ---------------------------------------------------------------------------
// Overlay geometry — the selection rectangle, the in-progress formula-range
// highlight, and the fill-drag preview are all a single absolutely
// positioned box computed from row/column INDICES, never a per-cell class
// toggle — the perf requirement (dragging must not re-render every cell per
// mousemove).
// ---------------------------------------------------------------------------

function rectStyleFor(range: SheetCellRange | null): Record<string, string> | null {
  if (!range) return null;
  return {
    top: `${(range.startRow - 1) * rowHeight}px`,
    left: `${rowHeaderWidth + (range.startColumn - 1) * columnWidth}px`,
    width: `${(range.endColumn - range.startColumn + 1) * columnWidth}px`,
    height: `${(range.endRow - range.startRow + 1) * rowHeight}px`,
  };
}

const selectionOverlayStyle = computed(() => rectStyleFor(selectionBounds.value));
const formulaRangeHighlightStyle = computed(() => rectStyleFor(formulaRangeHighlight.value));
const fillPreviewOverlayStyle = computed(() => rectStyleFor(fillPreviewRange.value));

const fillHandleStyle = computed(() => {
  if (!selectionBounds.value) return null;
  const bounds = selectionBounds.value;
  const top = bounds.endRow * rowHeight - FILL_HANDLE_SIZE_PX / 2;
  const left = rowHeaderWidth + bounds.endColumn * columnWidth - FILL_HANDLE_SIZE_PX / 2;
  return { top: `${top}px`, left: `${left}px` };
});

// ---------------------------------------------------------------------------
// Selection drag + click-to-select
// ---------------------------------------------------------------------------

function selectCell(target: SheetCellAddress, extend: boolean): void {
  activeCell.value = target;
  if (!extend) selectionAnchor.value = target;
  formulaBarValue.value = rawTextFor(target.row, target.column);
}

function onCellMouseDown(row: number, column: number, event: MouseEvent): void {
  if (!store.workbook || event.button !== 0) return;
  const target = resolveMergeTopLeft(row, column);

  if (isFormulaRangeDragContext.value) {
    // Clicking inside the cell currently being inline-edited should just
    // move the text cursor natively, not insert a self-reference.
    if (editingCell.value && isEditingInline(target.row, target.column)) return;
    event.preventDefault(); // keep focus on the formula input while picking a range
    beginFormulaRangeDrag(target.row, target.column);
    return;
  }

  if (editingCell.value && !isEditingInline(target.row, target.column)) commitInlineEdit();
  dragMode.value = 'select';
  selectCell(target, false);
  // `preventScroll` matters: without it, focusing a container taller than
  // the viewport can trigger a scroll-into-view that shifts the cell
  // sitting under the (still physically stationary) pointer mid-drag.
  scrollRef.value?.focus({ preventScroll: true });
}

function onCellMouseEnter(row: number, column: number): void {
  if (dragMode.value === 'select') {
    selectCell(resolveMergeTopLeft(row, column), true);
  } else if (dragMode.value === 'formula-range') {
    applyFormulaRangeReference(row, column);
  } else if (dragMode.value === 'fill') {
    updateFillDrag(row, column);
  }
}

function onGlobalMouseUp(): void {
  if (dragMode.value === 'formula-range') finalizeFormulaRangeDrag();
  else if (dragMode.value === 'fill') finalizeFillDrag();
  dragMode.value = 'none';
}

// ---------------------------------------------------------------------------
// Formula-range PICK drag (requirement #1 — the headline feature)
// ---------------------------------------------------------------------------

function beginFormulaRangeDrag(row: number, column: number): void {
  dragMode.value = 'formula-range';
  formulaRangeAnchorCell.value = { row, column };
  formulaRangeTarget.value = formulaBarFocused.value ? 'formula-bar' : 'inline';
  const input = formulaRangeTarget.value === 'formula-bar' ? formulaBarInputRef.value : currentInlineInputElement();
  const currentText = formulaRangeTarget.value === 'formula-bar' ? formulaBarValue.value : editingText.value;
  const cursorPosition = input?.selectionStart ?? currentText.length;
  formulaRangeState.value = beginRangeInsertion(cursorPosition);
  applyFormulaRangeReference(row, column);
}

function applyFormulaRangeReference(row: number, column: number): void {
  const anchor = formulaRangeAnchorCell.value;
  const state = formulaRangeState.value;
  if (!anchor || !state) return;
  const referenceText = rangeReferenceFor(anchor, { row, column });
  const currentText = formulaRangeTarget.value === 'formula-bar' ? formulaBarValue.value : editingText.value;
  const result = applyRangeInsertion(currentText, state, referenceText);
  formulaRangeState.value = result.state;
  if (formulaRangeTarget.value === 'formula-bar') formulaBarValue.value = result.text;
  else editingText.value = result.text;
  formulaRangeHighlight.value = normalizeRange(anchor, { row, column });
  nextTick(() => placeCursor(result.cursorPosition));
}

function placeCursor(position: number): void {
  const input = formulaRangeTarget.value === 'formula-bar' ? formulaBarInputRef.value : currentInlineInputElement();
  input?.setSelectionRange(position, position);
}

function finalizeFormulaRangeDrag(): void {
  const target = formulaRangeTarget.value;
  formulaRangeHighlight.value = null;
  formulaRangeState.value = null;
  formulaRangeAnchorCell.value = null;
  nextTick(() => {
    const input = target === 'formula-bar' ? formulaBarInputRef.value : currentInlineInputElement();
    input?.focus();
  });
}

// ---------------------------------------------------------------------------
// Fill-handle drag (requirement #2) — always fills FROM the single active
// cell, matching both the spec ("copies the source cell down/right") and the
// backend's `fill_from` contract (one source per target, never a pattern
// fill from a multi-cell source).
// ---------------------------------------------------------------------------

function onFillHandleMouseDown(event: MouseEvent): void {
  event.preventDefault();
  event.stopPropagation();
  if (!activeCell.value) return;
  dragMode.value = 'fill';
  const source: SheetCellRange = {
    startRow: activeCell.value.row,
    endRow: activeCell.value.row,
    startColumn: activeCell.value.column,
    endColumn: activeCell.value.column,
  };
  fillSourceBounds.value = source;
  fillPreviewRange.value = source;
}

function updateFillDrag(row: number, column: number): void {
  const source = fillSourceBounds.value;
  if (!source) return;
  const rowExtension = Math.max(0, row - source.endRow);
  const columnExtension = Math.max(0, column - source.endColumn);
  if (rowExtension === 0 && columnExtension === 0) {
    fillPreviewRange.value = source;
  } else if (rowExtension >= columnExtension) {
    fillPreviewRange.value = { ...source, endRow: source.endRow + rowExtension };
  } else {
    fillPreviewRange.value = { ...source, endColumn: source.endColumn + columnExtension };
  }
}

function finalizeFillDrag(): void {
  const source = fillSourceBounds.value;
  const target = fillPreviewRange.value;
  fillSourceBounds.value = null;
  fillPreviewRange.value = null;
  if (!source || !target) return;

  const sourceAddress = formatA1(source.startColumn, source.startRow);
  const targetAddresses: string[] = [];
  if (target.endRow > source.endRow) {
    for (let row = source.endRow + 1; row <= target.endRow; row += 1) {
      targetAddresses.push(formatA1(source.startColumn, row));
    }
  } else if (target.endColumn > source.endColumn) {
    for (let column = source.endColumn + 1; column <= target.endColumn; column += 1) {
      targetAddresses.push(formatA1(column, source.startRow));
    }
  }
  if (targetAddresses.length > 0) store.fillCells(store.activeSheetName, sourceAddress, targetAddresses);
}

// ---------------------------------------------------------------------------
// Cell editing
// ---------------------------------------------------------------------------

function onStartInlineEdit(row: number, column: number): void {
  if (!store.canEdit()) return;
  const target = resolveMergeTopLeft(row, column);
  activeCell.value = target;
  selectionAnchor.value = target;
  editingCell.value = target;
  editingText.value = rawTextFor(target.row, target.column);
  nextTick(() => currentInlineInputElement()?.focus());
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

function commitInlineEditAndMove(direction: 'down' | 'right'): void {
  const cell = editingCell.value;
  commitInlineEdit();
  if (!cell) return;
  const target = direction === 'down'
    ? { row: Math.min(totalRows, cell.row + 1), column: cell.column }
    : { row: cell.row, column: Math.min(totalColumns, cell.column + 1) };
  selectCell(resolveMergeTopLeft(target.row, target.column), false);
}

function cancelInlineEdit(): void {
  editingCell.value = null;
}

function commitFormulaBar(): void {
  if (!activeCell.value) return;
  const { row, column } = activeCell.value;
  store.setCellFromRawInput(store.activeSheetName, formatA1(column, row), formulaBarValue.value);
}

function cancelFormulaBarEdit(): void {
  if (activeCell.value) formulaBarValue.value = rawTextFor(activeCell.value.row, activeCell.value.column);
  formulaBarInputRef.value?.blur();
}

// ---------------------------------------------------------------------------
// Keyboard navigation (requirement #3) — arrows move, shift+arrows extend,
// Enter/Tab commit an in-progress inline edit, Escape cancels one. Only
// active when the grid container itself (not an `<input>` inside it) holds
// focus, so a formula being typed keeps native text-cursor arrow behaviour.
// ---------------------------------------------------------------------------

const ARROW_KEY_DELTAS: Record<string, { row: number; column: number }> = {
  ArrowUp: { row: -1, column: 0 },
  ArrowDown: { row: 1, column: 0 },
  ArrowLeft: { row: 0, column: -1 },
  ArrowRight: { row: 0, column: 1 },
};

function moveActiveCell(delta: { row: number; column: number }, extend: boolean): void {
  if (!activeCell.value) {
    activeCell.value = { row: 1, column: 1 };
    selectionAnchor.value = activeCell.value;
    return;
  }
  const nextRow = Math.min(totalRows, Math.max(1, activeCell.value.row + delta.row));
  const nextColumn = Math.min(totalColumns, Math.max(1, activeCell.value.column + delta.column));
  selectCell({ row: nextRow, column: nextColumn }, extend);
}

function onGridKeydown(event: KeyboardEvent): void {
  if (editingCell.value) return; // the inline `<input>` owns its own keys
  const arrowDelta = ARROW_KEY_DELTAS[event.key];
  if (arrowDelta) {
    event.preventDefault();
    moveActiveCell(arrowDelta, event.shiftKey);
    return;
  }
  if (event.key === 'Enter') {
    event.preventDefault();
    moveActiveCell(ARROW_KEY_DELTAS.ArrowDown, false);
  } else if (event.key === 'Tab') {
    event.preventDefault();
    moveActiveCell(ARROW_KEY_DELTAS.ArrowRight, false);
  }
}

// ---------------------------------------------------------------------------
// Toolbar actions (requirement #4)
// ---------------------------------------------------------------------------

function onToggleBold(): void {
  store.toggleBold(store.activeSheetName, selectionAddresses.value);
}

function onToggleItalic(): void {
  store.toggleItalic(store.activeSheetName, selectionAddresses.value);
}

function onSetAlign(align: OfficeSheetTextAlignment): void {
  store.setAlignment(store.activeSheetName, selectionAddresses.value, align);
}

function onSetFormat(format: OfficeSheetNumberFormat): void {
  store.setNumberFormat(store.activeSheetName, selectionAddresses.value, format);
}

function onAdjustDecimals(delta: number): void {
  store.adjustDecimals(store.activeSheetName, selectionAddresses.value, delta);
}

function onMerge(): void {
  if (!selectionBounds.value) return;
  const bounds = selectionBounds.value;
  const rangeText = `${formatA1(bounds.startColumn, bounds.startRow)}:${formatA1(bounds.endColumn, bounds.endRow)}`;
  store.mergeRange(store.activeSheetName, rangeText);
}

function onUnmerge(): void {
  if (!activeCell.value) return;
  store.unmergeAt(store.activeSheetName, formatA1(activeCell.value.column, activeCell.value.row));
}

function onPrint(): void {
  window.print();
}

async function onExportFormat(format: OfficeSheetExportFormat): Promise<void> {
  exportMenuOpen.value = false;
  await onExport(format);
}

function onSelectTab(sheetName: string): void {
  store.activeSheetName = sheetName;
  activeCell.value = null;
  selectionAnchor.value = null;
  formulaBarValue.value = '';
}

function onAddSheetDialogConfirm(sheetName: string): void {
  showAddSheetDialog.value = false;
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

async function onToggleAi(enabled: boolean): Promise<void> {
  await store.toggleAi(enabled);
}

async function onRunAiCapability(capability: OfficeSheetAiCapability, intent: string): Promise<void> {
  if (!activeCell.value) return;
  await store.runAiCapability(capability, {
    address: activeAddressLabel.value,
    rangeText: selectionRangeText.value ?? undefined,
    intent,
  });
}

async function onAcceptAiProposal(): Promise<void> {
  await store.acceptAiProposal();
}

async function onReload(): Promise<void> {
  await store.load(nodeId.value);
}

function onScroll(): void {
  if (scrollRef.value) scrollTop.value = scrollRef.value.scrollTop;
}

onMounted(async () => {
  window.addEventListener('mouseup', onGlobalMouseUp);
  await store.load(nodeId.value);
  // Wait for Vue to actually paint the loaded/loading-toggle DOM update
  // before measuring — reading `clientHeight` synchronously right after
  // `store.load` resolves can race Vue's own render flush.
  await nextTick();
  // Defensive ceiling. The host layout now bounds this page properly — the
  // route sets `meta.fullHeight`, and UserLayout's `.main-content--full-height`
  // gives it a real height with `overflow: hidden` — so `height: 100%` resolves
  // and virtualization works on its own. This clamp stays as a belt-and-braces
  // guard for any host that does NOT opt in (an embed, a future layout): without
  // a bound, `.office-sheet-grid` grows to fit all 500 rows and silently renders
  // ~13,000 cells instead of a window. `window.innerHeight` is always a correct
  // upper bound on what can be visible, so clamping to it can never render too
  // little — only avoid rendering absurdly too much.
  if (scrollRef.value) {
    const measuredHeight = scrollRef.value.clientHeight || DEFAULT_VIEWPORT_HEIGHT_PX;
    viewportHeight.value = Math.min(measuredHeight, window.innerHeight);
  }
});

onBeforeUnmount(() => {
  window.removeEventListener('mouseup', onGlobalMouseUp);
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
  /* A flex item will not shrink below its content unless min-width/height are
     cleared. Without these the 26-column grid (2 648px) stretches the page
     sideways instead of scrolling inside `.office-sheet-grid`. */
  min-width: 0;
  max-width: 100%;
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
.office-sheet-workspace {
  display: flex;
  flex: 1;
  min-height: 0;
}
.office-sheet-main-column {
  display: flex;
  flex-direction: column;
  flex: 1;
  min-width: 0;
  min-height: 0;
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
.office-sheet-toolbar {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 6px 16px;
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
  flex-wrap: wrap;
}
.office-sheet-toolbar-btn {
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: 4px;
  background: transparent;
  padding: 4px 8px;
  font-size: 0.75rem;
  line-height: 1;
  cursor: pointer;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-sheet-toolbar-btn:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
.office-sheet-toolbar-btn.is-active {
  background: var(--vbwd-color-primary, #3498db);
  color: #fff;
  border-color: var(--vbwd-color-primary, #3498db);
}
.office-sheet-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: var(--vbwd-color-border, #e9ecef);
  margin: 0 4px;
}
.office-sheet-export-menu {
  position: relative;
}
.office-sheet-export-menu-list {
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 4;
  display: flex;
  flex-direction: column;
  min-width: 140px;
  background: var(--vbwd-color-surface, #fff);
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: 4px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
}
.office-sheet-export-menu-item {
  border: none;
  background: transparent;
  text-align: left;
  padding: 6px 10px;
  font-size: 0.8rem;
  cursor: pointer;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-sheet-export-menu-item:hover {
  background: var(--vbwd-bg-hover, #f5f6f7);
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
  outline: none;
  /* Same reason as `.office-sheet-editor`: the scroll container must be allowed
     to be narrower than the grid it scrolls, in BOTH axes. */
  min-width: 0;
  min-height: 0;
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
.office-sheet-cell.is-merged-away {
  border-color: transparent;
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
.office-sheet-merge-overlay {
  position: absolute;
  z-index: 1;
  display: flex;
  align-items: center;
  padding: 0 6px;
  font-size: 0.8rem;
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
  background: var(--vbwd-color-surface, #fff);
  border: 1px solid var(--vbwd-border-light, #eee);
  pointer-events: none;
}
.office-sheet-selection {
  position: absolute;
  z-index: 2;
  border: 2px solid var(--vbwd-color-primary, #3498db);
  pointer-events: none;
}
.office-sheet-range-highlight {
  position: absolute;
  z-index: 2;
  border: 2px dashed var(--vbwd-color-warning, #f39c12);
  background: rgba(243, 156, 18, 0.08);
  pointer-events: none;
}
.office-sheet-fill-preview {
  position: absolute;
  z-index: 2;
  border: 1px dashed var(--vbwd-color-text-secondary, #666);
  pointer-events: none;
}
.office-sheet-fill-handle {
  position: absolute;
  z-index: 3;
  width: 8px;
  height: 8px;
  background: var(--vbwd-color-primary, #3498db);
  border: 1px solid #fff;
  cursor: crosshair;
}
</style>
