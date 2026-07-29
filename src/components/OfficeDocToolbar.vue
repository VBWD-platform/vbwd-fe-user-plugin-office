<template>
  <div
    class="office-doc-toolbar"
    data-testid="office-doc-toolbar"
  >
    <select
      class="office-doc-toolbar-select"
      data-testid="office-doc-tb-style"
      :disabled="!canEdit"
      :value="currentBlockStyle"
      @change="onStyleChange(($event.target as HTMLSelectElement).value)"
    >
      <option value="normal">
        {{ $t('office.doc.toolbar.styleNormal') }}
      </option>
      <option value="heading-1">
        {{ $t('office.doc.toolbar.styleHeading1') }}
      </option>
      <option value="heading-2">
        {{ $t('office.doc.toolbar.styleHeading2') }}
      </option>
      <option value="heading-3">
        {{ $t('office.doc.toolbar.styleHeading3') }}
      </option>
      <option value="blockquote">
        {{ $t('office.doc.toolbar.styleQuote') }}
      </option>
      <option value="codeBlock">
        {{ $t('office.doc.toolbar.styleCodeBlock') }}
      </option>
    </select>

    <select
      class="office-doc-toolbar-select"
      data-testid="office-doc-tb-font"
      :disabled="!canEdit"
      @change="onFontFamilyChange(($event.target as HTMLSelectElement).value)"
    >
      <option value="">
        {{ $t('office.doc.toolbar.fontDefault') }}
      </option>
      <option
        v-for="font in FONT_FAMILY_CHOICES"
        :key="font"
        :value="font"
      >
        {{ font }}
      </option>
    </select>

    <select
      class="office-doc-toolbar-select office-doc-toolbar-select--narrow"
      data-testid="office-doc-tb-size"
      :disabled="!canEdit"
      @change="onFontSizeChange(($event.target as HTMLSelectElement).value)"
    >
      <option value="">
        {{ $t('office.doc.toolbar.sizeDefault') }}
      </option>
      <option
        v-for="size in FONT_SIZE_CHOICES"
        :key="size"
        :value="size"
      >
        {{ size }}
      </option>
    </select>

    <span class="office-doc-toolbar-divider" />

    <button
      type="button"
      class="office-doc-toolbar-btn"
      :class="{ 'is-active': editor?.isActive('bold') }"
      data-testid="office-doc-tb-bold"
      :disabled="!canEdit"
      @click="editor?.chain().focus().toggleBold().run()"
    >
      <strong>B</strong>
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      :class="{ 'is-active': editor?.isActive('italic') }"
      data-testid="office-doc-tb-italic"
      :disabled="!canEdit"
      @click="editor?.chain().focus().toggleItalic().run()"
    >
      <em>I</em>
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      :class="{ 'is-active': editor?.isActive('underline') }"
      data-testid="office-doc-tb-underline"
      :disabled="!canEdit"
      @click="editor?.chain().focus().toggleUnderline().run()"
    >
      <u>U</u>
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      :class="{ 'is-active': editor?.isActive('strike') }"
      data-testid="office-doc-tb-strike"
      :disabled="!canEdit"
      @click="editor?.chain().focus().toggleStrike().run()"
    >
      <s>S</s>
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      data-testid="office-doc-tb-link"
      :disabled="!canEdit"
      :title="$t('office.doc.toolbar.insertLink')"
      @click="onInsertLink"
    >
      🔗
    </button>

    <span class="office-doc-toolbar-divider" />

    <button
      type="button"
      class="office-doc-toolbar-btn"
      :class="{ 'is-active': editor?.isActive('bulletList') }"
      data-testid="office-doc-tb-bullet"
      :disabled="!canEdit"
      @click="editor?.chain().focus().toggleBulletList().run()"
    >
      •≡
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      :class="{ 'is-active': editor?.isActive('orderedList') }"
      data-testid="office-doc-tb-numbered"
      :disabled="!canEdit"
      @click="editor?.chain().focus().toggleOrderedList().run()"
    >
      1≡
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      data-testid="office-doc-tb-outdent"
      :disabled="!canEdit || !editor?.can().liftListItem('listItem')"
      :title="$t('office.doc.toolbar.outdent')"
      @click="editor?.chain().focus().liftListItem('listItem').run()"
    >
      ⇤
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      data-testid="office-doc-tb-indent"
      :disabled="!canEdit || !editor?.can().sinkListItem('listItem')"
      :title="$t('office.doc.toolbar.indent')"
      @click="editor?.chain().focus().sinkListItem('listItem').run()"
    >
      ⇥
    </button>

    <span class="office-doc-toolbar-divider" />

    <button
      type="button"
      class="office-doc-toolbar-btn"
      data-testid="office-doc-tb-image"
      :disabled="!canEdit"
      :title="$t('office.doc.toolbar.insertImage')"
      @click="imageInputRef?.click()"
    >
      🖼
    </button>
    <input
      ref="imageInputRef"
      type="file"
      accept="image/png,image/jpeg,image/gif,image/webp"
      class="office-doc-toolbar-hidden-input"
      data-testid="office-doc-tb-image-input"
      @change="onImageInputChange"
    >

    <button
      type="button"
      class="office-doc-toolbar-btn"
      data-testid="office-doc-tb-table"
      :disabled="!canEdit"
      :title="$t('office.doc.toolbar.insertTable')"
      @click="onInsertEmptyTable"
    >
      ▦
    </button>
    <button
      type="button"
      class="office-doc-toolbar-btn"
      data-testid="office-doc-tb-table-from-file"
      :disabled="!canEdit"
      :title="$t('office.doc.toolbar.insertTableFromFile')"
      @click="tableFileInputRef?.click()"
    >
      ▦↑
    </button>
    <input
      ref="tableFileInputRef"
      type="file"
      accept=".csv,.xlsx"
      class="office-doc-toolbar-hidden-input"
      data-testid="office-doc-tb-table-from-file-input"
      @change="onTableFileInputChange"
    >

    <span class="office-doc-toolbar-divider" />

    <button
      type="button"
      class="office-doc-toolbar-btn"
      data-testid="office-doc-tb-print"
      :title="$t('office.doc.toolbar.print')"
      @click="onPrint"
    >
      🖨
    </button>

    <div class="office-doc-toolbar-export">
      <button
        type="button"
        class="office-doc-toolbar-btn"
        data-testid="office-doc-tb-export"
        @click="showExportMenu = !showExportMenu"
      >
        {{ $t('office.doc.toolbar.export') }} ▾
      </button>
      <div
        v-if="showExportMenu"
        class="office-doc-toolbar-export-menu"
        data-testid="office-doc-tb-export-menu"
      >
        <button
          type="button"
          class="office-doc-toolbar-export-item"
          data-testid="office-doc-tb-export-pdf"
          @click="onExport('pdf')"
        >
          {{ $t('office.doc.toolbar.exportPdf') }}
        </button>
        <button
          type="button"
          class="office-doc-toolbar-export-item"
          data-testid="office-doc-tb-export-docx"
          @click="onExport('docx')"
        >
          {{ $t('office.doc.toolbar.exportDocx') }}
        </button>
        <button
          type="button"
          class="office-doc-toolbar-export-item"
          data-testid="office-doc-tb-export-md"
          @click="onExport('md')"
        >
          {{ $t('office.doc.toolbar.exportMarkdown') }}
        </button>
      </div>
    </div>

    <div
      v-if="tableImportError"
      class="office-doc-toolbar-error"
      data-testid="office-doc-tb-table-import-error"
    >
      {{ tableImportError }}
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue';
import type { Editor } from '@tiptap/vue-3';

import { officeDocApi, type OfficeDocExportFormat } from '../api/officeApi';

const FONT_FAMILY_CHOICES = [
  'Arial',
  'Georgia',
  'Times New Roman',
  'Courier New',
  'Verdana',
] as const;

const FONT_SIZE_CHOICES = ['10px', '12px', '14px', '16px', '18px', '24px', '32px'] as const;

const props = defineProps<{
  editor: Editor | undefined;
  nodeId: string;
  docName: string;
  canEdit: boolean;
}>();

const imageInputRef = ref<HTMLInputElement | null>(null);
const tableFileInputRef = ref<HTMLInputElement | null>(null);
const showExportMenu = ref(false);
const tableImportError = ref<string | null>(null);

const currentBlockStyle = computed(() => {
  const editor = props.editor;
  if (!editor) return 'normal';
  if (editor.isActive('heading', { level: 1 })) return 'heading-1';
  if (editor.isActive('heading', { level: 2 })) return 'heading-2';
  if (editor.isActive('heading', { level: 3 })) return 'heading-3';
  if (editor.isActive('blockquote')) return 'blockquote';
  if (editor.isActive('codeBlock')) return 'codeBlock';
  return 'normal';
});

function onStyleChange(value: string): void {
  const editor = props.editor;
  if (!editor) return;
  const chain = editor.chain().focus();
  if (value === 'normal') chain.setParagraph().run();
  else if (value === 'heading-1') chain.setHeading({ level: 1 }).run();
  else if (value === 'heading-2') chain.setHeading({ level: 2 }).run();
  else if (value === 'heading-3') chain.setHeading({ level: 3 }).run();
  else if (value === 'blockquote') chain.setBlockquote().run();
  else if (value === 'codeBlock') chain.setCodeBlock().run();
}

function onFontFamilyChange(value: string): void {
  const editor = props.editor;
  if (!editor) return;
  if (value) editor.chain().focus().setFontFamily(value).run();
  else editor.chain().focus().unsetFontFamily().run();
}

function onFontSizeChange(value: string): void {
  const editor = props.editor;
  if (!editor) return;
  if (value) editor.chain().focus().setFontSize(value).run();
  else editor.chain().focus().unsetFontSize().run();
}

async function onImageInputChange(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (!file || !props.editor) return;

  const uploaded = await officeDocApi.uploadAsset(props.nodeId, file);
  props.editor.chain().focus().setImage({ src: uploaded.src, alt: file.name }).run();
}

function onInsertLink(): void {
  const editor = props.editor;
  if (!editor) return;
  // eslint-disable-next-line no-alert
  const url = window.prompt('URL (https://...)');
  if (!url) return;
  editor.chain().focus().extendMarkRange('link').setLink({ href: url }).run();
}

function onInsertEmptyTable(): void {
  const editor = props.editor;
  if (!editor) return;
  // eslint-disable-next-line no-alert
  const rowsInput = window.prompt('Rows', '3');
  if (!rowsInput) return;
  // eslint-disable-next-line no-alert
  const columnsInput = window.prompt('Columns', '3');
  if (!columnsInput) return;
  const rows = Math.max(1, parseInt(rowsInput, 10) || 0);
  const columns = Math.max(1, parseInt(columnsInput, 10) || 0);
  editor.chain().focus().insertTable({ rows, cols: columns, withHeaderRow: true }).run();
}

/** A cell's paragraph content — empty for a blank cell (a bare
 * ``{"type": "paragraph"}`` is a valid Doc content node, doc_content.py's
 * ``EMPTY_CONTENT_MODEL`` already relies on the same shape). */
function tableCellParagraph(text: string): Record<string, unknown> {
  return text
    ? { type: 'paragraph', content: [{ type: 'text', text }] }
    : { type: 'paragraph' };
}

/** Builds a Tiptap table content fragment from imported rows — the FIRST
 * row renders as header cells (the common shape for an imported .csv/.xlsx
 * with a header row), every other row as ordinary cells. */
function buildTableContent(rows: string[][]): Record<string, unknown> {
  return {
    type: 'table',
    content: rows.map((row, rowIndex) => ({
      type: 'tableRow',
      content: row.map((cellText) => ({
        type: rowIndex === 0 ? 'tableHeader' : 'tableCell',
        content: [tableCellParagraph(cellText)],
      })),
    })),
  };
}

async function onTableFileInputChange(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (!file || !props.editor) return;

  tableImportError.value = null;
  try {
    const result = await officeDocApi.importTableFromFile(props.nodeId, file);
    if (result.rows.length === 0) return;
    props.editor.chain().focus().insertContent(buildTableContent(result.rows)).run();
  } catch (caught) {
    tableImportError.value = (caught as Error).message;
  }
}

function onPrint(): void {
  window.print();
}

async function onExport(format: OfficeDocExportFormat): Promise<void> {
  showExportMenu.value = false;
  const { blob, filename } = await officeDocApi.exportDoc(
    props.nodeId,
    format,
    `${props.docName}.${format}`,
  );
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}
</script>

<style scoped>
.office-doc-toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 4px;
  padding: 8px 16px;
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
}
.office-doc-toolbar-divider {
  width: 1px;
  align-self: stretch;
  background: var(--vbwd-color-border, #e9ecef);
  margin: 2px 4px;
}
.office-doc-toolbar-select {
  height: 28px;
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: 4px;
  background: var(--vbwd-color-surface, #fff);
  color: var(--vbwd-color-text-primary, #2c3e50);
  font-size: 0.8rem;
  max-width: 140px;
}
.office-doc-toolbar-select--narrow {
  max-width: 80px;
}
.office-doc-toolbar-btn {
  min-width: 28px;
  height: 28px;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  font-size: 0.8rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
  padding: 0 6px;
}
.office-doc-toolbar-btn:hover:not(:disabled) {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-doc-toolbar-btn.is-active {
  background: var(--vbwd-color-primary, #3498db);
  color: #fff;
}
.office-doc-toolbar-btn:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
.office-doc-toolbar-hidden-input {
  display: none;
}
.office-doc-toolbar-export {
  position: relative;
}
.office-doc-toolbar-export-menu {
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 20;
  display: flex;
  flex-direction: column;
  min-width: 140px;
  background: var(--vbwd-color-surface, #fff);
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-card-radius, 8px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
  padding: 4px;
}
.office-doc-toolbar-export-item {
  background: none;
  border: none;
  text-align: left;
  padding: 8px 10px;
  font-size: 0.85rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
  border-radius: var(--vbwd-btn-radius, 4px);
  cursor: pointer;
}
.office-doc-toolbar-export-item:hover {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-doc-toolbar-error {
  flex-basis: 100%;
  font-size: 0.8rem;
  color: var(--vbwd-color-danger, #e74c3c);
}

/* The document's own print stylesheet (`OfficeDocEditor.vue`) hides
 * everything but `.office-doc-content` — this covers the toolbar too,
 * duplicated here since it is scoped to THIS component. */
@media print {
  .office-doc-toolbar {
    display: none !important;
  }
}
</style>
