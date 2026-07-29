import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';
import { Editor } from '@tiptap/vue-3';
import StarterKit from '@tiptap/starter-kit';
import { TextStyleKit } from '@tiptap/extension-text-style';
import { Table, TableRow, TableCell, TableHeader } from '@tiptap/extension-table';

import { OfficeAssetImage } from '../src/tiptap/OfficeAssetImage';

const mockUploadAsset = vi.fn();
const mockImportTableFromFile = vi.fn();
const mockExportDoc = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeDocApi: {
    uploadAsset: (...args: unknown[]) => mockUploadAsset(...args),
    importTableFromFile: (...args: unknown[]) => mockImportTableFromFile(...args),
    exportDoc: (...args: unknown[]) => mockExportDoc(...args),
  },
}));

import OfficeDocToolbar from '../src/components/OfficeDocToolbar.vue';

const tFallback = (key: string) => key;

function buildEditor(content?: Record<string, unknown>): Editor {
  return new Editor({
    content: content || { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'hello' }] }] },
    extensions: [
      StarterKit.configure({ link: { openOnClick: false } }),
      TextStyleKit.configure({ color: false, backgroundColor: false, lineHeight: false }),
      Table.configure({ resizable: false }),
      TableRow,
      TableHeader,
      TableCell,
      OfficeAssetImage,
    ],
  });
}

function mountToolbar(editor: Editor, canEdit = true) {
  return mount(OfficeDocToolbar, {
    props: { editor, nodeId: 'doc-1', docName: 'My Doc', canEdit },
    global: { mocks: { $t: tFallback } },
  });
}

describe('OfficeDocToolbar', () => {
  let editor: Editor;

  beforeEach(() => {
    vi.clearAllMocks();
    editor = buildEditor();
    // Select the paragraph's own text ("hello", positions 1..6) — NOT
    // `selectAll()`, which spans into the editor's auto-appended trailing
    // empty paragraph and leaves the selection outside the just-changed
    // node once a block-level command (setHeading/toggleBulletList) runs.
    editor.commands.setTextSelection({ from: 1, to: 6 });
  });

  it('renders every required toolbar control', () => {
    const wrapper = mountToolbar(editor);
    const expectedTestIds = [
      'office-doc-tb-style',
      'office-doc-tb-font',
      'office-doc-tb-size',
      'office-doc-tb-bold',
      'office-doc-tb-italic',
      'office-doc-tb-underline',
      'office-doc-tb-strike',
      'office-doc-tb-bullet',
      'office-doc-tb-numbered',
      'office-doc-tb-indent',
      'office-doc-tb-outdent',
      'office-doc-tb-image',
      'office-doc-tb-table',
      'office-doc-tb-table-from-file',
      'office-doc-tb-print',
      'office-doc-tb-export',
    ];
    for (const testId of expectedTestIds) {
      expect(wrapper.find(`[data-testid="${testId}"]`).exists(), testId).toBe(true);
    }
    wrapper.unmount();
  });

  it('toggles bold/italic/underline/strike marks on the editor', async () => {
    const wrapper = mountToolbar(editor);

    await wrapper.find('[data-testid="office-doc-tb-bold"]').trigger('click');
    expect(editor.isActive('bold')).toBe(true);

    await wrapper.find('[data-testid="office-doc-tb-italic"]').trigger('click');
    expect(editor.isActive('italic')).toBe(true);

    await wrapper.find('[data-testid="office-doc-tb-underline"]').trigger('click');
    expect(editor.isActive('underline')).toBe(true);

    await wrapper.find('[data-testid="office-doc-tb-strike"]').trigger('click');
    expect(editor.isActive('strike')).toBe(true);

    wrapper.unmount();
  });

  it('the style select sets a heading level via setHeading', async () => {
    const wrapper = mountToolbar(editor);
    const select = wrapper.get('[data-testid="office-doc-tb-style"]');

    await select.setValue('heading-2');

    const firstNode = editor.getJSON().content?.[0];
    expect(firstNode?.type).toBe('heading');
    expect(firstNode?.attrs?.level).toBe(2);
    wrapper.unmount();
  });

  it('the style select returns to a normal paragraph', async () => {
    editor.chain().focus().setHeading({ level: 1 }).run();
    const wrapper = mountToolbar(editor);
    const select = wrapper.get('[data-testid="office-doc-tb-style"]');

    await select.setValue('normal');

    expect(editor.getJSON().content?.[0]?.type).toBe('paragraph');
    wrapper.unmount();
  });

  it('the font select applies a fontFamily textStyle mark', async () => {
    const wrapper = mountToolbar(editor);
    const select = wrapper.get('[data-testid="office-doc-tb-font"]');

    await select.setValue('Georgia');

    expect(editor.getAttributes('textStyle').fontFamily).toBe('Georgia');
    wrapper.unmount();
  });

  it('the size select applies a fontSize textStyle mark', async () => {
    const wrapper = mountToolbar(editor);
    const select = wrapper.get('[data-testid="office-doc-tb-size"]');

    await select.setValue('18px');

    expect(editor.getAttributes('textStyle').fontSize).toBe('18px');
    wrapper.unmount();
  });

  it('bullet and numbered list buttons toggle the corresponding list', async () => {
    const wrapper = mountToolbar(editor);

    await wrapper.find('[data-testid="office-doc-tb-bullet"]').trigger('click');
    expect(editor.isActive('bulletList')).toBe(true);

    await wrapper.find('[data-testid="office-doc-tb-numbered"]').trigger('click');
    expect(editor.isActive('orderedList')).toBe(true);
    expect(editor.isActive('bulletList')).toBe(false);

    wrapper.unmount();
  });

  it('indent (sink) and outdent (lift) are disabled outside a list', () => {
    const wrapper = mountToolbar(editor);

    expect(
      (wrapper.get('[data-testid="office-doc-tb-indent"]').element as HTMLButtonElement).disabled,
    ).toBe(true);
    expect(
      (wrapper.get('[data-testid="office-doc-tb-outdent"]').element as HTMLButtonElement).disabled,
    ).toBe(true);

    wrapper.unmount();
  });

  it('indent sinks a list item deeper once a preceding sibling item exists', async () => {
    // sinkListItem needs a PRECEDING sibling to nest under — a lone list
    // item has nothing to become its parent, so this builds two items
    // UP FRONT (constructing a fresh editor, rather than swapping content
    // into the shared one via `setContent` — that route runs the trailing
    // -node housekeeping transaction and leaves the cursor past the list
    // entirely) and selects the SECOND item before indenting it.
    const listEditor = buildEditor({
      type: 'doc',
      content: [
        {
          type: 'bulletList',
          content: [
            { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'one' }] }] },
            { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'two' }] }] },
          ],
        },
      ],
    });
    listEditor.commands.focus('end');
    const wrapper = mountToolbar(listEditor);

    const indentButton = wrapper.get('[data-testid="office-doc-tb-indent"]');
    expect((indentButton.element as HTMLButtonElement).disabled).toBe(false);
    await indentButton.trigger('click');

    // The second item is now nested INSIDE the first item's own bulletList
    // — the top-level list has only one direct child left.
    const topLevelList = listEditor.getJSON().content?.[0];
    expect(topLevelList?.content?.length).toBe(1);
    wrapper.unmount();
  });

  it('uploading an image inserts an image node with the returned office-asset src', async () => {
    mockUploadAsset.mockResolvedValue({ node_id: 'asset-1', src: 'office-asset:asset-1' });
    const wrapper = mountToolbar(editor);
    const file = new File(['fake-bytes'], 'cat.png', { type: 'image/png' });
    const input = wrapper.get('[data-testid="office-doc-tb-image-input"]')
      .element as HTMLInputElement;
    Object.defineProperty(input, 'files', { value: [file] });

    await wrapper.find('[data-testid="office-doc-tb-image-input"]').trigger('change');
    await flushPromises();

    expect(mockUploadAsset).toHaveBeenCalledWith('doc-1', file);
    const json = editor.getJSON();
    const hasImage = JSON.stringify(json).includes('office-asset:asset-1');
    expect(hasImage).toBe(true);
    wrapper.unmount();
  });

  it('importing a table from a file inserts a table node built from the returned rows', async () => {
    mockImportTableFromFile.mockResolvedValue({
      rows: [
        ['Name', 'Age'],
        ['Ada', '36'],
      ],
      row_count: 2,
      column_count: 2,
    });
    const wrapper = mountToolbar(editor);
    const file = new File(['a,b'], 'people.csv', { type: 'text/csv' });
    const input = wrapper.get('[data-testid="office-doc-tb-table-from-file-input"]')
      .element as HTMLInputElement;
    Object.defineProperty(input, 'files', { value: [file] });

    await wrapper.find('[data-testid="office-doc-tb-table-from-file-input"]').trigger('change');
    await flushPromises();

    expect(mockImportTableFromFile).toHaveBeenCalledWith('doc-1', file);
    const html = editor.getHTML();
    expect(html).toContain('<table');
    expect(html).toContain('Ada');
    wrapper.unmount();
  });

  it('shows a table-import error message instead of throwing when the upload is refused', async () => {
    mockImportTableFromFile.mockRejectedValue(new Error('format must be one of'));
    const wrapper = mountToolbar(editor);
    const file = new File(['a,b'], 'people.pdf');
    const input = wrapper.get('[data-testid="office-doc-tb-table-from-file-input"]')
      .element as HTMLInputElement;
    Object.defineProperty(input, 'files', { value: [file] });

    await wrapper.find('[data-testid="office-doc-tb-table-from-file-input"]').trigger('change');
    await flushPromises();

    expect(wrapper.find('[data-testid="office-doc-tb-table-import-error"]').text()).toContain(
      'format must be one of',
    );
    wrapper.unmount();
  });

  it('the print button calls window.print', async () => {
    const printSpy = vi.spyOn(window, 'print').mockImplementation(() => {});
    const wrapper = mountToolbar(editor);

    await wrapper.find('[data-testid="office-doc-tb-print"]').trigger('click');

    expect(printSpy).toHaveBeenCalledOnce();
    wrapper.unmount();
    printSpy.mockRestore();
  });

  it('print and export remain enabled for a view-only (canEdit=false) document', () => {
    const wrapper = mountToolbar(editor, false);

    expect(
      (wrapper.get('[data-testid="office-doc-tb-print"]').element as HTMLButtonElement).disabled,
    ).toBe(false);
    expect(
      (wrapper.get('[data-testid="office-doc-tb-export"]').element as HTMLButtonElement).disabled,
    ).toBe(false);
    expect(
      (wrapper.get('[data-testid="office-doc-tb-bold"]').element as HTMLButtonElement).disabled,
    ).toBe(true);

    wrapper.unmount();
  });

  it('opens the export menu and exports the chosen format as a download', async () => {
    mockExportDoc.mockResolvedValue({ blob: new Blob(['%PDF-1.4']), filename: 'My Doc.pdf' });
    const createObjectURLSpy = vi
      .spyOn(URL, 'createObjectURL')
      .mockReturnValue('blob:fake-url');
    const revokeObjectURLSpy = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
    const wrapper = mountToolbar(editor);

    await wrapper.find('[data-testid="office-doc-tb-export"]').trigger('click');
    expect(wrapper.find('[data-testid="office-doc-tb-export-menu"]').exists()).toBe(true);

    await wrapper.find('[data-testid="office-doc-tb-export-pdf"]').trigger('click');
    await flushPromises();

    expect(mockExportDoc).toHaveBeenCalledWith('doc-1', 'pdf', 'My Doc.pdf');
    expect(createObjectURLSpy).toHaveBeenCalled();
    expect(revokeObjectURLSpy).toHaveBeenCalled();
    expect(wrapper.find('[data-testid="office-doc-tb-export-menu"]').exists()).toBe(false);

    wrapper.unmount();
    createObjectURLSpy.mockRestore();
    revokeObjectURLSpy.mockRestore();
  });
});
