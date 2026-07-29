// Extends Tiptap's official Image node with a custom Vue NodeView
// (`OfficeAssetImageNode.vue`) so an `office-asset:<id>` src — the ONLY src
// shape the backend's content schema allows (`doc_content.py`) — resolves
// through an authenticated fetch of `GET /docs/<id>/assets/<assetId>`
// rather than being handed to the browser as a literal `<img src>` (which
// it cannot load: it is not a URL scheme, and the route requires the
// session's Bearer header anyway). The STORED node attrs are unchanged —
// this only changes how the node is DISPLAYED.
import Image from '@tiptap/extension-image';
import { VueNodeViewRenderer } from '@tiptap/vue-3';

import OfficeAssetImageNode from '../components/OfficeAssetImageNode.vue';

export const OFFICE_ASSET_SRC_PREFIX = 'office-asset:';

export const OfficeAssetImage = Image.extend({
  addNodeView() {
    return VueNodeViewRenderer(OfficeAssetImageNode);
  },
});
