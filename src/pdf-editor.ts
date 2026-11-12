import { LitElement, css, html, nothing, svg, type PropertyValues, type SVGTemplateResult } from 'lit';
import { customElement, property, query, state } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import type { PDFDocumentProxy, PDFPageProxy, RenderTask } from 'pdfjs-dist';
import {
  iconChevron,
  iconComments,
  iconEmptyDoc,
  iconEye,
  iconFilter,
  iconHand,
  iconMenu,
  iconMinus,
  iconMultiSelect,
  iconOpen,
  iconPages,
  iconPlus,
  iconSave,
  iconSearch,
  iconSelect,
} from './icons';
import {
  addCheckBox,
  addTextField,
  applyMarks,
  findText,
  isRenderCancel,
  openDocument,
  pageSizeAtScale1,
  readOutline,
  pageTextLines,
  deleteDocumentPages,
  extractDocumentPages,
  insertBlankPage,
  pageFormFields,
  pageWordRects,
  paintPage,
  pdfRectFromPoints,
  projectRect,
  rebuildPages,
  rectsIntersect,
  replaceDocumentPage,
  replaceLineText,
  rotateDocumentPages,
  setTextField,
  toggleCheckBox,
  viewportToPdf,
  type DrawnMark,
  type OutlineEntry,
  type PdfRect,
  type TextMatch,
  type TextRun,
  type ViewportBox,
} from './pdf-engine';
import { renderSheetPages } from './sheet-preview';
import { renderSlidePages } from './slide-preview';
import { isZipDocument, renderDocxPages } from './word-preview';

export interface PdfLoadDetail {
  pages: number;
}

export interface PdfErrorDetail {
  message: string;
}

export type PdfSource = File | ArrayBuffer | Uint8Array | string;

type EditorMode = 'view' | 'annotate' | 'shapes' | 'insert' | 'measure' | 'edit' | 'fill-sign' | 'forms';
type PointerTool = 'hand' | 'select';
type MarkKind = 'highlight' | 'underline' | 'strike' | 'note' | 'freetext' | 'pen' | 'callout' | 'line' | 'arrow' | 'rect' | 'ellipse' | 'polygon' | 'text' | 'image' | 'link' | 'stamp' | 'distance' | 'perimeter' | 'area' | 'signature' | 'initials' | 'date';
type TextMarkKind = 'note' | 'freetext' | 'callout' | 'text' | 'link' | 'initials';
type ShapeKind = 'line' | 'arrow' | 'rect' | 'ellipse' | 'polygon';
type InsertKind = 'text' | 'image' | 'link' | 'stamp';
type MeasureKind = 'distance' | 'perimeter' | 'area';
type SignKind = 'signature' | 'initials' | 'date';
type FormKind = 'text-field' | 'checkbox';
type ArrowEnd = 'none' | 'start' | 'end';
type CommentStatus = 'none' | 'accepted' | 'rejected' | 'cancelled' | 'completed';
type CommentSort = 'position' | 'created' | 'modified' | 'status' | 'author' | 'type' | 'color';
type CommentAction = CommentStatus | 'marked' | 'unmarked';
type FilterTab = 'user' | 'color' | 'type' | 'status';

interface CommentFilterState {
  users: string[];
  colors: string[];
  types: string[];
  statuses: string[];
  includeReplies: boolean;
  filterDocument: boolean;
}

function emptyCommentFilter(): CommentFilterState {
  return { users: [], colors: [], types: [], statuses: [], includeReplies: true, filterDocument: false };
}

interface PdfPoint {
  x: number;
  y: number;
}

interface NoteReply {
  id: string;
  author: string;
  createdAt: number;
  text: string;
}

interface Mark {
  id: string;
  kind: MarkKind;
  page: number;
  rects: PdfRect[];
  text: string;
  sortY: number;
  points: PdfPoint[];
  src: string;
  color: string;
  opacity: number;
  weight: number;
  quote: string;
  author: string;
  createdAt: number;
  modifiedAt: number;
  status: CommentStatus;
  marked: boolean;
  replies: NoteReply[];
}

interface MarkBox extends ViewportBox {
  id: string;
  kind: MarkKind;
  selected: boolean;
  label: string;
  text: string;
  points: Array<{ x: number; y: number }>;
  src: string;
  color: string;
  opacity: number;
  weight: number;
}

interface TextEdit {
  page: number;
  rect: PdfRect;
  text: string;
  baseline: number;
  fontSize: number;
}

interface PageBookmark {
  id: string;
  page: number;
  title: string;
}

interface DocumentLayer {
  id: string;
  name: string;
  depth: number;
  visible: boolean;
}

interface DocumentFile {
  id: string;
  name: string;
  description: string;
  size: number | null;
}

interface EditDraft {
  page: number;
  rect: PdfRect;
  baseline: number;
  fontSize: number;
  left: number;
  top: number;
  text: string;
}

interface LineBox extends ViewportBox {
  text: string;
  rect: PdfRect;
  baseline: number;
  fontSize: number;
}

interface FieldBox extends ViewportBox {
  kind: 'text' | 'check';
  name: string;
  value: string;
  checked: boolean;
}

interface FormDraft {
  name: string;
  page: number;
  left: number;
  top: number;
  text: string;
}

interface NoteDraft {
  kind: TextMarkKind;
  page: number;
  pdfX: number;
  pdfY: number;
  left: number;
  top: number;
  text: string;
  tip: PdfPoint | null;
  rect: PdfRect | null;
  quote: string;
  rects: PdfRect[];
  author: string;
  createdAt: number;
}

interface CommentAnchor {
  left: number;
  top: number;
  fromX: number;
  fromY: number;
}

const COMMENT_POPUP_WIDTH = 272;

interface QuickSelection {
  page: number;
  rects: PdfRect[];
  quote: string;
  boxes: ViewportBox[];
}

const STYLE_SWATCHES = ['#d93025', '#e8710a', '#f6bf26', '#188038', '#129eaf', '#1a73e8', '#9334e6', '#202124'] as const;
const STYLE_WEIGHTS: ReadonlyArray<{ id: number; label: string }> = [
  { id: 1.25, label: 'Light' },
  { id: 2, label: 'Regular' },
  { id: 3.5, label: 'Heavy' },
];

const MARK_TOOLS: ReadonlyArray<{ id: MarkKind; label: string }> = [
  { id: 'highlight', label: 'Highlight' },
  { id: 'underline', label: 'Underline' },
  { id: 'strike', label: 'Strikethrough' },
  { id: 'note', label: 'Note' },
  { id: 'freetext', label: 'Free text' },
  { id: 'pen', label: 'Pen' },
  { id: 'callout', label: 'Callout' },
];

const SHAPE_TOOLS: ReadonlyArray<{ id: ShapeKind; label: string }> = [
  { id: 'line', label: 'Line' },
  { id: 'arrow', label: 'Arrow' },
  { id: 'rect', label: 'Rectangle' },
  { id: 'ellipse', label: 'Ellipse' },
  { id: 'polygon', label: 'Polygon' },
];

const COMMENT_STATUSES: ReadonlyArray<{ id: CommentStatus; label: string; color: string }> = [
  { id: 'accepted', label: 'Accepted', color: '#1e8e3e' },
  { id: 'rejected', label: 'Rejected', color: '#d93025' },
  { id: 'cancelled', label: 'Cancelled', color: '#80868b' },
  { id: 'completed', label: 'Completed', color: '#1a73e8' },
  { id: 'none', label: 'None', color: '#dadce0' },
];

const INSERT_TOOLS: ReadonlyArray<{ id: InsertKind; label: string }> = [
  { id: 'text', label: 'Text' },
  { id: 'image', label: 'Image' },
  { id: 'link', label: 'Link' },
  { id: 'stamp', label: 'Stamp' },
];

const MEASURE_TOOLS: ReadonlyArray<{ id: MeasureKind; label: string }> = [
  { id: 'distance', label: 'Distance' },
  { id: 'perimeter', label: 'Perimeter' },
  { id: 'area', label: 'Area' },
];

const SIGN_TOOLS: ReadonlyArray<{ id: SignKind; label: string }> = [
  { id: 'signature', label: 'Signature' },
  { id: 'initials', label: 'Initials' },
  { id: 'date', label: 'Date' },
];

const FORM_TOOLS: ReadonlyArray<{ id: FormKind; label: string }> = [
  { id: 'text-field', label: 'Text field' },
  { id: 'checkbox', label: 'Checkbox' },
];

const STAMPS = ['Approved', 'Draft', 'Final', 'Confidential'] as const;
type ZoomMode = 'percent' | 'fit-page' | 'fit-width';

const MODES: ReadonlyArray<{ id: EditorMode; label: string }> = [
  { id: 'view', label: 'View' },
  { id: 'annotate', label: 'Annotate' },
  { id: 'shapes', label: 'Shapes' },
  { id: 'insert', label: 'Insert' },
  { id: 'measure', label: 'Measure' },
  { id: 'edit', label: 'Edit' },
  { id: 'fill-sign', label: 'Fill and Sign' },
  { id: 'forms', label: 'Forms' },
];

const ZOOM_STEPS = [0.25, 0.5, 0.65, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 4];

const OPEN_ACCEPT =
  'application/pdf,.pdf,.doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,.ppt,.pptx,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation,.xls,.xlsx,.xlsm,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12,.png,.jpg,.jpeg,image/png,image/jpeg';
const WORD_ZOOM_LEVELS = [2, 1.5, 1, 0.75, 0.5] as const;

type WordFit = 'percent' | 'page' | 'width';
type WordZoomChoice = 'page' | 'width' | number;

function wordFormat(source: PdfSource): 'docx' | 'doc' | null {
  if (source instanceof File) {
    const name = source.name.toLowerCase();
    if (name.endsWith('.docx')) return 'docx';
    if (name.endsWith('.doc')) return 'doc';
    if (source.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') return 'docx';
    if (source.type === 'application/msword') return 'doc';
    return null;
  }
  if (typeof source === 'string') {
    const path = decodeURIComponent(source.split(/[?#]/)[0]).toLowerCase();
    if (path.endsWith('.docx')) return 'docx';
    if (path.endsWith('.doc')) return 'doc';
  }
  return null;
}

type SlideFormat = 'pptx' | 'ppt';

function slideFormat(source: PdfSource): SlideFormat | null {
  if (source instanceof File) {
    const name = source.name.toLowerCase();
    if (name.endsWith('.pptx')) return 'pptx';
    if (name.endsWith('.ppt')) return 'ppt';
    if (source.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') return 'pptx';
    if (source.type === 'application/vnd.ms-powerpoint') return 'ppt';
    return null;
  }
  if (typeof source === 'string') {
    const path = decodeURIComponent(source.split(/[?#]/)[0]).toLowerCase();
    if (path.endsWith('.pptx')) return 'pptx';
    if (path.endsWith('.ppt')) return 'ppt';
  }
  return null;
}

type SheetFormat = 'xlsx' | 'xlsm' | 'xls';

function sheetFormat(source: PdfSource): SheetFormat | null {
  if (source instanceof File) {
    const name = source.name.toLowerCase();
    if (name.endsWith('.xlsx')) return 'xlsx';
    if (name.endsWith('.xlsm')) return 'xlsm';
    if (name.endsWith('.xls')) return 'xls';
    if (source.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') return 'xlsx';
    if (source.type === 'application/vnd.ms-excel.sheet.macroEnabled.12') return 'xlsm';
    if (source.type === 'application/vnd.ms-excel') return 'xls';
    return null;
  }
  if (typeof source === 'string') {
    const path = decodeURIComponent(source.split(/[?#]/)[0]).toLowerCase();
    if (path.endsWith('.xlsx')) return 'xlsx';
    if (path.endsWith('.xlsm')) return 'xlsm';
    if (path.endsWith('.xls')) return 'xls';
  }
  return null;
}

type ImageFormat = 'png' | 'jpeg';

function imageFormat(source: PdfSource): ImageFormat | null {
  if (source instanceof File) {
    const name = source.name.toLowerCase();
    if (name.endsWith('.png')) return 'png';
    if (name.endsWith('.jpg') || name.endsWith('.jpeg')) return 'jpeg';
    if (source.type === 'image/png') return 'png';
    if (source.type === 'image/jpeg' || source.type === 'image/jpg') return 'jpeg';
    return null;
  }
  if (typeof source === 'string') {
    const path = decodeURIComponent(source.split(/[?#]/)[0]).toLowerCase();
    if (path.endsWith('.png')) return 'png';
    if (path.endsWith('.jpg') || path.endsWith('.jpeg')) return 'jpeg';
  }
  return null;
}

function looksLikeImage(bytes: Uint8Array, format: ImageFormat): boolean {
  if (format === 'png') {
    return bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;
  }
  return bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff;
}

function iconEnterFullScreen(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M9 4H5v4M15 4h4v4M9 20H5v-4M15 20h4v-4"></path>
    </svg>
  `;
}

function iconExitFullScreen(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M4 9h4V5M20 9h-4V5M4 15h4v4M20 15h-4v4"></path>
    </svg>
  `;
}

function iconPreviewDownload(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 5v10"></path>
      <path d="m8 11 4 4 4-4"></path>
      <path d="M6 19h12"></path>
    </svg>
  `;
}

function iconPreviewPrint(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round">
      <path d="M7 9V5h10v4"></path>
      <rect x="6" y="9" width="12" height="7" rx="1.4"></rect>
      <path d="M8 16h8v3H8z"></path>
    </svg>
  `;
}

function iconPreviewSettings(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round">
      <circle cx="12" cy="12" r="3"></circle>
      <path d="M12 4.5v2.2M12 17.3V19.5M4.5 12h2.2M17.3 12H19.5M6.6 6.6l1.6 1.6M15.8 15.8l1.6 1.6M17.4 6.6l-1.6 1.6M8.2 15.8l-1.6 1.6"></path>
    </svg>
  `;
}

function iconPaneThumbs(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7">
      <rect x="4" y="4" width="7" height="7" rx="1.2"></rect>
      <rect x="13" y="4" width="7" height="7" rx="1.2"></rect>
      <rect x="4" y="13" width="7" height="7" rx="1.2"></rect>
      <rect x="13" y="13" width="7" height="7" rx="1.2"></rect>
    </svg>
  `;
}

function iconPaneOutline(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round">
      <path d="M8 7h11M8 12h11M8 17h11"></path>
      <path d="M4.5 7h.01M4.5 12h.01M4.5 17h.01"></path>
    </svg>
  `;
}

function iconPaneBookmark(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round">
      <path d="M7 4.5h10V20l-5-3.2L7 20z"></path>
    </svg>
  `;
}

function iconPaneLayers(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round">
      <path d="M12 4.5 20 8.5 12 12.5 4 8.5z"></path>
      <path d="M4 12.2 12 16.2 20 12.2"></path>
      <path d="M4 15.6 12 19.6 20 15.6"></path>
    </svg>
  `;
}

function iconPaneClip(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M15.5 7.5 8.2 14.8a3 3 0 0 0 4.2 4.2l8-8a4.2 4.2 0 0 0-6-6L6.6 12.8a5.2 5.2 0 0 0 7.4 7.4l6.2-6.2"></path>
    </svg>
  `;
}

function iconTurnLeft(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M8 8H4.5V4.5"></path>
      <path d="M5 8a7 7 0 1 1-1.2 4"></path>
    </svg>
  `;
}

function iconTurnRight(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M16 8h3.5V4.5"></path>
      <path d="M19 8a7 7 0 1 0 1.2 4"></path>
    </svg>
  `;
}

function iconMoveUp(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 19V6"></path>
      <path d="M7 10.5 12 5.5l5 5"></path>
    </svg>
  `;
}

function iconMoveDown(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 5v13"></path>
      <path d="M7 13.5 12 18.5l5-5"></path>
    </svg>
  `;
}

function iconTrash(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M5 7.5h14"></path>
      <path d="M9 7.5V5.8h6V7.5"></path>
      <path d="M7.5 7.5 8.2 19h7.6l.7-11.5"></path>
    </svg>
  `;
}

function iconMore(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="currentColor">
      <circle cx="6" cy="12" r="1.4"></circle>
      <circle cx="12" cy="12" r="1.4"></circle>
      <circle cx="18" cy="12" r="1.4"></circle>
    </svg>
  `;
}

function iconPaneCheck(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M6 12.5 10 16.5 18 8"></path>
    </svg>
  `;
}

function iconPaneClear(): SVGTemplateResult {
  return svg`
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round">
      <path d="M7 7l10 10M17 7 7 17"></path>
    </svg>
  `;
}

interface PreviewSavePicker {
  showSaveFilePicker?: (options: {
    suggestedName?: string;
    types?: Array<{ description?: string; accept: Record<string, string[]> }>;
  }) => Promise<{
    name: string;
    createWritable: () => Promise<{ write: (data: Blob) => Promise<void>; close: () => Promise<void> }>;
  }>;
}

type SettingsSection = 'theme' | 'shortcuts' | 'advanced';
type ViewRead = 'page' | 'scroll';
type ViewSpread = 'single' | 'pair' | 'cover';
type PagePaneTab = 'thumbs' | 'outline' | 'bookmarks' | 'layers' | 'files';
type PageMove = 'start' | 'end' | 'up' | 'down';

const SETTINGS_SHORTCUTS: ReadonlyArray<{ keys: string; description: string; scope: 'all' | 'pdf' | 'preview' }> = [
  { keys: 'Ctrl + O', description: 'Open a file', scope: 'all' },
  { keys: 'Ctrl + S', description: 'Save the PDF', scope: 'pdf' },
  { keys: 'Ctrl + S', description: 'Download this file', scope: 'preview' },
  { keys: 'Ctrl + P', description: 'Print', scope: 'all' },
  { keys: 'Ctrl + F', description: 'Find in the document', scope: 'pdf' },
  { keys: 'Ctrl + C', description: 'Copy the selected text', scope: 'pdf' },
  { keys: 'Ctrl + +', description: 'Zoom in', scope: 'all' },
  { keys: 'Ctrl + -', description: 'Zoom out', scope: 'all' },
  { keys: 'Ctrl + 0', description: 'Zoom to 100%', scope: 'all' },
  { keys: 'Left / Right', description: 'Go to the previous or next page', scope: 'all' },
  { keys: 'Escape', description: 'Close a menu, dialog, or draft', scope: 'all' },
];

@customElement('pdf-editor')
export class PdfEditor extends LitElement {
  static override styles = css`
    :host {
      display: block;
      height: 100%;
      min-height: 480px;
      color: #202124;
      font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
    }

    * { box-sizing: border-box; }

    button, input, select { font-family: inherit; }

    .root {
      position: relative;
      height: 100%;
      min-height: 0;
      display: flex;
      flex-direction: column;
      background: #fff;
    }

    .toolbar {
      height: 48px;
      flex: none;
      display: flex;
      align-items: stretch;
      gap: 4px;
      position: relative;
      padding: 0 8px 0 4px;
      border-bottom: 1px solid #e0e3e7;
      background: #fff;
    }

    .cluster {
      display: flex;
      align-items: center;
      gap: 2px;
      flex: none;
    }

    .sep {
      width: 1px;
      height: 20px;
      margin: 0 6px;
      background: #e0e3e7;
    }

    .icon {
      width: 32px;
      height: 32px;
      padding: 0;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: #5f6368;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
    }

    .icon svg { width: 18px; height: 18px; display: block; }
    .icon:hover:not(:disabled) { background: #f1f3f4; }
    .icon:disabled { opacity: 0.38; cursor: default; }
    .icon[aria-pressed='true'] {
      color: #1a73e8;
      background: #e8f0fe;
    }

    .circle {
      width: 26px;
      height: 26px;
      border-radius: 50%;
      border: 1px solid #dadce0;
      background: #fff;
    }
    .circle:hover:not(:disabled) { background: #f8f9fa; }

    .anchor { position: relative; display: flex; align-items: center; }

    .zoom-button {
      height: 28px;
      min-width: 68px;
      padding: 0 4px 0 8px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      background: #fff;
      color: #3c4043;
      display: inline-flex;
      align-items: center;
      justify-content: space-between;
      gap: 2px;
      cursor: pointer;
      font-size: 13px;
    }
    .zoom-button:hover { background: #f8f9fa; }
    .zoom-button svg { width: 14px; height: 14px; color: #5f6368; }

    .tabs {
      flex: 1;
      min-width: 0;
      display: flex;
      align-items: stretch;
      overflow-x: auto;
      scrollbar-width: none;
      margin-left: 6px;
    }
    .tabs::-webkit-scrollbar { display: none; }

    .tab {
      border: none;
      border-bottom: 2px solid transparent;
      background: transparent;
      color: #5f6368;
      padding: 0 11px;
      font-size: 13.5px;
      font-weight: 500;
      white-space: nowrap;
      cursor: pointer;
    }
    .tab:hover { color: #202124; }
    .tab[aria-selected='true'] {
      color: #1a73e8;
      border-bottom-color: #1a73e8;
    }

    .workspace {
      flex: 1;
      min-height: 0;
      display: flex;
    }

    .viewer {
      position: relative;
      flex: 1;
      min-width: 0;
      overflow: auto;
      background: #e8eaed;
      overscroll-behavior: contain;
    }
    .viewer:focus { outline: none; }
    .viewer:focus-visible { outline: 2px solid #1a73e8; outline-offset: -2px; }
    .viewer.hand { cursor: grab; touch-action: none; }
    .viewer.hand.panning { cursor: grabbing; }

    .viewer::-webkit-scrollbar { width: 12px; height: 12px; }
    .viewer::-webkit-scrollbar-thumb {
      background: #c5c9ce;
      border-radius: 8px;
      border: 3px solid #e8eaed;
    }
    .viewer::-webkit-scrollbar-track { background: transparent; }

    .stage {
      box-sizing: border-box;
      display: flex;
      align-items: center;
      justify-content: center;
      width: max-content;
      min-width: 100%;
      height: max-content;
      min-height: 100%;
      padding: 28px 36px 40px;
    }
    .stage.is-flow {
      flex-direction: column;
      align-items: center;
      justify-content: flex-start;
      gap: 18px;
    }
    .sheet-row {
      display: flex;
      align-items: flex-start;
      justify-content: center;
      gap: 16px;
    }
    .sheet-hit {
      position: absolute;
      inset: 0;
      z-index: 4;
      padding: 0;
      border: none;
      background: transparent;
      cursor: pointer;
    }
    .sheet-hit:hover,
    .sheet-hit:focus-visible { box-shadow: inset 0 0 0 2px #1a73e8; }
    .viewer.hand .sheet-hit { cursor: grab; }

    .page-frame {
      position: relative;
      flex: none;
      background: #fff;
      line-height: 0;
      box-shadow:
        0 0 0 1px rgba(60, 64, 67, 0.06),
        0 1px 4px rgba(60, 64, 67, 0.22);
    }
    .page-frame.is-hidden { display: none; }
    .page-canvas { display: block; }
    .page-bookmark {
      position: absolute;
      top: 8px;
      right: 8px;
      z-index: 8;
      display: grid;
      place-items: center;
      width: 28px;
      height: 28px;
      padding: 0;
      border: none;
      border-radius: 6px;
      background: rgba(255, 255, 255, 0.94);
      color: #5f6368;
      cursor: pointer;
      box-shadow: 0 1px 3px rgba(60, 64, 67, 0.28);
    }
    .page-bookmark.marked { color: #1a73e8; }
    .page-bookmark:hover { background: #fff; }
    .page-bookmark svg { width: 16px; height: 16px; display: block; }

    .highlights {
      position: absolute;
      inset: 0;
      pointer-events: none;
    }
    .text-layer { position: absolute; inset: 0; }
    .text-line {
      position: absolute;
      margin: 0;
      padding: 0;
      border: 0;
      background: transparent;
      color: transparent;
      font-family: Helvetica, Arial, sans-serif;
      line-height: 1;
      text-align: left;
      white-space: nowrap;
      overflow: hidden;
      cursor: text;
    }
    .text-line:hover, .text-line:focus-visible {
      outline: 1px solid #1a73e8;
      background: rgba(26, 115, 232, 0.12);
    }
    .field-layer { position: absolute; inset: 0; }
    .field-target {
      position: absolute;
      margin: 0;
      padding: 0;
      border: 0;
      background: transparent;
      cursor: pointer;
    }
    .field-target:hover, .field-target:focus-visible,
    .field-target.highlighted {
      outline: 1px solid #1a73e8;
      background: rgba(26, 115, 232, 0.12);
    }
    .hit {
      position: absolute;
      background: rgba(255, 214, 0, 0.45);
      border-radius: 2px;
    }
    .hit.active { background: rgba(255, 145, 0, 0.55); }

    .marks { position: absolute; inset: 0; pointer-events: none; }
    .ink {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      overflow: visible;
      pointer-events: none;
    }
    .mark, .note-pin, .draft, .note-box { pointer-events: auto; }
    .mark.freetext, .mark.callout, .mark.text, .mark.stamp, .mark.distance, .mark.perimeter, .mark.area, .mark.initials, .mark.date {
      width: max-content;
      max-width: 16em;
      padding: 1px 4px;
      color: #202124;
      font-family: Arial, sans-serif;
      line-height: 1.3;
      white-space: pre-wrap;
      background: rgba(255, 255, 255, 0.92);
    }
    .mark.distance, .mark.perimeter, .mark.area {
      color: #1a73e8;
      font-weight: 700;
      transform: translate(-50%, -50%);
    }
    .mark.initials, .mark.date {
      color: #1a73e8;
      font-weight: 700;
    }
    .mark.initials { font-family: "Segoe Script", "Brush Script MT", cursive; }
    .mark.stamp {
      border: 2px solid #d93025;
      color: #d93025;
      font-weight: 700;
      letter-spacing: 0.04em;
      text-transform: uppercase;
    }
    .mark.link {
      border: 2px solid #1a73e8;
      background: transparent;
      cursor: pointer;
    }
    .mark.image { padding: 0; background: transparent; line-height: 0; }
    .mark.image img { display: block; width: 100%; height: 100%; pointer-events: none; }
    .stamp-menu {
      position: absolute;
      z-index: 3;
      display: flex;
      flex-direction: column;
      min-width: 140px;
      padding: 4px 0;
      border: 1px solid #e0e3e7;
      border-radius: 8px;
      background: #fff;
      box-shadow: 0 8px 24px rgba(60, 64, 67, 0.18);
    }
    .stamp-menu button {
      height: 32px;
      padding: 0 12px;
      border: none;
      background: transparent;
      color: #202124;
      font-size: 13px;
      text-align: left;
      cursor: pointer;
    }
    .stamp-menu button:hover { background: #f1f3f4; }
    .mark.callout {
      border: 1px solid #dadce0;
      border-radius: 4px;
    }
    .mark {
      position: absolute;
      box-sizing: border-box;
    }
    .mark.highlight { background: var(--mark-fill, rgba(246, 191, 38, 0.45)); }
    .mark.underline { border-bottom: var(--mark-weight, 2px) solid var(--mark-ink, #e37400); }
    .mark.strike::after {
      content: "";
      position: absolute;
      left: 0;
      right: 0;
      top: 50%;
      border-top: var(--mark-weight, 2px) solid var(--mark-ink, #d93025);
    }
    .mark.selected { outline: 2px solid #1a73e8; outline-offset: 1px; }
    .note-pin {
      position: absolute;
      width: 22px;
      height: 22px;
      margin: 0;
      padding: 0;
      border: none;
      border-radius: 4px 4px 4px 0;
      background: #f9ab00;
      color: #202124;
      font-size: 11px;
      font-weight: 700;
      cursor: pointer;
    }
    .note-pin.selected { outline: 2px solid #1a73e8; }
    .draft {
      position: absolute;
      border: 1px dashed #1a73e8;
      background: rgba(26, 115, 232, 0.12);
    }
    .note-box {
      position: absolute;
      z-index: 3;
      width: 220px;
      padding: 8px;
      border: 1px solid #e0e3e7;
      border-radius: 8px;
      background: #fff;
      box-shadow: 0 8px 24px rgba(60, 64, 67, 0.18);
    }
    .note-box textarea {
      width: 100%;
      min-height: 72px;
      resize: vertical;
      border: 1px solid #dadce0;
      border-radius: 4px;
      padding: 6px 8px;
      font-size: 13px;
    }
    .note-actions { display: flex; justify-content: flex-end; gap: 6px; margin-top: 8px; }
    .note-actions button {
      height: 28px;
      padding: 0 10px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      background: #fff;
      font-size: 12px;
      cursor: pointer;
    }
    .note-actions button[type='submit'] {
      border-color: #1a73e8;
      background: #1a73e8;
      color: #fff;
    }

    .annotate-bar {
      flex: none;
      display: flex;
      align-items: center;
      gap: 4px;
      height: 40px;
      padding: 0 8px;
      border-bottom: 1px solid #e0e3e7;
      background: #f8f9fa;
      overflow-x: auto;
    }
    .annotate-bar button {
      flex: none;
      height: 28px;
      padding: 0 10px;
      white-space: nowrap;
      border: 1px solid transparent;
      border-radius: 6px;
      background: transparent;
      color: #3c4043;
      font-size: 13px;
      cursor: pointer;
    }
    .note-actions button:disabled { opacity: 0.45; cursor: default; }

    .mark.note {
      background: rgba(249, 171, 0, 0.38);
      box-shadow: inset 0 -2px 0 #e37400;
    }
    .mark.note.pending { pointer-events: none; }

    .comment-link {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      overflow: visible;
      pointer-events: none;
      z-index: 2;
    }
    .comment-pop {
      position: absolute;
      z-index: 4;
      margin: 0;
      cursor: default;
      width: 272px;
      max-width: calc(100% - 16px);
      max-height: 360px;
      overflow: auto;
      padding: 12px;
      border: 1px solid #e0e3e7;
      border-top: 3px solid #e37400;
      border-radius: 10px;
      background: #fff;
      box-shadow: 0 10px 28px rgba(60, 64, 67, 0.2);
      line-height: normal;
    }
    .comment-person {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .avatar {
      width: 28px;
      height: 28px;
      flex: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 50%;
      background: #fde7c7;
      color: #8a4b08;
      font-size: 13px;
      font-weight: 700;
    }
    .avatar.small { width: 22px; height: 22px; font-size: 11px; }
    .person-name {
      display: block;
      color: #202124;
      font-size: 13px;
      font-weight: 600;
    }
    .person-time {
      display: block;
      margin-top: 1px;
      color: #80868b;
      font-size: 11px;
    }
    .comment-quote {
      display: block;
      margin: 10px 0 0;
      padding: 6px 8px;
      border-left: 3px solid #f0c36d;
      border-radius: 0 6px 6px 0;
      background: #fff8eb;
      color: #5f6368;
      font-size: 12px;
      line-height: 1.4;
      overflow-wrap: anywhere;
    }
    .comment-pop textarea {
      width: 100%;
      min-height: 72px;
      margin-top: 10px;
      resize: vertical;
      border: 1px solid #dadce0;
      border-radius: 6px;
      padding: 8px;
      font-size: 13px;
      line-height: 1.4;
      color: #202124;
    }
    .comment-pop textarea:focus {
      outline: none;
      border-color: #e37400;
      box-shadow: 0 0 0 2px rgba(227, 116, 0, 0.2);
    }
    .comment-body {
      display: block;
      margin: 10px 0 0;
      color: #202124;
      font-size: 13px;
      line-height: 1.4;
      white-space: pre-wrap;
    }
    .reply-list {
      list-style: none;
      margin: 10px 0 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .reply-list li {
      display: flex;
      gap: 8px;
      align-items: flex-start;
    }
    .reply-list p {
      margin: 4px 0 0;
      color: #202124;
      font-size: 13px;
      line-height: 1.35;
      white-space: pre-wrap;
    }
    .reply-form {
      display: flex;
      gap: 6px;
      margin-top: 10px;
    }
    .reply-form input {
      flex: 1;
      min-width: 0;
      height: 32px;
      padding: 0 10px;
      border: 1px solid #dadce0;
      border-radius: 16px;
      font-size: 13px;
      color: #202124;
    }
    .reply-form input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px #d2e3fc;
    }
    .reply-form button {
      flex: none;
      height: 32px;
      padding: 0 12px;
      border: none;
      border-radius: 16px;
      background: #1a73e8;
      color: #fff;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
    }
    .reply-form button:disabled { opacity: 0.45; cursor: default; }
    .annotate-bar button:hover { background: #e8eaed; }
    .annotate-bar button[aria-pressed='true'] {
      color: #1a73e8;
      background: #e8f0fe;
      border-color: #d2e3fc;
    }
    .viewer.markup { cursor: crosshair; }

    .comment-list { flex: 1; min-height: 0; overflow: auto; margin-top: 12px; }
    .comment-row {
      display: flex;
      align-items: flex-start;
      gap: 4px;
      border-bottom: 1px solid #eee;
    }
    .comment-row.selected { background: #e8f0fe; }
    .viewer.markup.note-tool { cursor: text; }
    .viewer.select-text { cursor: text; }

    .freetext-editor {
      position: absolute;
      z-index: 4;
      box-sizing: border-box;
      margin: 0;
      padding: 4px 6px;
      border: 1.5px dotted #5f6368;
      border-radius: 2px;
      background: rgba(255, 255, 255, 0.92);
      color: #202124;
      font-family: Arial, sans-serif;
      font-size: 14px;
      line-height: 1.35;
      resize: both;
      outline: none;
    }
    .freetext-editor::placeholder { color: #9aa0a6; }
    .mark.freetext.selected {
      border: 1.5px dotted #5f6368;
      background: rgba(255, 255, 255, 0.92);
    }
    .text-select {
      position: absolute;
      background: rgba(26, 115, 232, 0.2);
      border-radius: 2px;
      pointer-events: none;
    }
    .select-handle {
      position: absolute;
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #1a73e8;
      pointer-events: none;
    }
    .quick-bar, .quick-styles {
      position: absolute;
      z-index: 5;
      display: flex;
      align-items: center;
      gap: 4px;
      padding: 5px;
      border: 1px solid rgba(32, 33, 36, 0.08);
      border-radius: 16px;
      background: #fff;
      box-shadow:
        0 1px 2px rgba(32, 33, 36, 0.06),
        0 12px 28px rgba(32, 33, 36, 0.16);
      transform: translateX(-50%);
    }
    .quick-bar button, .quick-styles button {
      position: relative;
      height: 36px;
      border: none;
      border-radius: 11px;
      background: transparent;
      color: #3c4043;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      cursor: pointer;
    }
    .quick-bar button { width: 36px; padding: 0; }
    .quick-styles button {
      height: 34px;
      padding: 0 10px 0 8px;
      font-size: 13px;
      font-weight: 600;
      color: #202124;
    }
    .quick-sep {
      width: 1px;
      height: 18px;
      margin: 0 2px;
      background: #e6e8eb;
      flex: none;
    }
    .quick-bar .quick-comment { color: #e37400; }
    .quick-bar .quick-style { color: #7a4dff; }
    .quick-bar .quick-link { color: #1a73e8; }
    .quick-bar .quick-delete { color: #d93025; }
    .quick-bar button:hover,
    .quick-styles button:hover { background: #f4f6f8; }
    .quick-bar .quick-comment:hover { background: #fff4e5; }
    .quick-bar .quick-style:hover,
    .quick-bar .quick-style[aria-pressed='true'] { background: #f3eefe; }
    .quick-bar .quick-link:hover { background: #e8f0fe; }
    .quick-bar .quick-delete:hover:not(:disabled) { background: #fdecea; }
    .quick-bar button[aria-pressed='true'] { box-shadow: inset 0 0 0 1px currentColor; }
    .quick-bar button:disabled { color: #bdc1c6; cursor: default; }
    .quick-bar button:disabled:hover { background: transparent; }
    .quick-bar svg, .quick-styles svg { width: 18px; height: 18px; display: block; }
    .quick-bar button[data-tip]::after,
    .quick-bar button[data-tip]::before {
      position: absolute;
      left: 50%;
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.12s ease, transform 0.12s ease;
    }
    .quick-bar button[data-tip]::after {
      content: attr(data-tip);
      bottom: calc(100% + 10px);
      transform: translateX(-50%) translateY(4px);
      padding: 5px 8px;
      border-radius: 7px;
      background: #202124;
      color: #fff;
      font-size: 12px;
      font-weight: 600;
      line-height: 1;
      white-space: nowrap;
      box-shadow: 0 6px 16px rgba(32, 33, 36, 0.2);
    }
    .quick-bar button[data-tip]::before {
      content: "";
      bottom: calc(100% + 2px);
      transform: translateX(-50%) translateY(4px);
      border: 5px solid transparent;
      border-top-color: #202124;
    }
    .quick-bar button[data-tip]:hover::after,
    .quick-bar button[data-tip]:focus-visible::after,
    .quick-bar button[data-tip]:hover::before,
    .quick-bar button[data-tip]:focus-visible::before {
      opacity: 1;
      transform: translateX(-50%) translateY(0);
    }
    .quick-styles {
      flex-direction: column;
      align-items: stretch;
      min-width: 156px;
      padding: 6px;
    }
    .quick-styles button { justify-content: flex-start; }
    .style-swatch {
      width: 14px;
      height: 14px;
      border-radius: 4px;
      box-shadow: inset 0 0 0 1px rgba(32, 33, 36, 0.12);
    }
    .comment {
      flex: 1;
      min-width: 0;
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 2px;
      padding: 10px 8px;
      border: none;
      background: transparent;
      text-align: left;
      cursor: pointer;
      color: #202124;
    }
    .comment-kind { font-size: 12px; font-weight: 600; color: #1a73e8; }
    .comment-page { font-size: 11px; color: #80868b; }
    .comment-text { font-size: 13px; line-height: 1.35; white-space: pre-wrap; }
    .reply-count {
      margin-top: 6px;
      color: #80868b;
      font-size: 11px;
    }
    .comment-card {
      display: flex;
      align-items: flex-start;
      gap: 8px;
      padding: 10px 8px 12px;
      border-bottom: 1px solid #eee;
      border-left: 3px solid transparent;
    }
    .comment-card.selected { background: #f4f8fe; outline: 1px solid #8ab4f8; outline-offset: -1px; }
    .comment-card.is-marked:not(.tone-accepted):not(.tone-rejected):not(.tone-cancelled):not(.tone-completed) {
      border-left-color: #e37400;
      background: #fff8eb;
    }
    .comment-card.tone-accepted { border-left-color: #1e8e3e; background: #f2faf4; }
    .comment-card.tone-rejected { border-left-color: #d93025; background: #fdf3f2; }
    .comment-card.tone-cancelled { border-left-color: #80868b; background: #f5f6f7; }
    .comment-card.tone-completed { border-left-color: #1a73e8; background: #f2f6fe; }
    .comment-card.is-marked { box-shadow: inset 0 0 0 1px #f0c36d; }
    .comment-card .comment { padding: 0; }
    .comment-main { flex: 1; min-width: 0; }
    .comment-top { display: flex; align-items: flex-start; gap: 4px; }
    .comment-tools { display: flex; flex: none; gap: 2px; }
    .comment-menu-btn {
      width: 28px;
      height: 28px;
      padding: 0;
      border: 1px solid transparent;
      border-radius: 6px;
      background: transparent;
      color: #5f6368;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
    }
    .comment-menu-btn:hover,
    .comment-menu-btn[aria-expanded='true'] {
      background: #fff;
      border-color: #dadce0;
    }
    .status-dot {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      background: #dadce0;
      box-shadow: inset 0 0 0 1px rgba(32, 33, 36, 0.12);
    }
    .comment-menu {
      position: fixed;
      z-index: 40;
      min-width: 176px;
      margin: 0;
      padding: 6px;
      background: #fff;
      border: 1px solid #e0e3e7;
      border-radius: 10px;
      box-shadow: 0 10px 28px rgba(60, 64, 67, 0.18);
    }
    .comment-menu button {
      width: 100%;
      display: flex;
      align-items: center;
      gap: 8px;
      height: 32px;
      padding: 0 8px;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: #202124;
      font-size: 13px;
      text-align: left;
      cursor: pointer;
    }
    .comment-menu button:hover { background: #f1f3f4; }
    .comment-menu button[aria-checked='true'] { background: #eef3fd; }
    .comment-menu button.danger { color: #b3261e; }
    .comment-menu .swatch {
      width: 10px;
      height: 10px;
      flex: none;
      border-radius: 50%;
      box-shadow: inset 0 0 0 1px rgba(32, 33, 36, 0.12);
    }
    .comment-menu .swatch.marked { background: #e37400; }
    .menu-rule { height: 1px; margin: 4px 2px; background: #eee; border: none; }
    .chip-row { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 8px; }
    .status-chip {
      display: inline-flex;
      align-items: center;
      height: 18px;
      padding: 0 7px;
      border-radius: 9px;
      font-size: 11px;
      font-weight: 600;
    }
    .status-chip.tone-accepted { background: #e6f4ea; color: #137333; }
    .status-chip.tone-rejected { background: #fce8e6; color: #b3261e; }
    .status-chip.tone-cancelled { background: #f1f3f4; color: #5f6368; }
    .status-chip.tone-completed { background: #e8f0fe; color: #1967d2; }
    .status-chip.tone-marked { background: #fef3e2; color: #b06000; }
    .comment-pop.tone-accepted { border-top-color: #1e8e3e; }
    .comment-pop.tone-rejected { border-top-color: #d93025; }
    .comment-pop.tone-cancelled { border-top-color: #80868b; }
    .comment-pop.tone-completed { border-top-color: #1a73e8; }
    .comment-pop.is-marked { box-shadow: 0 10px 28px rgba(60, 64, 67, 0.2), inset 0 0 0 1px #f0c36d; }
    .comment-check { display: flex; padding-top: 2px; }
    .comment-check input { width: 15px; height: 15px; margin: 0; accent-color: #1a73e8; cursor: pointer; }
    .comment-edit { display: flex; flex-direction: column; gap: 8px; margin-top: 8px; }
    .comment-edit textarea {
      width: 100%;
      min-height: 64px;
      resize: vertical;
      border: 1px solid #dadce0;
      border-radius: 6px;
      padding: 6px 8px;
      font-size: 13px;
      color: #202124;
    }
    .comment-edit textarea:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px #d2e3fc;
    }
    .select-bar {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 10px;
      padding: 6px 8px;
      border-radius: 8px;
      background: #f8f9fa;
      color: #3c4043;
      font-size: 12px;
    }
    .select-bar label { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; }
    .select-bar input { width: 15px; height: 15px; margin: 0; accent-color: #1a73e8; }
    .select-bar button {
      margin-left: auto;
      height: 26px;
      padding: 0 10px;
      border: 1px solid #f3c1be;
      border-radius: 6px;
      background: #fff;
      color: #b3261e;
      font-size: 12px;
      cursor: pointer;
    }
    .select-bar button:disabled { opacity: 0.4; cursor: default; }
    .sort-actions .icon[aria-pressed='true'] {
      color: #1a73e8;
      background: #e8f0fe;
    }

    .filter-backdrop {
      position: absolute;
      inset: 0;
      z-index: 25;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 24px;
      background: rgba(32, 33, 36, 0.32);
    }
    .filter-dialog {
      width: 380px;
      max-width: 100%;
      max-height: 100%;
      display: flex;
      flex-direction: column;
      border-radius: 12px;
      background: #fff;
      box-shadow: 0 16px 48px rgba(60, 64, 67, 0.28);
      overflow: hidden;
    }
    .filter-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 14px 14px 10px 16px;
    }
    .filter-head h2 { margin: 0; font-size: 16px; font-weight: 600; }
    .filter-close {
      width: 32px;
      height: 32px;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: #5f6368;
      font-size: 20px;
      line-height: 1;
      cursor: pointer;
    }
    .filter-close:hover { background: #f1f3f4; }
    .filter-tabs {
      display: flex;
      gap: 2px;
      margin: 0 16px;
      padding: 3px;
      border: 1px solid #dadce0;
      border-radius: 8px;
    }
    .filter-tabs button {
      flex: 1;
      height: 30px;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: #3c4043;
      font-size: 12px;
      cursor: pointer;
    }
    .filter-tabs button[aria-selected='true'] {
      background: #1a73e8;
      color: #fff;
    }
    .filter-options {
      min-height: 148px;
      max-height: 220px;
      overflow: auto;
      margin: 8px 0 0;
      padding: 8px 16px 12px;
    }
    .filter-option {
      display: flex;
      align-items: center;
      gap: 8px;
      min-height: 32px;
      color: #202124;
      font-size: 13px;
      cursor: pointer;
    }
    .filter-option input { width: 15px; height: 15px; margin: 0; accent-color: #1a73e8; }
    .filter-swatch {
      width: 14px;
      height: 14px;
      border-radius: 50%;
      box-shadow: inset 0 0 0 1px rgba(32, 33, 36, 0.16);
    }
    .filter-empty { margin: 28px 0; color: #80868b; font-size: 13px; text-align: center; }
    .filter-settings {
      padding: 12px 16px;
      border-top: 1px solid #eee;
    }
    .filter-settings h3 {
      margin: 0 0 8px;
      color: #5f6368;
      font-size: 12px;
      font-weight: 600;
      text-align: center;
    }
    .filter-foot {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px 16px 16px;
      border-top: 1px solid #eee;
    }
    .filter-foot .text-btn {
      border: none;
      background: transparent;
      color: #1a73e8;
      font-size: 13px;
      cursor: pointer;
    }
    .filter-foot .text-btn:hover { text-decoration: underline; }
    .filter-foot .apply-btn {
      height: 32px;
      padding: 0 16px;
      border: none;
      border-radius: 6px;
      background: #1a73e8;
      color: #fff;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
    }

    .pages {
      position: relative;
      width: 248px;
      flex: none;
      min-height: 0;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      background: #f8f9fa;
      border-right: 1px solid #e0e3e7;
    }
    .pane-tabs {
      display: flex;
      gap: 2px;
      flex: none;
      padding: 8px 8px 0;
      border-bottom: 1px solid #e0e3e7;
      background: #fff;
    }
    .pane-tabs button {
      display: grid;
      place-items: center;
      width: 36px;
      height: 34px;
      border: none;
      border-bottom: 2px solid transparent;
      border-radius: 8px 8px 0 0;
      background: transparent;
      color: #5f6368;
      cursor: pointer;
    }
    .pane-tabs button:hover { background: #f1f3f4; color: #202124; }
    .pane-tabs button:focus-visible { outline: 2px solid #1a73e8; outline-offset: -2px; }
    .pane-tabs button[aria-selected='true'] {
      background: #e8f0fe;
      color: #1a73e8;
      border-bottom-color: #1a73e8;
    }
    .pane-tabs svg { width: 18px; height: 18px; display: block; }
    .pane-size {
      display: flex;
      align-items: center;
      gap: 6px;
      flex: none;
      padding: 8px 10px 2px;
    }
    .pane-size button {
      display: grid;
      place-items: center;
      width: 28px;
      height: 28px;
      border: none;
      border-radius: 8px;
      background: transparent;
      color: #5f6368;
      cursor: pointer;
    }
    .pane-size button:hover { background: #e8eaed; }
    .pane-size svg { width: 16px; height: 16px; display: block; }
    .pane-size input[type='range'] {
      flex: 1;
      min-width: 0;
      accent-color: #1a73e8;
    }
    .pane-list {
      flex: 1;
      min-height: 0;
      overflow: auto;
      padding: 8px 10px 16px;
    }
    .pane-soon {
      flex: 1;
      padding: 28px 18px;
      text-align: center;
    }
    .pane-soon h2 {
      margin: 0 0 6px;
      font-size: 15px;
      font-weight: 650;
    }
    .pane-soon p {
      margin: 0;
      color: #80868b;
      font-size: 13px;
      line-height: 1.45;
    }
    .pane-empty {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 14px;
      padding: 36px 24px;
      color: #80868b;
      text-align: center;
    }
    .pane-empty svg { width: 52px; height: 52px; }
    .pane-empty p { margin: 0; max-width: 180px; font-size: 13px; line-height: 1.4; }
    .layer-row {
      display: flex;
      align-items: flex-start;
      gap: 8px;
      padding: 7px 4px;
      color: #202124;
      font-size: 13px;
      line-height: 1.35;
      cursor: pointer;
    }
    .layer-row input { margin-top: 2px; accent-color: #1a73e8; }
    .file-row {
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 2px;
      width: 100%;
      padding: 8px 6px;
      border: none;
      border-radius: 8px;
      background: transparent;
      text-align: left;
      cursor: pointer;
      font: inherit;
    }
    .file-row:hover { background: #e8f0fe; }
    .file-name { color: #202124; font-size: 13px; }
    .file-meta { color: #80868b; font-size: 11px; }
    .pane-section-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex: none;
      padding: 12px 12px 4px;
    }
    .pane-section-head h2 { margin: 0; font-size: 15px; font-weight: 650; }
    .pane-section-head button,
    .pane-add {
      border: none;
      background: transparent;
      color: #1a73e8;
      font: inherit;
      font-size: 13px;
      font-weight: 650;
      cursor: pointer;
    }
    .pane-section-head button:hover,
    .pane-add:hover { text-decoration: underline; }
    .pane-toggle {
      display: flex;
      align-items: center;
      gap: 10px;
      flex: none;
      margin: 8px 12px 0;
      padding-bottom: 10px;
      border-bottom: 1px solid #e0e3e7;
      color: #202124;
      font-size: 13px;
    }
    .outline-row {
      display: flex;
      align-items: flex-start;
      gap: 2px;
      width: 100%;
    }
    .outline-twist {
      width: 18px;
      height: 28px;
      flex: none;
      border: none;
      background: transparent;
      color: #5f6368;
      cursor: pointer;
    }
    .outline-link,
    .bookmark-main {
      flex: 1;
      min-width: 0;
      border: none;
      background: transparent;
      text-align: left;
      cursor: pointer;
      color: #202124;
      font: inherit;
    }
    .outline-link {
      padding: 6px 4px;
      border-radius: 6px;
      font-size: 13px;
      line-height: 1.35;
    }
    .outline-link:hover:not(:disabled),
    .bookmark-main:hover { background: #e8f0fe; }
    .outline-link:disabled { color: #80868b; cursor: default; }
    .outline-link[aria-current='page'] { color: #1a73e8; font-weight: 650; }
    .bookmark-row { display: flex; align-items: flex-start; gap: 4px; }
    .bookmark-main {
      display: flex;
      flex-direction: column;
      gap: 1px;
      padding: 6px 6px;
      border-radius: 8px;
    }
    .bookmark-page { color: #80868b; font-size: 11px; font-weight: 650; }
    .bookmark-title { font-size: 13px; }
    .bookmark-empty { margin: 22px 8px; color: #80868b; font-size: 13px; text-align: center; }
    .bookmark-remove {
      width: 28px;
      height: 28px;
      flex: none;
      margin-top: 8px;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: #d93025;
      cursor: pointer;
    }
    .bookmark-remove:hover { background: #fdecea; }
    .bookmark-remove svg { width: 16px; height: 16px; display: block; }
    .bookmark-form {
      margin: 4px 2px 12px;
      padding: 12px;
      border-radius: 10px;
      background: #f1f3f4;
    }
    .bookmark-form p { margin: 0 0 8px; text-align: center; font-size: 13px; font-weight: 650; }
    .bookmark-form input {
      width: 100%;
      height: 34px;
      border: 1px solid #1a73e8;
      border-radius: 6px;
      padding: 0 8px;
      background: #fff;
      color: #202124;
      font: inherit;
      font-size: 13px;
    }
    .bookmark-form-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 8px; }
    .bookmark-form-actions button {
      height: 32px;
      padding: 0 12px;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: #1a73e8;
      font: inherit;
      font-size: 13px;
      font-weight: 650;
      cursor: pointer;
    }
    .bookmark-form-actions .primary { background: #1a73e8; color: #fff; }
    .bookmark-form-actions .primary:disabled { background: #e8eaed; color: #80868b; cursor: default; }
    .pane-add {
      flex: none;
      height: 44px;
      border-top: 1px solid #e0e3e7;
      background: #fff;
    }
    .pick-bar {
      display: flex;
      justify-content: center;
      gap: 2px;
      flex: none;
      margin: 0 10px 8px;
      padding: 4px;
      border-radius: 10px;
      background: #e8eaed;
    }
    .pick-bar button {
      display: grid;
      place-items: center;
      width: 32px;
      height: 32px;
      border: none;
      border-radius: 8px;
      background: transparent;
      color: #3c4043;
      cursor: pointer;
    }
    .pick-bar button:hover:not(:disabled) { background: #fff; }
    .pick-bar button:disabled { color: #bdc1c6; cursor: default; }
    .pick-bar svg { width: 16px; height: 16px; display: block; }
    .pane-foot {
      flex: none;
      padding: 8px 12px 12px;
      border-top: 1px solid #e0e3e7;
      background: #fff;
    }
    .pane-foot p {
      margin: 0 0 6px;
      color: #5f6368;
      font-size: 12px;
      text-align: center;
    }
    .pane-note {
      margin: 0 0 6px;
      color: #d93025;
      font-size: 12px;
      text-align: center;
    }
    .pane-foot-row { display: flex; gap: 6px; }
    .pane-foot input {
      flex: 1;
      min-width: 0;
      height: 32px;
      border: 1px solid #dadce0;
      border-radius: 8px;
      padding: 0 10px;
      background: #fff;
      color: #202124;
      font: inherit;
      font-size: 13px;
    }
    .pane-foot button {
      display: grid;
      place-items: center;
      width: 32px;
      height: 32px;
      flex: none;
      border: 1px solid #dadce0;
      border-radius: 8px;
      background: #fff;
      color: #1a73e8;
      cursor: pointer;
    }
    .pane-foot button:hover { background: #e8f0fe; }
    .pane-foot svg { width: 16px; height: 16px; display: block; }
    .page-ops {
      flex: none;
      z-index: 2;
      margin: 0 8px 8px;
      max-height: 240px;
      overflow: auto;
      padding: 4px 0;
      border: 1px solid #e0e3e7;
      border-radius: 10px;
      background: #fff;
      box-shadow: 0 8px 24px rgba(60, 64, 67, 0.18);
    }
    .page-ops .view-heading { padding: 6px 12px 2px; }
    .page-ops button {
      display: block;
      width: 100%;
      height: 32px;
      padding: 0 12px;
      border: none;
      background: transparent;
      color: #202124;
      font-size: 13px;
      text-align: left;
      cursor: pointer;
    }
    .page-ops button:hover:not(:disabled) { background: #f1f3f4; }
    .page-ops button:disabled { color: #bdc1c6; cursor: default; }
    .page-ops button.danger { color: #d93025; }
    .page-ops .rule { height: 1px; margin: 4px 0; background: #e0e3e7; }
    .style-panel {
      width: 252px;
      flex: none;
      min-height: 0;
      overflow: auto;
      display: flex;
      flex-direction: column;
      gap: 14px;
      padding: 14px 14px 20px;
      background: #fff;
      border-right: 1px solid #e0e3e7;
      color: #202124;
    }
    .style-head {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 8px;
    }
    .style-head h2 {
      margin: 0;
      font-size: 15px;
      font-weight: 650;
    }
    .style-head p {
      margin: 2px 0 0;
      color: #80868b;
      font-size: 12px;
    }
    .style-head .icon { width: 28px; height: 28px; }
    .style-kicker {
      margin: 0 0 8px;
      color: #80868b;
      font-size: 11px;
      font-weight: 650;
      letter-spacing: 0.06em;
      text-transform: uppercase;
    }
    .style-kinds, .style-weights {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }
    .style-kinds button, .style-weights button, .style-actions button, .style-more-toggle {
      border: 1px solid #e0e3e7;
      border-radius: 8px;
      background: #fff;
      color: #202124;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
    }
    .style-kinds button, .style-weights button {
      height: 30px;
      padding: 0 10px;
    }
    .style-kinds button[aria-checked='true'],
    .style-weights button[aria-checked='true'] {
      border-color: #1a73e8;
      background: #e8f0fe;
      color: #1a73e8;
    }
    .style-chips {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 8px;
    }
    .style-chip {
      height: 28px;
      padding: 0;
      border: 2px solid transparent;
      border-radius: 8px;
      cursor: pointer;
      box-shadow: inset 0 0 0 1px rgba(32, 33, 36, 0.12);
    }
    .style-chip[aria-checked='true'] {
      border-color: #1a73e8;
      box-shadow: inset 0 0 0 2px #fff;
    }
    .style-custom {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 8px;
      color: #5f6368;
      font-size: 12px;
      font-weight: 600;
    }
    .style-custom input {
      width: 28px;
      height: 28px;
      padding: 0;
      border: none;
      background: transparent;
      cursor: pointer;
    }
    .style-opacity {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .style-opacity input[type='range'] { flex: 1; accent-color: #1a73e8; }
    .style-opacity input[type='number'] {
      width: 58px;
      height: 30px;
      padding: 0 6px;
      border: 1px solid #dadce0;
      border-radius: 8px;
      color: #202124;
      font-size: 12px;
    }
    .style-actions { display: flex; gap: 8px; }
    .style-actions button {
      flex: 1;
      height: 32px;
    }
    .style-actions button:hover:not(:disabled),
    .style-kinds button:hover,
    .style-weights button:hover,
    .style-more-toggle:hover { background: #f4f6f8; }
    .style-actions button:disabled { color: #bdc1c6; cursor: default; }
    .style-actions .danger { color: #d93025; }
    .style-actions .danger:disabled { color: #bdc1c6; }
    .style-more-toggle {
      width: 100%;
      height: 32px;
      text-align: left;
      padding: 0 10px;
    }
    .style-more { display: flex; flex-direction: column; gap: 12px; margin-top: 12px; }
    .pages-empty {
      margin: 24px 8px;
      color: #80868b;
      font-size: 12px;
      text-align: center;
    }
    .thumb-card { margin: 0 0 12px; }
    .thumb {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 6px;
      width: 100%;
      padding: 0;
      border: none;
      background: transparent;
      color: #5f6368;
      font-size: 12px;
      cursor: pointer;
    }
    .thumb-frame {
      background: #fff;
      line-height: 0;
      box-shadow: 0 0 0 1px rgba(60, 64, 67, 0.18);
    }
    .thumb[aria-current='page'] .thumb-frame,
    .thumb-card.picked .thumb-frame {
      box-shadow: 0 0 0 2px #1a73e8;
    }
    .thumb[aria-current='page'] .thumb-num { color: #1a73e8; font-weight: 700; }
    .thumb-card.picked .thumb-num::before {
      content: "";
      display: inline-block;
      width: 6px;
      height: 6px;
      margin-right: 5px;
      border-radius: 50%;
      background: #1a73e8;
      vertical-align: 1px;
    }
    .thumb canvas { display: block; height: auto; }
    .thumb-actions {
      display: none;
      justify-content: center;
      gap: 2px;
      margin-top: 4px;
    }
    .thumb-card:hover .thumb-actions,
    .thumb-card.current .thumb-actions,
    .thumb-card.picked .thumb-actions,
    .thumb-card.menu .thumb-actions { display: flex; }
    .thumb-actions button {
      display: grid;
      place-items: center;
      width: 28px;
      height: 28px;
      border: none;
      border-radius: 8px;
      background: transparent;
      color: #5f6368;
      cursor: pointer;
    }
    .thumb-actions button:hover:not(:disabled) { background: #e8eaed; }
    .thumb-actions button:disabled { color: #bdc1c6; cursor: default; }
    .thumb-actions .danger { color: #d93025; }
    .thumb-actions .danger:hover:not(:disabled) { background: #fdecea; }
    .thumb-actions svg { width: 16px; height: 16px; display: block; }

    .status {
      position: absolute;
      inset: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 24px;
      text-align: center;
      color: #5f6368;
      font-size: 14px;
      pointer-events: none;
    }
    .status.error { color: #b3261e; }

    .drop-hint {
      position: absolute;
      inset: 16px;
      border: 2px dashed #1a73e8;
      border-radius: 8px;
      background: rgba(26, 115, 232, 0.08);
      color: #1a73e8;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 15px;
      pointer-events: none;
    }

    .comments {
      width: 320px;
      flex: none;
      min-height: 0;
      display: flex;
      flex-direction: column;
      padding: 16px 16px 20px;
      border-left: 1px solid #e0e3e7;
      background: #fff;
    }

    .comment-search {
      display: flex;
      align-items: center;
      gap: 8px;
      height: 34px;
      padding: 0 10px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      color: #80868b;
    }
    .comment-search:focus-within {
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px #d2e3fc;
    }
    .comment-search svg { width: 16px; height: 16px; flex: none; }
    .comment-search input {
      flex: 1;
      min-width: 0;
      border: none;
      outline: none;
      font-size: 13px;
      color: #202124;
      background: transparent;
    }

    .comments h2 {
      margin: 18px 0 0;
      font-size: 15px;
      font-weight: 600;
    }

    .sort-row {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 14px;
      color: #5f6368;
      font-size: 13px;
    }
    .sort-row select {
      flex: 1;
      min-width: 0;
      max-width: 156px;
      height: 28px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      background: #fff;
      color: #202124;
      font-size: 13px;
      padding: 0 4px;
    }
    .sort-actions { margin-left: auto; display: flex; gap: 2px; }
    .sort-actions .icon { width: 28px; height: 28px; }
    .sort-actions svg { width: 16px; height: 16px; }

    .empty {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 14px;
      padding: 0 28px 40px;
      text-align: center;
      color: #80868b;
      font-size: 13px;
      line-height: 1.45;
    }
    .empty svg { width: 46px; height: 46px; }

    .popup {
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      z-index: 30;
      min-width: 196px;
      margin: 0;
      padding: 6px 0;
      list-style: none;
      background: #fff;
      border: 1px solid #e0e3e7;
      border-radius: 8px;
      box-shadow: 0 8px 24px rgba(60, 64, 67, 0.18);
    }
    .popup.right { left: auto; right: 0; }
    .popup button {
      width: 100%;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      padding: 7px 14px;
      border: none;
      background: transparent;
      color: #202124;
      font-size: 13px;
      text-align: left;
      cursor: pointer;
    }
    .popup button:hover:not(:disabled) { background: #f1f3f4; }
    .popup button:disabled { color: #9aa0a6; cursor: default; }
    .popup button[aria-checked='true'] { color: #1a73e8; }
    .menu-label {
      padding: 4px 14px 8px;
      color: #80868b;
      font-size: 12px;
    }
    .rule { height: 1px; margin: 6px 0; background: #eee; border: none; }

    .search-pop {
      position: absolute;
      top: calc(100% + 6px);
      right: 0;
      z-index: 30;
      padding: 8px;
      background: #fff;
      border: 1px solid #e0e3e7;
      border-radius: 8px;
      box-shadow: 0 8px 24px rgba(60, 64, 67, 0.18);
    }
    .search-pop input {
      width: 220px;
      height: 32px;
      padding: 0 10px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      font-size: 13px;
      color: #202124;
    }
    .search-pop input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px #d2e3fc;
    }
    .search-meta {
      display: flex;
      align-items: center;
      gap: 4px;
      margin-top: 8px;
      color: #5f6368;
      font-size: 12px;
    }
    .search-meta span { flex: 1; min-width: 0; }
    .match-nav {
      width: 28px;
      height: 28px;
      border: none;
      border-radius: 4px;
      background: transparent;
      color: #3c4043;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
    }
    .match-nav svg { width: 16px; height: 16px; }
    .match-nav.prev svg { transform: rotate(90deg); }
    .match-nav.next svg { transform: rotate(-90deg); }
    .match-nav:hover:not(:disabled) { background: #f1f3f4; }
    .match-nav:disabled { opacity: 0.35; cursor: default; }

    .file-input { display: none; }

    @media (max-width: 840px) {
      .comments { width: 260px; }
    }

    .word-stage {
      position: relative;
      flex: 1;
      min-height: 0;
      overflow: auto;
      background: #e7e9ee;
      display: flex;
      justify-content: safe center;
      align-items: flex-start;
      padding: 28px 20px 36px;
    }

    .word-stage:focus { outline: none; }

    .word-stage > .status {
      background: #e7e9ee;
      pointer-events: auto;
    }

    .word-frame {
      position: relative;
      flex: none;
      margin: 0 auto;
      background: #fff;
      box-shadow: 0 1px 2px rgba(32, 33, 36, 0.16), 0 10px 28px rgba(32, 33, 36, 0.14);
      overflow: hidden;
    }

    .word-pages {
      position: absolute;
      top: 0;
      left: 0;
      transform-origin: top left;
    }

    .word-pages .docx-wrapper {
      background: transparent !important;
      padding: 0 !important;
      display: block !important;
      width: max-content !important;
      min-height: 0 !important;
    }

    .word-pages section.docx {
      margin: 0 !important;
      box-shadow: none !important;
    }

    .word-pages section.docx[hidden] { display: none !important; }

    .word-pages img.preview-image {
      display: block;
      width: auto;
      height: auto;
      max-width: none;
      background: #fff;
    }

    .word-pages .pptx-preview-wrapper {
      background: transparent !important;
      width: max-content !important;
      height: auto !important;
      margin: 0 !important;
      overflow: visible !important;
    }

    .word-pages .preview-slide {
      margin: 0 !important;
    }

    .word-pages .preview-slide[hidden] { display: none !important; }

    .word-pages .preview-sheet {
      box-sizing: border-box;
      min-width: 640px;
      padding: 28px 32px 36px;
      background: #fff;
      color: #202124;
    }

    .word-pages .preview-sheet[hidden] { display: none !important; }

    .word-pages .sheet-name {
      margin: 0 0 16px;
      font-size: 15px;
      font-weight: 650;
    }

    .word-pages .sheet-table {
      border-collapse: collapse;
      font-size: 13px;
      line-height: 1.35;
    }

    .word-pages .sheet-table th,
    .word-pages .sheet-table td {
      border: 1px solid #e0e3e7;
      padding: 6px 10px;
      min-width: 28px;
      max-width: 280px;
      white-space: pre-wrap;
      vertical-align: top;
      text-align: left;
    }

    .word-pages .sheet-table th {
      min-width: 72px;
      background: #f8f9fa;
      color: #5f6368;
      font-weight: 650;
      text-align: center;
    }

    .word-pages .sheet-table tbody th { min-width: 36px; }

    .word-pages .sheet-table td {
      background: #fff;
      color: #202124;
    }

    .word-pages .sheet-note,
    .word-pages .sheet-empty {
      margin: 12px 0 0;
      color: #5f6368;
      font-size: 13px;
    }

    .word-bar {
      flex: none;
      height: 44px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 0 12px 0 16px;
      border-top: 1px solid #e0e3e7;
      background: #fff;
      z-index: 2;
    }

    .word-page {
      display: flex;
      align-items: center;
      gap: 8px;
      min-width: 0;
      color: #3c4043;
    }

    .word-page svg { width: 16px; height: 16px; flex: none; }

    .word-page-label {
      font-size: 12px;
      font-weight: 650;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      white-space: nowrap;
    }

    .word-actions { display: flex; align-items: center; gap: 2px; flex: none; }

    .word-zoom {
      height: 32px;
      min-width: 76px;
      padding: 0 6px 0 10px;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: #202124;
      font-size: 13px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 2px;
      cursor: pointer;
    }

    .word-zoom svg { width: 16px; height: 16px; }
    .word-zoom:hover:not(:disabled) { background: #f1f3f4; }
    .word-zoom:disabled { opacity: 0.38; cursor: default; }

    .word-menu-root { position: relative; display: flex; align-items: center; }

    .word-pop {
      position: absolute;
      right: 0;
      bottom: calc(100% + 8px);
      min-width: 220px;
      padding: 6px;
      background: #fff;
      border: 1px solid #e0e3e7;
      border-radius: 10px;
      box-shadow: 0 8px 24px rgba(32, 33, 36, 0.16);
      z-index: 6;
    }

    .word-pop button {
      width: 100%;
      display: flex;
      align-items: center;
      gap: 10px;
      text-align: left;
      border: none;
      background: transparent;
      border-radius: 6px;
      padding: 8px 10px;
      color: #202124;
      font-size: 13px;
      cursor: pointer;
    }

    .word-pop button svg { width: 16px; height: 16px; flex: none; }
    .word-pop button:hover:not(:disabled) { background: #f1f3f4; }
    .word-pop button:disabled { opacity: 0.4; cursor: default; }
    .word-pop .rule { margin: 4px 6px; }

    .settings-label {
      margin: 2px 2px 4px;
      font-size: 12px;
      font-weight: 650;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      color: #5f6368;
    }

    .preview-dark { background: #202124; }
    .preview-dark .word-stage { background: #202124; }
    .preview-dark .word-stage > .status { background: #202124; color: #e8eaed; }
    .preview-dark .word-bar { background: #292a2d; border-top-color: #3c4043; }
    .preview-dark .word-page,
    .preview-dark .word-zoom,
    .preview-dark .icon { color: #e8eaed; }
    .preview-dark .word-zoom:hover:not(:disabled),
    .preview-dark .icon:hover:not(:disabled) { background: #3c4043; }
    .preview-dark .word-pop { background: #292a2d; border-color: #3c4043; }
    .preview-dark .word-pop button { color: #e8eaed; }
    .preview-dark .word-pop button:hover:not(:disabled) { background: #3c4043; }
    .preview-dark .word-pop .rule { background: #3c4043; }
    .preview-dark .toolbar,
    .preview-dark .annotate-bar,
    .preview-dark .popup,
    .preview-dark .search-pop {
      background: #292a2d;
      border-color: #3c4043;
      color: #e8eaed;
    }
    .preview-dark .viewer,
    .preview-dark .pages { background: #202124; border-right-color: #3c4043; }
    .preview-dark .pane-tabs,
    .preview-dark .pane-foot,
    .preview-dark .page-ops { background: #292a2d; border-color: #3c4043; color: #e8eaed; }
    .preview-dark .pane-tabs { border-bottom-color: #3c4043; }
    .preview-dark .pane-tabs button { color: #9aa0a6; }
    .preview-dark .pane-tabs button:hover { background: #3c4043; color: #e8eaed; }
    .preview-dark .pane-tabs button[aria-selected='true'] { background: #174ea6; color: #fff; }
    .preview-dark .pane-size button,
    .preview-dark .thumb-actions button,
    .preview-dark .pick-bar button { color: #e8eaed; }
    .preview-dark .pane-size button:hover,
    .preview-dark .thumb-actions button:hover:not(:disabled) { background: #3c4043; }
    .preview-dark .pick-bar { background: #3c4043; }
    .preview-dark .pick-bar button:hover:not(:disabled) { background: #202124; }
    .preview-dark .pane-foot { border-top-color: #3c4043; }
    .preview-dark .pane-foot p,
    .preview-dark .pane-soon p,
    .preview-dark .pane-empty,
    .preview-dark .file-meta,
    .preview-dark .bookmark-empty,
    .preview-dark .bookmark-page,
    .preview-dark .thumb { color: #9aa0a6; }
    .preview-dark .pane-section-head h2,
    .preview-dark .pane-toggle,
    .preview-dark .outline-link,
    .preview-dark .bookmark-title { color: #e8eaed; }
    .preview-dark .outline-link:hover:not(:disabled),
    .preview-dark .bookmark-main:hover,
    .preview-dark .file-row:hover { background: #3c4043; }
    .preview-dark .layer-row,
    .preview-dark .file-name { color: #e8eaed; }
    .preview-dark .outline-link[aria-current='page'] { color: #8ab4f8; }
    .preview-dark .pane-toggle,
    .preview-dark .pane-add { border-color: #3c4043; }
    .preview-dark .pane-add,
    .preview-dark .bookmark-form { background: #292a2d; }
    .preview-dark .bookmark-form input {
      background: #202124;
      border-color: #8ab4f8;
      color: #e8eaed;
    }
    .preview-dark .page-bookmark { background: rgba(32, 33, 36, 0.92); color: #e8eaed; }
    .preview-dark .page-bookmark.marked { color: #8ab4f8; }
    .preview-dark .pane-foot input {
      background: #202124;
      border-color: #3c4043;
      color: #e8eaed;
    }
    .preview-dark .pane-foot button { background: #202124; border-color: #3c4043; }
    .preview-dark .page-ops button { color: #e8eaed; }
    .preview-dark .page-ops button:hover:not(:disabled) { background: #3c4043; }
    .preview-dark .page-ops button.danger { color: #f28b82; }
    .preview-dark .page-ops .rule { background: #3c4043; }
    .preview-dark .thumb-frame { background: #fff; }
    .preview-dark .style-panel {
      background: #292a2d;
      border-right-color: #3c4043;
      color: #e8eaed;
    }
    .preview-dark .style-head p,
    .preview-dark .style-kicker,
    .preview-dark .style-custom { color: #9aa0a6; }
    .preview-dark .style-kinds button,
    .preview-dark .style-weights button,
    .preview-dark .style-actions button,
    .preview-dark .style-more-toggle,
    .preview-dark .style-opacity input[type='number'] {
      background: #202124;
      border-color: #3c4043;
      color: #e8eaed;
    }
    .preview-dark .style-kinds button[aria-checked='true'],
    .preview-dark .style-weights button[aria-checked='true'] {
      background: #3c4043;
      border-color: #8ab4f8;
      color: #8ab4f8;
    }
    .preview-dark .style-actions .danger { color: #f28b82; }
    .preview-dark .style-kinds button:hover,
    .preview-dark .style-weights button:hover,
    .preview-dark .style-actions button:hover:not(:disabled),
    .preview-dark .style-more-toggle:hover { background: #3c4043; }
    .preview-dark .toolbar { border-bottom-color: #3c4043; }
    .preview-dark .sep,
    .preview-dark .popup .rule,
    .preview-dark .annotate-bar { background-color: #3c4043; }
    .preview-dark .annotate-bar { background: #292a2d; }
    .preview-dark .tab,
    .preview-dark .popup button,
    .preview-dark .zoom-button,
    .preview-dark .menu-label { color: #e8eaed; }
    .preview-dark .popup button:hover:not(:disabled),
    .preview-dark .tab:hover { background: #3c4043; }
    .preview-dark .icon[aria-pressed='true'] { color: #8ab4f8; background: #3c4043; }
    .popup.file-menu { min-width: 220px; }
    .popup.view-menu {
      min-width: 236px;
      max-height: min(72vh, 540px);
      overflow: auto;
    }
    .view-heading {
      margin: 0;
      padding: 8px 14px 2px;
      color: #80868b;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.06em;
      text-transform: uppercase;
    }
    .popup.view-menu button[aria-checked='true'] { background: #eef3fd; }
    .popup.view-menu button[aria-checked='true']:hover:not(:disabled) { background: #e4edfc; }
    .preview-dark .view-heading { color: #9aa0a6; }
    .preview-dark .popup button[aria-checked='true'] { color: #8ab4f8; }
    .preview-dark .popup.view-menu button[aria-checked='true'] { background: #30343a; }
    .preview-dark .popup.view-menu button[aria-checked='true']:hover:not(:disabled) { background: #3c4043; }
    .preview-dark .sheet-hit:hover,
    .preview-dark .sheet-hit:focus-visible { box-shadow: inset 0 0 0 2px #8ab4f8; }
    .menu-item { display: inline-flex; align-items: center; gap: 10px; }
    .popup button svg,
    .menu-item svg { width: 16px; height: 16px; flex: none; }

    .word-modal {
      position: absolute;
      inset: 0;
      z-index: 20;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 16px;
      background: rgba(32, 33, 36, 0.32);
    }

    .settings-shell {
      width: min(760px, 100%);
      height: min(540px, 100%);
      display: flex;
      flex-direction: column;
      background: #fff;
      color: #202124;
      border-radius: 14px;
      box-shadow: 0 18px 50px rgba(32, 33, 36, 0.28);
      overflow: hidden;
    }
    .settings-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px 16px;
      border-bottom: 1px solid #e8eaed;
    }
    .settings-head h2 { margin: 0; font-size: 18px; font-weight: 650; color: #202124; }
    .settings-close {
      width: 32px;
      height: 32px;
      border: none;
      border-radius: 8px;
      background: transparent;
      color: #5f6368;
      font-size: 22px;
      line-height: 1;
      cursor: pointer;
    }
    .settings-close:hover { background: #f1f3f4; color: #202124; }
    .settings-search { padding: 12px 16px 0; }
    .settings-search input {
      width: 100%;
      height: 36px;
      border: 1px solid #dadce0;
      border-radius: 8px;
      padding: 0 12px;
      color: #202124;
      font-size: 14px;
    }
    .settings-search input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px #d2e3fc;
    }
    .settings-body { flex: 1; min-height: 0; display: flex; }
    .settings-nav {
      width: 176px;
      flex: none;
      display: flex;
      flex-direction: column;
      gap: 4px;
      padding: 12px;
      border-right: 1px solid #e8eaed;
    }
    .settings-nav button {
      height: 36px;
      border: none;
      border-radius: 8px;
      background: transparent;
      color: #3c4043;
      padding: 0 10px;
      text-align: left;
      font-size: 14px;
      cursor: pointer;
    }
    .settings-nav button:hover { background: #f4f6f8; }
    .settings-nav button[aria-current='true'] { background: #e8f0fe; color: #1a73e8; font-weight: 650; }
    .settings-pane { flex: 1; min-width: 0; overflow: auto; padding: 16px 18px 20px; }
    .settings-pane h3 { margin: 0 0 6px; font-size: 15px; }
    .settings-note { margin: 0 0 14px; color: #5f6368; font-size: 13px; line-height: 1.4; }
    .settings-empty { margin: 24px 0; color: #80868b; font-size: 14px; }
    .theme-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; }
    .theme-card {
      width: 100%;
      padding: 0;
      border: 1px solid #dadce0;
      border-radius: 12px;
      background: #fff;
      overflow: hidden;
      cursor: pointer;
      text-align: left;
      color: inherit;
    }
    .theme-card:hover { border-color: #8ab4f8; }
    .theme-card[aria-pressed='true'] { border-color: #1a73e8; box-shadow: 0 0 0 1px #1a73e8; }
    .theme-preview { height: 96px; display: flex; flex-direction: column; }
    .theme-preview.light { background: #eef1f4; }
    .theme-preview.dark { background: #202124; }
    .theme-chrome { height: 16px; display: flex; align-items: center; gap: 4px; padding: 0 8px; }
    .theme-chrome i { width: 6px; height: 6px; border-radius: 50%; background: #c4c7cc; display: block; }
    .theme-preview.dark .theme-chrome i { background: #5f6368; }
    .theme-body { flex: 1; display: flex; min-height: 0; }
    .theme-side { width: 36%; }
    .theme-preview.light .theme-side { background: #fff; border-right: 1px solid #e4e7eb; }
    .theme-preview.dark .theme-side { background: #303134; }
    .theme-main { flex: 1; margin: 8px; border-radius: 4px; }
    .theme-preview.light .theme-main { background: #fff; }
    .theme-preview.dark .theme-main { background: #3c4043; }
    .theme-caption { display: flex; align-items: center; gap: 8px; padding: 8px 10px; font-size: 13px; font-weight: 600; }
    .theme-radio {
      width: 14px;
      height: 14px;
      border-radius: 50%;
      border: 2px solid #5f6368;
      box-sizing: border-box;
    }
    .theme-card[aria-pressed='true'] .theme-radio { border-color: #1a73e8; box-shadow: inset 0 0 0 3px #fff, inset 0 0 0 7px #1a73e8; }
    .shortcut-table { width: 100%; border-collapse: collapse; }
    .shortcut-table th {
      text-align: left;
      padding: 8px 10px;
      color: #5f6368;
      font-size: 12px;
      font-weight: 650;
      border-bottom: 1px solid #e8eaed;
    }
    .shortcut-table td { padding: 10px; font-size: 13px; border-bottom: 1px solid #f1f3f4; vertical-align: middle; }
    .shortcut-keys { font-weight: 650; white-space: nowrap; }
    .adv-row {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 16px;
      padding: 12px 0;
      border-bottom: 1px solid #eee;
    }
    .adv-row h3 { margin: 0 0 4px; font-size: 14px; }
    .adv-row p { margin: 0; max-width: 420px; color: #5f6368; font-size: 12px; line-height: 1.45; }
    .switch { position: relative; width: 40px; height: 22px; flex: none; }
    .switch input { position: absolute; opacity: 0; width: 0; height: 0; }
    .switch span {
      position: absolute;
      inset: 0;
      border-radius: 11px;
      background: #dadce0;
      cursor: pointer;
      transition: background 0.15s ease;
    }
    .switch span::after {
      content: "";
      position: absolute;
      top: 2px;
      left: 2px;
      width: 18px;
      height: 18px;
      border-radius: 50%;
      background: #fff;
      box-shadow: 0 1px 2px rgba(32, 33, 36, 0.28);
      transition: transform 0.15s ease;
    }
    .switch input:checked + span { background: #1a73e8; }
    .switch input:checked + span::after { transform: translateX(18px); }
    .switch input:focus-visible + span { box-shadow: 0 0 0 2px #d2e3fc; }
    .preview-dark .settings-shell { background: #292a2d; color: #e8eaed; }
    .preview-dark .settings-head,
    .preview-dark .settings-nav { border-color: #3c4043; }
    .preview-dark .settings-head h2,
    .preview-dark .settings-nav button,
    .preview-dark .shortcut-table td,
    .preview-dark .adv-row h3,
    .preview-dark .theme-caption { color: #e8eaed; }
    .preview-dark .settings-close { color: #e8eaed; }
    .preview-dark .settings-close:hover,
    .preview-dark .settings-nav button:hover { background: #3c4043; }
    .preview-dark .settings-nav button[aria-current='true'] { background: #394457; color: #8ab4f8; }
    .preview-dark .settings-search input { background: #202124; border-color: #3c4043; color: #e8eaed; }
    .preview-dark .settings-note,
    .preview-dark .adv-row p,
    .preview-dark .shortcut-table th { color: #9aa0a6; }
    .preview-dark .shortcut-table td,
    .preview-dark .adv-row { border-bottom-color: #3c4043; }
    .preview-dark .theme-card { background: #202124; border-color: #3c4043; }
    .preview-dark .theme-card[aria-pressed='true'] { border-color: #8ab4f8; box-shadow: 0 0 0 1px #8ab4f8; }

    .word-dialog {
      width: min(300px, 100%);
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 12px 40px rgba(32, 33, 36, 0.28);
      padding: 18px 18px 14px;
    }

    .word-dialog h2 {
      margin: 0 0 8px;
      font-size: 18px;
      font-weight: 650;
      color: #1a73e8;
    }

    .zoom-option {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 6px 2px;
      font-size: 14px;
      color: #202124;
      cursor: pointer;
    }

    .zoom-option input { accent-color: #1a73e8; margin: 0; }

    .word-dialog-actions {
      display: flex;
      justify-content: flex-end;
      gap: 8px;
      margin-top: 14px;
    }

    .word-dialog-actions button {
      min-width: 78px;
      height: 34px;
      border-radius: 6px;
      border: 1px solid #dadce0;
      background: #fff;
      color: #202124;
      font-size: 14px;
      cursor: pointer;
    }

    .word-dialog-actions button:hover { background: #f8f9fa; }

    .word-dialog-actions .primary {
      background: #1a73e8;
      border-color: #1a73e8;
      color: #fff;
    }

    .word-dialog-actions .primary:hover { background: #1967d2; }
    .word-dialog-actions button:disabled { opacity: 0.45; cursor: default; }

    .word-dialog.link-dialog,
    .word-dialog.link-confirm {
      width: min(440px, 100%);
      max-height: calc(100% - 32px);
      overflow: auto;
    }

    .link-field-label {
      margin: 14px 0 6px;
      font-size: 13px;
      font-weight: 650;
      color: #202124;
    }

    .link-bound {
      margin: 0;
      padding: 8px 10px;
      border-radius: 8px;
      background: #f4f8fe;
      color: #1a73e8;
      font-size: 14px;
      line-height: 1.4;
      text-decoration: underline;
      text-underline-offset: 2px;
      white-space: pre-wrap;
    }

    .link-bound.muted {
      background: #f8f9fa;
      color: #80868b;
      text-decoration: none;
    }

    .link-field {
      display: flex;
      flex-direction: column;
      gap: 4px;
      margin-top: 12px;
      color: #3c4043;
      font-size: 13px;
      font-weight: 600;
    }

    .link-field input {
      height: 36px;
      border: 1px solid #dadce0;
      border-radius: 8px;
      padding: 0 10px;
      color: #202124;
      font-size: 14px;
      font-weight: 400;
    }

    .link-field input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px #d2e3fc;
    }

    .link-confirm-copy {
      margin: 8px 0 0;
      color: #3c4043;
      font-size: 14px;
      font-weight: 400;
      line-height: 1.45;
      word-break: break-word;
    }

    :host(:fullscreen) { height: 100%; background: #e7e9ee; }
  `;

  /** URL of a PDF to open. Relative URLs resolve against the host page. */
  @property({ type: String })
  src = '';

  /** Name stored on new notes and replies. */
  @property({ type: String })
  author = 'Guest';

  @state() private mode: EditorMode = 'view';
  @state() private tool: PointerTool = 'select';
  @state() private zoomMode: ZoomMode = 'fit-page';
  @state() private zoomPercent = 1;
  @state() private scale = 1;
  @state() private currentPage = 1;
  @state() private totalPages = 0;
  @state() private loading = false;
  @state() private errorMessage = '';
  @state() private pageReady = false;
  @state() private documentKind: 'pdf' | 'word' | 'image' | 'slides' | 'sheets' = 'pdf';
  @state() private wordFit: WordFit = 'percent';
  @state() private wordPercent = 1;
  @state() private wordScaleLabel = 100;
  @state() private wordZoomOpen = false;
  @state() private wordZoomDraft: WordZoomChoice = 1;
  @state() private wordMenuOpen = false;
  @state() private wordFullScreen = false;
  @state() private previewSettingsOpen = false;
  @state() private settingsSection: SettingsSection = 'theme';
  @state() private settingsQuery = '';
  @state() private askBeforeLinks = true;
  @state() private showSelectionTools = true;
  @state() private highlightFormFields = true;
  @state() private viewMenuOpen = false;
  @state() private viewRead: ViewRead = 'page';
  @state() private viewSpread: ViewSpread = 'single';
  @state() private viewTurn = 0;
  @state() private previewTheme: 'light' | 'dark' = 'light';
  @state() private menuOpen = false;
  @state() private zoomOpen = false;
  @state() private searchOpen = false;
  @state() private commentsOpen = true;
  @state() private pagesPressed = false;
  @state() private pageTab: PagePaneTab = 'thumbs';
  @state() private thumbWidth = 112;
  @state() private pagePicks: number[] = [];
  @state() private pageListDraft = '';
  @state() private pageNote = '';
  @state() private pageMenu: { page: number; scope: 'one' | 'many' } | null = null;
  @state() private pageBusy = false;
  @state() private outlineItems: OutlineEntry[] = [];
  @state() private outlineOpen: string[] = [];
  @state() private pageBookmarks: PageBookmark[] = [];
  @state() private showBookmarksOnPage = false;
  @state() private bookmarkEditing = false;
  @state() private bookmarkDraft: { page: number; title: string } | null = null;
  @state() private documentLayers: DocumentLayer[] = [];
  @state() private documentFiles: DocumentFile[] = [];
  @state() private dragOver = false;
  @state() private panning = false;
  @state() private findQuery = '';
  @state() private commentQuery = '';
  @state() private sortBy: CommentSort = 'position';
  @state() private selectingComments = false;
  @state() private checkedCommentIds: string[] = [];
  @state() private commentMenuId: string | null = null;
  @state() private statusMenuId: string | null = null;
  @state() private menuAnchor: { top: number; left: number } | null = null;
  @state() private editingCommentId: string | null = null;
  @state() private editingCommentText = '';
  @state() private filterOpen = false;
  @state() private filterTab: FilterTab = 'user';
  @state() private appliedFilter: CommentFilterState = emptyCommentFilter();
  @state() private draftFilter: CommentFilterState = emptyCommentFilter();
  @state() private highlights: Array<ViewportBox & { active: boolean }> = [];
  @state() private matchIndex = -1;
  @state() private matchCount = 0;
  @state() private searching = false;
  @state() private marks: Mark[] = [];
  @state() private selectedMarkId: string | null = null;
  @state() private markupKind: MarkKind | null = null;
  @state() private shapeKind: ShapeKind | null = null;
  @state() private insertKind: InsertKind | null = null;
  @state() private measureKind: MeasureKind | null = null;
  @state() private editText = false;
  @state() private signKind: SignKind | null = null;
  @state() private formKind: FormKind | null = null;
  @state() private editDraft: EditDraft | null = null;
  @state() private lineBoxes: LineBox[] = [];
  @state() private fieldBoxes: FieldBox[] = [];
  @state() private formDraft: FormDraft | null = null;
  @state() private stampMenu: { page: number; pdfX: number; pdfY: number; left: number; top: number } | null = null;
  @state() private markBoxes: MarkBox[] = [];
  @state() private draftBox: ViewportBox | null = null;
  @state() private draftStroke: Array<{ x: number; y: number }> = [];
  @state() private draftArrow: ArrowEnd = 'none';
  @state() private polygonHover: { x: number; y: number } | null = null;
  @state() private noteDraft: NoteDraft | null = null;
  @state() private linkPageInput = '';
  @state() private linkConfirmId: string | null = null;
  @state() private notePreview: ViewportBox[] = [];
  @state() private freeTextBox: { left: number; top: number; width: number; height: number } | null = null;
  @state() private textSelection: QuickSelection | null = null;
  @state() private styleMenuOpen = false;
  @state() private styleMoreOpen = false;
  @state() private styleTouched = false;
  @state() private styleColor = '#f6bf26';
  @state() private styleOpacity = 0.45;
  @state() private styleWeight = 2;
  @state() private replyText = '';

  @query('.viewer') private viewerEl!: HTMLElement;
  @query('.page-canvas') private canvasEl!: HTMLCanvasElement;
  @query('.file-input') private fileInput!: HTMLInputElement;
  @query('.image-input') private imageInput!: HTMLInputElement;
  @query('.menu-root') private menuRoot!: HTMLElement;
  @query('.zoom-root') private zoomRoot!: HTMLElement;
  @query('.search-root') private searchRoot!: HTMLElement;
  @query('.view-root') private viewRoot!: HTMLElement;
  @query('.pane-list') private pageListEl!: HTMLElement | null;

  private pdf: PDFDocumentProxy | null = null;
  private renderTask: RenderTask | null = null;
  private sideTasks: RenderTask[] = [];
  private renderToken = 0;
  private sheetScrollLock = false;
  private scrollSpyFrame = 0;
  private loadId = 0;
  private dragDepth = 0;
  private resetScroll = false;
  private resizeObserver: ResizeObserver | null = null;
  private textRunCache = new Map<number, TextRun[]>();
  private markSeq = 0;
  private dragStart: { x: number; y: number } | null = null;
  private textSelectStart: { x: number; y: number } | null = null;
  private quickToken = 0;
  private selectionToken = 0;
  private dragToken = 0;
  private imagePoint: { page: number; x: number; y: number } | null = null;
  private pdfBytes: Uint8Array | null = null;
  private sourceUrl = '';
  private downloadName = 'document.pdf';
  private wordBytes: Uint8Array | null = null;
  private wordGeneration = 0;
  private imageUrl: string | null = null;
  @state() private saving = false;
  private textEdits: TextEdit[] = [];
  private panPoint: { x: number; y: number; left: number; top: number } | null = null;
  private textMatches: TextMatch[] = [];
  private searchToken = 0;
  private searchTimer = 0;
  private thumbObserver: IntersectionObserver | null = null;
  private thumbTasks = new Map<number, RenderTask>();
  private thumbsPainted = new Set<number>();
  private thumbsPending = new Set<number>();
  private thumbToken = 0;
  private outlineToken = 0;
  private layerToken = 0;
  private fileToken = 0;
  private layerConfig: Awaited<ReturnType<PDFDocumentProxy['getOptionalContentConfig']>> | null = null;

  /** 1-based page currently shown. */
  get page(): number {
    return this.currentPage;
  }

  /** Number of pages in the open document. Zero when nothing is open. */
  get pageCount(): number {
    return this.totalPages;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.addEventListener('keydown', this.onKeydown);
    window.addEventListener('pointerdown', this.onWindowPointerDown);
    document.addEventListener('fullscreenchange', this.onFullScreenChange);
  }

  override disconnectedCallback(): void {
    this.removeEventListener('keydown', this.onKeydown);
    window.removeEventListener('pointerdown', this.onWindowPointerDown);
    document.removeEventListener('fullscreenchange', this.onFullScreenChange);
    if (document.fullscreenElement === this) void document.exitFullscreen();
    this.resizeObserver?.disconnect();
    window.clearTimeout(this.searchTimer);
    this.cancelThumbnails();
    this.renderToken += 1;
    this.renderTask?.cancel();
    this.cancelSideSheets();
    if (this.scrollSpyFrame) cancelAnimationFrame(this.scrollSpyFrame);
    void this.pdf?.loadingTask.destroy();
    this.pdf = null;
    this.clearImagePreview();
    super.disconnectedCallback();
  }

  protected override firstUpdated(): void {
    this.resizeObserver = new ResizeObserver(() => {
      if (this.isReadOnlyPreview()) {
        this.layoutWord();
        return;
      }
      if (!this.pdf || this.zoomMode === 'percent') return;
      this.scheduleRender();
    });
    if (this.viewerEl) this.resizeObserver.observe(this.viewerEl);
  }

  protected override updated(changed: PropertyValues): void {
    if (changed.has('src') && this.src) {
      void this.load(this.src);
    }
    if (changed.has('documentKind')) {
      const target = this.isReadOnlyPreview()
        ? this.renderRoot.querySelector('.word-stage')
        : this.renderRoot.querySelector('.viewer');
      if (target instanceof HTMLElement) this.resizeObserver?.observe(target);
    }
    if (
      this.isReadOnlyPreview() &&
      (changed.has('documentKind') ||
        changed.has('currentPage') ||
        changed.has('wordFit') ||
        changed.has('wordPercent') ||
        changed.has('totalPages') ||
        changed.has('pageReady'))
    ) {
      this.layoutWord();
      requestAnimationFrame(() => this.layoutWord());
    }
    if (changed.has('wordZoomOpen') && this.wordZoomOpen) {
      void this.updateComplete.then(() => {
        const ok = this.renderRoot.querySelector('.word-dialog .primary');
        if (ok instanceof HTMLElement) ok.focus();
      });
    }
    if (changed.has('viewTurn') && changed.get('viewTurn') !== undefined) {
      this.thumbsPainted.clear();
      this.thumbsPending.clear();
      this.cancelThumbnails();
      if (this.pagesPressed) {
        void this.updateComplete.then(() => this.observeThumbnails());
      }
    }
    if (
      this.pdf &&
      !this.isReadOnlyPreview() &&
      (changed.has('viewRead') || changed.has('viewSpread') || changed.has('viewTurn'))
    ) {
      this.scheduleRender();
    }
    if (changed.has('pagesPressed') || changed.has('totalPages') || changed.has('thumbWidth') || (changed.has('pageTab') && this.pageTab === 'thumbs')) {
      this.thumbsPainted.clear();
      this.thumbsPending.clear();
      this.cancelThumbnails();
      if (this.pagesPressed) {
        void this.updateComplete.then(() => {
          this.observeThumbnails();
          this.scrollActiveThumb();
        });
      }
    }
    if (changed.has('currentPage') && this.pagesPressed) {
      void this.updateComplete.then(() => this.scrollActiveThumb());
    }
    if (changed.has('noteDraft') && this.noteDraft && !changed.get('noteDraft')) {
      void this.updateComplete.then(() => {
        const field = this.renderRoot.querySelector('.freetext-editor, .comment-pop textarea, .note-box textarea');
        if (field instanceof HTMLTextAreaElement) field.focus();
        this.renderRoot.querySelector('.comment-pop')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      });
    }
    if (changed.has('selectedMarkId') && this.selectedMarkId) {
      void this.updateComplete.then(() => {
        this.renderRoot.querySelector('.comment-card.selected')?.scrollIntoView({ block: 'nearest' });
        if (this.selectedNote) {
          this.renderRoot.querySelector('.comment-pop')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }
      });
    }
    if (changed.has('editingCommentId') && this.editingCommentId) {
      void this.updateComplete.then(() => {
        const field = this.renderRoot.querySelector('.comment-edit textarea');
        if (field instanceof HTMLTextAreaElement) field.focus();
      });
    }
    if (
      this.selectingComments &&
      (changed.has('checkedCommentIds') || changed.has('marks') || changed.has('commentQuery') || changed.has('selectingComments') || changed.has('sortBy'))
    ) {
      void this.updateComplete.then(() => this.syncSelectAll());
    }
    if (changed.has('formDraft') && this.formDraft && !changed.get('formDraft')) {
      void this.updateComplete.then(() => {
        const field = this.renderRoot.querySelector('.form-box textarea');
        if (field instanceof HTMLTextAreaElement) field.focus();
      });
    }
  }

  /** Open a PDF, or preview a Word document, presentation, spreadsheet, or image, from a URL, a File, or binary data. */
  async load(source: PdfSource): Promise<void> {
    if (imageFormat(source)) {
      await this.loadImage(source);
      return;
    }
    if (slideFormat(source)) {
      await this.loadSlides(source);
      return;
    }
    if (sheetFormat(source)) {
      await this.loadSheets(source);
      return;
    }
    if (wordFormat(source)) {
      await this.loadWord(source);
      return;
    }
    if (this.documentKind !== 'pdf') this.documentKind = 'pdf';
    this.wordBytes = null;
    this.clearImagePreview();
    this.wordMenuOpen = false;
    this.wordZoomOpen = false;
    const id = ++this.loadId;
    this.loading = true;
    this.errorMessage = '';
    this.pageReady = false;
    this.clearSearch();
    this.marks = [];
    this.selectedMarkId = null;
    this.markBoxes = [];
    this.noteDraft = null;
    this.linkPageInput = '';
    this.linkConfirmId = null;
    this.notePreview = [];
    this.freeTextBox = null;
    this.textSelection = null;
    this.styleMenuOpen = false;
    this.replyText = '';
    this.selectingComments = false;
    this.checkedCommentIds = [];
    this.commentMenuId = null;
    this.statusMenuId = null;
    this.menuAnchor = null;
    this.editingCommentId = null;
    this.editingCommentText = '';
    this.filterOpen = false;
    this.appliedFilter = emptyCommentFilter();
    this.draftFilter = emptyCommentFilter();
    this.selectionToken += 1;
    this.stampMenu = null;
    this.imagePoint = null;
    this.draftBox = null;
    this.draftStroke = [];
    this.draftArrow = 'none';
    this.polygonHover = null;
    this.dragStart = null;
    this.editDraft = null;
    this.formDraft = null;
    this.editText = false;
    this.formKind = null;
    this.lineBoxes = [];
    this.fieldBoxes = [];
    this.textEdits = [];
    this.pdfBytes = null;
    this.sourceUrl = '';
    this.textRunCache.clear();
    this.thumbsPainted.clear();
    this.thumbsPending.clear();
    this.cancelThumbnails();
    this.renderToken += 1;
    this.renderTask?.cancel();
    this.renderTask = null;
    this.cancelSideSheets();
    const previous = this.pdf;
    this.pdf = null;
    this.totalPages = 0;
    await previous?.loadingTask.destroy();

    try {
      const bytes = await this.readSource(source);
      if (id !== this.loadId) return;
      const opened = await this.openKept(bytes);
      if (id !== this.loadId) {
        await opened.loadingTask.destroy();
        return;
      }
      this.pdf = opened;
      this.downloadName = this.fileName(source);
      this.totalPages = opened.numPages;
      this.currentPage = 1;
      this.pagePicks = [];
      this.pageListDraft = '';
      this.pageNote = '';
      this.pageMenu = null;
      this.pageBookmarks = [];
      this.bookmarkDraft = null;
      this.bookmarkEditing = false;
      void this.refreshOutline();
      void this.refreshLayers();
      void this.refreshFiles();
      this.loading = false;
      this.resetScroll = true;
      this.dispatchEvent(
        new CustomEvent<PdfLoadDetail>('pdf-load', {
          detail: { pages: opened.numPages },
          bubbles: true,
          composed: true,
        }),
      );
      this.scheduleRender();
      if (this.findQuery.trim()) this.queueSearch();
    } catch (error) {
      if (id !== this.loadId) return;
      this.loading = false;
      this.pdf = null;
      this.totalPages = 0;
      this.pageReady = false;
      const message = error instanceof Error ? error.message : 'Could not open this PDF.';
      this.errorMessage = message;
      this.dispatchEvent(
        new CustomEvent<PdfErrorDetail>('pdf-error', {
          detail: { message },
          bubbles: true,
          composed: true,
        }),
      );
    }
  }

  private async loadWord(source: PdfSource): Promise<void> {
    const id = ++this.loadId;
    const format = wordFormat(source);
    this.documentKind = 'word';
    this.loading = true;
    this.errorMessage = '';
    this.pageReady = false;
    this.wordMenuOpen = false;
    this.wordZoomOpen = false;
    this.wordFit = 'percent';
    this.wordPercent = 1;
    this.wordScaleLabel = 100;
    this.wordZoomDraft = 1;
    this.totalPages = 0;
    this.currentPage = 1;
    this.wordBytes = null;
    this.clearImagePreview();
    this.clearSearch();
    this.releasePdf();

    try {
      const bytes = await this.sourceBytes(source);
      if (id !== this.loadId) return;
      if (!isZipDocument(bytes)) {
        throw new Error(
          format === 'doc'
            ? 'Older .doc files can’t be previewed. Save the document as .docx and open that file.'
            : 'Could not open this Word document.',
        );
      }
      this.wordBytes = bytes;
      this.downloadName = this.wordFileName(source);
      const generation = ++this.wordGeneration;
      await this.updateComplete;
      if (id !== this.loadId || this.documentKind !== 'word') return;
      await this.paintWord(generation);
      if (id !== this.loadId) return;
      this.loading = false;
      this.dispatchEvent(
        new CustomEvent<PdfLoadDetail>('pdf-load', {
          detail: { pages: this.totalPages },
          bubbles: true,
          composed: true,
        }),
      );
    } catch (error) {
      if (id !== this.loadId) return;
      this.loading = false;
      this.pageReady = false;
      this.wordBytes = null;
      this.totalPages = 0;
      const message = this.wordErrorText(error);
      this.errorMessage = message;
      this.dispatchEvent(
        new CustomEvent<PdfErrorDetail>('pdf-error', {
          detail: { message },
          bubbles: true,
          composed: true,
        }),
      );
    }
  }

  private async loadSlides(source: PdfSource): Promise<void> {
    const id = ++this.loadId;
    const format = slideFormat(source);
    this.documentKind = 'slides';
    this.loading = true;
    this.errorMessage = '';
    this.pageReady = false;
    this.wordMenuOpen = false;
    this.wordZoomOpen = false;
    this.wordFit = 'percent';
    this.wordPercent = 1;
    this.wordScaleLabel = 100;
    this.wordZoomDraft = 1;
    this.totalPages = 0;
    this.currentPage = 1;
    this.wordBytes = null;
    this.clearImagePreview();
    this.clearSearch();
    this.releasePdf();

    try {
      const bytes = await this.sourceBytes(source, 'Could not read this presentation.');
      if (id !== this.loadId) return;
      if (!isZipDocument(bytes)) {
        throw new Error(
          format === 'ppt'
            ? 'Older .ppt files can’t be previewed. Save the presentation as .pptx and open that file.'
            : 'Could not open this presentation.',
        );
      }
      this.wordBytes = bytes;
      this.downloadName = this.slideFileName(source);
      const generation = ++this.wordGeneration;
      await this.updateComplete;
      if (id !== this.loadId || this.documentKind !== 'slides') return;
      await this.paintSlides(generation);
      if (id !== this.loadId) return;
      this.loading = false;
      this.dispatchEvent(
        new CustomEvent<PdfLoadDetail>('pdf-load', {
          detail: { pages: this.totalPages },
          bubbles: true,
          composed: true,
        }),
      );
    } catch (error) {
      if (id !== this.loadId) return;
      this.loading = false;
      this.pageReady = false;
      this.wordBytes = null;
      this.totalPages = 0;
      const message = this.slideErrorText(error);
      this.errorMessage = message;
      this.dispatchEvent(
        new CustomEvent<PdfErrorDetail>('pdf-error', {
          detail: { message },
          bubbles: true,
          composed: true,
        }),
      );
    }
  }

  private async loadSheets(source: PdfSource): Promise<void> {
    const id = ++this.loadId;
    const format = sheetFormat(source);
    this.documentKind = 'sheets';
    this.loading = true;
    this.errorMessage = '';
    this.pageReady = false;
    this.wordMenuOpen = false;
    this.wordZoomOpen = false;
    this.wordFit = 'percent';
    this.wordPercent = 1;
    this.wordScaleLabel = 100;
    this.wordZoomDraft = 1;
    this.totalPages = 0;
    this.currentPage = 1;
    this.wordBytes = null;
    this.clearImagePreview();
    this.clearSearch();
    this.releasePdf();

    try {
      const bytes = await this.sourceBytes(source, 'Could not read this spreadsheet.');
      if (id !== this.loadId) return;
      if ((format === 'xlsx' || format === 'xlsm') && !isZipDocument(bytes)) {
        throw new Error('Could not open this spreadsheet.');
      }
      this.wordBytes = bytes;
      this.downloadName = this.sheetFileName(source);
      const generation = ++this.wordGeneration;
      await this.updateComplete;
      if (id !== this.loadId || this.documentKind !== 'sheets') return;
      await this.paintSheets(generation);
      if (id !== this.loadId) return;
      this.loading = false;
      this.dispatchEvent(
        new CustomEvent<PdfLoadDetail>('pdf-load', {
          detail: { pages: this.totalPages },
          bubbles: true,
          composed: true,
        }),
      );
    } catch (error) {
      if (id !== this.loadId) return;
      this.loading = false;
      this.pageReady = false;
      this.wordBytes = null;
      this.totalPages = 0;
      const message = this.sheetErrorText(error);
      this.errorMessage = message;
      this.dispatchEvent(
        new CustomEvent<PdfErrorDetail>('pdf-error', {
          detail: { message },
          bubbles: true,
          composed: true,
        }),
      );
    }
  }

  private async loadImage(source: PdfSource): Promise<void> {
    const format = imageFormat(source);
    const id = ++this.loadId;
    const generation = ++this.wordGeneration;
    this.documentKind = 'image';
    this.loading = true;
    this.errorMessage = '';
    this.pageReady = false;
    this.wordMenuOpen = false;
    this.wordZoomOpen = false;
    this.wordFit = 'percent';
    this.wordPercent = 1;
    this.wordScaleLabel = 100;
    this.wordZoomDraft = 1;
    this.totalPages = 0;
    this.currentPage = 1;
    this.wordBytes = null;
    this.clearImagePreview();
    this.clearSearch();
    this.releasePdf();
    let url: string | null = null;

    try {
      if (!format) throw new Error('Could not open this image.');
      const bytes = new Uint8Array(await this.sourceBytes(source, 'Could not read this image.'));
      if (id !== this.loadId) return;
      if (!looksLikeImage(bytes, format)) throw new Error('Could not open this image.');
      const type = format === 'png' ? 'image/png' : 'image/jpeg';
      url = URL.createObjectURL(new Blob([bytes], { type }));
      const image = await this.decodePreviewImage(url);
      if (id !== this.loadId || this.documentKind !== 'image' || generation !== this.wordGeneration) {
        URL.revokeObjectURL(url);
        return;
      }
      this.imageUrl = url;
      url = null;
      image.alt = this.imageFileName(source, format);
      image.draggable = false;
      this.wordBytes = bytes;
      this.downloadName = image.alt;
      await this.updateComplete;
      if (id !== this.loadId || this.documentKind !== 'image') return;
      const body = this.renderRoot.querySelector('.word-pages');
      if (!(body instanceof HTMLElement)) throw new Error('Could not open this image.');
      body.replaceChildren(image);
      this.totalPages = 1;
      this.currentPage = 1;
      this.pageReady = true;
      this.loading = false;
      await this.updateComplete;
      this.layoutWord();
      this.dispatchEvent(
        new CustomEvent<PdfLoadDetail>('pdf-load', {
          detail: { pages: 1 },
          bubbles: true,
          composed: true,
        }),
      );
    } catch (error) {
      if (url) URL.revokeObjectURL(url);
      if (id !== this.loadId) return;
      this.loading = false;
      this.pageReady = false;
      this.wordBytes = null;
      this.clearImagePreview();
      this.totalPages = 0;
      const message = this.imageErrorText(error);
      this.errorMessage = message;
      this.dispatchEvent(
        new CustomEvent<PdfErrorDetail>('pdf-error', {
          detail: { message },
          bubbles: true,
          composed: true,
        }),
      );
    }
  }

  private isReadOnlyPreview(): boolean {
    return this.documentKind === 'word' || this.documentKind === 'image' || this.documentKind === 'slides' || this.documentKind === 'sheets';
  }

  private clearImagePreview(): void {
    if (!this.imageUrl) return;
    URL.revokeObjectURL(this.imageUrl);
    this.imageUrl = null;
  }

  private decodePreviewImage(url: string): Promise<HTMLImageElement> {
    const image = new Image();
    image.className = 'preview-image';
    return new Promise((resolve, reject) => {
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error('Could not open this image.'));
      image.src = url;
    });
  }

  private imageFileName(source: PdfSource, format: ImageFormat): string {
    if (typeof source === 'string') {
      const part = decodeURIComponent(source.split(/[?#]/)[0]).split(/[/\\]/).pop() ?? '';
      if (/\.(png|jpe?g)$/i.test(part)) return part;
    }
    if (source instanceof File && /\.(png|jpe?g)$/i.test(source.name)) return source.name;
    return format === 'png' ? 'image.png' : 'image.jpg';
  }

  private imageErrorText(error: unknown): string {
    if (error instanceof Error && (error.message.startsWith('Could not') || error.message.includes('image'))) {
      return error.message;
    }
    return 'Could not open this image.';
  }

  private releasePdf(): void {
    this.renderToken += 1;
    this.renderTask?.cancel();
    this.renderTask = null;
    this.cancelSideSheets();
    const previous = this.pdf;
    this.pdf = null;
    void previous?.loadingTask.destroy();
  }

  private async sourceBytes(source: PdfSource, failure = 'Could not read this Word document.'): Promise<Uint8Array> {
    const read = await this.readSource(source);
    if (typeof read !== 'string') return read;
    const response = await fetch(read);
    if (!response.ok) throw new Error(failure);
    return new Uint8Array(await response.arrayBuffer());
  }

  private slideFileName(source: PdfSource): string {
    if (typeof source === 'string') {
      const part = decodeURIComponent(source.split(/[?#]/)[0]).split(/[/\\]/).pop() ?? '';
      if (/\.pptx?$/i.test(part)) return part;
    }
    if (source instanceof File && /\.pptx?$/i.test(source.name)) return source.name;
    return 'presentation.pptx';
  }

  private slideErrorText(error: unknown): string {
    if (
      error instanceof Error &&
      (error.message.startsWith('Could not') || error.message.includes('.ppt') || error.message.includes('no slides'))
    ) {
      return error.message;
    }
    return 'Could not open this presentation.';
  }

  private sheetFileName(source: PdfSource): string {
    if (typeof source === 'string') {
      const part = decodeURIComponent(source.split(/[?#]/)[0]).split(/[/\\]/).pop() ?? '';
      if (/\.xls[xm]?$/i.test(part)) return part;
    }
    if (source instanceof File && /\.xls[xm]?$/i.test(source.name)) return source.name;
    return 'workbook.xlsx';
  }

  private sheetErrorText(error: unknown): string {
    if (
      error instanceof Error &&
      (error.message.startsWith('Could not') || error.message.includes('.xls') || error.message.includes('no sheets'))
    ) {
      return error.message;
    }
    return 'Could not open this spreadsheet.';
  }

  private wordFileName(source: PdfSource): string {
    if (typeof source === 'string') {
      const part = decodeURIComponent(source.split(/[?#]/)[0]).split(/[/\\]/).pop() ?? '';
      if (/\.docx?$/i.test(part)) return part;
    }
    if (source instanceof File && /\.docx?$/i.test(source.name)) return source.name;
    return 'document.docx';
  }

  private wordErrorText(error: unknown): string {
    if (
      error instanceof Error &&
      (error.message.startsWith('Could not') || error.message.includes('.doc') || error.message.includes('no pages'))
    ) {
      return error.message;
    }
    return 'Could not open this Word document.';
  }

  private async paintWord(generation: number): Promise<void> {
    if (!this.wordBytes) throw new Error('Could not open this Word document.');
    const scratch = document.createElement('div');
    const count = await renderDocxPages(this.wordBytes, scratch);
    if (generation !== this.wordGeneration || this.documentKind !== 'word') return;
    const body = this.renderRoot.querySelector('.word-pages');
    if (!(body instanceof HTMLElement)) throw new Error('Could not open this Word document.');
    if (count < 1) throw new Error('This Word document has no pages to preview.');
    body.replaceChildren(...scratch.childNodes);
    this.totalPages = count;
    this.currentPage = 1;
    this.pageReady = true;
    await this.updateComplete;
    this.layoutWord();
  }

  private async paintSlides(generation: number): Promise<void> {
    if (!this.wordBytes) throw new Error('Could not open this presentation.');
    const scratch = document.createElement('div');
    const count = await renderSlidePages(this.wordBytes, scratch);
    if (generation !== this.wordGeneration || this.documentKind !== 'slides') return;
    const body = this.renderRoot.querySelector('.word-pages');
    if (!(body instanceof HTMLElement)) throw new Error('Could not open this presentation.');
    if (count < 1) throw new Error('This presentation has no slides to preview.');
    body.replaceChildren(...scratch.childNodes);
    this.totalPages = count;
    this.currentPage = 1;
    this.pageReady = true;
    await this.updateComplete;
    this.layoutWord();
  }

  private async paintSheets(generation: number): Promise<void> {
    if (!this.wordBytes) throw new Error('Could not open this spreadsheet.');
    const scratch = document.createElement('div');
    const count = renderSheetPages(this.wordBytes, scratch);
    if (generation !== this.wordGeneration || this.documentKind !== 'sheets') return;
    const body = this.renderRoot.querySelector('.word-pages');
    if (!(body instanceof HTMLElement)) throw new Error('Could not open this spreadsheet.');
    if (count < 1) throw new Error('This spreadsheet has no sheets to preview.');
    body.replaceChildren(...scratch.childNodes);
    this.totalPages = count;
    this.currentPage = 1;
    this.pageReady = true;
    await this.updateComplete;
    this.layoutWord();
  }

  private layoutWord(): void {
    if (!this.isReadOnlyPreview()) return;
    const stage = this.renderRoot.querySelector('.word-stage');
    const frame = this.renderRoot.querySelector('.word-frame');
    const host = this.renderRoot.querySelector('.word-pages');
    if (!(stage instanceof HTMLElement) || !(frame instanceof HTMLElement) || !(host instanceof HTMLElement)) return;
    if (stage.clientWidth < 40 || stage.clientHeight < 40) return;
    host.style.transform = 'none';
    if (this.documentKind === 'image') {
      const image = host.querySelector('img.preview-image');
      if (!(image instanceof HTMLImageElement)) return;
      this.applyPreviewScale(stage, frame, host, image.naturalWidth, image.naturalHeight);
      return;
    }
    const pages = this.previewPageNodes(host);
    if (pages.length === 0) return;
    const index = Math.min(Math.max(this.currentPage, 1), pages.length) - 1;
    pages.forEach((page, pageIndex) => {
      page.hidden = pageIndex !== index;
    });
    const page = pages[index];
    this.applyPreviewScale(stage, frame, host, page.offsetWidth, page.offsetHeight);
  }

  private previewPageNodes(host: ParentNode): HTMLElement[] {
    const selector =
      this.documentKind === 'slides' ? '.preview-slide' : this.documentKind === 'sheets' ? '.preview-sheet' : 'section.docx';
    return [...host.querySelectorAll(selector)] as HTMLElement[];
  }

  private applyPreviewScale(stage: HTMLElement, frame: HTMLElement, host: HTMLElement, naturalW: number, naturalH: number): void {
    if (naturalW < 2 || naturalH < 2) return;
    let scale = this.wordPercent;
    if (this.wordFit === 'page') {
      scale = Math.min((stage.clientWidth - 48) / naturalW, (stage.clientHeight - 48) / naturalH);
    } else if (this.wordFit === 'width') {
      scale = (stage.clientWidth - 48) / naturalW;
    }
    scale = Math.min(Math.max(scale, 0.2), 4);
    frame.style.width = `${Math.ceil(naturalW * scale)}px`;
    frame.style.height = `${Math.ceil(naturalH * scale)}px`;
    host.style.width = `${naturalW}px`;
    host.style.height = `${naturalH}px`;
    host.style.transform = `scale(${scale})`;
    const label = Math.round(scale * 100);
    if (label !== this.wordScaleLabel) this.wordScaleLabel = label;
  }

  private wordGo(delta: number): void {
    if (this.totalPages < 1) return;
    const next = Math.min(this.totalPages, Math.max(1, this.currentPage + delta));
    this.wordMenuOpen = false;
    if (next !== this.currentPage) {
      this.currentPage = next;
      const stage = this.renderRoot.querySelector('.word-stage');
      if (stage instanceof HTMLElement) stage.scrollTop = 0;
    }
  }

  private openWordZoom(): void {
    this.wordMenuOpen = false;
    this.previewSettingsOpen = false;
    if (this.wordFit === 'page' || this.wordFit === 'width') this.wordZoomDraft = this.wordFit;
    else this.wordZoomDraft = this.nearestWordPercent(this.wordPercent);
    this.wordZoomOpen = true;
  }

  private nearestWordPercent(value: number): number {
    return WORD_ZOOM_LEVELS.reduce((best, step) => (Math.abs(step - value) < Math.abs(best - value) ? step : best), 1);
  }

  private applyWordZoom(): void {
    const choice = this.wordZoomDraft;
    if (choice === 'page' || choice === 'width') this.wordFit = choice;
    else {
      this.wordFit = 'percent';
      this.wordPercent = choice;
    }
    this.wordZoomOpen = false;
  }

  private onWordZoomKey(event: KeyboardEvent): void {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    event.stopPropagation();
    this.applyWordZoom();
  }

  private toggleWordMenu(): void {
    this.wordMenuOpen = !this.wordMenuOpen;
    this.wordZoomOpen = false;
    this.previewSettingsOpen = false;
  }

  private openPreviewSettings(): void {
    this.wordMenuOpen = false;
    this.wordZoomOpen = false;
    this.menuOpen = false;
    this.zoomOpen = false;
    this.searchOpen = false;
    this.viewMenuOpen = false;
    this.previewSettingsOpen = true;
  }

  private setPreviewTheme(theme: 'light' | 'dark'): void {
    this.previewTheme = theme;
  }

  private toggleWordFullScreen(): void {
    this.wordMenuOpen = false;
    this.menuOpen = false;
    this.viewMenuOpen = false;
    if (document.fullscreenElement === this) {
      void document.exitFullscreen();
      return;
    }
    void this.requestFullscreen().catch(() => undefined);
  }

  private onFullScreenChange = (): void => {
    this.wordFullScreen = document.fullscreenElement === this;
  };

  private previewMime(name: string): string {
    const lower = name.toLowerCase();
    if (lower.endsWith('.png')) return 'image/png';
    if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
    if (lower.endsWith('.pptx')) return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    if (lower.endsWith('.ppt')) return 'application/vnd.ms-powerpoint';
    if (lower.endsWith('.xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if (lower.endsWith('.xlsm')) return 'application/vnd.ms-excel.sheet.macroEnabled.12';
    if (lower.endsWith('.xls')) return 'application/vnd.ms-excel';
    if (lower.endsWith('.doc') && !lower.endsWith('.docx')) return 'application/msword';
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  }

  private previewExtension(name: string): string {
    const dot = name.lastIndexOf('.');
    return dot >= 0 ? name.slice(dot) : '';
  }

  private ensureExtension(name: string, extension: string): string {
    if (!extension) return name;
    return name.toLowerCase().endsWith(extension.toLowerCase()) ? name : `${name}${extension}`;
  }

  private downloadOriginal(): void {
    if (!this.wordBytes) return;
    this.wordMenuOpen = false;
    this.downloadBytes(this.downloadName);
  }

  private downloadBytes(name: string): void {
    if (!this.wordBytes) return;
    const type = this.previewMime(name);
    const copy = new Uint8Array(this.wordBytes);
    const url = URL.createObjectURL(new Blob([copy.buffer], { type }));
    const link = document.createElement('a');
    link.href = url;
    link.download = name;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  private async savePreviewAs(): Promise<void> {
    if (!this.wordBytes) return;
    this.wordMenuOpen = false;
    const suggested = this.downloadName;
    const type = this.previewMime(suggested);
    const extension = this.previewExtension(suggested);
    const picker = (window as unknown as PreviewSavePicker).showSaveFilePicker;
    if (typeof picker === 'function') {
      try {
        const handle = await picker({
          suggestedName: suggested,
          types: extension ? [{ description: 'Document', accept: { [type]: [extension] } }] : undefined,
        });
        const writable = await handle.createWritable();
        const copy = new Uint8Array(this.wordBytes);
        await writable.write(new Blob([copy.buffer], { type }));
        await writable.close();
        if (handle.name) this.downloadName = handle.name;
        return;
      } catch (error) {
        if (error instanceof DOMException && error.name === 'AbortError') return;
      }
    }
    const chosen = window.prompt('Save as', suggested);
    if (!chosen?.trim()) return;
    const name = this.ensureExtension(chosen.trim(), extension);
    this.downloadName = name;
    this.downloadBytes(name);
  }

  private printPreview(): void {
    if (!this.wordBytes) return;
    this.wordMenuOpen = false;
    const iframe = document.createElement('iframe');
    iframe.setAttribute('aria-hidden', 'true');
    iframe.style.cssText = 'position:fixed;width:0;height:0;border:0;right:0;bottom:0';
    document.body.append(iframe);
    const doc = iframe.contentDocument;
    const frameWindow = iframe.contentWindow;
    if (!doc || !frameWindow) {
      iframe.remove();
      return;
    }
    const style = doc.createElement('style');
    style.textContent = `
      @page { margin: 12mm; }
      html, body { margin: 0; background: #fff; }
      section.docx { break-after: page; box-shadow: none !important; margin: 0 auto !important; }
      img.preview-image { display: block; max-width: 100%; height: auto; margin: 0 auto; }
      .preview-slide { break-after: page; margin: 0 auto 12px !important; position: relative !important; }
      .preview-sheet { break-after: page; }
      .sheet-name { font: 650 15px sans-serif; margin: 0 0 12px; }
      .sheet-table { border-collapse: collapse; font: 13px sans-serif; }
      .sheet-table th, .sheet-table td { border: 1px solid #ccc; padding: 4px 8px; vertical-align: top; }
      .sheet-note, .sheet-empty { color: #5f6368; font: 13px sans-serif; }
    `;
    doc.head.append(style);
    if (this.documentKind === 'image') {
      const image = this.renderRoot.querySelector('img.preview-image');
      if (image) doc.body.append(doc.importNode(image, true));
    } else if (this.documentKind === 'slides' || this.documentKind === 'sheets') {
      const frameElement = doc.defaultView?.Element;
      const selector = this.documentKind === 'slides' ? '.word-pages .preview-slide' : '.word-pages .preview-sheet';
      this.renderRoot.querySelectorAll(selector).forEach((node) => {
        const copy = this.clonePreviewNode(node, doc);
        if (frameElement && copy instanceof frameElement) copy.removeAttribute('hidden');
        doc.body.append(copy);
      });
    } else {
      this.renderRoot.querySelectorAll('.word-pages style').forEach((node) => {
        doc.head.append(doc.importNode(node, true));
      });
      const frameElement = doc.defaultView?.Element;
      this.renderRoot.querySelectorAll('.word-pages section.docx').forEach((node) => {
        const copy = doc.importNode(node, true);
        if (frameElement && copy instanceof frameElement) copy.removeAttribute('hidden');
        doc.body.append(copy);
      });
    }
    const cleanup = (): void => {
      frameWindow.removeEventListener('afterprint', cleanup);
      iframe.remove();
    };
    frameWindow.addEventListener('afterprint', cleanup);
    window.setTimeout(cleanup, 120000);
    requestAnimationFrame(() => {
      frameWindow.focus();
      frameWindow.print();
    });
  }

  private clonePreviewNode(node: Element, doc: Document): Node {
    const copy = doc.importNode(node, true);
    if (!(node instanceof HTMLElement) || !(copy instanceof HTMLElement)) return copy;
    const sources = [...node.querySelectorAll('canvas')];
    const targets = [...copy.querySelectorAll('canvas')];
    sources.forEach((source, index) => {
      const target = targets[index];
      if (!(target instanceof HTMLCanvasElement)) return;
      target.width = source.width;
      target.height = source.height;
      target.getContext('2d')?.drawImage(source, 0, 0);
    });
    return copy;
  }

  private focusWordStage(event: PointerEvent): void {
    const target = event.target;
    if (target instanceof Element && target.closest('button, a, input, textarea, select, label')) return;
    const stage = event.currentTarget;
    if (stage instanceof HTMLElement) stage.focus({ preventScroll: true });
  }

  private stepPreviewZoom(direction: 1 | -1): void {
    const current = this.wordFit === 'percent' ? this.wordPercent : this.wordScaleLabel / 100;
    const steps = [...WORD_ZOOM_LEVELS].sort((left, right) => left - right);
    const next =
      direction > 0
        ? (steps.find((step) => step > current + 0.02) ?? steps[steps.length - 1])
        : ([...steps].reverse().find((step) => step < current - 0.02) ?? steps[0]);
    this.wordFit = 'percent';
    this.wordPercent = next;
  }

  private handleChord(event: KeyboardEvent): boolean {
    if (!(event.ctrlKey || event.metaKey) || event.altKey) return false;
    if (this.previewSettingsOpen || this.wordZoomOpen || this.linkConfirmId) return false;
    const target = event.composedPath()[0];
    if (target instanceof HTMLElement && target.closest('input, textarea, select')) return false;
    const key = event.key.toLowerCase();
    if (key === 'o') {
      event.preventDefault();
      this.chooseFile();
      return true;
    }
    if (key === 'p') {
      event.preventDefault();
      if (this.isReadOnlyPreview()) this.printPreview();
      else if (this.fileReady) void this.printPdf();
      return true;
    }
    if (key === 's') {
      event.preventDefault();
      if (this.isReadOnlyPreview()) this.downloadOriginal();
      else if (this.fileReady) void this.save();
      return true;
    }
    if (key === 'f' && !this.isReadOnlyPreview()) {
      event.preventDefault();
      this.menuOpen = false;
      this.zoomOpen = false;
      this.viewMenuOpen = false;
      this.searchOpen = true;
      void this.updateComplete.then(() => {
        const input = this.renderRoot.querySelector('.search-pop input');
        if (input instanceof HTMLInputElement) input.focus();
      });
      return true;
    }
    if (key === '=' || key === '+' || key === 'add') {
      event.preventDefault();
      if (this.isReadOnlyPreview()) this.stepPreviewZoom(1);
      else this.stepZoom(1);
      return true;
    }
    if (key === '-' || key === '_' || key === 'subtract') {
      event.preventDefault();
      if (this.isReadOnlyPreview()) this.stepPreviewZoom(-1);
      else this.stepZoom(-1);
      return true;
    }
    if (key === '0') {
      event.preventDefault();
      if (this.isReadOnlyPreview()) {
        this.wordFit = 'percent';
        this.wordPercent = 1;
      } else this.setPercent(1);
      return true;
    }
    if (key === 'c' && !this.isReadOnlyPreview()) {
      const quote = this.textSelection?.quote.trim();
      if (!quote) return false;
      event.preventDefault();
      void navigator.clipboard.writeText(quote);
      return true;
    }
    return false;
  }

  private onWordKeydown(event: KeyboardEvent): void {
    if (this.previewSettingsOpen || this.wordZoomOpen) {
      if (event.key === 'Escape') {
        if (this.settingsQuery) this.settingsQuery = '';
        else {
          this.previewSettingsOpen = false;
          this.wordZoomOpen = false;
        }
        event.preventDefault();
      }
      return;
    }
    if (event.key === 'Escape') {
      if (this.wordMenuOpen) {
        this.wordMenuOpen = false;
        event.preventDefault();
        return;
      }
      if (document.fullscreenElement === this) {
        void document.exitFullscreen();
        event.preventDefault();
      }
      return;
    }
    const target = event.composedPath()[0];
    if (target instanceof HTMLElement && target.closest('input, textarea, select')) return;
    if (this.handleChord(event)) return;
    if (event.key === 'ArrowLeft' || event.key === 'ArrowUp' || event.key === 'PageUp') {
      this.wordGo(-1);
      event.preventDefault();
    } else if (event.key === 'ArrowRight' || event.key === 'ArrowDown' || event.key === 'PageDown') {
      this.wordGo(1);
      event.preventDefault();
    }
  }

  private previewStatus(): string {
    if (this.errorMessage) return this.errorMessage;
    if (this.loading) {
      if (this.documentKind === 'image') return 'Opening image…';
      if (this.documentKind === 'slides') return 'Opening presentation…';
      if (this.documentKind === 'sheets') return 'Opening spreadsheet…';
      return 'Opening document…';
    }
    if (this.documentKind === 'image') return 'Open an image to preview it.';
    if (this.documentKind === 'slides') return 'Open a presentation to preview it.';
    if (this.documentKind === 'sheets') return 'Open a spreadsheet to preview it.';
    return 'Open a Word document to preview it.';
  }

  private get fileReady(): boolean {
    return this.pdf !== null && !this.saving;
  }

  private shortcutRows() {
    const preview = this.isReadOnlyPreview();
    return SETTINGS_SHORTCUTS.filter((item) => item.scope === 'all' || (preview ? item.scope === 'preview' : item.scope === 'pdf'));
  }

  private listedShortcuts() {
    const rows = this.shortcutRows();
    const query = this.settingsQuery.trim().toLowerCase();
    if (!query) return rows;
    const matched = rows.filter((row) => `${row.keys} ${row.description}`.toLowerCase().includes(query));
    if (matched.length > 0) return matched;
    return 'shortcuts keyboard'.includes(query) ? rows : matched;
  }

  private advancedOptions() {
    const rows = [
      {
        title: 'Ask before opening links',
        detail: 'Show a confirmation before a link opens a new tab or moves to another page.',
        checked: this.askBeforeLinks,
        change: (checked: boolean) => {
          this.askBeforeLinks = checked;
        },
      },
      {
        title: 'Selection toolbar',
        detail: 'Show Comment, Style, Link, and Delete when text is selected.',
        checked: this.showSelectionTools,
        change: (checked: boolean) => {
          this.showSelectionTools = checked;
        },
      },
      {
        title: 'Form field highlights',
        detail: 'Outline form fields so they are easier to find. On by default.',
        checked: this.highlightFormFields,
        change: (checked: boolean) => {
          this.highlightFormFields = checked;
        },
      },
    ];
    const query = this.settingsQuery.trim().toLowerCase();
    if (!query || 'advanced settings'.includes(query)) return rows;
    return rows.filter((row) => `${row.title} ${row.detail}`.toLowerCase().includes(query));
  }

  private settingsSections(): Array<{ id: SettingsSection; label: string }> {
    const query = this.settingsQuery.trim().toLowerCase();
    const sections: Array<{ id: SettingsSection; label: string; text: string }> = [
      { id: 'theme', label: 'Theme', text: 'theme light dark mode appearance' },
      {
        id: 'shortcuts',
        label: 'Shortcuts',
        text: `shortcuts keyboard ${this.shortcutRows().map((row) => `${row.keys} ${row.description}`).join(' ')}`,
      },
      {
        id: 'advanced',
        label: 'Advanced',
        text: `advanced ${this.advancedOptions().map((row) => `${row.title} ${row.detail}`).join(' ')}`,
      },
    ];
    if (!query) return sections;
    return sections.filter((section) => `${section.label} ${section.text}`.toLowerCase().includes(query));
  }

  private currentSettingsSection(): SettingsSection | null {
    const sections = this.settingsSections();
    if (sections.some((section) => section.id === this.settingsSection)) return this.settingsSection;
    return sections[0]?.id ?? null;
  }

  private themeCard(theme: 'light' | 'dark', label: string) {
    return html`
      <button
        class="theme-card"
        type="button"
        aria-pressed=${this.previewTheme === theme}
        @click=${() => this.setPreviewTheme(theme)}
      >
        <span class="theme-preview ${theme}" aria-hidden="true">
          <span class="theme-chrome"><i></i><i></i><i></i></span>
          <span class="theme-body"><span class="theme-side"></span><span class="theme-main"></span></span>
        </span>
        <span class="theme-caption"><span class="theme-radio"></span>${label}</span>
      </button>
    `;
  }

  private settingsPane(section: SettingsSection) {
    if (section === 'theme') {
      return html`
        <h3>Theme</h3>
        <p class="settings-note">Choose how the editor looks. The document page stays white.</p>
        <div class="theme-grid">
          ${this.themeCard('light', 'Light')}
          ${this.themeCard('dark', 'Dark')}
        </div>
      `;
    }
    if (section === 'shortcuts') {
      const rows = this.listedShortcuts();
      return html`
        <h3>Shortcuts</h3>
        <p class="settings-note">These shortcuts work while you are viewing a document.</p>
        ${rows.length === 0
          ? html`<p class="settings-empty">No shortcuts match that search.</p>`
          : html`
              <table class="shortcut-table">
                <thead>
                  <tr>
                    <th>Command</th>
                    <th>Description</th>
                  </tr>
                </thead>
                <tbody>
                  ${rows.map(
                    (row) => html`
                      <tr>
                        <td class="shortcut-keys">${row.keys}</td>
                        <td>${row.description}</td>
                      </tr>
                    `,
                  )}
                </tbody>
              </table>
            `}
      `;
    }
    const rows = this.advancedOptions();
    return html`
      <h3>Advanced</h3>
      <p class="settings-note">A few viewing options. More will be added here.</p>
      ${rows.length === 0
        ? html`<p class="settings-empty">No advanced settings match that search.</p>`
        : rows.map(
            (row) => html`
              <div class="adv-row">
                <div>
                  <h3>${row.title}</h3>
                  <p>${row.detail}</p>
                </div>
                <label class="switch">
                  <input
                    type="checkbox"
                    .checked=${row.checked}
                    aria-label=${row.title}
                    @change=${(event: Event) => row.change((event.target as HTMLInputElement).checked)}
                  />
                  <span></span>
                </label>
              </div>
            `,
          )}
    `;
  }

  private settingsDialog() {
    if (!this.previewSettingsOpen) return null;
    const sections = this.settingsSections();
    const section = this.currentSettingsSection();
    return html`
      <div
        class="word-modal"
        @click=${(event: Event) => {
          if (event.target === event.currentTarget) this.previewSettingsOpen = false;
        }}
      >
        <div class="settings-shell" role="dialog" aria-modal="true" aria-label="Settings">
          <div class="settings-head">
            <h2>Settings</h2>
            <button class="settings-close" type="button" aria-label="Close" @click=${() => { this.previewSettingsOpen = false; }}>×</button>
          </div>
          <div class="settings-search">
            <input
              type="search"
              placeholder="Search settings"
              aria-label="Search settings"
              .value=${this.settingsQuery}
              @input=${(event: Event) => {
                this.settingsQuery = (event.target as HTMLInputElement).value;
              }}
            />
          </div>
          <div class="settings-body">
            <nav class="settings-nav" aria-label="Settings sections">
              ${sections.map(
                (item) => html`
                  <button
                    type="button"
                    aria-current=${section === item.id ? 'true' : nothing}
                    @click=${() => {
                      this.settingsSection = item.id;
                    }}
                  >
                    ${item.label}
                  </button>
                `,
              )}
            </nav>
            <div class="settings-pane">
              ${section ? this.settingsPane(section) : html`<p class="settings-empty">No settings match that search.</p>`}
            </div>
          </div>
        </div>
      </div>
    `;
  }

  private wordView() {
    const pageLabel = this.totalPages > 0 ? `Page ${this.currentPage} of ${this.totalPages}` : 'No document';
    const zoomChoices: Array<{ label: string; value: WordZoomChoice }> = [
      { label: 'Full Page', value: 'page' },
      { label: 'Page Width', value: 'width' },
      ...WORD_ZOOM_LEVELS.map((level) => ({ label: `${Math.round(level * 100)}%`, value: level as WordZoomChoice })),
    ];
    return html`
      <div
        class="root ${this.previewTheme === 'dark' ? 'preview-dark' : ''}"
        @dragenter=${this.onDragEnter}
        @dragover=${this.onDragOver}
        @dragleave=${this.onDragLeave}
        @drop=${this.onDrop}
      >
        <div class="word-stage" tabindex="-1" @pointerdown=${this.focusWordStage}>
          <div class="word-frame"><div class="word-pages"></div></div>
          ${this.pageReady
            ? null
            : html`
                <div class="status ${this.errorMessage ? 'error' : ''}">
                  ${this.previewStatus()}
                </div>
              `}
          ${this.dragOver ? html`<div class="drop-hint">Drop a PDF, Word document, presentation, spreadsheet, or image</div>` : null}
        </div>
        <footer class="word-bar">
          <div class="word-page">
            ${iconPages}
            <span class="word-page-label">${pageLabel}</span>
          </div>
          <div class="word-actions">
            <button
              class="word-zoom"
              type="button"
              aria-label="Zoom"
              aria-haspopup="dialog"
              ?disabled=${this.totalPages < 1}
              @click=${() => this.openWordZoom()}
            >
              <span>${this.wordScaleLabel}%</span>
              ${iconChevron}
            </button>
            <div class="word-menu-root">
              <button
                class="icon"
                type="button"
                aria-label="Menu"
                title="Menu"
                aria-haspopup="menu"
                aria-expanded=${this.wordMenuOpen}
                @click=${() => this.toggleWordMenu()}
              >
                ${iconMenu}
              </button>
              ${this.wordMenuOpen
                ? html`
                    <div class="word-pop" role="menu">
                      <button type="button" role="menuitem" @click=${() => this.chooseFile()}>
                        ${iconOpen}<span>Open</span>
                      </button>
                      <button type="button" role="menuitem" ?disabled=${!this.wordBytes} @click=${() => this.downloadOriginal()}>
                        ${iconPreviewDownload()}<span>Download</span>
                      </button>
                      <button type="button" role="menuitem" @click=${() => this.toggleWordFullScreen()}>
                        ${this.wordFullScreen ? iconExitFullScreen() : iconEnterFullScreen()}
                        <span>${this.wordFullScreen ? 'Exit Full Screen' : 'Enter Full Screen'}</span>
                      </button>
                      <button type="button" role="menuitem" ?disabled=${!this.wordBytes} @click=${() => void this.savePreviewAs()}>
                        ${iconSave}<span>Save As</span>
                      </button>
                      <button type="button" role="menuitem" ?disabled=${!this.wordBytes} @click=${() => this.printPreview()}>
                        ${iconPreviewPrint()}<span>Print</span>
                      </button>
                      <div class="rule"></div>
                      <button type="button" role="menuitem" @click=${() => this.openPreviewSettings()}>
                        ${iconPreviewSettings()}<span>Settings</span>
                      </button>
                    </div>
                  `
                : null}
            </div>
            <button
              class="icon"
              type="button"
              aria-label=${this.wordFullScreen ? 'Exit full screen' : 'Full Screen'}
              title=${this.wordFullScreen ? 'Exit full screen' : 'Full Screen'}
              aria-pressed=${this.wordFullScreen}
              @click=${() => this.toggleWordFullScreen()}
            >
              ${this.wordFullScreen ? iconExitFullScreen() : iconEnterFullScreen()}
            </button>
          </div>
        </footer>
        ${this.wordZoomOpen
          ? html`
              <div
                class="word-modal"
                @click=${(event: Event) => {
                  if (event.target === event.currentTarget) this.wordZoomOpen = false;
                }}
              >
                <div class="word-dialog" role="dialog" aria-label="Zoom to" @keydown=${this.onWordZoomKey}>
                  <h2>Zoom to</h2>
                  ${zoomChoices.map(
                    (choice) => html`
                      <label class="zoom-option">
                        <input
                          type="radio"
                          name="word-zoom"
                          .checked=${this.wordZoomDraft === choice.value}
                          @change=${() => {
                            this.wordZoomDraft = choice.value;
                          }}
                        />
                        <span>${choice.label}</span>
                      </label>
                    `,
                  )}
                  <div class="word-dialog-actions">
                    <button class="primary" type="button" @click=${() => this.applyWordZoom()}>OK</button>
                    <button type="button" @click=${() => { this.wordZoomOpen = false; }}>Cancel</button>
                  </div>
                </div>
              </div>
            `
          : null}
        ${this.settingsDialog()}
        <input class="file-input" type="file" accept=${OPEN_ACCEPT} @change=${this.onFileChosen} />
      </div>
    `;
  }

  private async openKept(source: Uint8Array | string): Promise<PDFDocumentProxy> {
    if (typeof source === 'string') {
      this.sourceUrl = source;
      this.pdfBytes = null;
      return openDocument(source);
    }
    this.sourceUrl = '';
    this.pdfBytes = new Uint8Array(source);
    return openDocument(this.pdfBytes.slice());
  }

  private async ensureBytes(): Promise<Uint8Array> {
    if (!this.pdfBytes) {
      if (!this.sourceUrl) throw new Error('Could not read this PDF.');
      const response = await fetch(this.sourceUrl);
      if (!response.ok) throw new Error('Could not read this PDF.');
      this.pdfBytes = new Uint8Array(await response.arrayBuffer());
    }
    return this.pdfBytes.slice();
  }

  private fileName(source: PdfSource): string {
    if (typeof source === 'string') {
      const path = decodeURIComponent(source.split(/[?#]/)[0]);
      const part = path.split(/[/\\]/).pop() ?? '';
      if (part.toLowerCase().endsWith('.pdf')) return part;
    }
    if (source instanceof File && source.name.toLowerCase().endsWith('.pdf')) return source.name;
    return 'document.pdf';
  }

  private async exportedPdfBytes(): Promise<Uint8Array> {
    const drawn: DrawnMark[] = [];
    for (const mark of this.marks) {
      const image = mark.kind === 'image' && mark.src ? await this.imageForSave(mark.src) : null;
      drawn.push({
        kind: mark.kind,
        page: mark.page,
        rects: mark.rects,
        text: mark.text,
        points: mark.points,
        image: image?.bytes,
        imageType: image?.type,
        color: mark.color,
        opacity: mark.opacity,
        weight: mark.weight,
      });
    }
    return applyMarks(await this.ensureBytes(), drawn);
  }

  private async save(): Promise<void> {
    if (!this.pdf || this.saving) return;
    this.menuOpen = false;
    this.saving = true;
    try {
      this.downloadFile(await this.exportedPdfBytes(), this.downloadName);
    } catch (error) {
      this.reportError('Could not save this PDF.', error);
    } finally {
      this.saving = false;
    }
  }

  private downloadPdf(): void {
    void this.save();
  }

  private async savePdfAs(): Promise<void> {
    if (!this.pdf || this.saving) return;
    this.menuOpen = false;
    const suggested = this.downloadName.toLowerCase().endsWith('.pdf') ? this.downloadName : 'document.pdf';
    const picker = (window as unknown as PreviewSavePicker).showSaveFilePicker;
    if (typeof picker === 'function') {
      try {
        const handle = await picker({
          suggestedName: suggested,
          types: [{ description: 'PDF', accept: { 'application/pdf': ['.pdf'] } }],
        });
        this.saving = true;
        const bytes = await this.exportedPdfBytes();
        const writable = await handle.createWritable();
        const copy = new Uint8Array(bytes);
        await writable.write(new Blob([copy.buffer], { type: 'application/pdf' }));
        await writable.close();
        if (handle.name) this.downloadName = handle.name;
        this.saving = false;
        return;
      } catch (error) {
        this.saving = false;
        if (error instanceof DOMException && error.name === 'AbortError') return;
        this.reportError('Could not save this PDF.', error);
        return;
      }
    }
    const chosen = window.prompt('Save as', suggested);
    if (!chosen?.trim()) return;
    const name = chosen.trim().toLowerCase().endsWith('.pdf') ? chosen.trim() : `${chosen.trim()}.pdf`;
    this.downloadName = name;
    await this.save();
  }

  private async printPdf(): Promise<void> {
    if (!this.pdf || this.saving) return;
    this.menuOpen = false;
    const pageCount = this.totalPages;
    this.saving = true;
    try {
      const bytes = await this.exportedPdfBytes();
      const printed = await openDocument(bytes.slice());
      const iframe = document.createElement('iframe');
      iframe.setAttribute('aria-hidden', 'true');
      iframe.style.cssText = 'position:fixed;width:0;height:0;border:0;right:0;bottom:0';
      document.body.append(iframe);
      const doc = iframe.contentDocument;
      const frameWindow = iframe.contentWindow;
      if (!doc || !frameWindow) {
        iframe.remove();
        await printed.loadingTask.destroy();
        return;
      }
      const style = doc.createElement('style');
      style.textContent = '@page { margin: 12mm; } body { margin: 0; } canvas { display: block; width: 100%; height: auto; break-after: page; }';
      doc.head.append(style);
      for (let page = 1; page <= pageCount; page += 1) {
        const canvas = doc.createElement('canvas');
        doc.body.append(canvas);
        const task = await paintPage(printed, page, canvas, 1.5);
        await task.promise;
      }
      await printed.loadingTask.destroy();
      const cleanup = (): void => {
        frameWindow.removeEventListener('afterprint', cleanup);
        iframe.remove();
      };
      frameWindow.addEventListener('afterprint', cleanup);
      window.setTimeout(cleanup, 120000);
      frameWindow.focus();
      frameWindow.print();
    } catch (error) {
      this.reportError('Could not print this PDF.', error);
    } finally {
      this.saving = false;
    }
  }

  private downloadFile(bytes: Uint8Array, name: string, type = 'application/pdf'): void {
    const copy = new Uint8Array(bytes);
    const url = URL.createObjectURL(new Blob([copy.buffer], { type }));
    const link = document.createElement('a');
    link.href = url;
    link.download = name;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  private async imageForSave(src: string): Promise<{ bytes: Uint8Array; type: 'png' | 'jpg' }> {
    if (src.startsWith('data:image/png') || src.startsWith('data:image/jpeg')) {
      return { bytes: this.bytesFromDataUrl(src), type: src.startsWith('data:image/png') ? 'png' : 'jpg' };
    }
    return { bytes: await this.pngFromImage(src), type: 'png' };
  }

  private bytesFromDataUrl(src: string): Uint8Array {
    const encoded = src.slice(src.indexOf(',') + 1);
    const binary = atob(encoded);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  private pngFromImage(src: string): Promise<Uint8Array> {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(image.naturalWidth, 1);
        canvas.height = Math.max(image.naturalHeight, 1);
        const context = canvas.getContext('2d');
        if (!context) {
          reject(new Error('Could not read this image.'));
          return;
        }
        context.drawImage(image, 0, 0);
        canvas.toBlob((blob) => {
          if (!blob) {
            reject(new Error('Could not read this image.'));
            return;
          }
          void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), () => reject(new Error('Could not read this image.')));
        }, 'image/png');
      };
      image.onerror = () => reject(new Error('Could not read this image.'));
      image.src = src;
    });
  }

  private async swapDocument(bytes: Uint8Array): Promise<void> {
    const opened = await openDocument(bytes.slice());
    const previous = this.pdf;
    this.pdf = opened;
    await previous?.loadingTask.destroy();
    this.textRunCache.clear();
    this.cancelThumbnails();
    this.thumbsPainted.clear();
    this.thumbsPending.clear();
    this.scheduleRender();
    if (this.findQuery.trim()) this.queueSearch();
    void this.refreshOutline();
    void this.refreshLayers();
    void this.refreshFiles();
  }

  private async readSource(source: PdfSource): Promise<Uint8Array | string> {
    if (typeof source === 'string') return source;
    if (source instanceof Uint8Array) return source;
    if (source instanceof ArrayBuffer) return new Uint8Array(source);
    return new Uint8Array(await source.arrayBuffer());
  }

  private scheduleRender(): void {
    const token = ++this.renderToken;
    this.renderTask?.cancel();
    this.renderTask = null;
    this.cancelSideSheets();
    requestAnimationFrame(() => {
      void this.renderCurrent(token);
    });
  }

  private async renderCurrent(token: number): Promise<void> {
    const pdf = this.pdf;
    const viewer = this.viewerEl;
    const canvas = this.canvasEl;
    if (!pdf || !viewer || !canvas || token !== this.renderToken) return;
    if (viewer.clientWidth < 40 || viewer.clientHeight < 40) return;

    try {
      const size = await pageSizeAtScale1(pdf, this.currentPage, this.viewTurn);
      if (token !== this.renderToken) return;
      const scale = this.scaleFor(size.width, size.height, viewer);
      this.scale = scale;
      const task = await paintPage(pdf, this.currentPage, canvas, scale, this.viewTurn, this.layerPromise());
      if (token !== this.renderToken) {
        task.cancel();
        return;
      }
      this.renderTask = task;
      await task.promise;
      if (token !== this.renderToken) return;
      this.highlights = await this.boxesForCurrentPage();
      if (token !== this.renderToken) return;
      await this.refreshMarkBoxes();
      if (token !== this.renderToken) return;
      await this.refreshTextLines();
      if (token !== this.renderToken) return;
      await this.refreshFormFields();
      if (token !== this.renderToken) return;
      await this.paintSideSheets(pdf, scale, token);
      if (token !== this.renderToken) return;
      canvas.dataset.sheetStamp = this.viewStamp(scale);
      this.pageReady = true;
      this.renderTask = null;
      if (this.resetScroll) {
        viewer.scrollTop = 0;
        viewer.scrollLeft = 0;
        this.resetScroll = false;
      }
    } catch (error) {
      if (isRenderCancel(error) || token !== this.renderToken) return;
      this.pageReady = false;
      const message = error instanceof Error ? error.message : 'Could not draw this page.';
      this.errorMessage = message;
      this.dispatchEvent(
        new CustomEvent<PdfErrorDetail>('pdf-error', {
          detail: { message },
          bubbles: true,
          composed: true,
        }),
      );
    }
  }

  private scaleFor(width: number, height: number, viewer: HTMLElement): number {
    if (this.zoomMode === 'percent') return this.zoomPercent;
    const columns = Math.max(1, this.spreadColumns());
    const gap = columns > 1 ? 16 * (columns - 1) : 0;
    const availableWidth = Math.max((viewer.clientWidth - 72 - gap) / columns, 80);
    const availableHeight = Math.max(viewer.clientHeight - 68, 120);
    if (this.zoomMode === 'fit-width') return availableWidth / width;
    return Math.min(availableWidth / width, availableHeight / height);
  }

  private viewportOf(page: PDFPageProxy, scale = this.scale) {
    const rotation = (((page.rotate || 0) + this.viewTurn * 90) % 360 + 360) % 360;
    return page.getViewport({ scale, rotation });
  }

  private sheetRows(): number[][] {
    const total = this.totalPages;
    if (total < 1) return [];
    if (this.viewSpread === 'single') return Array.from({ length: total }, (_, index) => [index + 1]);
    const rows: number[][] = [];
    let page = 1;
    if (this.viewSpread === 'cover') {
      rows.push([1]);
      page = 2;
    }
    for (; page <= total; page += 2) {
      const row = [page];
      if (page + 1 <= total) row.push(page + 1);
      rows.push(row);
    }
    return rows;
  }

  private visibleRows(): number[][] {
    const rows = this.sheetRows();
    if (rows.length === 0) return [[Math.max(this.currentPage, 1)]];
    if (this.viewRead === 'scroll') return rows;
    const index = rows.findIndex((row) => row.includes(this.currentPage));
    return [rows[index < 0 ? 0 : index]];
  }

  private spreadColumns(): number {
    return this.sheetRows().reduce((max, row) => Math.max(max, row.length), 1);
  }

  private cancelSideSheets(): void {
    for (const task of this.sideTasks) {
      try {
        task.cancel();
      } catch {
        /* The task may already have finished. */
      }
    }
    this.sideTasks = [];
  }

  private async paintSideSheets(pdf: PDFDocumentProxy, scale: number, token: number): Promise<void> {
    const stamp = this.viewStamp(scale);
    const canvases = this.renderRoot.querySelectorAll('canvas.sheet-canvas');
    for (const node of canvases) {
      if (!(node instanceof HTMLCanvasElement)) continue;
      if (node.dataset.sheetStamp === stamp && node.width > 0) continue;
      const page = Number(node.dataset.sheetCanvas);
      if (!page || token !== this.renderToken) return;
      try {
        const task = await paintPage(pdf, page, node, scale, this.viewTurn, this.layerPromise());
        if (token !== this.renderToken) {
          task.cancel();
          return;
        }
        this.sideTasks.push(task);
        await task.promise;
        node.dataset.sheetStamp = stamp;
        this.sideTasks = this.sideTasks.filter((item) => item !== task);
      } catch (error) {
        if (isRenderCancel(error) || token !== this.renderToken) return;
      }
    }
  }

  private onViewerScroll = (): void => {
    if (this.viewRead !== 'scroll' || this.sheetScrollLock) return;
    if (this.scrollSpyFrame) return;
    this.scrollSpyFrame = requestAnimationFrame(() => {
      this.scrollSpyFrame = 0;
      this.syncPageFromScroll();
    });
  };

  private syncPageFromScroll(): void {
    if (this.viewRead !== 'scroll' || this.sheetScrollLock || !this.viewerEl) return;
    const view = this.viewerEl.getBoundingClientRect();
    let best = this.currentPage;
    let bestOverlap = 0;
    this.renderRoot.querySelectorAll('[data-sheet]').forEach((node) => {
      if (!(node instanceof HTMLElement)) return;
      const page = Number(node.dataset.sheet);
      if (!page) return;
      const box = node.getBoundingClientRect();
      const overlap = Math.min(box.bottom, view.bottom) - Math.max(box.top, view.top);
      if (overlap > bestOverlap) {
        best = page;
        bestOverlap = overlap;
      }
    });
    if (best !== this.currentPage && bestOverlap > 0) {
      this.currentPage = best;
      this.scheduleRender();
    }
  }

  private setViewRead(mode: ViewRead): void {
    this.viewMenuOpen = false;
    if (this.viewRead === mode) return;
    this.viewRead = mode;
    if (mode === 'page') this.resetScroll = true;
  }

  private setViewSpread(mode: ViewSpread): void {
    this.viewMenuOpen = false;
    if (this.viewSpread === mode) return;
    this.viewSpread = mode;
    if (this.viewRead === 'page') this.resetScroll = true;
  }

  private turnView(direction: 1 | -1 | 0): void {
    this.viewMenuOpen = false;
    const next = direction === 0 ? 0 : (this.viewTurn + direction + 4) % 4;
    if (next === this.viewTurn) return;
    this.viewTurn = next;
    if (this.viewRead === 'page') this.resetScroll = true;
  }

  private chooseFit(mode: 'fit-page' | 'fit-width'): void {
    this.viewMenuOpen = false;
    this.setFit(mode);
  }

  private actualSize(): void {
    this.viewMenuOpen = false;
    this.setPercent(1);
  }

  private togglePageList(): void {
    this.pagesPressed = !this.pagesPressed;
    this.viewMenuOpen = false;
  }

  private setPercent(value: number): void {
    this.zoomMode = 'percent';
    this.zoomPercent = value;
    this.zoomOpen = false;
    this.scheduleRender();
  }

  private setFit(mode: 'fit-page' | 'fit-width'): void {
    this.zoomMode = mode;
    this.zoomOpen = false;
    this.scheduleRender();
  }

  private stepZoom(direction: 1 | -1): void {
    const current = this.zoomMode === 'percent' ? this.zoomPercent : this.scale;
    const next =
      direction > 0
        ? (ZOOM_STEPS.find((step) => step > current + 0.001) ?? Math.min(current * 1.25, 8))
        : ([...ZOOM_STEPS].reverse().find((step) => step < current - 0.001) ?? Math.max(current / 1.25, 0.1));
    this.setPercent(next);
  }

  private goToPage(page: number): void {
    if (!this.pdf) return;
    const next = Math.min(this.totalPages, Math.max(1, page));
    this.menuOpen = false;
    this.viewMenuOpen = false;
    if (next === this.currentPage) return;
    const staying =
      this.viewRead === 'scroll' ||
      this.sheetRows().some((row) => row.includes(this.currentPage) && row.includes(next));
    this.highlights = [];
    this.noteDraft = null;
    this.linkPageInput = '';
    this.notePreview = [];
    this.freeTextBox = null;
    this.textSelection = null;
    this.styleMenuOpen = false;
    this.textSelectStart = null;
    this.selectionToken += 1;
    this.stampMenu = null;
    this.imagePoint = null;
    this.dragStart = null;
    this.draftBox = null;
    this.draftStroke = [];
    this.draftArrow = 'none';
    this.polygonHover = null;
    this.currentPage = next;
    if (staying) {
      this.resetScroll = false;
      if (this.viewRead === 'scroll') {
        this.sheetScrollLock = true;
        const target = next;
        void this.updateComplete.then(() => {
          this.renderRoot.querySelector(`[data-sheet="${target}"]`)?.scrollIntoView({ block: 'start', inline: 'nearest' });
          window.setTimeout(() => {
            this.sheetScrollLock = false;
          }, 80);
        });
      }
    } else {
      this.pageReady = false;
      this.resetScroll = true;
    }
    this.scheduleRender();
  }

  private clearSearch(): void {
    window.clearTimeout(this.searchTimer);
    this.searchToken += 1;
    this.textMatches = [];
    this.matchCount = 0;
    this.matchIndex = -1;
    this.highlights = [];
    this.searching = false;
  }

  private queueSearch(): void {
    window.clearTimeout(this.searchTimer);
    this.searchTimer = window.setTimeout(() => {
      void this.runSearch();
    }, 200);
  }

  private onFindInput(event: Event): void {
    this.findQuery = (event.target as HTMLInputElement).value;
    if (!this.findQuery.trim()) {
      this.clearSearch();
      return;
    }
    this.searching = true;
    this.queueSearch();
  }

  private async runSearch(): Promise<void> {
    const token = ++this.searchToken;
    const query = this.findQuery.trim();
    const pdf = this.pdf;
    if (!pdf || !query) {
      this.clearSearch();
      return;
    }
    this.searching = true;
    try {
      const found = this.withEdits(await findText(pdf, query), query);
      if (token !== this.searchToken || this.pdf !== pdf) return;
      this.textMatches = found;
      this.matchCount = found.length;
      this.searching = false;
      if (found.length === 0) {
        this.matchIndex = -1;
        this.highlights = [];
        return;
      }
      this.matchIndex = 0;
      if (found[0].page !== this.currentPage) this.goToPage(found[0].page);
      else this.highlights = await this.boxesForCurrentPage();
    } catch {
      if (token !== this.searchToken) return;
      this.searching = false;
    }
  }

  private showMatch(index: number): void {
    if (this.textMatches.length === 0) return;
    const next = (index + this.textMatches.length) % this.textMatches.length;
    this.matchIndex = next;
    const page = this.textMatches[next].page;
    if (page !== this.currentPage) this.goToPage(page);
    else void this.boxesForCurrentPage().then((boxes) => {
      if (this.matchIndex === next) this.highlights = boxes;
    });
  }

  private onFindKeydown(event: KeyboardEvent): void {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    event.stopPropagation();
    this.showMatch(this.matchIndex + (event.shiftKey ? -1 : 1));
  }

  private async boxesForCurrentPage(): Promise<Array<ViewportBox & { active: boolean }>> {
    const pdf = this.pdf;
    if (!pdf || this.textMatches.length === 0) return [];
    const page = await pdf.getPage(this.currentPage);
    const viewport = this.viewportOf(page);
    const boxes: Array<ViewportBox & { active: boolean }> = [];
    this.textMatches.forEach((match, index) => {
      if (match.page !== this.currentPage) return;
      for (const rect of match.rects) {
        boxes.push({ ...projectRect(viewport, rect), active: index === this.matchIndex });
      }
    });
    return boxes;
  }

  private cancelThumbnails(): void {
    this.thumbToken += 1;
    this.thumbsPending.clear();
    this.thumbObserver?.disconnect();
    this.thumbObserver = null;
    for (const task of this.thumbTasks.values()) task.cancel();
    this.thumbTasks.clear();
  }

  private observeThumbnails(): void {
    const root = this.pageListEl;
    if (!this.pagesPressed || this.pageTab !== 'thumbs' || !root || !this.pdf) return;
    this.thumbObserver?.disconnect();
    this.thumbObserver = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const page = Number((entry.target as HTMLElement).dataset.page);
          if (!Number.isFinite(page)) continue;
          void this.paintThumbnail(page, entry.target as HTMLElement);
        }
      },
      { root, rootMargin: '240px 0px' },
    );
    root.querySelectorAll('.thumb').forEach((node) => this.thumbObserver?.observe(node));
  }

  private async paintThumbnail(pageNumber: number, card: HTMLElement): Promise<void> {
    const pdf = this.pdf;
    if (!pdf || this.thumbsPainted.has(pageNumber) || this.thumbsPending.has(pageNumber)) return;
    const canvas = card.querySelector('canvas');
    if (!(canvas instanceof HTMLCanvasElement)) return;
    const token = this.thumbToken;
    this.thumbsPending.add(pageNumber);
    try {
      const size = await pageSizeAtScale1(pdf, pageNumber, this.viewTurn);
      if (token !== this.thumbToken || this.pdf !== pdf) return;
      const task = await paintPage(pdf, pageNumber, canvas, this.thumbWidth / size.width, this.viewTurn, this.layerPromise());
      if (token !== this.thumbToken || this.pdf !== pdf) {
        task.cancel();
        return;
      }
      this.thumbTasks.set(pageNumber, task);
      this.thumbsPainted.add(pageNumber);
      await task.promise;
      this.thumbTasks.delete(pageNumber);
    } catch (error) {
      this.thumbTasks.delete(pageNumber);
      if (!isRenderCancel(error)) this.thumbsPainted.delete(pageNumber);
    } finally {
      if (token === this.thumbToken) this.thumbsPending.delete(pageNumber);
    }
  }

  private scrollActiveThumb(): void {
    this.pageListEl?.querySelector('.thumb[aria-current="page"]')?.scrollIntoView({ block: 'nearest' });
  }

  private chooseFile(): void {
    this.menuOpen = false;
    this.wordMenuOpen = false;
    this.fileInput.click();
  }

  private onFileChosen(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) void this.load(file);
  }

  private onKeydown = (event: KeyboardEvent): void => {
    if (this.isReadOnlyPreview()) {
      this.onWordKeydown(event);
      return;
    }
    if (event.key === 'Escape') {
      if (this.linkConfirmId) {
        this.linkConfirmId = null;
        event.preventDefault();
        return;
      }
      if (this.previewSettingsOpen) {
        if (this.settingsQuery) this.settingsQuery = '';
        else this.previewSettingsOpen = false;
        event.preventDefault();
        return;
      }
      if (this.filterOpen) {
        this.closeFilter();
        event.preventDefault();
        return;
      }
      if (this.commentMenuId || this.statusMenuId) {
        this.closeCommentMenus();
        event.preventDefault();
        return;
      }
      if (this.pageMenu) {
        this.pageMenu = null;
        event.preventDefault();
        return;
      }
      if (this.editingCommentId) {
        this.cancelEditComment();
        event.preventDefault();
        return;
      }
      if (this.noteDraft || this.editDraft || this.formDraft || this.draftStroke.length > 0 || this.stampMenu) {
        this.cancelDraft();
        event.preventDefault();
        return;
      }
      this.menuOpen = false;
      this.zoomOpen = false;
      this.searchOpen = false;
      this.viewMenuOpen = false;
      if (this.styleMenuOpen) {
        this.styleMenuOpen = false;
        this.styleMoreOpen = false;
      }
      return;
    }
    if (this.previewSettingsOpen) return;
    const target = event.composedPath()[0];
    if (target instanceof HTMLElement && target.closest('input, textarea, select')) return;
    if (this.handleChord(event)) return;
    if (event.key === 'Enter' && this.shapeKind === 'polygon' && this.draftStroke.length >= 3) {
      event.preventDefault();
      this.finishPolygon();
      return;
    }
    if (event.key === 'Enter' && this.measureKind === 'perimeter' && this.draftStroke.length >= 2) {
      event.preventDefault();
      this.finishMeasure('perimeter');
      return;
    }
    if (event.key === 'Enter' && this.measureKind === 'area' && this.draftStroke.length >= 3) {
      event.preventDefault();
      this.finishMeasure('area');
      return;
    }
    if (event.key === 'ArrowLeft') {
      this.goToPage(this.currentPage - 1);
      event.preventDefault();
    } else if (event.key === 'ArrowRight') {
      this.goToPage(this.currentPage + 1);
      event.preventDefault();
    }
  };

  private onWindowPointerDown = (event: PointerEvent): void => {
    const path = event.composedPath();
    if (this.wordMenuOpen) {
      const menu = this.renderRoot.querySelector('.word-menu-root');
      if (menu && !path.includes(menu)) this.wordMenuOpen = false;
    }
    if (this.menuOpen && this.menuRoot && !path.includes(this.menuRoot)) this.menuOpen = false;
    if (this.zoomOpen && this.zoomRoot && !path.includes(this.zoomRoot)) this.zoomOpen = false;
    if (this.searchOpen && this.searchRoot && !path.includes(this.searchRoot)) this.searchOpen = false;
    if (this.viewMenuOpen && this.viewRoot && !path.includes(this.viewRoot)) this.viewMenuOpen = false;
    if (this.pageMenu) {
      const inside = path.some(
        (node) =>
          node instanceof HTMLElement &&
          (node.classList.contains('page-ops') || node.classList.contains('thumb-more') || node.classList.contains('pick-more')),
      );
      if (!inside) this.pageMenu = null;
    }
    if (this.commentMenuId || this.statusMenuId) {
      const inside = path.some(
        (node) => node instanceof HTMLElement && (node.classList.contains('comment-menu') || node.classList.contains('comment-menu-btn')),
      );
      if (!inside) this.closeCommentMenus();
    }
  };

  private onPointerDown(event: PointerEvent): void {
    if (event.button !== 0) return;
    if (this.eventIn(event, '.sheet-hit') && this.tool !== 'hand') return;
    if (
      this.eventIn(event, '.note-box') ||
      this.eventIn(event, '.freetext-editor') ||
      this.eventIn(event, '.quick-bar') ||
      this.eventIn(event, '.quick-styles') ||
      this.eventIn(event, '.stamp-menu') ||
      this.eventIn(event, '.text-line') ||
      this.eventIn(event, '.field-target')
    ) {
      return;
    }
    if (this.noteDraft?.kind === 'freetext') {
      this.commitFreeText();
      return;
    }
    this.stampMenu = null;
    if (this.tool === 'hand') {
      this.panPoint = {
        x: event.clientX,
        y: event.clientY,
        left: this.viewerEl.scrollLeft,
        top: this.viewerEl.scrollTop,
      };
      this.panning = true;
      this.capturePointer(event);
      return;
    }
    const kind = this.drawKind;
    if (this.canDraw && (kind === 'polygon' || kind === 'perimeter' || kind === 'area')) {
      const point = this.canvasPoint(event);
      if (!point) return;
      this.dragStart = point;
      this.capturePointer(event);
      return;
    }
    if (this.canDraw && kind === 'freetext') {
      const point = this.canvasPoint(event);
      if (!point) return;
      this.dragStart = point;
      this.capturePointer(event);
      return;
    }
    if (this.canDraw && (kind === 'text' || kind === 'image' || kind === 'stamp' || kind === 'initials' || kind === 'date')) {
      const point = this.canvasPoint(event);
      if (!point) return;
      this.dragStart = point;
      this.capturePointer(event);
      return;
    }
    if (this.canDraw && kind) {
      const point = this.canvasPoint(event);
      if (!point) return;
      this.cancelDraft();
      this.dragStart = point;
      this.dragToken = this.selectionToken;
      if (kind === 'note') {
        this.selectedMarkId = null;
        this.replyText = '';
      }
      if (kind === 'pen' || kind === 'signature') this.draftStroke = [point];
      if (kind === 'callout') {
        this.draftStroke = [point, point];
        this.draftArrow = 'start';
      }
      if (kind === 'line' || kind === 'arrow' || kind === 'distance') {
        this.draftStroke = [point, point];
        this.draftArrow = kind === 'arrow' ? 'end' : 'none';
      }
      this.capturePointer(event);
      return;
    }
    if (!this.canDraw && !this.editText && this.pdf) {
      const hitId = this.markIdFromEvent(event);
      const hit = hitId ? this.marks.find((item) => item.id === hitId) : undefined;
      if (hit?.kind === 'link') {
        this.selectedMarkId = hit.id;
        this.textSelection = null;
        this.styleMenuOpen = false;
        this.textSelectStart = null;
        void this.refreshMarkBoxes();
        if (this.askBeforeLinks) this.linkConfirmId = hit.id;
        else this.followLink(hit);
        return;
      }
      const point = this.canvasPoint(event);
      if (point) {
        this.quickToken += 1;
        this.textSelectStart = point;
        this.textSelection = null;
        this.styleMenuOpen = false;
        this.capturePointer(event);
      }
    }
    const markId = this.markIdFromEvent(event);
    if (markId !== this.selectedMarkId) this.replyText = '';
    this.selectedMarkId = markId;
    void this.refreshMarkBoxes();
  }

  private onPointerMove(event: PointerEvent): void {
    if (this.textSelectStart && !this.canDraw) {
      const point = this.canvasPoint(event);
      if (!point) return;
      this.draftBox = this.boxBetween(this.textSelectStart, point);
      return;
    }
    if (this.dragStart && this.canDraw && (this.drawKind === 'pen' || this.drawKind === 'signature')) {
      const point = this.canvasPoint(event);
      if (!point) return;
      const last = this.draftStroke[this.draftStroke.length - 1];
      if (!last || Math.hypot(point.x - last.x, point.y - last.y) >= 2) {
        this.draftStroke = [...this.draftStroke, point];
      }
      return;
    }
    if (this.dragStart && this.canDraw && (this.drawKind === 'callout' || this.drawKind === 'line' || this.drawKind === 'arrow' || this.drawKind === 'distance')) {
      const point = this.canvasPoint(event);
      if (!point) return;
      this.draftStroke = [this.dragStart, point];
      this.draftArrow = this.drawKind === 'callout' ? 'start' : this.drawKind === 'arrow' ? 'end' : 'none';
      return;
    }
    if (this.canDraw && (this.drawKind === 'polygon' || this.drawKind === 'perimeter' || this.drawKind === 'area') && this.draftStroke.length > 0) {
      const point = this.canvasPoint(event);
      if (point) this.polygonHover = point;
      return;
    }
    if (
      this.dragStart &&
      this.canDraw &&
      this.drawKind !== 'text' &&
      this.drawKind !== 'image' &&
      this.drawKind !== 'stamp' &&
      this.drawKind !== 'initials' &&
      this.drawKind !== 'date' &&
      this.drawKind !== 'polygon' &&
      this.drawKind !== 'perimeter' &&
      this.drawKind !== 'area'
    ) {
      const point = this.canvasPoint(event);
      if (!point) return;
      this.draftBox = this.boxBetween(this.dragStart, point);
      return;
    }
    if (!this.panPoint) return;
    this.viewerEl.scrollLeft = this.panPoint.left - (event.clientX - this.panPoint.x);
    this.viewerEl.scrollTop = this.panPoint.top - (event.clientY - this.panPoint.y);
  }

  private onPointerUp(event: PointerEvent): void {
    if (this.textSelectStart && !this.canDraw) {
      const start = this.textSelectStart;
      const token = this.quickToken;
      const point = this.canvasPoint(event) ?? start;
      this.textSelectStart = null;
      this.draftBox = null;
      if (Math.hypot(point.x - start.x, point.y - start.y) >= 4) void this.openQuickSelection(start, point, token);
      return;
    }
    if (this.dragStart && this.canDraw && this.drawKind) {
      const start = this.dragStart;
      const point = this.canvasPoint(event) ?? start;
      const kind = this.drawKind;
      this.dragStart = null;
      const distance = Math.hypot(point.x - start.x, point.y - start.y);
      if (kind === 'note') {
        this.draftBox = null;
        if (distance >= 4) void this.openNoteFromSelection(start, point, this.dragToken);
        return;
      }
      if (kind === 'freetext') {
        this.draftBox = null;
        this.openFreeText(start, point);
        return;
      }
      if (kind === 'text' || kind === 'initials') {
        this.draftBox = null;
        if (distance < 8) this.openText(kind, start);
        return;
      }
      if (kind === 'date') {
        this.draftBox = null;
        if (distance < 8) void this.placeDate(start);
        return;
      }
      if (kind === 'image') {
        this.draftBox = null;
        if (distance < 8 && this.pdf) {
          this.imagePoint = { page: this.currentPage, x: start.x, y: start.y };
          this.imageInput.click();
        }
        return;
      }
      if (kind === 'stamp') {
        this.draftBox = null;
        if (distance < 8) this.openStampMenu(start);
        return;
      }
      if (kind === 'link') {
        if (distance < 4) {
          this.draftBox = null;
          return;
        }
        this.openLink(start, point);
        return;
      }
      this.draftBox = null;
      if (kind === 'pen' || kind === 'signature') {
        const stroke = this.draftStroke;
        this.draftStroke = [];
        if (this.strokeLength(stroke) >= 4) void this.createPen(kind, stroke);
        return;
      }
      if (kind === 'callout') {
        if (distance < 4) {
          this.draftStroke = [];
          this.draftArrow = 'none';
          return;
        }
        this.openCallout(start, point);
        return;
      }
      if (kind === 'line' || kind === 'arrow') {
        this.draftStroke = [];
        this.draftArrow = 'none';
        if (distance >= 4) void this.addShape(kind, [start, point]);
        return;
      }
      if (kind === 'distance') {
        this.draftStroke = [];
        this.draftArrow = 'none';
        if (distance >= 4) void this.addMeasure('distance', [start, point]);
        return;
      }
      if (kind === 'rect' || kind === 'ellipse') {
        if (distance >= 4) void this.addShape(kind, [start, point]);
        return;
      }
      if (kind === 'text-field' || kind === 'checkbox') {
        if (distance >= 4) void this.addFormField(kind, start, point);
        return;
      }
      if (kind === 'polygon' || kind === 'perimeter' || kind === 'area') {
        if (distance < 8) this.addPolygonCorner(start);
        return;
      }
      if (kind === 'highlight' || kind === 'underline' || kind === 'strike') {
        if (distance >= 4) void this.createMarkup(kind, start, point);
      }
      return;
    }
    this.panPoint = null;
    this.panning = false;
  }

  private onPointerCancel(): void {
    this.dragStart = null;
    this.textSelectStart = null;
    this.draftBox = null;
    this.panPoint = null;
    this.panning = false;
    if (this.shapeKind === 'polygon' || this.measureKind === 'perimeter' || this.measureKind === 'area') return;
    this.draftStroke = [];
    this.draftArrow = 'none';
    this.polygonHover = null;
  }

  private get drawKind(): MarkKind | FormKind | null {
    if (this.mode === 'annotate') return this.markupKind;
    if (this.mode === 'shapes') return this.shapeKind;
    if (this.mode === 'insert') return this.insertKind;
    if (this.mode === 'measure') return this.measureKind;
    if (this.mode === 'fill-sign') return this.signKind;
    if (this.mode === 'forms') return this.formKind;
    return null;
  }

  private get canDraw(): boolean {
    return this.drawKind !== null && this.tool !== 'hand' && this.pdf !== null;
  }

  private chooseMarkup(kind: MarkKind): void {
    this.cancelDraft();
    this.shapeKind = null;
    this.insertKind = null;
    this.measureKind = null;
    this.editText = false;
    this.signKind = null;
    this.formKind = null;
    this.markupKind = this.markupKind === kind ? null : kind;
    if (this.markupKind) this.tool = 'select';
  }

  private chooseShape(kind: ShapeKind): void {
    this.cancelDraft();
    this.markupKind = null;
    this.insertKind = null;
    this.measureKind = null;
    this.editText = false;
    this.signKind = null;
    this.formKind = null;
    this.shapeKind = this.shapeKind === kind ? null : kind;
    if (this.shapeKind) this.tool = 'select';
  }

  private chooseInsert(kind: InsertKind): void {
    this.cancelDraft();
    this.markupKind = null;
    this.shapeKind = null;
    this.measureKind = null;
    this.editText = false;
    this.signKind = null;
    this.formKind = null;
    this.insertKind = this.insertKind === kind ? null : kind;
    if (this.insertKind) this.tool = 'select';
  }

  private chooseMeasure(kind: MeasureKind): void {
    this.cancelDraft();
    this.markupKind = null;
    this.shapeKind = null;
    this.insertKind = null;
    this.editText = false;
    this.signKind = null;
    this.formKind = null;
    this.measureKind = this.measureKind === kind ? null : kind;
    if (this.measureKind) this.tool = 'select';
  }

  private chooseEdit(): void {
    this.cancelDraft();
    this.markupKind = null;
    this.shapeKind = null;
    this.insertKind = null;
    this.measureKind = null;
    this.signKind = null;
    this.formKind = null;
    this.editText = !this.editText;
    if (this.editText) this.tool = 'select';
    if (!this.editText) this.lineBoxes = [];
    else void this.refreshTextLines();
  }

  private chooseSign(kind: SignKind): void {
    this.cancelDraft();
    this.markupKind = null;
    this.shapeKind = null;
    this.insertKind = null;
    this.measureKind = null;
    this.editText = false;
    this.lineBoxes = [];
    this.formKind = null;
    this.signKind = this.signKind === kind ? null : kind;
    if (this.signKind) this.tool = 'select';
  }

  private chooseForm(kind: FormKind): void {
    this.cancelDraft();
    this.markupKind = null;
    this.shapeKind = null;
    this.insertKind = null;
    this.measureKind = null;
    this.editText = false;
    this.lineBoxes = [];
    this.signKind = null;
    this.formKind = this.formKind === kind ? null : kind;
    if (this.formKind) this.tool = 'select';
  }

  private get canEditText(): boolean {
    return this.mode === 'edit' && this.editText && this.tool !== 'hand' && this.pdf !== null;
  }

  private async refreshTextLines(): Promise<void> {
    const pdf = this.pdf;
    if (!pdf || !this.canEditText) {
      this.lineBoxes = [];
      return;
    }
    const pageNumber = this.currentPage;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf || this.currentPage !== pageNumber || !this.canEditText) return;
    const viewport = this.viewportOf(page);
    const lines = await pageTextLines(pdf, pageNumber);
    if (this.pdf !== pdf || this.currentPage !== pageNumber || !this.canEditText) return;
    this.lineBoxes = lines.map((line) => {
      const edit = [...this.textEdits].reverse().find((item) => item.page === pageNumber && rectsIntersect(item.rect, line.rect));
      return {
        ...projectRect(viewport, line.rect),
        text: edit?.text ?? line.text,
        rect: line.rect,
        baseline: line.baseline,
        fontSize: line.fontSize,
      };
    });
  }

  private async refreshFormFields(): Promise<void> {
    const pdf = this.pdf;
    const visible = (): boolean => this.pdf === pdf && this.mode === 'forms' && this.tool !== 'hand';
    if (!pdf || !visible()) {
      if (this.fieldBoxes.length > 0) this.fieldBoxes = [];
      return;
    }
    const pageNumber = this.currentPage;
    const scale = this.scale;
    const page = await pdf.getPage(pageNumber);
    if (!visible() || this.currentPage !== pageNumber || this.scale !== scale) return;
    const viewport = this.viewportOf(page, scale);
    const fields = await pageFormFields(pdf, pageNumber);
    if (!visible() || this.currentPage !== pageNumber || this.scale !== scale) return;
    this.fieldBoxes = fields.map((field) => ({ ...projectRect(viewport, field.rect), kind: field.kind, name: field.name, value: field.value, checked: field.checked }));
  }

  private openField(field: FieldBox): void {
    this.cancelDraft();
    if (field.kind === 'check') {
      void this.toggleField(field.name);
      return;
    }
    this.formDraft = {
      name: field.name,
      page: this.currentPage,
      left: field.left,
      top: field.top,
      text: field.value,
    };
  }

  private reportError(fallback: string, error: unknown): void {
    const message = error instanceof Error ? error.message : fallback;
    this.errorMessage = message;
    this.dispatchEvent(
      new CustomEvent<PdfErrorDetail>('pdf-error', {
        detail: { message },
        bubbles: true,
        composed: true,
      }),
    );
  }

  private async commitForm(event: Event): Promise<void> {
    event.preventDefault();
    const draft = this.formDraft;
    if (!draft || !this.pdf) return;
    try {
      const next = await setTextField(await this.ensureBytes(), draft.name, draft.text);
      this.pdfBytes = new Uint8Array(next);
      this.formDraft = null;
      await this.swapDocument(this.pdfBytes);
    } catch (error) {
      this.reportError('Could not update this field.', error);
    }
  }

  private async toggleField(name: string): Promise<void> {
    if (!this.pdf) return;
    try {
      const next = await toggleCheckBox(await this.ensureBytes(), name);
      this.pdfBytes = new Uint8Array(next);
      await this.swapDocument(this.pdfBytes);
    } catch (error) {
      this.reportError('Could not update this field.', error);
    }
  }

  private async addFormField(kind: FormKind, start: { x: number; y: number }, end: { x: number; y: number }): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf || this.formKind !== kind) return;
    const from = viewportToPdf(this.viewportOf(page), start.x, start.y);
    const to = viewportToPdf(this.viewportOf(page), end.x, end.y);
    const rect = pdfRectFromPoints(from.x, from.y, to.x, to.y);
    try {
      const bytes = await this.ensureBytes();
      if (this.pdf !== pdf || this.formKind !== kind) return;
      const next = kind === 'text-field' ? await addTextField(bytes, pageNumber, rect) : await addCheckBox(bytes, pageNumber, rect);
      this.pdfBytes = new Uint8Array(next);
      await this.swapDocument(this.pdfBytes);
    } catch (error) {
      this.reportError('Could not add this field.', error);
    }
  }

  private openLine(line: LineBox): void {
    this.cancelDraft();
    this.editDraft = {
      page: this.currentPage,
      rect: line.rect,
      baseline: line.baseline,
      fontSize: line.fontSize,
      left: line.left,
      top: line.top,
      text: line.text,
    };
  }

  private async commitEdit(event: Event): Promise<void> {
    event.preventDefault();
    const draft = this.editDraft;
    if (!draft || !this.pdf) return;
    const text = draft.text.trim();
    if (!text) return;
    try {
      const next = await replaceLineText(await this.ensureBytes(), draft.page, draft.rect, draft.baseline, draft.fontSize, text);
      this.pdfBytes = new Uint8Array(next);
      this.textEdits = [
        ...this.textEdits.filter((item) => !(item.page === draft.page && rectsIntersect(item.rect, draft.rect))),
        { page: draft.page, rect: draft.rect, text, baseline: draft.baseline, fontSize: draft.fontSize },
      ];
      this.editDraft = null;
      await this.swapDocument(this.pdfBytes);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Could not edit this text.';
      this.errorMessage = message;
      this.dispatchEvent(
        new CustomEvent<PdfErrorDetail>('pdf-error', {
          detail: { message },
          bubbles: true,
          composed: true,
        }),
      );
    }
  }

  private withEdits(found: TextMatch[], query: string): TextMatch[] {
    const needle = query.trim().toLowerCase();
    const kept = found.filter(
      (match) => !this.textEdits.some((edit) => edit.page === match.page && match.rects.some((rect) => rectsIntersect(rect, edit.rect))),
    );
    const extras: TextMatch[] = [];
    for (const edit of this.textEdits) {
      if (needle && edit.text.toLowerCase().includes(needle)) extras.push({ page: edit.page, rects: [edit.rect] });
    }
    return [...kept, ...extras].sort((left, right) => left.page - right.page || (right.rects[0]?.y ?? 0) - (left.rects[0]?.y ?? 0));
  }

  private capturePointer(event: PointerEvent): void {
    try {
      this.viewerEl.setPointerCapture(event.pointerId);
    } catch {
      // Capture is only available for a real, active pointer.
    }
  }

  private canvasPoint(event: PointerEvent): { x: number; y: number } | null {
    const canvas = this.canvasEl;
    if (!canvas) return null;
    const bounds = canvas.getBoundingClientRect();
    return { x: event.clientX - bounds.left, y: event.clientY - bounds.top };
  }

  private boxBetween(start: { x: number; y: number }, end: { x: number; y: number }): ViewportBox {
    return {
      left: Math.min(start.x, end.x),
      top: Math.min(start.y, end.y),
      width: Math.abs(end.x - start.x),
      height: Math.abs(end.y - start.y),
    };
  }

  private eventIn(event: Event, selector: string): boolean {
    const target = event.composedPath()[0];
    return target instanceof Element && Boolean(target.closest(selector));
  }

  private markIdFromEvent(event: Event): string | null {
    for (const node of event.composedPath()) {
      if (node instanceof HTMLElement && node.dataset.mark) return node.dataset.mark;
      if (node instanceof SVGElement && node.dataset.mark) return node.dataset.mark;
    }
    return null;
  }

  private async wordsForPage(pageNumber: number): Promise<TextRun[]> {
    const cached = this.textRunCache.get(pageNumber);
    if (cached) return cached;
    const pdf = this.pdf;
    if (!pdf) return [];
    const runs = await pageWordRects(pdf, pageNumber);
    if (this.pdf === pdf) this.textRunCache.set(pageNumber, runs);
    return runs;
  }

  private async createMarkup(
    kind: 'highlight' | 'underline' | 'strike',
    start: { x: number; y: number },
    end: { x: number; y: number },
  ): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    const page = await pdf.getPage(pageNumber);
    const viewport = this.viewportOf(page);
    const from = viewportToPdf(viewport, start.x, start.y);
    const to = viewportToPdf(viewport, end.x, end.y);
    const drag = pdfRectFromPoints(from.x, from.y, to.x, to.y);
    const probe = {
      x: drag.x,
      y: drag.y - 4,
      width: drag.width,
      height: drag.height + 8,
    };
    const runs = await this.wordsForPage(pageNumber);
    if (this.pdf !== pdf) return;
    const hit = runs.filter((run) => rectsIntersect(probe, run.rect));
    const rects = hit.length > 0 ? hit.map((run) => run.rect) : [drag];
    const mark = this.makeMark(kind, pageNumber, rects, hit.map((run) => run.text).join(' '));
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    if (this.currentPage === pageNumber) await this.refreshMarkBoxes();
  }

  private async openNoteFromSelection(
    start: { x: number; y: number },
    end: { x: number; y: number },
    token: number,
  ): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf || this.markupKind !== 'note') return;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf || token !== this.selectionToken || this.markupKind !== 'note') return;
    const viewport = this.viewportOf(page);
    const from = viewportToPdf(viewport, start.x, start.y);
    const to = viewportToPdf(viewport, end.x, end.y);
    const drag = pdfRectFromPoints(from.x, from.y, to.x, to.y);
    const probe = {
      x: drag.x,
      y: drag.y - 4,
      width: drag.width,
      height: drag.height + 8,
    };
    const runs = await this.wordsForPage(pageNumber);
    if (this.pdf !== pdf || token !== this.selectionToken || this.markupKind !== 'note') return;
    const hit = runs.filter((run) => rectsIntersect(probe, run.rect));
    if (hit.length === 0) return;
    const rects = hit.map((run) => run.rect);
    const quote = hit.map((run) => run.text).join(' ');
    this.commentsOpen = true;
    this.replyText = '';
    this.selectedMarkId = null;
    this.notePreview = rects.map((rect) => projectRect(viewport, rect));
    this.noteDraft = {
      kind: 'note',
      page: pageNumber,
      pdfX: rects[0].x,
      pdfY: rects[0].y + rects[0].height,
      left: this.notePreview[0].left,
      top: this.notePreview[0].top,
      text: '',
      tip: null,
      rect: null,
      quote,
      rects,
      author: this.authorName,
      createdAt: Date.now(),
    };
  }

  private openFreeText(start: { x: number; y: number }, end: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    const distance = Math.hypot(end.x - start.x, end.y - start.y);
    const left = distance >= 8 ? Math.min(start.x, end.x) : start.x;
    const top = distance >= 8 ? Math.min(start.y, end.y) : start.y;
    const width = distance >= 8 ? Math.max(Math.abs(end.x - start.x), 96) : 180;
    const height = distance >= 8 ? Math.max(Math.abs(end.y - start.y), 32) : 40;
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber || this.markupKind !== 'freetext') return;
      const pdfPoint = viewportToPdf(this.viewportOf(page), left, top);
      this.freeTextBox = { left, top, width, height };
      this.noteDraft = {
        kind: 'freetext',
        page: pageNumber,
        pdfX: pdfPoint.x,
        pdfY: pdfPoint.y,
        left,
        top,
        text: '',
        tip: null,
        rect: null,
        quote: '',
        rects: [],
        author: '',
        createdAt: 0,
      };
    });
  }

  private commitFreeText(): void {
    const draft = this.noteDraft;
    const box = this.freeTextBox;
    if (!draft || draft.kind !== 'freetext') return;
    const text = draft.text.trim();
    this.noteDraft = null;
    this.freeTextBox = null;
    if (!text) return;
    const width = Math.max((box?.width ?? 120) / this.scale, 1);
    const height = Math.max((box?.height ?? 24) / this.scale, 1);
    const mark = this.makeMark('freetext', draft.page, [{ x: draft.pdfX, y: draft.pdfY - height, width, height }], text);
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    void this.refreshMarkBoxes();
  }

  private async openQuickSelection(
    start: { x: number; y: number },
    end: { x: number; y: number },
    token: number,
  ): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf || token !== this.quickToken) return;
    const viewport = this.viewportOf(page);
    const from = viewportToPdf(viewport, start.x, start.y);
    const to = viewportToPdf(viewport, end.x, end.y);
    const drag = pdfRectFromPoints(from.x, from.y, to.x, to.y);
    const probe = { x: drag.x, y: drag.y - 4, width: drag.width, height: drag.height + 8 };
    const runs = await this.wordsForPage(pageNumber);
    if (this.pdf !== pdf || token !== this.quickToken) return;
    const hit = runs.filter((run) => rectsIntersect(probe, run.rect));
    if (hit.length === 0) {
      this.textSelection = null;
      return;
    }
    const rects = hit.map((run) => run.rect);
    this.selectedMarkId = null;
    this.styleMenuOpen = false;
    this.textSelection = {
      page: pageNumber,
      rects,
      quote: hit.map((run) => run.text).join(' '),
      boxes: rects.map((rect) => projectRect(viewport, rect)),
    };
  }

  private selectionMarkupIds(): string[] {
    const selection = this.textSelection;
    if (!selection) return [];
    return this.marks
      .filter(
        (mark) =>
          mark.page === selection.page &&
          mark.rects.some((rect) => selection.rects.some((hit) => rectsIntersect(rect, hit))),
      )
      .map((mark) => mark.id);
  }

  private commentOnSelection(): void {
    const selection = this.textSelection;
    if (!selection || selection.rects.length === 0) return;
    const first = selection.rects[0];
    this.commentsOpen = true;
    this.replyText = '';
    this.selectedMarkId = null;
    this.styleMenuOpen = false;
    this.notePreview = selection.boxes;
    this.noteDraft = {
      kind: 'note',
      page: selection.page,
      pdfX: first.x,
      pdfY: first.y + first.height,
      left: selection.boxes[0]?.left ?? 0,
      top: selection.boxes[0]?.top ?? 0,
      text: '',
      tip: null,
      rect: null,
      quote: selection.quote,
      rects: selection.rects,
      author: this.authorName,
      createdAt: Date.now(),
    };
    this.textSelection = null;
  }

  private linkOnSelection(): void {
    const selection = this.textSelection;
    if (!selection || selection.rects.length === 0) return;
    const x1 = Math.min(...selection.rects.map((rect) => rect.x));
    const y1 = Math.min(...selection.rects.map((rect) => rect.y));
    const x2 = Math.max(...selection.rects.map((rect) => rect.x + rect.width));
    const y2 = Math.max(...selection.rects.map((rect) => rect.y + rect.height));
    const anchor = selection.boxes[selection.boxes.length - 1];
    this.noteDraft = {
      kind: 'link',
      page: selection.page,
      pdfX: x1,
      pdfY: y1,
      left: anchor?.left ?? 0,
      top: (anchor?.top ?? 0) + (anchor?.height ?? 0) + 8,
      text: '',
      tip: null,
      rect: { x: x1, y: y1, width: Math.max(x2 - x1, 1), height: Math.max(y2 - y1, 1) },
      quote: selection.quote,
      rects: selection.rects,
      author: '',
      createdAt: 0,
    };
    this.linkPageInput = '';
    this.linkConfirmId = null;
    this.textSelection = null;
    this.styleMenuOpen = false;
  }

  private deleteOnSelection(): void {
    const ids = new Set(this.selectionMarkupIds());
    if (ids.size === 0) return;
    this.marks = this.marks.filter((mark) => !ids.has(mark.id));
    if (this.selectedMarkId && ids.has(this.selectedMarkId)) this.selectedMarkId = null;
    this.checkedCommentIds = this.checkedCommentIds.filter((id) => !ids.has(id));
    this.textSelection = null;
    this.styleMenuOpen = false;
    void this.refreshMarkBoxes();
  }

  private freeTextEditor() {
    const draft = this.noteDraft;
    const box = this.freeTextBox;
    if (!draft || draft.kind !== 'freetext' || draft.page !== this.currentPage || !box) return nothing;
    return html`
      <textarea
        class="freetext-editor"
        placeholder="Enter text here"
        aria-label="Free text"
        style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px"
        .value=${draft.text}
        @input=${(event: Event) => {
          if (!this.noteDraft || this.noteDraft.kind !== 'freetext') return;
          this.noteDraft = { ...this.noteDraft, text: (event.target as HTMLTextAreaElement).value };
        }}
        @keydown=${(event: KeyboardEvent) => {
          if (event.key !== 'Escape') return;
          event.preventDefault();
          event.stopPropagation();
          this.cancelDraft();
        }}
      ></textarea>
    `;
  }

  private defaultInk(kind: MarkKind): { color: string; opacity: number; weight: number } {
    if (kind === 'highlight') return { color: '#f6bf26', opacity: 0.45, weight: 2 };
    if (kind === 'underline') return { color: '#e37400', opacity: 1, weight: 2 };
    if (kind === 'strike') return { color: '#d93025', opacity: 1, weight: 2 };
    return { color: '#202124', opacity: 1, weight: 2 };
  }

  private fadeColor(hex: string, opacity: number): string {
    const match = /^#([0-9a-f]{6})$/i.exec(hex);
    const amount = Math.min(1, Math.max(0, opacity));
    if (!match) return hex;
    const value = Number.parseInt(match[1], 16);
    return `rgba(${(value >> 16) & 255}, ${(value >> 8) & 255}, ${value & 255}, ${amount})`;
  }

  private isTextStyle(kind: MarkKind): kind is 'highlight' | 'underline' | 'strike' {
    return kind === 'highlight' || kind === 'underline' || kind === 'strike';
  }

  private styleTarget(): Mark | undefined {
    const marks = this.marks.filter((mark) => this.selectionMarkupIds().includes(mark.id) && this.isTextStyle(mark.kind));
    return marks.find((mark) => mark.id === this.selectedMarkId) ?? marks.at(-1);
  }

  private styleLook(): { kind: 'highlight' | 'underline' | 'strike' | null; color: string; opacity: number; weight: number } {
    const mark = this.styleTarget();
    if (mark && this.isTextStyle(mark.kind)) {
      return { kind: mark.kind, color: mark.color, opacity: mark.opacity, weight: mark.weight };
    }
    return { kind: null, color: this.styleColor, opacity: this.styleOpacity, weight: this.styleWeight };
  }

  private updateStyleMark(patch: Partial<Pick<Mark, 'kind' | 'color' | 'opacity' | 'weight'>>): void {
    const target = this.styleTarget();
    if (!target) return;
    this.marks = this.marks.map((mark) => (mark.id === target.id ? { ...mark, ...patch, modifiedAt: Date.now() } : mark));
    void this.refreshMarkBoxes();
  }

  private chooseStyleKind(kind: 'highlight' | 'underline' | 'strike'): void {
    const selection = this.textSelection;
    if (!selection) return;
    if (this.styleTarget()) {
      this.updateStyleMark({ kind });
      return;
    }
    const mark = this.makeMark(kind, selection.page, selection.rects, selection.quote);
    mark.quote = selection.quote;
    if (this.styleTouched) {
      mark.color = this.styleColor;
      mark.opacity = this.styleOpacity;
      mark.weight = this.styleWeight;
    }
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    void this.refreshMarkBoxes();
  }

  private setStyleColor(color: string): void {
    const next = color.toLowerCase();
    if (!this.styleTarget()) {
      this.styleTouched = true;
      this.styleColor = next;
      return;
    }
    this.updateStyleMark({ color: next });
  }

  private setStyleOpacity(opacity: number): void {
    if (!Number.isFinite(opacity)) return;
    const next = Math.min(1, Math.max(0.05, opacity));
    if (!this.styleTarget()) {
      this.styleTouched = true;
      this.styleOpacity = next;
      return;
    }
    this.updateStyleMark({ opacity: next });
  }

  private setStyleWeight(weight: number): void {
    if (!this.styleTarget()) {
      this.styleTouched = true;
      this.styleWeight = weight;
      return;
    }
    this.updateStyleMark({ weight });
  }

  private copyStyleMark(): void {
    const target = this.styleTarget();
    if (!target) return;
    const rects = target.rects.map((rect) => ({ ...rect, y: rect.y - 4 }));
    const copy: Mark = {
      ...target,
      id: `mark-${++this.markSeq}`,
      rects,
      points: target.points.map((point) => ({ ...point })),
      sortY: rects.length > 0 ? Math.max(...rects.map((rect) => rect.y + rect.height)) : target.sortY,
      replies: [],
      createdAt: Date.now(),
      modifiedAt: Date.now(),
    };
    this.marks = [...this.marks, copy];
    this.selectedMarkId = copy.id;
    void this.refreshMarkBoxes();
  }

  private styleKindLabel(kind: 'highlight' | 'underline' | 'strike' | null): string {
    if (kind === 'highlight') return 'Highlight';
    if (kind === 'underline') return 'Underline';
    if (kind === 'strike') return 'Strikethrough';
    return 'Choose a mark';
  }

  private stylePanel() {
    if (!this.styleMenuOpen || !this.textSelection) return null;
    const look = this.styleLook();
    const percent = Math.round(look.opacity * 100);
    const hasMark = look.kind !== null;
    return html`
      <aside class="style-panel" aria-label="Text style">
        <div class="style-head">
          <div>
            <h2>Text style</h2>
            <p>${this.styleKindLabel(look.kind)}</p>
          </div>
          <button
            class="icon"
            type="button"
            aria-label="Close style"
            @click=${() => {
              this.styleMenuOpen = false;
              this.styleMoreOpen = false;
            }}
          >
            <svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round">
              <path d="M7 7l10 10M17 7 7 17"></path>
            </svg>
          </button>
        </div>
        <section>
          <p class="style-kicker">Mark</p>
          <div class="style-kinds" role="radiogroup" aria-label="Mark">
            ${MARK_TOOLS.filter((item) => this.isTextStyle(item.id)).map(
              (item) => html`
                <button
                  type="button"
                  role="radio"
                  aria-checked=${look.kind === item.id}
                  @click=${() => this.chooseStyleKind(item.id as 'highlight' | 'underline' | 'strike')}
                >
                  ${item.label}
                </button>
              `,
            )}
          </div>
        </section>
        <section>
          <p class="style-kicker">Color</p>
          <div class="style-chips" role="radiogroup" aria-label="Color">
            ${STYLE_SWATCHES.map(
              (color) => html`
                <button
                  class="style-chip"
                  type="button"
                  role="radio"
                  aria-label=${color}
                  aria-checked=${look.color.toLowerCase() === color}
                  style="background:${color}"
                  @click=${() => this.setStyleColor(color)}
                ></button>
              `,
            )}
          </div>
          <label class="style-custom">
            <input
              type="color"
              aria-label="Custom color"
              .value=${/^#[0-9a-f]{6}$/i.test(look.color) ? look.color : '#f6bf26'}
              @input=${(event: Event) => this.setStyleColor((event.target as HTMLInputElement).value)}
            />
            Custom
          </label>
        </section>
        <section>
          <p class="style-kicker">Opacity</p>
          <div class="style-opacity">
            <input
              type="range"
              min="5"
              max="100"
              aria-label="Opacity"
              .value=${String(percent)}
              @input=${(event: Event) => this.setStyleOpacity(Number((event.target as HTMLInputElement).value) / 100)}
            />
            <input
              type="number"
              min="5"
              max="100"
              aria-label="Opacity percent"
              .value=${String(percent)}
              @input=${(event: Event) => this.setStyleOpacity(Number((event.target as HTMLInputElement).value) / 100)}
            />
          </div>
        </section>
        <div class="style-actions">
          <button class="danger" type="button" ?disabled=${!hasMark} @click=${() => {
            const target = this.styleTarget();
            if (target) this.deleteMark(target.id);
          }}>Remove</button>
          <button type="button" ?disabled=${!hasMark} @click=${() => this.copyStyleMark()}>Copy</button>
        </div>
        <section>
          <button
            class="style-more-toggle"
            type="button"
            aria-expanded=${this.styleMoreOpen}
            @click=${() => {
              this.styleMoreOpen = !this.styleMoreOpen;
            }}
          >
            ${this.styleMoreOpen ? 'Fewer options' : 'More options'}
          </button>
          ${this.styleMoreOpen
            ? html`
                <div class="style-more">
                  <div>
                    <p class="style-kicker">Line weight</p>
                    <div class="style-weights" role="radiogroup" aria-label="Line weight">
                      ${STYLE_WEIGHTS.map(
                        (item) => html`
                          <button
                            type="button"
                            role="radio"
                            aria-checked=${Math.abs(look.weight - item.id) < 0.01}
                            ?disabled=${look.kind === 'highlight'}
                            @click=${() => this.setStyleWeight(item.id)}
                          >
                            ${item.label}
                          </button>
                        `,
                      )}
                    </div>
                  </div>
                </div>
              `
            : null}
        </section>
      </aside>
    `;
  }

  private quickSelectionUi() {
    const selection = this.textSelection;
    if (!selection || selection.page !== this.currentPage || selection.boxes.length === 0) return nothing;
    const first = selection.boxes[0];
    const last = selection.boxes[selection.boxes.length - 1];
    const left = Math.min(...selection.boxes.map((box) => box.left));
    const right = Math.max(...selection.boxes.map((box) => box.left + box.width));
    const bottom = Math.max(...selection.boxes.map((box) => box.top + box.height));
    const center = (left + right) / 2;
    const pageWidth = this.canvasEl?.clientWidth ?? 0;
    const place = (preferred: number, half: number) => {
      if (pageWidth <= half * 2 + 12) return pageWidth / 2;
      return Math.min(pageWidth - half - 6, Math.max(half + 6, preferred));
    };
    const barLeft = place(center, 90);
    const canDelete = this.selectionMarkupIds().length > 0;
    return html`
      ${selection.boxes.map(
        (box) => html`
          <div class="text-select" style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px"></div>
        `,
      )}
      <span class="select-handle" style="left:${first.left - 4}px;top:${first.top + first.height / 2 - 4}px"></span>
      <span class="select-handle" style="left:${last.left + last.width - 4}px;top:${last.top + last.height / 2 - 4}px"></span>
      ${this.showSelectionTools
        ? html`
      <div class="quick-bar" style="left:${barLeft}px;top:${bottom + 10}px">
        <button class="quick-comment" type="button" aria-label="Comment" data-tip="Comment" @click=${() => this.commentOnSelection()}>
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path fill="currentColor" d="M6.2 5.2h11.6a2.2 2.2 0 0 1 2.2 2.2v5.7a2.2 2.2 0 0 1-2.2 2.2H12l-3.1 2.5a.7.7 0 0 1-1.1-.6v-1.9H6.2A2.2 2.2 0 0 1 4 13.1V7.4a2.2 2.2 0 0 1 2.2-2.2z"></path>
            <path d="M8 9.1h8M8 12h5" fill="none" stroke="#fff" stroke-width="1.5" stroke-linecap="round"></path>
          </svg>
        </button>
        <button
          class="quick-style"
          type="button"
          aria-label="Style"
          data-tip="Style"
          aria-pressed=${this.styleMenuOpen}
          @click=${() => {
            this.styleMenuOpen = !this.styleMenuOpen;
            this.styleMoreOpen = false;
            if (!this.styleMenuOpen) return;
            this.styleTouched = false;
            const target = this.styleTarget();
            if (target) this.selectedMarkId = target.id;
          }}
        >
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path fill="currentColor" d="M12 3.2a8.8 8.8 0 1 0 1.4 17.4c.7.1 1.3-.4 1.4-1.1.1-.8-.2-1.4-.8-1.8-.5-.3-.7-.8-.6-1.3.2-.8.9-1.3 1.7-1.3H17a4.8 4.8 0 0 0 4.8-5.2A8.6 8.6 0 0 0 12 3.2z"></path>
            <circle cx="8.2" cy="9.2" r="1.35" fill="#f9ab00"></circle>
            <circle cx="12" cy="7.3" r="1.35" fill="#1a73e8"></circle>
            <circle cx="15.6" cy="9.4" r="1.35" fill="#d93025"></circle>
            <circle cx="9.4" cy="13.2" r="1.35" fill="#1e8e3e"></circle>
          </svg>
        </button>
        <button class="quick-link" type="button" aria-label="Add link" data-tip="Add link" @click=${() => this.linkOnSelection()}>
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" aria-hidden="true">
            <path d="M10 13.6a3.6 3.6 0 0 0 5.1.4l1.7-1.7a3.6 3.6 0 0 0-5.1-5.1L10.6 8.3"></path>
            <path d="M14 10.4a3.6 3.6 0 0 0-5.1-.4l-1.7 1.7a3.6 3.6 0 0 0 5.1 5.1l1.1-1.1"></path>
          </svg>
        </button>
        <span class="quick-sep"></span>
        <button
          class="quick-delete"
          type="button"
          aria-label="Delete"
          data-tip=${canDelete ? 'Delete' : 'Nothing to delete'}
          ?disabled=${!canDelete}
          @click=${() => this.deleteOnSelection()}
        >
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path fill="currentColor" d="M9.2 4.4h5.6l.5 1.5H19v1.7H5V5.9h3.7l.5-1.5z"></path>
            <path fill="currentColor" d="M6.6 9.2h10.8l-.7 9.6a1.6 1.6 0 0 1-1.6 1.5H8.9a1.6 1.6 0 0 1-1.6-1.5l-.7-9.6z"></path>
            <path d="M10 12.1v5.2M14 12.1v5.2" fill="none" stroke="#fff" stroke-width="1.4" stroke-linecap="round"></path>
          </svg>
        </button>
      </div>
      `
        : nothing}
    `;
  }

  private openText(kind: 'text' | 'initials', point: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber) return;
      if (kind === 'text' && this.insertKind !== 'text') return;
      if (kind === 'initials' && this.signKind !== 'initials') return;
      const pdfPoint = viewportToPdf(this.viewportOf(page), point.x, point.y);
      this.noteDraft = {
        kind,
        page: pageNumber,
        pdfX: pdfPoint.x,
        pdfY: pdfPoint.y,
        left: point.x,
        top: point.y,
        text: '',
        tip: null,
        rect: null,
        quote: '',
        rects: [],
        author: '',
        createdAt: 0,
      };
    });
  }

  private openCallout(start: { x: number; y: number }, end: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber) return;
      const viewport = this.viewportOf(page);
      const tip = viewportToPdf(viewport, start.x, start.y);
      const label = viewportToPdf(viewport, end.x, end.y);
      this.noteDraft = {
        kind: 'callout',
        page: pageNumber,
        pdfX: label.x,
        pdfY: label.y,
        left: end.x,
        top: end.y,
        text: '',
        tip,
        rect: null,
        quote: '',
        rects: [],
        author: '',
        createdAt: 0,
      };
    });
  }

  private async createPen(kind: 'pen' | 'signature', points: Array<{ x: number; y: number }>): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf || points.length < 2) return;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf) return;
    const viewport = this.viewportOf(page);
    const pdfPoints = points.map((point) => viewportToPdf(viewport, point.x, point.y));
    const xs = pdfPoints.map((point) => point.x);
    const ys = pdfPoints.map((point) => point.y);
    const rect = {
      x: Math.min(...xs),
      y: Math.min(...ys),
      width: Math.max(Math.max(...xs) - Math.min(...xs), 0.5),
      height: Math.max(Math.max(...ys) - Math.min(...ys), 0.5),
    };
    const text = kind === 'signature' ? 'Signature' : 'Pen';
    const mark = this.makeMark(kind, pageNumber, [rect], text, pdfPoints);
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    if (this.currentPage === pageNumber) await this.refreshMarkBoxes();
  }

  private placeDate(point: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    const text = new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber || this.signKind !== 'date') return;
      const pdfPoint = viewportToPdf(this.viewportOf(page), point.x, point.y);
      const mark = this.makeMark('date', pageNumber, [{ x: pdfPoint.x, y: pdfPoint.y, width: 0.01, height: 0.01 }], text);
      this.marks = [...this.marks, mark];
      this.selectedMarkId = mark.id;
      void this.refreshMarkBoxes();
    });
  }

  private addPolygonCorner(point: { x: number; y: number }): void {
    const last = this.draftStroke[this.draftStroke.length - 1];
    if (last && Math.hypot(point.x - last.x, point.y - last.y) < 4) return;
    this.draftStroke = [...this.draftStroke, point];
    this.polygonHover = point;
  }

  private finishPolygon(): void {
    const corners = this.draftStroke.slice();
    if (corners.length < 3) return;
    this.draftStroke = [];
    this.polygonHover = null;
    void this.addShape('polygon', corners);
  }

  private finishMeasure(kind: 'perimeter' | 'area'): void {
    const corners = this.draftStroke.slice();
    if (corners.length < (kind === 'area' ? 3 : 2)) return;
    this.draftStroke = [];
    this.polygonHover = null;
    void this.addMeasure(kind, corners);
  }

  private async addShape(
    kind: ShapeKind,
    canvasPoints: Array<{ x: number; y: number }>,
  ): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf || canvasPoints.length === 0) return;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf) return;
    const viewport = this.viewportOf(page);
    const pdfPoints = canvasPoints.map((point) => viewportToPdf(viewport, point.x, point.y));
    const xs = pdfPoints.map((point) => point.x);
    const ys = pdfPoints.map((point) => point.y);
    const rect = {
      x: Math.min(...xs),
      y: Math.min(...ys),
      width: Math.max(Math.max(...xs) - Math.min(...xs), 0.5),
      height: Math.max(Math.max(...ys) - Math.min(...ys), 0.5),
    };
    const text = SHAPE_TOOLS.find((item) => item.id === kind)?.label ?? kind;
    const points = kind === 'rect' || kind === 'ellipse' ? [] : pdfPoints;
    const mark = this.makeMark(kind, pageNumber, [rect], text, points);
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    if (this.currentPage === pageNumber) await this.refreshMarkBoxes();
  }

  private async addMeasure(
    kind: MeasureKind,
    canvasPoints: Array<{ x: number; y: number }>,
  ): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf || canvasPoints.length === 0) return;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf) return;
    const viewport = this.viewportOf(page);
    const pdfPoints = canvasPoints.map((point) => viewportToPdf(viewport, point.x, point.y));
    const anchor = this.measureAnchor(kind, pdfPoints);
    const text = kind === 'area' ? this.formatArea(pdfPoints) : this.formatLength(pdfPoints);
    const mark = this.makeMark(kind, pageNumber, [{ x: anchor.x, y: anchor.y, width: 0.01, height: 0.01 }], text, pdfPoints);
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    if (this.currentPage === pageNumber) await this.refreshMarkBoxes();
  }

  private formatLength(points: PdfPoint[]): string {
    let length = 0;
    for (let index = 1; index < points.length; index += 1) {
      length += Math.hypot(points[index].x - points[index - 1].x, points[index].y - points[index - 1].y);
    }
    return `${(length / 72).toFixed(2)} in`;
  }

  private formatArea(points: PdfPoint[]): string {
    let sum = 0;
    for (let index = 0; index < points.length; index += 1) {
      const next = points[(index + 1) % points.length];
      sum += points[index].x * next.y - next.x * points[index].y;
    }
    return `${(Math.abs(sum) / 2 / (72 * 72)).toFixed(2)} sq in`;
  }

  private measureAnchor(kind: MeasureKind, points: PdfPoint[]): PdfPoint {
    if (kind === 'distance') {
      return { x: (points[0].x + points[1].x) / 2, y: (points[0].y + points[1].y) / 2 };
    }
    if (kind === 'perimeter') {
      const last = points[points.length - 1];
      return { x: last.x + 8, y: last.y - 8 };
    }
    let twice = 0;
    let cx = 0;
    let cy = 0;
    for (let index = 0; index < points.length; index += 1) {
      const next = points[(index + 1) % points.length];
      const cross = points[index].x * next.y - next.x * points[index].y;
      twice += cross;
      cx += (points[index].x + next.x) * cross;
      cy += (points[index].y + next.y) * cross;
    }
    if (Math.abs(twice) < 0.001) {
      return {
        x: points.reduce((sum, point) => sum + point.x, 0) / points.length,
        y: points.reduce((sum, point) => sum + point.y, 0) / points.length,
      };
    }
    return { x: cx / (3 * twice), y: cy / (3 * twice) };
  }

  private commitNote(event: Event): void {
    event.preventDefault();
    const draft = this.noteDraft;
    if (!draft) return;
    const text = draft.text.trim();
    if (!text) return;
    const anchor = { x: draft.pdfX, y: draft.pdfY, width: 0.01, height: 0.01 };
    let mark: Mark;
    if (draft.kind === 'note') {
      if (draft.rects.length === 0) {
        this.cancelDraft();
        return;
      }
      mark = this.makeMark('note', draft.page, draft.rects, text);
      mark.quote = draft.quote;
      mark.author = draft.author;
      mark.createdAt = draft.createdAt;
      mark.replies = [];
    } else if (draft.kind === 'callout' && draft.tip) {
      mark = this.makeMark('callout', draft.page, [anchor], text, [draft.tip, { x: draft.pdfX, y: draft.pdfY }]);
    } else if (draft.kind === 'link' && draft.rect) {
      mark = this.makeMark('link', draft.page, [draft.rect], text);
    } else if (draft.kind === 'text') {
      mark = this.makeMark('text', draft.page, [anchor], text);
    } else if (draft.kind === 'initials') {
      mark = this.makeMark('initials', draft.page, [anchor], text);
    } else {
      mark = this.makeMark('freetext', draft.page, [anchor], text);
    }
    this.cancelDraft();
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    this.replyText = '';
    this.commentsOpen = true;
    void this.refreshMarkBoxes();
  }

  private cancelDraft(): void {
    this.selectionToken += 1;
    this.noteDraft = null;
    this.notePreview = [];
    this.freeTextBox = null;
    this.linkPageInput = '';
    this.textSelection = null;
    this.styleMenuOpen = false;
    this.editDraft = null;
    this.formDraft = null;
    this.stampMenu = null;
    this.imagePoint = null;
    this.draftBox = null;
    this.draftStroke = [];
    this.draftArrow = 'none';
    this.polygonHover = null;
  }

  private openLink(start: { x: number; y: number }, end: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    this.draftBox = null;
    void pdf.getPage(pageNumber).then(async (page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber || this.insertKind !== 'link') return;
      const viewport = this.viewportOf(page);
      const startPdf = viewportToPdf(viewport, start.x, start.y);
      const endPdf = viewportToPdf(viewport, end.x, end.y);
      const rect = pdfRectFromPoints(startPdf.x, startPdf.y, endPdf.x, endPdf.y);
      const runs = await this.wordsForPage(pageNumber);
      if (this.pdf !== pdf || this.currentPage !== pageNumber || this.insertKind !== 'link') return;
      const hit = runs.filter((run) => rectsIntersect(rect, run.rect));
      this.linkPageInput = '';
      this.linkConfirmId = null;
      this.noteDraft = {
        kind: 'link',
        page: pageNumber,
        pdfX: rect.x,
        pdfY: rect.y,
        left: end.x,
        top: end.y,
        text: '',
        tip: null,
        rect,
        quote: hit.map((run) => run.text).join(' '),
        rects: hit.map((run) => run.rect),
        author: '',
        createdAt: 0,
      };
    });
  }

  private linkTarget(mark: Mark): { url: string } | { page: number } | null {
    const pageMatch = /^page:(\d+)$/.exec(mark.text.trim());
    if (pageMatch) {
      const page = Number(pageMatch[1]);
      if (page >= 1) return { page };
    }
    const url = mark.text.trim();
    return url ? { url } : null;
  }

  private linkSummary(mark: Mark): string {
    const target = this.linkTarget(mark);
    if (!target) return mark.text;
    return 'page' in target ? `Page ${target.page}` : target.url;
  }

  private safeHttpUrl(value: string): string | null {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const withScheme = /^[a-z][a-z0-9+.-]*:/i.test(trimmed) ? trimmed : `https://${trimmed}`;
    try {
      const url = new URL(withScheme);
      if (url.protocol !== 'http:' && url.protocol !== 'https:') return null;
      return url.href;
    } catch {
      return null;
    }
  }

  private linkPageNumber(): number | null {
    const raw = this.linkPageInput.trim();
    if (!/^\d+$/.test(raw)) return null;
    const page = Number(raw);
    if (page < 1 || page > this.totalPages) return null;
    return page;
  }

  private linkDraftReady(): boolean {
    const draft = this.noteDraft;
    if (!draft || draft.kind !== 'link') return false;
    if (draft.text.trim()) return this.safeHttpUrl(draft.text) !== null;
    return this.linkPageNumber() !== null;
  }

  private onLinkUrlInput(event: Event): void {
    const draft = this.noteDraft;
    if (!draft || draft.kind !== 'link') return;
    const text = (event.target as HTMLInputElement).value;
    this.noteDraft = { ...draft, text };
    if (text.trim()) this.linkPageInput = '';
  }

  private onLinkPageInput(event: Event): void {
    const draft = this.noteDraft;
    if (!draft || draft.kind !== 'link') return;
    const text = (event.target as HTMLInputElement).value;
    this.linkPageInput = text;
    if (text.trim()) this.noteDraft = { ...draft, text: '' };
  }

  private commitLink(event: Event): void {
    event.preventDefault();
    const draft = this.noteDraft;
    if (!draft || draft.kind !== 'link' || !draft.rect) return;
    const safeUrl = this.safeHttpUrl(draft.text);
    const page = this.linkPageNumber();
    const stored = safeUrl ?? (page ? `page:${page}` : '');
    if (!stored) return;
    const rects = draft.rects.length > 0 ? draft.rects : [draft.rect];
    const mark = this.makeMark('link', draft.page, rects, stored);
    mark.quote = draft.quote;
    this.cancelDraft();
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    this.replyText = '';
    this.commentsOpen = true;
    void this.refreshMarkBoxes();
  }

  private confirmOpenLink(): void {
    const mark = this.marks.find((item) => item.id === this.linkConfirmId && item.kind === 'link');
    this.linkConfirmId = null;
    if (mark) this.followLink(mark);
  }

  private followLink(mark: Mark): void {
    const target = this.linkTarget(mark);
    if (!target) return;
    if ('page' in target) {
      this.goToPage(target.page);
      return;
    }
    const safeUrl = this.safeHttpUrl(target.url);
    if (!safeUrl) return;
    window.open(safeUrl, '_blank', 'noopener,noreferrer');
  }

  private linkDialog() {
    const draft = this.noteDraft;
    if (!draft || draft.kind !== 'link') return null;
    const bound = draft.quote.trim();
    return html`
      <div
        class="word-modal"
        @click=${(event: Event) => {
          if (event.target === event.currentTarget) this.cancelDraft();
        }}
      >
        <form class="word-dialog link-dialog" role="dialog" aria-label="Add link" @submit=${this.commitLink}>
          <h2>Add link</h2>
          <p class="link-field-label">Link Text</p>
          <p class="link-bound ${bound ? '' : 'muted'}">${bound || 'No text selected'}</p>
          <label class="link-field">
            URL
            <input
              type="text"
              placeholder="Enter URL"
              autofocus
              .value=${draft.text}
              @input=${this.onLinkUrlInput}
            />
          </label>
          <label class="link-field">
            Page
            <input
              type="text"
              inputmode="numeric"
              placeholder="Page number"
              .value=${this.linkPageInput}
              @input=${this.onLinkPageInput}
            />
          </label>
          <div class="word-dialog-actions">
            <button type="button" @click=${() => this.cancelDraft()}>Cancel</button>
            <button class="primary" type="submit" ?disabled=${!this.linkDraftReady()}>Link</button>
          </div>
        </form>
      </div>
    `;
  }

  private linkConfirmDialog() {
    const mark = this.marks.find((item) => item.id === this.linkConfirmId && item.kind === 'link');
    if (!mark) return null;
    const target = this.linkTarget(mark);
    if (!target) return null;
    const destination = 'page' in target ? `page ${target.page}` : target.url;
    const message =
      'page' in target
        ? `This will take you to ${destination}. Are you sure you want to open it?`
        : `This will open a new tab and take you to the ${destination}. Are you sure you want to open it?`;
    return html`
      <div
        class="word-modal"
        @click=${(event: Event) => {
          if (event.target === event.currentTarget) this.linkConfirmId = null;
        }}
      >
        <div class="word-dialog link-confirm" role="dialog" aria-label="Open link">
          <h2>Open link</h2>
          <p class="link-confirm-copy">${message}</p>
          <div class="word-dialog-actions">
            <button type="button" @click=${() => { this.linkConfirmId = null; }}>Cancel</button>
            <button class="primary" type="button" @click=${() => this.confirmOpenLink()}>Confirm</button>
          </div>
        </div>
      </div>
    `;
  }

  private openStampMenu(point: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber || this.insertKind !== 'stamp') return;
      const pdfPoint = viewportToPdf(this.viewportOf(page), point.x, point.y);
      this.stampMenu = { page: pageNumber, pdfX: pdfPoint.x, pdfY: pdfPoint.y, left: point.x, top: point.y };
    });
  }

  private placeStamp(word: string): void {
    const menu = this.stampMenu;
    if (!menu) return;
    const mark = this.makeMark('stamp', menu.page, [{ x: menu.pdfX, y: menu.pdfY, width: 0.01, height: 0.01 }], word);
    this.stampMenu = null;
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    void this.refreshMarkBoxes();
  }

  private onImageChosen(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    const point = this.imagePoint;
    this.imagePoint = null;
    if (!file || !point || this.insertKind !== 'image') return;
    if (!['image/png', 'image/jpeg', 'image/gif', 'image/webp'].includes(file.type)) return;
    const pdf = this.pdf;
    if (!pdf) return;
    void pdf.getPage(point.page).then((page) => {
      if (this.pdf !== pdf) return;
      const pdfPoint = viewportToPdf(this.viewportOf(page), point.x, point.y);
      const reader = new FileReader();
      reader.onload = () => {
        const src = typeof reader.result === 'string' ? reader.result : '';
        if (!src) return;
        const image = new Image();
        image.onload = () => {
          const width = 144;
          const ratio = image.naturalWidth > 0 ? image.naturalHeight / image.naturalWidth : 1;
          const height = Math.max(width * ratio, 1);
          const rect = { x: pdfPoint.x, y: pdfPoint.y - height, width, height };
          const mark = this.makeMark('image', point.page, [rect], file.name, [], src);
          this.marks = [...this.marks, mark];
          this.selectedMarkId = mark.id;
          if (this.currentPage === point.page) void this.refreshMarkBoxes();
        };
        image.src = src;
      };
      reader.readAsDataURL(file);
    });
  }

  private strokeLength(points: Array<{ x: number; y: number }>): number {
    let length = 0;
    for (let index = 1; index < points.length; index += 1) {
      length += Math.hypot(points[index].x - points[index - 1].x, points[index].y - points[index - 1].y);
    }
    return length;
  }

  private makeMark(kind: MarkKind, page: number, rects: PdfRect[], text: string, points: PdfPoint[] = [], src = ''): Mark {
    const tops = [...rects.map((rect) => rect.y + rect.height), ...points.map((point) => point.y)];
    return {
      id: `mark-${++this.markSeq}`,
      kind,
      page,
      rects,
      text,
      sortY: Math.max(...tops),
      points,
      src,
      ...this.defaultInk(kind),
      quote: '',
      author: this.authorName,
      createdAt: Date.now(),
      modifiedAt: Date.now(),
      status: 'none',
      marked: false,
      replies: [],
    };
  }

  private async refreshMarkBoxes(): Promise<void> {
    const pdf = this.pdf;
    if (!pdf) {
      this.markBoxes = [];
      return;
    }
    const pageNumber = this.currentPage;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf || this.currentPage !== pageNumber) return;
    const viewport = this.viewportOf(page);
    const noteNumbers = new Map<string, number>();
    let noteNumber = 0;
    for (const mark of this.marks) {
      if (mark.kind === 'note') noteNumbers.set(mark.id, ++noteNumber);
    }
    const boxes: MarkBox[] = [];
    for (const mark of this.marks) {
      if (mark.page !== pageNumber) continue;
      if (this.appliedFilter.filterDocument && !this.matchesFilter(mark)) continue;
      const selected = mark.id === this.selectedMarkId;
      const label = mark.kind === 'note' ? String(noteNumbers.get(mark.id) ?? '') : '';
      if (mark.kind === 'pen' || mark.kind === 'signature' || mark.kind === 'callout' || mark.kind === 'freetext' || mark.kind === 'text' || mark.kind === 'stamp' || mark.kind === 'initials' || mark.kind === 'date' || mark.kind === 'line' || mark.kind === 'arrow' || mark.kind === 'polygon' || mark.kind === 'distance' || mark.kind === 'perimeter' || mark.kind === 'area') {
        const points = mark.points.map((point) => {
          const [x, y] = viewport.convertToViewportPoint(point.x, point.y);
          return { x, y };
        });
        const anchor = mark.rects[0] ?? { x: 0, y: 0, width: 0.01, height: 0.01 };
        boxes.push({
          ...projectRect(viewport, anchor),
          id: mark.id,
          kind: mark.kind,
          selected,
          label,
          text: mark.text,
          points,
          src: mark.src,
          color: mark.color,
          opacity: mark.opacity,
          weight: mark.weight,
        });
        continue;
      }
      for (const rect of mark.rects) {
        boxes.push({
          ...projectRect(viewport, rect),
          id: mark.id,
          kind: mark.kind,
          selected,
          label,
          text: mark.kind === 'link' ? mark.text : '',
          points: [],
          src: mark.src,
          color: mark.color,
          opacity: mark.opacity,
          weight: mark.weight,
        });
      }
    }
    this.markBoxes = boxes;
    const draft = this.noteDraft;
    if (draft?.kind === 'note' && draft.page === pageNumber && draft.rects.length > 0) {
      this.notePreview = draft.rects.map((rect) => projectRect(viewport, rect));
    } else if (this.notePreview.length > 0) {
      this.notePreview = [];
    }
    if (this.textSelection?.page === pageNumber) {
      this.textSelection = {
        ...this.textSelection,
        boxes: this.textSelection.rects.map((rect) => projectRect(viewport, rect)),
      };
    }
  }

  private focusMark(id: string): void {
    const mark = this.marks.find((item) => item.id === id);
    if (!mark) return;
    if (this.selectedMarkId !== id) this.replyText = '';
    this.selectedMarkId = id;
    if (mark.page !== this.currentPage) this.goToPage(mark.page);
    else void this.refreshMarkBoxes();
  }

  private deleteMark(id: string): void {
    this.marks = this.marks.filter((mark) => mark.id !== id);
    if (this.selectedMarkId === id) this.selectedMarkId = null;
    if (this.editingCommentId === id) this.cancelEditComment();
    this.checkedCommentIds = this.checkedCommentIds.filter((item) => item !== id);
    this.closeCommentMenus();
    void this.refreshMarkBoxes();
  }

  private compareComments(left: Mark, right: Mark): number {
    const position = left.page - right.page || right.sortY - left.sortY || left.id.localeCompare(right.id);
    switch (this.sortBy) {
      case 'created':
        return right.createdAt - left.createdAt || position;
      case 'modified':
        return right.modifiedAt - left.modifiedAt || position;
      case 'status':
        return this.statusRank(left) - this.statusRank(right) || position;
      case 'author':
        return (left.author || 'Guest').localeCompare(right.author || 'Guest') || position;
      case 'type':
        return this.markLabel(left.kind).localeCompare(this.markLabel(right.kind)) || position;
      case 'color':
        return this.colorRank(left) - this.colorRank(right) || position;
      default:
        return position;
    }
  }

  private statusRank(mark: Mark): number {
    const order = ['accepted', 'rejected', 'cancelled', 'completed', 'none'];
    const rank = order.indexOf(mark.status);
    return (rank < 0 ? order.length : rank) * 2 + (mark.marked ? 0 : 1);
  }

  private colorRank(mark: Mark): number {
    if (mark.status === 'accepted') return 0;
    if (mark.status === 'completed') return 1;
    if (mark.status === 'rejected') return 2;
    if (mark.status === 'cancelled') return 3;
    if (mark.marked) return 4;
    switch (mark.kind) {
      case 'highlight':
        return 5;
      case 'note':
      case 'underline':
        return 6;
      case 'strike':
        return 7;
      case 'pen':
      case 'signature':
      case 'freetext':
      case 'text':
      case 'callout':
      case 'stamp':
      case 'initials':
      case 'date':
      case 'line':
      case 'arrow':
      case 'rect':
      case 'ellipse':
      case 'polygon':
        return 8;
      default:
        return 9;
    }
  }

  private get listedMarks(): Mark[] {
    const query = this.commentQuery.trim().toLowerCase();
    const sorted = [...this.marks].sort((left, right) => this.compareComments(left, right)).filter((mark) => this.matchesFilter(mark));
    if (!query) return sorted;
    return sorted.filter((mark) => {
      const status = COMMENT_STATUSES.find((item) => item.id === mark.status)?.label ?? '';
      const haystack = [mark.text, mark.quote, mark.author, status, mark.marked ? 'marked' : '', ...mark.replies.map((reply) => reply.text)]
        .join(' ')
        .toLowerCase();
      return haystack.includes(query);
    });
  }

  private markLabel(kind: MarkKind): string {
    return MARK_TOOLS.find((item) => item.id === kind)?.label
      ?? SHAPE_TOOLS.find((item) => item.id === kind)?.label
      ?? INSERT_TOOLS.find((item) => item.id === kind)?.label
      ?? MEASURE_TOOLS.find((item) => item.id === kind)?.label
      ?? SIGN_TOOLS.find((item) => item.id === kind)?.label
      ?? kind;
  }

  private draftLabel(kind: TextMarkKind): string {
    if (kind === 'freetext') return 'Free text';
    if (kind === 'callout') return 'Callout text';
    if (kind === 'text') return 'Text';
    if (kind === 'link') return 'Link URL';
    if (kind === 'initials') return 'Initials';
    return 'Comment';
  }

  private get authorName(): string {
    const name = this.author.trim();
    return name || 'Guest';
  }

  private get selectedNote(): Mark | undefined {
    return this.marks.find((mark) => mark.id === this.selectedMarkId && mark.kind === 'note');
  }

  private formatWhen(time: number): string {
    if (!time) return '';
    return new Date(time).toLocaleString('en-US', {
      month: 'short',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
    });
  }

  private nameInitial(name: string): string {
    const trimmed = name.trim();
    return trimmed ? trimmed.charAt(0).toUpperCase() : 'G';
  }

  private commentAnchor(): CommentAnchor | null {
    const drafting = this.noteDraft?.kind === 'note' && this.noteDraft.page === this.currentPage;
    const boxes = drafting
      ? this.notePreview
      : this.selectedNote?.page === this.currentPage
        ? this.markBoxes.filter((box) => box.id === this.selectedNote?.id)
        : [];
    if (boxes.length === 0) return null;
    const last = boxes[boxes.length - 1];
    const fromX = last.left + last.width;
    const fromY = last.top + last.height / 2;
    const pageWidth = this.canvasEl?.clientWidth ?? 0;
    const pageHeight = this.canvasEl?.clientHeight ?? 0;
    let left = fromX + 16;
    let top = Math.max(8, boxes[0].top - 8);
    if (pageWidth > 0 && left + COMMENT_POPUP_WIDTH > pageWidth - 8) {
      left = Math.max(8, boxes[0].left - COMMENT_POPUP_WIDTH - 16);
    }
    if (pageHeight > 0) top = Math.min(top, Math.max(8, pageHeight - 160));
    return { left, top, fromX, fromY };
  }

  private addReply(event: Event, markId: string): void {
    event.preventDefault();
    const text = this.replyText.trim();
    if (!text) return;
    const reply: NoteReply = {
      id: `reply-${++this.markSeq}`,
      author: this.authorName,
      createdAt: Date.now(),
      text,
    };
    this.marks = this.marks.map((mark) =>
      mark.id === markId ? { ...mark, replies: [...mark.replies, reply], modifiedAt: Date.now() } : mark,
    );
    this.replyText = '';
  }

  private onReplyInput(event: Event): void {
    this.replyText = (event.target as HTMLInputElement).value;
  }

  private personHeading(name: string, time: number, detail = '') {
    return html`
      <span class="comment-person">
        <span class="avatar">${this.nameInitial(name)}</span>
        <span>
          <span class="person-name">${name}</span>
          <span class="person-time">${this.formatWhen(time)}${detail}</span>
        </span>
      </span>
    `;
  }

  private replyList(replies: NoteReply[]) {
    if (replies.length === 0) return nothing;
    return html`
      <ul class="reply-list">
        ${replies.map(
          (reply) => html`
            <li>
              <span class="avatar small">${this.nameInitial(reply.author)}</span>
              <div>
                <span class="person-name">${reply.author}</span>
                <span class="person-time">${this.formatWhen(reply.createdAt)}</span>
                <p>${reply.text}</p>
              </div>
            </li>
          `,
        )}
      </ul>
    `;
  }

  private replyForm(markId: string) {
    return html`
      <form class="reply-form" @submit=${(event: Event) => this.addReply(event, markId)}>
        <input
          type="text"
          placeholder="Add a reply"
          aria-label="Add a reply"
          .value=${this.replyText}
          @input=${this.onReplyInput}
        />
        <button type="submit" ?disabled=${this.replyText.trim().length === 0}>Reply</button>
      </form>
    `;
  }

  private noteLink() {
    const drafting = this.noteDraft?.kind === 'note' && this.noteDraft.page === this.currentPage;
    const saved = !drafting && this.selectedNote?.page === this.currentPage;
    if (!drafting && !saved) return nothing;
    const anchor = this.commentAnchor();
    if (!anchor) return nothing;
    const toX = anchor.left >= anchor.fromX ? anchor.left : anchor.left + COMMENT_POPUP_WIDTH;
    const toY = anchor.top + 28;
    return svg`
      <svg class="comment-link" aria-hidden="true">
        <line
          x1=${anchor.fromX}
          y1=${anchor.fromY}
          x2=${toX}
          y2=${toY}
          stroke="#c26401"
          stroke-width="1.25"
        ></line>
      </svg>
    `;
  }

  private noteComposer() {
    const draft = this.noteDraft;
    const anchor = this.commentAnchor();
    if (!draft || draft.kind !== 'note' || draft.page !== this.currentPage || !anchor) return nothing;
    return html`
      <form
        class="note-box comment-pop"
        style="left:${anchor.left}px;top:${anchor.top}px"
        @submit=${this.commitNote}
      >
        ${this.personHeading(draft.author, draft.createdAt)}
        <blockquote class="comment-quote">${draft.quote}</blockquote>
        <textarea
          aria-label="Comment"
          placeholder="Write a comment"
          autofocus
          .value=${draft.text}
          @input=${(event: Event) => {
            if (!this.noteDraft) return;
            this.noteDraft = { ...this.noteDraft, text: (event.target as HTMLTextAreaElement).value };
          }}
        ></textarea>
        <div class="note-actions">
          <button type="button" @click=${() => this.cancelDraft()}>Cancel</button>
          <button type="submit" ?disabled=${draft.text.trim().length === 0}>Save</button>
        </div>
      </form>
    `;
  }

  private savedNotePopup() {
    if (this.noteDraft?.kind === 'note') return nothing;
    const mark = this.selectedNote;
    const anchor = this.commentAnchor();
    if (!mark || mark.page !== this.currentPage || !anchor) return nothing;
    if (this.appliedFilter.filterDocument && !this.matchesFilter(mark)) return nothing;
    const tone = mark.status !== 'none' ? `tone-${mark.status}` : '';
    return html`
      <div class="note-box comment-pop ${tone} ${mark.marked ? 'is-marked' : ''}" style="left:${anchor.left}px;top:${anchor.top}px">
        ${this.personHeading(mark.author || 'Guest', mark.createdAt)}
        ${this.commentChips(mark)}
        ${mark.quote ? html`<blockquote class="comment-quote">${mark.quote}</blockquote>` : nothing}
        <p class="comment-body">${mark.text}</p>
        ${this.replyList(mark.replies)}
        ${this.replyForm(mark.id)}
      </div>
    `;
  }

  private cloneFilter(filter: CommentFilterState): CommentFilterState {
    return {
      users: [...filter.users],
      colors: [...filter.colors],
      types: [...filter.types],
      statuses: [...filter.statuses],
      includeReplies: filter.includeReplies,
      filterDocument: filter.filterDocument,
    };
  }

  private filterCount(filter: CommentFilterState): number {
    return filter.users.length + filter.colors.length + filter.types.length + filter.statuses.length;
  }

  private commentSwatch(mark: Mark): string {
    if (mark.status === 'accepted') return '#1e8e3e';
    if (mark.status === 'rejected') return '#d93025';
    if (mark.status === 'cancelled') return '#80868b';
    if (mark.status === 'completed') return '#1a73e8';
    if (mark.marked) return '#e37400';
    switch (mark.kind) {
      case 'highlight':
      case 'underline':
      case 'strike':
        return mark.color;
      case 'note':
        return '#e37400';
      case 'signature':
      case 'distance':
      case 'perimeter':
      case 'area':
      case 'link':
        return '#1a73e8';
      default:
        return '#202124';
    }
  }

  private matchesFilter(mark: Mark, filter = this.appliedFilter): boolean {
    if (filter.users.length > 0) {
      const names = [mark.author || 'Guest'];
      if (filter.includeReplies) {
        for (const reply of mark.replies) names.push(reply.author || 'Guest');
      }
      if (!names.some((name) => filter.users.includes(name))) return false;
    }
    if (filter.colors.length > 0 && !filter.colors.includes(this.commentSwatch(mark))) return false;
    if (filter.types.length > 0 && !filter.types.includes(mark.kind)) return false;
    if (filter.statuses.length > 0 && !filter.statuses.includes(mark.status)) return false;
    return true;
  }

  private filterChoices(tab: FilterTab): Array<{ id: string; label: string; color: string }> {
    if (tab === 'user') {
      const names = new Set<string>();
      for (const mark of this.marks) {
        names.add(mark.author || 'Guest');
        if (this.draftFilter.includeReplies) {
          for (const reply of mark.replies) names.add(reply.author || 'Guest');
        }
      }
      return [...names].sort((left, right) => left.localeCompare(right)).map((name) => ({ id: name, label: name, color: '' }));
    }
    if (tab === 'color') {
      const colors = new Set<string>();
      for (const mark of this.marks) colors.add(this.commentSwatch(mark));
      return [...colors].sort().map((color) => ({ id: color, label: color, color }));
    }
    if (tab === 'type') {
      const kinds = new Map<string, string>();
      for (const mark of this.marks) kinds.set(mark.kind, this.markLabel(mark.kind));
      return [...kinds.entries()]
        .sort((left, right) => left[1].localeCompare(right[1]))
        .map(([id, label]) => ({ id, label, color: '' }));
    }
    return COMMENT_STATUSES.map((item) => ({ id: item.id, label: item.label, color: item.color }));
  }

  private filterSelection(tab: FilterTab): string[] {
    if (tab === 'user') return this.draftFilter.users;
    if (tab === 'color') return this.draftFilter.colors;
    if (tab === 'type') return this.draftFilter.types;
    return this.draftFilter.statuses;
  }

  private toggleFilterChoice(tab: FilterTab, id: string): void {
    const draft = this.cloneFilter(this.draftFilter);
    const key = tab === 'user' ? 'users' : tab === 'color' ? 'colors' : tab === 'type' ? 'types' : 'statuses';
    const current = draft[key];
    draft[key] = current.includes(id) ? current.filter((item) => item !== id) : [...current, id];
    this.draftFilter = draft;
  }

  private setIncludeReplies(includeReplies: boolean): void {
    const draft = this.cloneFilter(this.draftFilter);
    draft.includeReplies = includeReplies;
    if (!includeReplies) {
      const authors = new Set(this.marks.map((mark) => mark.author || 'Guest'));
      draft.users = draft.users.filter((name) => authors.has(name));
    }
    this.draftFilter = draft;
  }

  private openFilter(): void {
    this.closeCommentMenus();
    this.draftFilter = this.cloneFilter(this.appliedFilter);
    this.filterTab = 'user';
    this.filterOpen = true;
  }

  private closeFilter(): void {
    this.filterOpen = false;
  }

  private clearFilterDraft(): void {
    this.draftFilter = {
      ...emptyCommentFilter(),
      includeReplies: this.draftFilter.includeReplies,
      filterDocument: this.draftFilter.filterDocument,
    };
  }

  private applyFilter(): void {
    this.appliedFilter = this.cloneFilter(this.draftFilter);
    this.filterOpen = false;
    void this.refreshMarkBoxes();
  }

  private filterDialog() {
    if (!this.filterOpen) return nothing;
    const choices = this.filterChoices(this.filterTab);
    const selected = this.filterSelection(this.filterTab);
    const count = this.filterCount(this.draftFilter);
    const tabs: FilterTab[] = ['user', 'color', 'type', 'status'];
    return html`
      <div class="filter-backdrop" @click=${this.closeFilter}>
        <div
          class="filter-dialog"
          role="dialog"
          aria-modal="true"
          aria-labelledby="filter-title"
          @click=${(event: Event) => event.stopPropagation()}
        >
          <div class="filter-head">
            <h2 id="filter-title">Filters${count > 0 ? ` (${count})` : ''}</h2>
            <button class="filter-close" type="button" aria-label="Close filters" @click=${this.closeFilter}>×</button>
          </div>
          <div class="filter-tabs" role="tablist" aria-label="Filter by">
            ${tabs.map(
              (tab) => html`
                <button
                  type="button"
                  role="tab"
                  aria-selected=${this.filterTab === tab}
                  @click=${() => {
                    this.filterTab = tab;
                  }}
                >
                  ${tab === 'user' ? 'User' : tab === 'color' ? 'Color' : tab === 'type' ? 'Type' : 'Status'}
                </button>
              `,
            )}
          </div>
          <div class="filter-options" role="tabpanel">
            ${choices.length === 0
              ? html`<p class="filter-empty">No comments to filter yet.</p>`
              : choices.map(
                  (choice) => html`
                    <label class="filter-option">
                      <input
                        type="checkbox"
                        .checked=${selected.includes(choice.id)}
                        aria-label=${this.filterTab === 'color' ? `Color ${choice.color}` : choice.label}
                        @change=${() => this.toggleFilterChoice(this.filterTab, choice.id)}
                      />
                      ${choice.color ? html`<span class="filter-swatch" style="background:${choice.color}"></span>` : nothing}
                      ${this.filterTab === 'color' ? nothing : choice.label}
                    </label>
                  `,
                )}
          </div>
          <div class="filter-settings">
            <h3>Filter settings</h3>
            <label class="filter-option">
              <input
                type="checkbox"
                .checked=${this.draftFilter.includeReplies}
                @change=${(event: Event) => this.setIncludeReplies((event.target as HTMLInputElement).checked)}
              />
              Include replies
            </label>
            <label class="filter-option">
              <input
                type="checkbox"
                .checked=${this.draftFilter.filterDocument}
                @change=${(event: Event) => {
                  this.draftFilter = { ...this.cloneFilter(this.draftFilter), filterDocument: (event.target as HTMLInputElement).checked };
                }}
              />
              Filter document and comments panel
            </label>
          </div>
          <div class="filter-foot">
            <button class="text-btn" type="button" @click=${this.clearFilterDraft}>Clear all</button>
            <button class="apply-btn" type="button" @click=${this.applyFilter}>Apply</button>
          </div>
        </div>
      </div>
    `;
  }

  private closeCommentMenus(): void {
    this.commentMenuId = null;
    this.statusMenuId = null;
    this.menuAnchor = null;
  }

  private openCommentMenu(event: Event, kind: 'status' | 'actions', id: string): void {
    const openId = kind === 'status' ? this.statusMenuId : this.commentMenuId;
    if (openId === id) {
      this.closeCommentMenus();
      return;
    }
    const button = event.currentTarget;
    if (!(button instanceof HTMLElement)) return;
    const rect = button.getBoundingClientRect();
    const menuWidth = kind === 'status' ? 196 : 148;
    const menuHeight = kind === 'status' ? 292 : 88;
    let left = Math.max(8, rect.right - menuWidth);
    let top = rect.bottom + 6;
    if (top + menuHeight > window.innerHeight - 8) top = Math.max(8, rect.top - menuHeight - 6);
    this.menuAnchor = { top, left };
    this.statusMenuId = kind === 'status' ? id : null;
    this.commentMenuId = kind === 'actions' ? id : null;
  }

  private applyCommentAction(id: string, action: CommentAction): void {
    const modifiedAt = Date.now();
    this.marks = this.marks.map((mark) => {
      if (mark.id !== id) return mark;
      if (action === 'marked' || action === 'unmarked') return { ...mark, marked: action === 'marked', modifiedAt };
      return { ...mark, status: action, modifiedAt };
    });
    this.closeCommentMenus();
  }

  private startEditComment(mark: Mark): void {
    this.closeCommentMenus();
    this.editingCommentId = mark.id;
    this.editingCommentText = mark.text;
    this.focusMark(mark.id);
  }

  private saveCommentEdit(event: Event): void {
    event.preventDefault();
    const id = this.editingCommentId;
    const text = this.editingCommentText.trim();
    if (!id || !text) return;
    this.marks = this.marks.map((mark) => (mark.id === id ? { ...mark, text, modifiedAt: Date.now() } : mark));
    this.editingCommentId = null;
    this.editingCommentText = '';
  }

  private cancelEditComment(): void {
    this.editingCommentId = null;
    this.editingCommentText = '';
  }

  private toggleSelecting(): void {
    this.selectingComments = !this.selectingComments;
    this.checkedCommentIds = [];
    this.closeCommentMenus();
  }

  private toggleChecked(id: string): void {
    this.checkedCommentIds = this.checkedCommentIds.includes(id)
      ? this.checkedCommentIds.filter((item) => item !== id)
      : [...this.checkedCommentIds, id];
  }

  private toggleCheckAll(): void {
    const ids = this.listedMarks.map((mark) => mark.id);
    const allChecked = ids.length > 0 && ids.every((id) => this.checkedCommentIds.includes(id));
    this.checkedCommentIds = allChecked ? this.checkedCommentIds.filter((id) => !ids.includes(id)) : [...new Set([...this.checkedCommentIds, ...ids])];
  }

  private deleteChecked(): void {
    const ids = new Set(this.checkedCommentIds);
    if (ids.size === 0) return;
    this.marks = this.marks.filter((mark) => !ids.has(mark.id));
    if (this.selectedMarkId && ids.has(this.selectedMarkId)) this.selectedMarkId = null;
    if (this.editingCommentId && ids.has(this.editingCommentId)) this.cancelEditComment();
    this.checkedCommentIds = [];
    this.closeCommentMenus();
    void this.refreshMarkBoxes();
  }

  private onSortChange(event: Event): void {
    const value = (event.target as HTMLSelectElement).value;
    if (
      value === 'position' ||
      value === 'created' ||
      value === 'modified' ||
      value === 'status' ||
      value === 'author' ||
      value === 'type' ||
      value === 'color'
    ) {
      this.sortBy = value;
    }
  }

  private syncSelectAll(): void {
    const box = this.renderRoot.querySelector('.select-all');
    if (!(box instanceof HTMLInputElement)) return;
    const ids = this.listedMarks.map((mark) => mark.id);
    const count = ids.filter((id) => this.checkedCommentIds.includes(id)).length;
    box.indeterminate = count > 0 && count < ids.length;
  }

  private statusColor(mark: Mark): string {
    if (mark.status === 'accepted') return '#1e8e3e';
    if (mark.status === 'rejected') return '#d93025';
    if (mark.status === 'cancelled') return '#80868b';
    if (mark.status === 'completed') return '#1a73e8';
    if (mark.marked) return '#e37400';
    return '#dadce0';
  }

  private commentChips(mark: Mark) {
    const status = COMMENT_STATUSES.find((item) => item.id === mark.status);
    if ((!status || status.id === 'none') && !mark.marked) return nothing;
    return html`
      <span class="chip-row">
        ${status && status.id !== 'none' ? html`<span class="status-chip tone-${status.id}">${status.label}</span>` : nothing}
        ${mark.marked ? html`<span class="status-chip tone-marked">Marked</span>` : nothing}
      </span>
    `;
  }

  private commentCard(mark: Mark) {
    const selected = mark.id === this.selectedMarkId;
    const editing = this.editingCommentId === mark.id;
    const note = mark.kind === 'note';
    const tone = mark.status !== 'none' ? `tone-${mark.status}` : '';
    return html`
      <article class="comment-card ${selected ? 'selected' : ''} ${tone} ${mark.marked ? 'is-marked' : ''}">
        ${this.selectingComments
          ? html`
              <label class="comment-check">
                <input
                  type="checkbox"
                  .checked=${this.checkedCommentIds.includes(mark.id)}
                  aria-label=${`Select comment ${mark.text || this.markLabel(mark.kind)}`}
                  @change=${() => this.toggleChecked(mark.id)}
                />
              </label>
            `
          : nothing}
        <div class="comment-main">
          <div class="comment-top">
            <button class="comment" type="button" @click=${() => this.focusMark(mark.id)}>
              ${note
                ? this.personHeading(mark.author || 'Guest', mark.createdAt, ` · Page ${mark.page}`)
                : html`
                    <span class="comment-kind">${this.markLabel(mark.kind)}</span>
                    <span class="comment-page">Page ${mark.page}${mark.author ? ` · ${mark.author}` : ''}</span>
                  `}
              ${note && mark.quote ? html`<span class="comment-quote">${mark.quote}</span>` : nothing}
              ${editing ? nothing : html`<span class=${note ? 'comment-body' : 'comment-text'}>${mark.kind === 'link' ? this.linkSummary(mark) : mark.text}</span>`}
            </button>
            <div class="comment-tools">
              <button
                class="comment-menu-btn"
                type="button"
                aria-label="Set comment status"
                aria-haspopup="menu"
                aria-expanded=${this.statusMenuId === mark.id}
                @click=${(event: Event) => this.openCommentMenu(event, 'status', mark.id)}
              >
                <span class="status-dot" style="background:${this.statusColor(mark)}"></span>
              </button>
              <button
                class="comment-menu-btn"
                type="button"
                aria-label="Comment actions"
                aria-haspopup="menu"
                aria-expanded=${this.commentMenuId === mark.id}
                @click=${(event: Event) => this.openCommentMenu(event, 'actions', mark.id)}
              >
                <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" fill="currentColor">
                  <circle cx="12" cy="6" r="1.5"></circle>
                  <circle cx="12" cy="12" r="1.5"></circle>
                  <circle cx="12" cy="18" r="1.5"></circle>
                </svg>
              </button>
            </div>
          </div>
          ${this.commentChips(mark)}
          ${editing
            ? html`
                <form class="comment-edit" @submit=${this.saveCommentEdit}>
                  <textarea
                    aria-label="Edit comment"
                    .value=${this.editingCommentText}
                    @input=${(event: Event) => {
                      this.editingCommentText = (event.target as HTMLTextAreaElement).value;
                    }}
                  ></textarea>
                  <div class="note-actions">
                    <button type="button" @click=${this.cancelEditComment}>Cancel</button>
                    <button type="submit" ?disabled=${this.editingCommentText.trim().length === 0}>Save</button>
                  </div>
                </form>
              `
            : nothing}
          ${note && !editing
            ? selected
              ? html`${this.replyList(mark.replies)} ${this.replyForm(mark.id)}`
              : mark.replies.length > 0
                ? html`<span class="reply-count">${mark.replies.length === 1 ? '1 reply' : `${mark.replies.length} replies`}</span>`
                : nothing
            : nothing}
        </div>
      </article>
    `;
  }

  private statusMenu() {
    const mark = this.marks.find((item) => item.id === this.statusMenuId);
    const anchor = this.menuAnchor;
    if (!mark || !anchor) return nothing;
    return html`
      <div class="comment-menu" role="menu" style="top:${anchor.top}px;left:${anchor.left}px">
        ${COMMENT_STATUSES.map(
          (item) => html`
            <button
              type="button"
              role="menuitemradio"
              aria-checked=${mark.status === item.id}
              @click=${() => this.applyCommentAction(mark.id, item.id)}
            >
              <span class="swatch" style="background:${item.color}"></span>
              <span>${item.label}</span>
            </button>
          `,
        )}
        <hr class="menu-rule" />
        <button type="button" role="menuitemcheckbox" aria-checked=${mark.marked} @click=${() => this.applyCommentAction(mark.id, 'marked')}>
          <span class="swatch marked"></span>
          <span>Marked</span>
        </button>
        <button
          type="button"
          role="menuitemcheckbox"
          aria-checked=${!mark.marked}
          @click=${() => this.applyCommentAction(mark.id, 'unmarked')}
        >
          <span class="swatch" style="background:#fff"></span>
          <span>Unmarked</span>
        </button>
      </div>
    `;
  }

  private actionMenu() {
    const mark = this.marks.find((item) => item.id === this.commentMenuId);
    const anchor = this.menuAnchor;
    if (!mark || !anchor) return nothing;
    return html`
      <div class="comment-menu" role="menu" style="top:${anchor.top}px;left:${anchor.left}px">
        <button type="button" role="menuitem" @click=${() => this.startEditComment(mark)}>Edit</button>
        <button type="button" role="menuitem" class="danger" @click=${() => this.deleteMark(mark.id)}>Delete</button>
      </div>
    `;
  }

  private markFace(box: MarkBox) {
    if (box.kind === 'pen' || box.kind === 'signature' || box.kind === 'line' || box.kind === 'arrow' || box.kind === 'rect' || box.kind === 'ellipse' || box.kind === 'polygon') {
      return nothing;
    }
    if (box.kind === 'freetext' || box.kind === 'callout' || box.kind === 'text' || box.kind === 'stamp' || box.kind === 'distance' || box.kind === 'perimeter' || box.kind === 'area' || box.kind === 'initials' || box.kind === 'date') {
      const fontSize = box.kind === 'initials' ? 18 : 12;
      return html`
        <div
          class="mark ${box.kind} ${box.selected ? 'selected' : ''}"
          data-mark=${box.id}
          style="left:${box.left}px;top:${box.top}px;font-size:${fontSize * this.scale}px"
        >
          ${box.text}
        </div>
      `;
    }
    if (box.kind === 'image') {
      return html`
        <div
          class="mark image ${box.selected ? 'selected' : ''}"
          data-mark=${box.id}
          style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px"
        >
          <img src=${box.src} alt="" />
        </div>
      `;
    }
    if (box.kind === 'link') {
      return html`
        <div
          class="mark link ${box.selected ? 'selected' : ''}"
          data-mark=${box.id}
          style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px"
        ></div>
      `;
    }
    const inked = box.kind === 'highlight' || box.kind === 'underline' || box.kind === 'strike';
    const tint = inked ? this.fadeColor(box.color, box.opacity) : '';
    const thickness = inked ? Math.max(box.weight, box.weight * this.scale) : 0;
    return html`
      <div
        class="mark ${box.kind} ${box.selected ? 'selected' : ''}"
        data-mark=${box.id}
        style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px${inked ? `;--mark-fill:${tint};--mark-ink:${tint};--mark-weight:${thickness}px` : ''}"
      ></div>
    `;
  }

  private inkSvg() {
    const width = Math.max(1, 1.5 * this.scale);
    const signWidth = Math.max(1.5, 2.25 * this.scale);
    const preview =
      (this.shapeKind === 'polygon' || this.measureKind === 'perimeter' || this.measureKind === 'area') && this.polygonHover && this.draftStroke.length > 0
        ? [...this.draftStroke, this.polygonHover]
        : this.draftStroke;
    const draftColor = this.measureKind || this.signKind === 'signature' ? '#1a73e8' : '#202124';
    const draftWidth = this.signKind === 'signature' ? signWidth : width;
    const draftShape = this.shapeKind === 'rect' || this.shapeKind === 'ellipse' ? this.shapeKind : null;
    return svg`
      <svg class="ink">
        <defs>
          <marker id="ink-arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto">
            <path d="M0,0 L10,5 L0,10 Z" fill="#202124"></path>
          </marker>
        </defs>
        ${preview.length > 1 ? this.inkLine(preview, null, this.draftArrow, draftWidth, draftColor) : nothing}
        ${this.draftBox && draftShape ? this.inkOutline(this.draftBox, draftShape, null, width) : nothing}
        ${this.markBoxes.map((box) => {
          if (box.kind === 'rect' || box.kind === 'ellipse') return this.inkOutline(box, box.kind, box, width);
          if (box.kind === 'area' && box.points.length > 2) return this.inkPolygon(box, width, '#1a73e8', 'rgba(26, 115, 232, 0.12)');
          if (box.kind === 'polygon' && box.points.length > 2) return this.inkPolygon(box, width);
          if (box.points.length > 1) {
            const arrow: ArrowEnd = box.kind === 'callout' ? 'start' : box.kind === 'arrow' ? 'end' : 'none';
            const color = box.kind === 'distance' || box.kind === 'perimeter' || box.kind === 'signature' ? '#1a73e8' : '#202124';
            const stroke = box.kind === 'signature' ? signWidth : width;
            return this.inkLine(box.points, box, arrow, stroke, color);
          }
          return nothing;
        })}
      </svg>
    `;
  }

  private inkLine(
    points: Array<{ x: number; y: number }>,
    box: MarkBox | null,
    arrow: ArrowEnd,
    width: number,
    color = '#202124',
  ): SVGTemplateResult {
    const drawn = arrow === 'start' && points.length >= 2 ? [points[points.length - 1], points[0]] : points;
    const pts = drawn.map((point) => `${point.x},${point.y}`).join(' ');
    const visible = arrow === 'none'
      ? svg`<polyline points=${pts} fill="none" stroke=${color} stroke-width=${width} stroke-linecap="round" stroke-linejoin="round" pointer-events="none"></polyline>`
      : svg`<polyline points=${pts} fill="none" stroke=${color} stroke-width=${width} stroke-linecap="round" stroke-linejoin="round" pointer-events="none" marker-end="url(#ink-arrow)"></polyline>`;
    return svg`
      ${box?.selected
        ? svg`<polyline points=${pts} fill="none" stroke="#1a73e8" stroke-width=${width + 4} stroke-linecap="round" stroke-linejoin="round" pointer-events="none"></polyline>`
        : nothing}
      ${box
        ? svg`<polyline points=${pts} fill="none" stroke="transparent" stroke-width=${Math.max(12, width * 4)} stroke-linecap="round" stroke-linejoin="round" pointer-events="stroke" data-mark=${box.id}></polyline>`
        : nothing}
      ${visible}
    `;
  }

  private inkPolygon(box: MarkBox, width: number, stroke = '#202124', fill = 'transparent'): SVGTemplateResult {
    const pts = box.points.map((point) => `${point.x},${point.y}`).join(' ');
    const closed = `${pts} ${box.points[0].x},${box.points[0].y}`;
    return svg`
      ${box.selected
        ? svg`<polyline points=${closed} fill="none" stroke="#1a73e8" stroke-width=${width + 4} stroke-linejoin="round" stroke-linecap="round" pointer-events="none"></polyline>`
        : nothing}
      <polygon points=${pts} fill=${fill} stroke="none" pointer-events="none"></polygon>
      <polygon points=${pts} fill="transparent" stroke="transparent" stroke-width=${Math.max(12, width * 4)} stroke-linejoin="round" pointer-events="all" data-mark=${box.id}></polygon>
      <polyline points=${closed} fill="none" stroke=${stroke} stroke-width=${width} stroke-linejoin="round" stroke-linecap="round" pointer-events="none"></polyline>
    `;
  }

  private inkOutline(
    box: ViewportBox,
    kind: 'rect' | 'ellipse',
    mark: { id: string; selected: boolean } | null,
    width: number,
  ): SVGTemplateResult {
    const selected = mark?.selected ?? false;
    if (kind === 'ellipse') {
      const rx = Math.max(box.width / 2, 0.5);
      const ry = Math.max(box.height / 2, 0.5);
      const cx = box.left + box.width / 2;
      const cy = box.top + box.height / 2;
      return svg`
        ${selected
          ? svg`<ellipse cx=${cx} cy=${cy} rx=${rx} ry=${ry} fill="none" stroke="#1a73e8" stroke-width=${width + 4} pointer-events="none"></ellipse>`
          : nothing}
        ${mark
          ? svg`<ellipse cx=${cx} cy=${cy} rx=${rx} ry=${ry} fill="transparent" stroke="transparent" pointer-events="all" data-mark=${mark.id}></ellipse>`
          : nothing}
        <ellipse cx=${cx} cy=${cy} rx=${rx} ry=${ry} fill="none" stroke="#202124" stroke-width=${width} pointer-events="none"></ellipse>
      `;
    }
    return svg`
      ${selected
        ? svg`<rect x=${box.left} y=${box.top} width=${box.width} height=${box.height} fill="none" stroke="#1a73e8" stroke-width=${width + 4} pointer-events="none"></rect>`
        : nothing}
      ${mark
        ? svg`<rect x=${box.left} y=${box.top} width=${box.width} height=${box.height} fill="transparent" stroke="transparent" pointer-events="all" data-mark=${mark.id}></rect>`
        : nothing}
      <rect x=${box.left} y=${box.top} width=${box.width} height=${box.height} fill="none" stroke="#202124" stroke-width=${width} pointer-events="none"></rect>
    `;
  }

  private hasFiles(event: DragEvent): boolean {
    return Array.from(event.dataTransfer?.types ?? []).includes('Files');
  }

  private onDragEnter(event: DragEvent): void {
    if (!this.hasFiles(event)) return;
    event.preventDefault();
    this.dragDepth += 1;
    this.dragOver = true;
  }

  private onDragOver(event: DragEvent): void {
    if (!this.hasFiles(event)) return;
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
  }

  private onDragLeave(): void {
    this.dragDepth = Math.max(0, this.dragDepth - 1);
    if (this.dragDepth === 0) this.dragOver = false;
  }

  private onDrop(event: DragEvent): void {
    if (!this.hasFiles(event)) return;
    event.preventDefault();
    this.dragDepth = 0;
    this.dragOver = false;
    const file = event.dataTransfer?.files?.[0];
    if (file) void this.load(file);
  }

  private get zoomLabel(): string {
    return `${Math.round(this.scale * 100)}%`;
  }

  private get matchLabel(): string {
    if (this.searching) return 'Searching…';
    if (!this.findQuery.trim()) return '';
    if (this.matchCount === 0) return 'No matches';
    return `${this.matchIndex + 1} of ${this.matchCount}`;
  }

  private cleanPages(pages: number[]): number[] {
    return [...new Set(pages.filter((page) => page >= 1 && page <= this.totalPages))].sort((a, b) => a - b);
  }

  private identityPages(): Map<number, number> {
    const mapping = new Map<number, number>();
    for (let page = 1; page <= this.totalPages; page += 1) mapping.set(page, page);
    return mapping;
  }

  private parsePageList(value: string): { pages: number[]; skipped: boolean } {
    const pages = new Set<number>();
    let skipped = false;
    for (const part of value.split(',')) {
      const piece = part.trim();
      if (!piece) continue;
      const range = /^(\d+)\s*-\s*(\d+)$/.exec(piece);
      if (range) {
        let start = Number(range[1]);
        let end = Number(range[2]);
        if (start > end) [start, end] = [end, start];
        if (end - start > 5000) {
          skipped = true;
          continue;
        }
        for (let page = start; page <= end; page += 1) {
          if (page >= 1 && page <= this.totalPages) pages.add(page);
          else skipped = true;
        }
        continue;
      }
      if (!/^\d+$/.test(piece)) {
        skipped = true;
        continue;
      }
      const page = Number(piece);
      if (page >= 1 && page <= this.totalPages) pages.add(page);
      else skipped = true;
    }
    return { pages: [...pages].sort((a, b) => a - b), skipped };
  }

  private applyPageList(): void {
    const { pages, skipped } = this.parsePageList(this.pageListDraft);
    if (!this.pageListDraft.trim() || pages.length === 0) {
      this.pagePicks = [];
      this.pageNote = this.pageListDraft.trim() ? 'Use page numbers such as 1, 3, 5-10.' : '';
      return;
    }
    this.pagePicks = pages;
    this.pageNote = skipped ? 'Some numbers are outside this document.' : '';
    this.pageMenu = null;
    this.goToPage(pages[0]);
  }

  private clearPagePicks(): void {
    this.pagePicks = [];
    this.pageListDraft = '';
    this.pageNote = '';
    this.pageMenu = null;
  }

  private setThumbWidth(next: number): void {
    const width = Math.min(200, Math.max(72, Math.round(next)));
    if (width !== this.thumbWidth) this.thumbWidth = width;
  }

  private togglePageMenu(page: number, scope: 'one' | 'many'): void {
    if (this.pageMenu?.page === page && this.pageMenu.scope === scope) this.pageMenu = null;
    else this.pageMenu = { page, scope };
  }

  private menuPages(): number[] {
    if (!this.pageMenu) return [];
    if (this.pageMenu.scope === 'many') return this.cleanPages(this.pagePicks);
    return this.cleanPages([this.pageMenu.page]);
  }

  private orderFor(selected: number[], where: PageMove): number[] {
    const all = Array.from({ length: this.totalPages }, (_, index) => index + 1);
    const chosen = new Set(selected);
    const rest = all.filter((page) => !chosen.has(page));
    if (where === 'start') return [...selected, ...rest];
    if (where === 'end') return [...rest, ...selected];
    const first = selected[0];
    const last = selected[selected.length - 1];
    if (where === 'up') {
      if (first <= 1) return all;
      const before = rest.filter((page) => page < first);
      const after = rest.filter((page) => page > first);
      return [...before.slice(0, -1), ...selected, before[before.length - 1], ...after];
    }
    if (last >= this.totalPages) return all;
    const before = rest.filter((page) => page < last);
    const after = rest.filter((page) => page > last);
    return [...before, after[0], ...selected, ...after.slice(1)];
  }

  private async commitPages(bytes: Uint8Array, mapping: Map<number, number>, nextPage: number): Promise<void> {
    const hadPicks = this.pagePicks.length > 0;
    this.pdfBytes = new Uint8Array(bytes);
    this.marks = this.marks.flatMap((mark) => {
      const page = mapping.get(mark.page);
      return page ? [{ ...mark, page }] : [];
    });
    this.textEdits = this.textEdits.flatMap((edit) => {
      const page = mapping.get(edit.page);
      return page ? [{ ...edit, page }] : [];
    });
    this.pageBookmarks = this.pageBookmarks.flatMap((item) => {
      const page = mapping.get(item.page);
      return page ? [{ ...item, page }] : [];
    });
    if (this.bookmarkDraft) {
      const page = mapping.get(this.bookmarkDraft.page);
      this.bookmarkDraft = page ? { ...this.bookmarkDraft, page } : null;
    }
    this.pagePicks = [...new Set(this.pagePicks.flatMap((page) => {
      const next = mapping.get(page);
      return next ? [next] : [];
    }))].sort((a, b) => a - b);
    if (hadPicks) this.pageListDraft = this.pagePicks.join(', ');
    await this.swapDocument(this.pdfBytes);
    this.totalPages = this.pdf?.numPages ?? 0;
    const focus = Math.min(Math.max(nextPage, 1), Math.max(this.totalPages, 1));
    if (focus !== this.currentPage) this.goToPage(focus);
    void this.updateComplete.then(() => this.observeThumbnails());
  }

  private async runPageEdit(work: () => Promise<void>, fallback: string): Promise<void> {
    if (this.pageBusy || !this.pdf) return;
    this.pageBusy = true;
    this.pageMenu = null;
    try {
      await work();
      this.pageNote = '';
    } catch (error) {
      this.reportError(fallback, error);
    } finally {
      this.pageBusy = false;
    }
  }

  private async turnPages(pages: number[], direction: 1 | -1): Promise<void> {
    const selected = this.cleanPages(pages);
    if (!selected.length) return;
    await this.runPageEdit(async () => {
      const bytes = await rotateDocumentPages(await this.ensureBytes(), selected, direction);
      await this.commitPages(bytes, this.identityPages(), this.currentPage);
    }, 'Could not turn those pages.');
  }

  private async movePages(pages: number[], where: PageMove): Promise<void> {
    const selected = this.cleanPages(pages);
    if (!selected.length) return;
    const order = this.orderFor(selected, where);
    if (order.every((page, index) => page === index + 1)) {
      this.pageMenu = null;
      return;
    }
    const mapping = new Map<number, number>();
    order.forEach((page, index) => mapping.set(page, index + 1));
    await this.runPageEdit(async () => {
      const bytes = await rebuildPages(await this.ensureBytes(), order);
      await this.commitPages(bytes, mapping, mapping.get(this.currentPage) ?? 1);
    }, 'Could not move those pages.');
  }

  private async removePages(pages: number[]): Promise<void> {
    const selected = this.cleanPages(pages);
    if (!selected.length) return;
    if (selected.length >= this.totalPages) {
      this.pageMenu = null;
      this.pageNote = 'Keep at least one page in the document.';
      return;
    }
    const drop = new Set(selected);
    const mapping = new Map<number, number>();
    let next = 1;
    for (let page = 1; page <= this.totalPages; page += 1) {
      if (!drop.has(page)) mapping.set(page, next++);
    }
    let focus = mapping.get(this.currentPage) ?? 0;
    if (!focus) {
      const following = Array.from({ length: this.totalPages }, (_, index) => index + 1).find((page) => page > this.currentPage && mapping.has(page));
      const prior = Array.from({ length: this.currentPage - 1 }, (_, index) => this.currentPage - 1 - index).find((page) => mapping.has(page));
      focus = (following ? mapping.get(following) : undefined) ?? (prior ? mapping.get(prior) : undefined) ?? 1;
    }
    await this.runPageEdit(async () => {
      const bytes = await deleteDocumentPages(await this.ensureBytes(), selected);
      await this.commitPages(bytes, mapping, focus);
    }, 'Could not remove those pages.');
  }

  private async insertBlank(pages: number[]): Promise<void> {
    const selected = this.cleanPages(pages);
    if (!selected.length) return;
    const after = selected[selected.length - 1];
    const mapping = new Map<number, number>();
    for (let page = 1; page <= this.totalPages; page += 1) mapping.set(page, page <= after ? page : page + 1);
    const focus = this.currentPage <= after ? this.currentPage : this.currentPage + 1;
    await this.runPageEdit(async () => {
      const bytes = await insertBlankPage(await this.ensureBytes(), after);
      await this.commitPages(bytes, mapping, focus);
    }, 'Could not insert a page.');
  }

  private chooseReplacement(page: number): void {
    this.pageMenu = null;
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'application/pdf,.pdf';
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      if (file) void this.replacePageFromFile(page, file);
    });
    input.click();
  }

  private async replacePageFromFile(page: number, file: File): Promise<void> {
    if (!file.name.toLowerCase().endsWith('.pdf') && file.type !== 'application/pdf') {
      this.pageNote = 'Choose a PDF to replace this page.';
      return;
    }
    await this.runPageEdit(async () => {
      const source = new Uint8Array(await file.arrayBuffer());
      const bytes = await replaceDocumentPage(await this.ensureBytes(), page, source);
      const mapping = this.identityPages();
      mapping.delete(page);
      await this.commitPages(bytes, mapping, page);
    }, 'Could not replace that page.');
  }

  private async savePageCopy(pages: number[]): Promise<void> {
    const selected = this.cleanPages(pages);
    if (!selected.length) return;
    this.pageMenu = null;
    this.pageBusy = true;
    try {
      const bytes = await extractDocumentPages(await this.ensureBytes(), selected);
      const base = this.downloadName.replace(/\.pdf$/i, '') || 'document';
      const label = selected.length === 1 ? `page-${selected[0]}` : 'pages';
      this.downloadFile(bytes, `${base}-${label}.pdf`);
      this.pageNote = '';
    } catch (error) {
      this.reportError('Could not save those pages.', error);
    } finally {
      this.pageBusy = false;
    }
  }

  private pageOps(pages: number[]) {
    const locked = this.pageBusy || !this.pdf;
    const replacePage = pages[0];
    return html`
      <div class="page-ops" role="menu" aria-label="Page actions">
        <div class="view-heading" role="presentation">${pages.length > 1 ? `${pages.length} pages` : `Page ${pages[0]}`}</div>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => void this.movePages(pages, 'start')}>Move to the start</button>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => void this.movePages(pages, 'end')}>Move to the end</button>
        <div class="rule"></div>
        <div class="view-heading" role="presentation">Turn</div>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => void this.turnPages(pages, 1)}>Turn clockwise</button>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => void this.turnPages(pages, -1)}>Turn counterclockwise</button>
        <div class="rule"></div>
        <div class="view-heading" role="presentation">Pages</div>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => void this.insertBlank(pages)}>Insert a blank page</button>
        <button type="button" role="menuitem" ?disabled=${locked || pages.length !== 1} @click=${() => this.chooseReplacement(replacePage)}>Replace from a PDF</button>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => void this.savePageCopy(pages)}>Save a copy</button>
        <button type="button" role="menuitem" class="danger" ?disabled=${locked || pages.length >= this.totalPages} @click=${() => void this.removePages(pages)}>Remove</button>
      </div>
    `;
  }

  private async refreshOutline(): Promise<void> {
    const pdf = this.pdf;
    const token = ++this.outlineToken;
    if (!pdf) {
      this.outlineItems = [];
      this.outlineOpen = [];
      return;
    }
    try {
      const items = await readOutline(pdf);
      if (token !== this.outlineToken || this.pdf !== pdf) return;
      this.outlineItems = items;
      this.outlineOpen = this.outlineBranchIds(items);
    } catch {
      if (token === this.outlineToken) {
        this.outlineItems = [];
        this.outlineOpen = [];
      }
    }
  }

  private outlineBranchIds(items: OutlineEntry[]): string[] {
    const ids: string[] = [];
    const walk = (nodes: OutlineEntry[]) => {
      for (const node of nodes) {
        if (node.items.length === 0) continue;
        ids.push(node.id);
        walk(node.items);
      }
    };
    walk(items);
    return ids;
  }

  private toggleOutline(id: string): void {
    this.outlineOpen = this.outlineOpen.includes(id)
      ? this.outlineOpen.filter((item) => item !== id)
      : [...this.outlineOpen, id];
  }

  private openOutlineEntry(entry: OutlineEntry): void {
    if (entry.page) {
      this.goToPage(entry.page);
      return;
    }
    if (entry.url && /^https?:/i.test(entry.url)) window.open(entry.url, '_blank', 'noopener');
  }

  private outlineTree(items: OutlineEntry[], depth: number): unknown {
    return items.map((item) => {
      const open = this.outlineOpen.includes(item.id);
      return html`
        <div class="outline-row" style=${`padding-left:${depth * 14}px`}>
          ${item.items.length
            ? html`
                <button type="button" class="outline-twist" aria-label=${open ? 'Collapse' : 'Expand'} @click=${() => this.toggleOutline(item.id)}>
                  ${open ? '▾' : '▸'}
                </button>
              `
            : html`<span class="outline-twist"></span>`}
          <button
            type="button"
            class="outline-link"
            aria-current=${item.page === this.currentPage ? 'page' : nothing}
            ?disabled=${item.page == null && !item.url}
            @click=${() => this.openOutlineEntry(item)}
          >
            ${item.title}
          </button>
        </div>
        ${open ? this.outlineTree(item.items, depth + 1) : null}
      `;
    });
  }

  private outlinePane() {
    return html`
      <div class="pane-section-head"><h2>Outline</h2></div>
      <div class="pane-list">
        ${this.outlineItems.length
          ? this.outlineTree(this.outlineItems, 0)
          : html`<p class="bookmark-empty">No outline in this document.</p>`}
      </div>
    `;
  }

  private startBookmark(page = this.currentPage): void {
    this.bookmarkEditing = false;
    this.bookmarkDraft = { page, title: '' };
  }

  private addBookmark(): void {
    const draft = this.bookmarkDraft;
    const title = draft?.title.trim() ?? '';
    if (!draft || !title) return;
    this.pageBookmarks = [...this.pageBookmarks, { id: `bookmark-${Date.now()}`, page: draft.page, title }];
    this.bookmarkDraft = null;
  }

  private removeBookmark(id: string): void {
    this.pageBookmarks = this.pageBookmarks.filter((item) => item.id !== id);
    if (this.pageBookmarks.length === 0) this.bookmarkEditing = false;
  }

  private pageBookmarkButton(page: number) {
    if (!this.showBookmarksOnPage || this.isReadOnlyPreview()) return null;
    const marked = this.pageBookmarks.some((item) => item.page === page);
    return html`
      <button
        class="page-bookmark ${marked ? 'marked' : ''}"
        type="button"
        aria-label=${marked ? `Bookmarks on page ${page}` : `Bookmark page ${page}`}
        @pointerdown=${(event: Event) => event.stopPropagation()}
        @click=${() => {
          this.pagesPressed = true;
          this.pageTab = 'bookmarks';
          if (page !== this.currentPage) this.goToPage(page);
          if (!this.pageBookmarks.some((item) => item.page === page)) this.startBookmark(page);
        }}
      >
        ${iconPaneBookmark()}
      </button>
    `;
  }

  private bookmarkPane() {
    const marks = [...this.pageBookmarks].sort((a, b) => a.page - b.page || a.title.localeCompare(b.title));
    const draft = this.bookmarkDraft;
    return html`
      <div class="pane-section-head">
        <h2>Bookmarks</h2>
        <button
          type="button"
          ?disabled=${marks.length === 0 && !this.bookmarkEditing}
          @click=${() => {
            this.bookmarkEditing = !this.bookmarkEditing;
            if (this.bookmarkEditing) this.bookmarkDraft = null;
          }}
        >
          ${this.bookmarkEditing ? 'Done' : 'Edit'}
        </button>
      </div>
      <label class="pane-toggle">
        <span class="switch">
          <input
            type="checkbox"
            aria-label="Show on the page"
            .checked=${this.showBookmarksOnPage}
            @change=${(event: Event) => {
              this.showBookmarksOnPage = (event.target as HTMLInputElement).checked;
            }}
          />
          <span></span>
        </span>
        Show on the page
      </label>
      <div class="pane-list">
        ${draft
          ? html`
              <form
                class="bookmark-form"
                @submit=${(event: Event) => {
                  event.preventDefault();
                  this.addBookmark();
                }}
              >
                <p>Page ${draft.page}</p>
                <input
                  type="text"
                  aria-label="Bookmark title"
                  placeholder="Title"
                  .value=${draft.title}
                  @input=${(event: Event) => {
                    if (!this.bookmarkDraft) return;
                    this.bookmarkDraft = { ...this.bookmarkDraft, title: (event.target as HTMLInputElement).value };
                  }}
                />
                <div class="bookmark-form-actions">
                  <button type="button" @click=${() => { this.bookmarkDraft = null; }}>Cancel</button>
                  <button class="primary" type="submit" ?disabled=${draft.title.trim().length === 0}>Add</button>
                </div>
              </form>
            `
          : null}
        ${marks.length === 0 && !draft ? html`<p class="bookmark-empty">No bookmarks yet.</p>` : null}
        ${marks.map(
          (item) => html`
            <div class="bookmark-row">
              <button type="button" class="bookmark-main" @click=${() => this.goToPage(item.page)}>
                <span class="bookmark-page">Page ${item.page}</span>
                <span class="bookmark-title">${item.title}</span>
              </button>
              ${this.bookmarkEditing
                ? html`
                    <button type="button" class="bookmark-remove" aria-label=${`Remove ${item.title}`} @click=${() => this.removeBookmark(item.id)}>
                      ${iconTrash()}
                    </button>
                  `
                : null}
            </div>
          `,
        )}
      </div>
      <button type="button" class="pane-add" @click=${() => this.startBookmark()}>+ Add a bookmark</button>
    `;
  }

  private paneTabs() {
    const tabs: Array<{ id: PagePaneTab; label: string; icon: () => SVGTemplateResult }> = [
      { id: 'thumbs', label: 'Thumbnails', icon: iconPaneThumbs },
      { id: 'outline', label: 'Outline', icon: iconPaneOutline },
      { id: 'bookmarks', label: 'Bookmarks', icon: iconPaneBookmark },
      { id: 'layers', label: 'Layers', icon: iconPaneLayers },
      { id: 'files', label: 'Attachments', icon: iconPaneClip },
    ];
    return html`
      <div class="pane-tabs" role="tablist" aria-label="Side panel">
        ${tabs.map(
          (tab) => html`
            <button
              type="button"
              role="tab"
              aria-label=${tab.label}
              aria-selected=${this.pageTab === tab.id}
              title=${tab.label}
              @click=${() => {
                this.pageTab = tab.id;
                this.pageMenu = null;
              }}
            >
              ${tab.icon()}
            </button>
          `,
        )}
      </div>
    `;
  }

  private viewStamp(scale: number): string {
    return `${this.viewTurn}:${scale.toFixed(4)}:${this.layerConfig?.getHash() ?? ''}`;
  }

  private layerPromise(): Promise<NonNullable<typeof this.layerConfig>> | null {
    return this.layerConfig ? Promise.resolve(this.layerConfig) : null;
  }

  private layersFrom(config: NonNullable<typeof this.layerConfig>): DocumentLayer[] {
    const order = config.getOrder();
    if (!Array.isArray(order)) return [];
    const layers: DocumentLayer[] = [];
    const seen = new Set<string>();
    const add = (id: string, depth: number): void => {
      if (seen.has(id)) return;
      const group = config.getGroup(id);
      if (!group) return;
      seen.add(id);
      const name = typeof group.name === 'string' && group.name.trim() ? group.name.trim() : 'Layer';
      layers.push({ id, name, depth, visible: Boolean(group.visible) });
    };
    const walk = (items: unknown[], depth: number): void => {
      for (const item of items) {
        if (typeof item === 'string') {
          add(item, depth);
          continue;
        }
        if (Array.isArray(item)) {
          walk(item, depth + 1);
          continue;
        }
        if (item && typeof item === 'object' && 'order' in item && Array.isArray(item.order)) {
          const folder = item as { name?: unknown; order: unknown[] };
          if (typeof folder.name === 'string' && folder.name.trim()) {
            layers.push({ id: `folder-${layers.length}`, name: folder.name.trim(), depth, visible: true });
          }
          walk(folder.order, typeof folder.name === 'string' && folder.name.trim() ? depth + 1 : depth);
          continue;
        }
        const id = this.layerId(item);
        if (id) add(id, depth);
      }
    };
    walk(order, 0);
    return layers;
  }

  private layerId(item: unknown): string | null {
    if (typeof item === 'string' && item) return item;
    if (item && typeof item === 'object' && 'id' in item && typeof item.id === 'string' && item.id) return item.id;
    return null;
  }

  private async refreshLayers(): Promise<void> {
    const pdf = this.pdf;
    const token = ++this.layerToken;
    if (!pdf) {
      this.layerConfig = null;
      this.documentLayers = [];
      return;
    }
    try {
      const config = await pdf.getOptionalContentConfig();
      if (token !== this.layerToken || this.pdf !== pdf) return;
      const layers = this.layersFrom(config);
      this.layerConfig = layers.length ? config : null;
      this.documentLayers = layers;
      if (layers.length) this.scheduleRender();
    } catch {
      if (token !== this.layerToken) return;
      this.layerConfig = null;
      this.documentLayers = [];
    }
  }

  private setLayerVisible(id: string, visible: boolean): void {
    const config = this.layerConfig;
    if (!config) return;
    config.setVisibility(id, visible, true);
    this.documentLayers = this.documentLayers.map((layer) => ({
      ...layer,
      visible: Boolean(config.getGroup(layer.id)?.visible),
    }));
    this.thumbsPainted.clear();
    this.thumbsPending.clear();
    this.cancelThumbnails();
    this.scheduleRender();
    if (this.pagesPressed && this.pageTab === 'thumbs') {
      void this.updateComplete.then(() => this.observeThumbnails());
    }
  }

  private async refreshFiles(): Promise<void> {
    const pdf = this.pdf;
    const token = ++this.fileToken;
    if (!pdf) {
      this.documentFiles = [];
      return;
    }
    try {
      const map = await pdf.getAttachments();
      if (token !== this.fileToken || this.pdf !== pdf) return;
      const files: DocumentFile[] = [];
      map?.forEach((file, id) => {
        const name = file.filename || file.rawFilename || 'Attachment';
        files.push({
          id,
          name,
          description: file.description || '',
          size: file.content?.byteLength ?? null,
        });
      });
      files.sort((a, b) => a.name.localeCompare(b.name));
      this.documentFiles = files;
    } catch {
      if (token === this.fileToken) this.documentFiles = [];
    }
  }

  private async saveAttachment(file: DocumentFile): Promise<void> {
    const pdf = this.pdf;
    if (!pdf) return;
    try {
      const bytes = await pdf.getAttachmentContent(file.id);
      if (!bytes?.byteLength) {
        this.reportError('Could not save this attachment.', new Error('This attachment is empty.'));
        return;
      }
      this.downloadFile(bytes, file.name, 'application/octet-stream');
    } catch (error) {
      this.reportError('Could not save this attachment.', error);
    }
  }

  private fileSizeLabel(size: number | null): string {
    if (size == null) return 'Save';
    if (size < 1024) return `${size} B`;
    if (size < 1024 * 1024) return `${Math.max(1, Math.round(size / 1024))} KB`;
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }

  private layerPane() {
    return this.documentLayers.length
      ? html`
          <div class="pane-section-head"><h2>Layers</h2></div>
          <div class="pane-list">
            ${this.documentLayers.map(
              (layer) => html`
                ${layer.id.startsWith('folder-')
                  ? html`<div class="layer-row" style=${`padding-left:${8 + layer.depth * 16}px`}><span>${layer.name}</span></div>`
                  : html`
                      <label class="layer-row" style=${`padding-left:${8 + layer.depth * 16}px`}>
                        <input
                          type="checkbox"
                          .checked=${layer.visible}
                          aria-label=${layer.name}
                          @change=${(event: Event) => this.setLayerVisible(layer.id, (event.target as HTMLInputElement).checked)}
                        />
                        <span>${layer.name}</span>
                      </label>
                    `}
              `,
            )}
          </div>
        `
      : html`
          <div class="pane-empty">
            ${iconPaneLayers()}
            <p>This document has no layers.</p>
          </div>
        `;
  }

  private filePane() {
    return this.documentFiles.length
      ? html`
          <div class="pane-section-head"><h2>Attachments</h2></div>
          <div class="pane-list">
            ${this.documentFiles.map(
              (file) => html`
                <button type="button" class="file-row" aria-label=${`Save ${file.name}`} @click=${() => void this.saveAttachment(file)}>
                  <span class="file-name">${file.name}</span>
                  <span class="file-meta">${file.description ? `${file.description} · ` : ''}${this.fileSizeLabel(file.size)}</span>
                </button>
              `,
            )}
          </div>
        `
      : html`
          <div class="pane-empty">
            ${iconPaneClip()}
            <p>This document has no attachments.</p>
          </div>
        `;
  }

  private thumbnailPane() {
    const pages = Array.from({ length: this.totalPages }, (_, index) => index + 1);
    const menuPages = this.menuPages();
    const locked = this.pageBusy || !this.pdf;
    return html`
      <div class="pane-size">
        <button type="button" aria-label="Smaller thumbnails" ?disabled=${this.thumbWidth <= 72} @click=${() => this.setThumbWidth(this.thumbWidth - 16)}>
          ${iconMinus}
        </button>
        <input
          type="range"
          min="72"
          max="200"
          step="8"
          .value=${String(this.thumbWidth)}
          aria-label="Thumbnail size"
          @input=${(event: Event) => this.setThumbWidth(Number((event.target as HTMLInputElement).value))}
        />
        <button type="button" aria-label="Larger thumbnails" ?disabled=${this.thumbWidth >= 200} @click=${() => this.setThumbWidth(this.thumbWidth + 16)}>
          ${iconPlus}
        </button>
      </div>
      <div class="pane-list">
        ${pages.length === 0 ? html`<p class="pages-empty">Open a PDF to see pages.</p>` : null}
        ${repeat(
          pages,
          (page) => page,
          (page) => html`
            <div class="thumb-card ${page === this.currentPage ? 'current' : ''} ${this.pagePicks.includes(page) ? 'picked' : ''} ${this.pageMenu?.scope === 'one' && this.pageMenu.page === page ? 'menu' : ''}">
              <button
                class="thumb"
                type="button"
                data-page=${page}
                aria-label=${`Page ${page}`}
                aria-current=${page === this.currentPage ? 'page' : nothing}
                @click=${() => this.goToPage(page)}
              >
                <span class="thumb-frame" style=${`width:${this.thumbWidth}px`}><canvas></canvas></span>
                <span class="thumb-num">${page}</span>
              </button>
              <div class="thumb-actions">
                <button type="button" aria-label="Turn counterclockwise" ?disabled=${locked} @click=${() => void this.turnPages([page], -1)}>
                  ${iconTurnLeft()}
                </button>
                <button type="button" class="danger" aria-label="Remove page" ?disabled=${locked || this.totalPages < 2} @click=${() => void this.removePages([page])}>
                  ${iconTrash()}
                </button>
                <button
                  class="thumb-more"
                  type="button"
                  aria-label="Page actions"
                  aria-expanded=${this.pageMenu?.scope === 'one' && this.pageMenu.page === page}
                  ?disabled=${locked}
                  @click=${() => this.togglePageMenu(page, 'one')}
                >
                  ${iconMore()}
                </button>
              </div>
            </div>
          `,
        )}
      </div>
      ${menuPages.length ? this.pageOps(menuPages) : null}
      ${this.pagePicks.length
        ? html`
            <div class="pick-bar">
              <button type="button" aria-label="Turn counterclockwise" ?disabled=${locked} @click=${() => void this.turnPages(this.pagePicks, -1)}>${iconTurnLeft()}</button>
              <button type="button" aria-label="Turn clockwise" ?disabled=${locked} @click=${() => void this.turnPages(this.pagePicks, 1)}>${iconTurnRight()}</button>
              <button type="button" aria-label="Move earlier" ?disabled=${locked} @click=${() => void this.movePages(this.pagePicks, 'up')}>${iconMoveUp()}</button>
              <button type="button" aria-label="Move later" ?disabled=${locked} @click=${() => void this.movePages(this.pagePicks, 'down')}>${iconMoveDown()}</button>
              <button
                class="pick-more"
                type="button"
                aria-label="Actions for selected pages"
                aria-expanded=${this.pageMenu?.scope === 'many'}
                ?disabled=${locked}
                @click=${() => this.togglePageMenu(0, 'many')}
              >
                ${iconMore()}
              </button>
            </div>
          `
        : null}
      <div class="pane-foot">
        ${this.pageNote ? html`<p class="pane-note">${this.pageNote}</p>` : null}
        <p>Pages to use, such as 1, 3, 5-10</p>
        <div class="pane-foot-row">
          <input
            type="text"
            .value=${this.pageListDraft}
            placeholder="1, 3, 5-10"
            aria-label="Pages to use"
            @input=${(event: Event) => {
              this.pageListDraft = (event.target as HTMLInputElement).value;
            }}
            @keydown=${(event: KeyboardEvent) => {
              if (event.key !== 'Enter') return;
              event.preventDefault();
              this.applyPageList();
            }}
          />
          <button type="button" aria-label="Use these pages" @click=${() => this.applyPageList()}>${iconPaneCheck()}</button>
          ${this.pagePicks.length || this.pageListDraft
            ? html`<button type="button" aria-label="Clear pages" @click=${() => this.clearPagePicks()}>${iconPaneClear()}</button>`
            : null}
        </div>
      </div>
    `;
  }

  private get pagesPanel() {
    if (!this.pagesPressed) return null;
    return html`
      <aside class="pages" aria-label="Thumbnails">
        ${this.paneTabs()}
        ${this.pageTab === 'thumbs'
          ? this.thumbnailPane()
          : this.pageTab === 'outline'
            ? this.outlinePane()
            : this.pageTab === 'bookmarks'
              ? this.bookmarkPane()
              : this.pageTab === 'layers'
                ? this.layerPane()
                : this.filePane()}
      </aside>
    `;
  }

  private viewChoice(label: string, checked: boolean, action: () => void, disabled = false) {
    return html`
      <button type="button" role="menuitemradio" aria-checked=${checked} ?disabled=${disabled} @click=${action}>
        <span>${label}</span>
        <span>${checked ? '✓' : ''}</span>
      </button>
    `;
  }

  private viewMenu() {
    const locked = !this.pdf;
    const actual = this.zoomMode === 'percent' && Math.abs(this.zoomPercent - 1) < 0.001;
    return html`
      <div class="popup view-menu" role="menu" aria-label="View">
        <div class="view-heading" role="presentation">Sheets</div>
        ${this.viewChoice('One page wide', this.viewSpread === 'single', () => this.setViewSpread('single'), locked)}
        ${this.viewChoice('Two pages wide', this.viewSpread === 'pair', () => this.setViewSpread('pair'), locked)}
        ${this.viewChoice('First page alone', this.viewSpread === 'cover', () => this.setViewSpread('cover'), locked)}
        <div class="rule"></div>
        <div class="view-heading" role="presentation">Motion</div>
        ${this.viewChoice('Stay on the page', this.viewRead === 'page', () => this.setViewRead('page'), locked)}
        ${this.viewChoice('Move through pages', this.viewRead === 'scroll', () => this.setViewRead('scroll'), locked)}
        <div class="rule"></div>
        <div class="view-heading" role="presentation">Angle</div>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => this.turnView(1)}>
          <span>Quarter turn right</span>
        </button>
        <button type="button" role="menuitem" ?disabled=${locked} @click=${() => this.turnView(-1)}>
          <span>Quarter turn left</span>
        </button>
        <button type="button" role="menuitem" ?disabled=${locked || this.viewTurn === 0} @click=${() => this.turnView(0)}>
          <span>Original angle</span>
        </button>
        <div class="rule"></div>
        <div class="view-heading" role="presentation">Display</div>
        <button type="button" role="menuitemradio" aria-checked=${this.zoomMode === 'fit-page'} ?disabled=${locked} @click=${() => this.chooseFit('fit-page')}>
          <span>Fit the page</span>
          <span>${this.zoomMode === 'fit-page' ? '✓' : ''}</span>
        </button>
        <button type="button" role="menuitemradio" aria-checked=${this.zoomMode === 'fit-width'} ?disabled=${locked} @click=${() => this.chooseFit('fit-width')}>
          <span>Fit the width</span>
          <span>${this.zoomMode === 'fit-width' ? '✓' : ''}</span>
        </button>
        <button type="button" role="menuitemradio" aria-checked=${actual} ?disabled=${locked} @click=${() => this.actualSize()}>
          <span>Actual size</span>
          <span>${actual ? '✓' : ''}</span>
        </button>
        <button type="button" role="menuitemcheckbox" aria-checked=${this.pagesPressed} ?disabled=${locked} @click=${() => this.togglePageList()}>
          <span>Thumbnails</span>
          <span>${this.pagesPressed ? '✓' : ''}</span>
        </button>
        <button type="button" role="menuitemcheckbox" aria-checked=${this.wordFullScreen} @click=${() => this.toggleWordFullScreen()}>
          <span>${this.wordFullScreen ? 'Leave full screen' : 'Full screen'}</span>
          <span>${this.wordFullScreen ? '✓' : ''}</span>
        </button>
      </div>
    `;
  }

  private zoomChecked(step: number): boolean {
    return this.zoomMode === 'percent' && Math.abs(this.zoomPercent - step) < 0.001;
  }

  override render() {
    if (this.isReadOnlyPreview()) return this.wordView();
    return html`
      <div
        class="root ${this.previewTheme === 'dark' ? 'preview-dark' : ''}"
        @dragenter=${this.onDragEnter}
        @dragover=${this.onDragOver}
        @dragleave=${this.onDragLeave}
        @drop=${this.onDrop}
      >
        <header class="toolbar">
          <div class="cluster">
            <div class="anchor menu-root">
              <button
                class="icon"
                type="button"
                aria-label="Menu"
                aria-haspopup="menu"
                aria-expanded=${this.menuOpen}
                @click=${() => {
                  this.menuOpen = !this.menuOpen;
                  this.zoomOpen = false;
                  this.searchOpen = false;
                  this.viewMenuOpen = false;
                }}
              >
                ${iconMenu}
              </button>
              ${this.menuOpen
                ? html`
                    <div class="popup file-menu" role="menu">
                      <button type="button" role="menuitem" @click=${() => this.chooseFile()}>
                        <span class="menu-item">${iconOpen}<span>Open</span></span>
                      </button>
                      <button type="button" role="menuitem" ?disabled=${!this.fileReady} @click=${() => this.downloadPdf()}>
                        <span class="menu-item">${iconPreviewDownload()}<span>Download</span></span>
                      </button>
                      <button type="button" role="menuitem" @click=${() => this.toggleWordFullScreen()}>
                        <span class="menu-item">
                          ${this.wordFullScreen ? iconExitFullScreen() : iconEnterFullScreen()}
                          <span>${this.wordFullScreen ? 'Exit Full Screen' : 'Enter Full Screen'}</span>
                        </span>
                      </button>
                      <button type="button" role="menuitem" ?disabled=${!this.fileReady} @click=${() => void this.savePdfAs()}>
                        <span class="menu-item">${iconSave}<span>Save As</span></span>
                      </button>
                      <button type="button" role="menuitem" ?disabled=${!this.fileReady} @click=${() => void this.printPdf()}>
                        <span class="menu-item">${iconPreviewPrint()}<span>Print</span></span>
                      </button>
                      <div class="rule"></div>
                      <button type="button" role="menuitem" @click=${() => this.openPreviewSettings()}>
                        <span class="menu-item">${iconPreviewSettings()}<span>Settings</span></span>
                      </button>
                    </div>
                  `
                : null}
            </div>
            <button
              class="icon"
              type="button"
              aria-label="Thumbnails"
              aria-pressed=${this.pagesPressed}
              @click=${() => {
                this.pagesPressed = !this.pagesPressed;
              }}
            >
              ${iconPages}
            </button>
            <div class="anchor view-root">
              <button
                class="icon"
                type="button"
                aria-label="View"
                aria-haspopup="menu"
                aria-pressed=${this.viewMenuOpen}
                aria-expanded=${this.viewMenuOpen}
                @click=${() => {
                  this.viewMenuOpen = !this.viewMenuOpen;
                  this.menuOpen = false;
                  this.zoomOpen = false;
                  this.searchOpen = false;
                }}
              >
                ${iconEye}
              </button>
              ${this.viewMenuOpen ? this.viewMenu() : null}
            </div>
            <button class="icon" type="button" aria-label="Save" ?disabled=${!this.pdf || this.saving} @click=${() => void this.save()}>${iconSave}</button>
            <span class="sep"></span>
            <div class="anchor zoom-root">
              <button
                class="zoom-button"
                type="button"
                aria-label="Zoom"
                aria-haspopup="menu"
                aria-expanded=${this.zoomOpen}
                @click=${() => {
                  this.zoomOpen = !this.zoomOpen;
                  this.menuOpen = false;
                  this.searchOpen = false;
                  this.viewMenuOpen = false;
                }}
              >
                <span>${this.zoomLabel}</span>
                ${iconChevron}
              </button>
              ${this.zoomOpen
                ? html`
                    <div class="popup" role="menu">
                      ${ZOOM_STEPS.map(
                        (step) => html`
                          <button
                            type="button"
                            role="menuitemradio"
                            aria-checked=${this.zoomChecked(step)}
                            @click=${() => this.setPercent(step)}
                          >
                            <span>${Math.round(step * 100)}%</span>
                            <span>${this.zoomChecked(step) ? '✓' : ''}</span>
                          </button>
                        `,
                      )}
                      <div class="rule"></div>
                      <button
                        type="button"
                        role="menuitemradio"
                        aria-checked=${this.zoomMode === 'fit-width'}
                        @click=${() => this.setFit('fit-width')}
                      >
                        <span>Fit width</span>
                        <span>${this.zoomMode === 'fit-width' ? '✓' : ''}</span>
                      </button>
                      <button
                        type="button"
                        role="menuitemradio"
                        aria-checked=${this.zoomMode === 'fit-page'}
                        @click=${() => this.setFit('fit-page')}
                      >
                        <span>Fit page</span>
                        <span>${this.zoomMode === 'fit-page' ? '✓' : ''}</span>
                      </button>
                    </div>
                  `
                : null}
            </div>
            <button class="icon circle" type="button" aria-label="Zoom out" @click=${() => this.stepZoom(-1)}>
              ${iconMinus}
            </button>
            <button class="icon circle" type="button" aria-label="Zoom in" @click=${() => this.stepZoom(1)}>
              ${iconPlus}
            </button>
            <span class="sep"></span>
            <button
              class="icon"
              type="button"
              aria-label="Pan"
              aria-pressed=${this.tool === 'hand'}
              @click=${() => {
                this.tool = 'hand';
                this.markupKind = null;
                this.shapeKind = null;
                this.insertKind = null;
                this.measureKind = null;
                this.editText = false;
                this.signKind = null;
                this.formKind = null;
                this.lineBoxes = [];
                this.fieldBoxes = [];
                this.cancelDraft();
              }}
            >
              ${iconHand}
            </button>
            <button
              class="icon"
              type="button"
              aria-label="Select"
              aria-pressed=${this.tool === 'select' && !this.markupKind && !this.shapeKind && !this.insertKind && !this.measureKind && !this.editText && !this.signKind && !this.formKind}
              @click=${() => {
                this.tool = 'select';
                this.markupKind = null;
                this.shapeKind = null;
                this.insertKind = null;
                this.measureKind = null;
                this.editText = false;
                this.signKind = null;
                this.formKind = null;
                this.lineBoxes = [];
                this.cancelDraft();
                if (this.mode === 'forms') void this.refreshFormFields();
              }}
            >
              ${iconSelect}
            </button>
          </div>
          <div class="tabs" role="tablist" aria-label="Editor modes">
            ${MODES.map(
              (mode) => html`
                <button
                  class="tab"
                  type="button"
                  role="tab"
                  aria-selected=${this.mode === mode.id}
                  @click=${() => {
                    if (mode.id === this.mode) return;
                    this.cancelDraft();
                    this.mode = mode.id;
                    if (mode.id !== 'annotate') this.markupKind = null;
                    if (mode.id !== 'shapes') this.shapeKind = null;
                    if (mode.id !== 'insert') this.insertKind = null;
                    if (mode.id !== 'measure') this.measureKind = null;
                    if (mode.id !== 'edit') {
                      this.editText = false;
                      this.lineBoxes = [];
                    }
                    if (mode.id !== 'fill-sign') this.signKind = null;
                    if (mode.id !== 'forms') {
                      this.formKind = null;
                      this.fieldBoxes = [];
                    } else void this.refreshFormFields();
                  }}
                >
                  ${mode.label}
                </button>
              `,
            )}
          </div>
          <div class="cluster">
            <div class="anchor search-root">
              <button
                class="icon"
                type="button"
                aria-label="Find in document"
                aria-expanded=${this.searchOpen}
                @click=${() => {
                  this.searchOpen = !this.searchOpen;
                  this.menuOpen = false;
                  this.zoomOpen = false;
                  this.viewMenuOpen = false;
                  if (this.searchOpen) {
                    void this.updateComplete.then(() => {
                      this.searchRoot?.querySelector('input')?.focus();
                    });
                  }
                }}
              >
                ${iconSearch}
              </button>
              ${this.searchOpen
                ? html`
                    <div class="search-pop">
                      <input
                        type="search"
                        placeholder="Find in document"
                        aria-label="Find in document"
                        .value=${this.findQuery}
                        @input=${this.onFindInput}
                        @keydown=${this.onFindKeydown}
                      />
                      <div class="search-meta">
                        <span>${this.matchLabel}</span>
                        <button
                          class="match-nav prev"
                          type="button"
                          aria-label="Previous match"
                          ?disabled=${this.matchCount === 0}
                          @click=${() => this.showMatch(this.matchIndex - 1)}
                        >
                          ${iconChevron}
                        </button>
                        <button
                          class="match-nav next"
                          type="button"
                          aria-label="Next match"
                          ?disabled=${this.matchCount === 0}
                          @click=${() => this.showMatch(this.matchIndex + 1)}
                        >
                          ${iconChevron}
                        </button>
                      </div>
                    </div>
                  `
                : null}
            </div>
            <button
              class="icon"
              type="button"
              aria-label="Comments"
              aria-pressed=${this.commentsOpen}
              @click=${() => {
                this.commentsOpen = !this.commentsOpen;
              }}
            >
              ${iconComments}
            </button>
          </div>
        </header>
        ${this.mode === 'annotate'
          ? html`
              <div class="annotate-bar" role="toolbar" aria-label="Annotate">
                ${MARK_TOOLS.map(
                  (item) => html`
                    <button
                      type="button"
                      aria-pressed=${this.markupKind === item.id}
                      @click=${() => this.chooseMarkup(item.id)}
                    >
                      ${item.label}
                    </button>
                  `,
                )}
              </div>
            `
          : null}
        ${this.mode === 'shapes'
          ? html`
              <div class="annotate-bar" role="toolbar" aria-label="Shapes">
                ${SHAPE_TOOLS.map(
                  (item) => html`
                    <button
                      type="button"
                      aria-pressed=${this.shapeKind === item.id}
                      @click=${() => this.chooseShape(item.id)}
                    >
                      ${item.label}
                    </button>
                  `,
                )}
              </div>
            `
          : null}
        ${this.mode === 'insert'
          ? html`
              <div class="annotate-bar" role="toolbar" aria-label="Insert">
                ${INSERT_TOOLS.map(
                  (item) => html`
                    <button
                      type="button"
                      aria-pressed=${this.insertKind === item.id}
                      @click=${() => this.chooseInsert(item.id)}
                    >
                      ${item.label}
                    </button>
                  `,
                )}
              </div>
            `
          : null}
        ${this.mode === 'measure'
          ? html`
              <div class="annotate-bar" role="toolbar" aria-label="Measure">
                ${MEASURE_TOOLS.map(
                  (item) => html`
                    <button
                      type="button"
                      aria-pressed=${this.measureKind === item.id}
                      @click=${() => this.chooseMeasure(item.id)}
                    >
                      ${item.label}
                    </button>
                  `,
                )}
              </div>
            `
          : null}
        ${this.mode === 'edit'
          ? html`
              <div class="annotate-bar" role="toolbar" aria-label="Edit">
                <button type="button" aria-pressed=${this.editText} @click=${() => this.chooseEdit()}>Edit text</button>
              </div>
            `
          : null}
        ${this.mode === 'fill-sign'
          ? html`
              <div class="annotate-bar" role="toolbar" aria-label="Fill and Sign">
                ${SIGN_TOOLS.map(
                  (item) => html`
                    <button
                      type="button"
                      aria-pressed=${this.signKind === item.id}
                      @click=${() => this.chooseSign(item.id)}
                    >
                      ${item.label}
                    </button>
                  `,
                )}
              </div>
            `
          : null}
        ${this.mode === 'forms'
          ? html`
              <div class="annotate-bar" role="toolbar" aria-label="Forms">
                ${FORM_TOOLS.map(
                  (item) => html`
                    <button
                      type="button"
                      aria-pressed=${this.formKind === item.id}
                      @click=${() => this.chooseForm(item.id)}
                    >
                      ${item.label}
                    </button>
                  `,
                )}
              </div>
            `
          : null}
        <div class="workspace">
          ${this.stylePanel()}
          ${this.pagesPanel}
          <div
            class="viewer ${this.tool === 'hand' ? 'hand' : ''} ${this.panning ? 'panning' : ''} ${this.canDraw ? 'markup' : ''} ${this.markupKind === 'note' ? 'note-tool' : ''} ${!this.canDraw && this.tool === 'select' && !this.editText ? 'select-text' : ''}"
            tabindex="0"
            @pointerdown=${this.onPointerDown}
            @pointermove=${this.onPointerMove}
            @pointerup=${this.onPointerUp}
            @pointercancel=${this.onPointerCancel}
            @scroll=${this.onViewerScroll}
          >
            <div class="stage ${this.viewRead === 'scroll' ? 'is-flow' : ''}">
              ${repeat(
                this.visibleRows(),
                (row) => row.join('-'),
                (row) => html`
                  <div class="sheet-row">
                    ${repeat(
                      row,
                      (page) => page,
                      (page) => html`
                        <div class="page-frame ${this.pageReady ? '' : 'is-hidden'}" data-sheet=${page}>
                          <canvas
                            class=${page === this.currentPage ? 'page-canvas' : 'sheet-canvas'}
                            data-sheet-canvas=${page}
                          ></canvas>
                          ${page === this.currentPage
                            ? html`
                <div class="marks">
                  ${this.inkSvg()}
                  ${this.markBoxes.map((box) => this.markFace(box))}
                  ${this.draftBox && this.shapeKind !== 'rect' && this.shapeKind !== 'ellipse'
                    ? html`
                        <div
                          class="draft"
                          style="left:${this.draftBox.left}px;top:${this.draftBox.top}px;width:${this.draftBox.width}px;height:${this.draftBox.height}px"
                        ></div>
                      `
                    : null}
                </div>
                ${this.notePreview.map(
                  (box) => html`
                    <div
                      class="mark note pending"
                      style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px"
                    ></div>
                  `,
                )}
                ${this.canEditText
                  ? html`
                      <div class="text-layer">
                        ${this.lineBoxes.map(
                          (line) => html`
                            <button
                              class="text-line"
                              type="button"
                              style="left:${line.left}px;top:${line.top}px;width:${line.width}px;height:${line.height}px;font-size:${line.fontSize * this.scale}px"
                              @click=${() => this.openLine(line)}
                            >
                              ${line.text}
                            </button>
                          `,
                        )}
                      </div>
                    `
                  : null}
                ${this.mode === 'forms' && this.tool !== 'hand'
                  ? html`
                      <div class="field-layer">
                        ${this.fieldBoxes.map(
                          (field) => html`
                            <button
                              class="field-target ${this.highlightFormFields ? 'highlighted' : ''}"
                              type="button"
                              aria-label=${field.name}
                              style="left:${field.left}px;top:${field.top}px;width:${field.width}px;height:${field.height}px"
                              @click=${() => this.openField(field)}
                            ></button>
                          `,
                        )}
                      </div>
                    `
                  : null}
                <div class="highlights">
                  ${this.highlights.map(
                    (box) => html`
                      <div
                        class="hit ${box.active ? 'active' : ''}"
                        style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px"
                      ></div>
                    `,
                  )}
                </div>
                ${this.noteLink()}
                ${this.noteComposer()}
                ${this.savedNotePopup()}
                ${this.freeTextEditor()}
                ${this.quickSelectionUi()}
                ${this.noteDraft && this.noteDraft.kind !== 'note' && this.noteDraft.kind !== 'freetext' && this.noteDraft.kind !== 'link' && this.noteDraft.page === this.currentPage
                  ? html`
                      <form
                        class="note-box"
                        style="left:${this.noteDraft.left}px;top:${this.noteDraft.top}px"
                        @submit=${this.commitNote}
                      >
                        <textarea
                          aria-label=${this.draftLabel(this.noteDraft.kind)}
                          autofocus
                          .value=${this.noteDraft.text}
                          @input=${(event: Event) => {
                            if (!this.noteDraft) return;
                            this.noteDraft = { ...this.noteDraft, text: (event.target as HTMLTextAreaElement).value };
                          }}
                        ></textarea>
                        <div class="note-actions">
                          <button type="button" @click=${() => this.cancelDraft()}>Cancel</button>
                          <button type="submit" ?disabled=${this.noteDraft.text.trim().length === 0}>Save</button>
                        </div>
                      </form>
                    `
                  : null}
                ${this.editDraft && this.editDraft.page === this.currentPage
                  ? html`
                      <form class="note-box" style="left:${this.editDraft.left}px;top:${this.editDraft.top}px" @submit=${this.commitEdit}>
                        <textarea
                          aria-label="Edit text"
                          autofocus
                          .value=${this.editDraft.text}
                          @input=${(event: Event) => {
                            if (!this.editDraft) return;
                            this.editDraft = { ...this.editDraft, text: (event.target as HTMLTextAreaElement).value };
                          }}
                        ></textarea>
                        <div class="note-actions">
                          <button type="button" @click=${() => this.cancelDraft()}>Cancel</button>
                          <button type="submit" ?disabled=${this.editDraft.text.trim().length === 0}>Save</button>
                        </div>
                      </form>
                    `
                  : null}
                ${this.formDraft && this.formDraft.page === this.currentPage
                  ? html`
                      <form class="note-box form-box" style="left:${this.formDraft.left}px;top:${this.formDraft.top}px" @submit=${this.commitForm}>
                        <textarea
                          aria-label="Field value"
                          autofocus
                          .value=${this.formDraft.text}
                          @input=${(event: Event) => {
                            if (!this.formDraft) return;
                            this.formDraft = { ...this.formDraft, text: (event.target as HTMLTextAreaElement).value };
                          }}
                        ></textarea>
                        <div class="note-actions">
                          <button type="button" @click=${() => this.cancelDraft()}>Cancel</button>
                          <button type="submit">Save</button>
                        </div>
                      </form>
                    `
                  : null}
                ${this.stampMenu && this.stampMenu.page === this.currentPage
                  ? html`
                      <div class="stamp-menu" style="left:${this.stampMenu.left}px;top:${this.stampMenu.top}px" role="menu">
                        ${STAMPS.map(
                          (word) => html`
                            <button type="button" role="menuitem" @click=${() => this.placeStamp(word)}>${word}</button>
                          `,
                        )}
                      </div>
                    `
                  : null}
                            `
                            : html`
                                <button
                                  class="sheet-hit"
                                  type="button"
                                  aria-label=${`Show page ${page}`}
                                  @click=${(event: Event) => {
                                    if (this.tool === 'hand') {
                                      event.preventDefault();
                                      return;
                                    }
                                    this.goToPage(page);
                                  }}
                                ></button>
                              `}
                          ${this.pageBookmarkButton(page)}
                        </div>
                      `,
                    )}
                  </div>
                `,
              )}
            </div>
            ${this.pageReady
              ? null
              : html`
                  <div class="status ${this.errorMessage ? 'error' : ''}">
                    ${this.errorMessage || (this.loading ? 'Opening PDF…' : 'Open a PDF, Word document, presentation, spreadsheet, or image to start viewing.')}
                  </div>
                `}
            ${this.dragOver ? html`<div class="drop-hint">Drop a PDF, Word document, presentation, spreadsheet, or image</div>` : null}
          </div>
          ${this.commentsOpen
            ? html`
                <aside class="comments">
                  <label class="comment-search">
                    ${iconSearch}
                    <input
                      type="search"
                      placeholder="Search"
                      aria-label="Search comments"
                      .value=${this.commentQuery}
                      @input=${(event: Event) => {
                        this.commentQuery = (event.target as HTMLInputElement).value;
                      }}
                    />
                  </label>
                  <h2>Comments (${this.marks.length})</h2>
                  <div class="sort-row">
                    <span>Sort:</span>
                    <select aria-label="Sort comments" .value=${this.sortBy} @change=${this.onSortChange}>
                      <option value="position">Position</option>
                      <option value="created">Created Date</option>
                      <option value="modified">Modified Date</option>
                      <option value="status">Status</option>
                      <option value="author">Author</option>
                      <option value="type">Type</option>
                      <option value="color">Color</option>
                    </select>
                    <div class="sort-actions">
                      <button
                        class="icon"
                        type="button"
                        aria-label="Select comments"
                        aria-pressed=${this.selectingComments}
                        @click=${this.toggleSelecting}
                      >
                        ${iconMultiSelect}
                      </button>
                      <button class="icon" type="button" aria-label="Comment list options">${iconOpen}</button>
                      <button
                        class="icon"
                        type="button"
                        aria-label="Filter comments"
                        aria-pressed=${this.filterCount(this.appliedFilter) > 0}
                        @click=${this.openFilter}
                      >
                        ${iconFilter}
                      </button>
                    </div>
                  </div>
                  ${this.selectingComments && this.marks.length > 0
                    ? html`
                        <div class="select-bar">
                          <label>
                            <input
                              class="select-all"
                              type="checkbox"
                              .checked=${this.listedMarks.length > 0 && this.listedMarks.every((mark) => this.checkedCommentIds.includes(mark.id))}
                              aria-label="Select all comments"
                              @change=${this.toggleCheckAll}
                            />
                            All
                          </label>
                          <span>${this.checkedCommentIds.length} selected</span>
                          <button type="button" ?disabled=${this.checkedCommentIds.length === 0} @click=${this.deleteChecked}>Delete</button>
                        </div>
                      `
                    : nothing}
                  ${this.marks.length === 0
                    ? html`
                        <div class="empty">
                          ${iconEmptyDoc}
                          <span>Start making annotations to leave a comment.</span>
                        </div>
                      `
                    : html`
                        <div class="comment-list" @scroll=${this.closeCommentMenus}>
                          ${this.listedMarks.length === 0
                            ? html`<div class="empty"><span>${this.commentQuery.trim() ? 'No comments match that search.' : 'No comments match the filter.'}</span></div>`
                            : this.listedMarks.map((mark) => this.commentCard(mark))}
                        </div>
                      `}
                  ${this.statusMenu()}
                  ${this.actionMenu()}
                </aside>
              `
            : null}
        </div>
        <input class="file-input" type="file" accept=${OPEN_ACCEPT} @change=${this.onFileChosen} />
        <input
          class="image-input file-input"
          type="file"
          accept="image/png,image/jpeg,image/gif,image/webp"
          @change=${this.onImageChosen}
        />
        ${this.linkDialog()}
        ${this.linkConfirmDialog()}
        ${this.settingsDialog()}
        ${this.filterDialog()}
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'pdf-editor': PdfEditor;
  }
}
