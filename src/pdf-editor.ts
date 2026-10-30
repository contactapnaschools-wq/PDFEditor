import { LitElement, css, html, nothing, svg, type PropertyValues, type SVGTemplateResult } from 'lit';
import { customElement, property, query, state } from 'lit/decorators.js';
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist';
import {
  iconChevron,
  iconComments,
  iconEmptyDoc,
  iconFilter,
  iconHand,
  iconMenu,
  iconMinus,
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
  pageTextLines,
  pageFormFields,
  pageWordRects,
  paintPage,
  pdfRectFromPoints,
  projectRect,
  rectsIntersect,
  replaceLineText,
  setTextField,
  toggleCheckBox,
  viewportToPdf,
  type DrawnMark,
  type PdfRect,
  type TextMatch,
  type TextRun,
  type ViewportBox,
} from './pdf-engine';

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

interface PdfPoint {
  x: number;
  y: number;
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
}

interface MarkBox extends ViewportBox {
  id: string;
  kind: MarkKind;
  selected: boolean;
  label: string;
  text: string;
  points: Array<{ x: number; y: number }>;
  src: string;
}

interface TextEdit {
  page: number;
  rect: PdfRect;
  text: string;
  baseline: number;
  fontSize: number;
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
}

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
    .field-target:hover, .field-target:focus-visible {
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
    .mark.highlight { background: rgba(255, 214, 0, 0.45); }
    .mark.underline { border-bottom: 2px solid #e37400; }
    .mark.strike::after {
      content: "";
      position: absolute;
      left: 0;
      right: 0;
      top: 50%;
      border-top: 2px solid #d93025;
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
    .comment-text { font-size: 13px; line-height: 1.35; }
    .comment-delete {
      margin: 8px 8px 0 0;
      height: 28px;
      padding: 0 8px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      background: #fff;
      color: #b3261e;
      font-size: 12px;
      cursor: pointer;
    }

    .pages {
      width: 168px;
      flex: none;
      min-height: 0;
      overflow: auto;
      background: #f8f9fa;
      border-right: 1px solid #e0e3e7;
      padding: 12px 10px 24px;
    }
    .pages-empty {
      margin: 24px 8px;
      color: #80868b;
      font-size: 12px;
      text-align: center;
    }
    .thumb {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 6px;
      width: 100%;
      margin: 0 0 14px;
      padding: 0;
      border: none;
      background: transparent;
      color: #5f6368;
      font-size: 12px;
      cursor: pointer;
    }
    .thumb-frame {
      width: 120px;
      background: #fff;
      line-height: 0;
      box-shadow: 0 0 0 1px rgba(60, 64, 67, 0.18);
    }
    .thumb[aria-current='page'] .thumb-frame {
      box-shadow: 0 0 0 2px #1a73e8;
    }
    .thumb canvas { display: block; max-width: 120px; height: auto; }

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
  `;

  /** URL of a PDF to open. Relative URLs resolve against the host page. */
  @property({ type: String })
  src = '';

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
  @state() private menuOpen = false;
  @state() private zoomOpen = false;
  @state() private searchOpen = false;
  @state() private commentsOpen = true;
  @state() private pagesPressed = false;
  @state() private dragOver = false;
  @state() private panning = false;
  @state() private findQuery = '';
  @state() private commentQuery = '';
  @state() private sortBy = 'position';
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

  @query('.viewer') private viewerEl!: HTMLElement;
  @query('.page-canvas') private canvasEl!: HTMLCanvasElement;
  @query('.file-input') private fileInput!: HTMLInputElement;
  @query('.image-input') private imageInput!: HTMLInputElement;
  @query('.menu-root') private menuRoot!: HTMLElement;
  @query('.zoom-root') private zoomRoot!: HTMLElement;
  @query('.search-root') private searchRoot!: HTMLElement;
  @query('.pages') private pagesEl!: HTMLElement;

  private pdf: PDFDocumentProxy | null = null;
  private renderTask: RenderTask | null = null;
  private renderToken = 0;
  private loadId = 0;
  private dragDepth = 0;
  private resetScroll = false;
  private resizeObserver: ResizeObserver | null = null;
  private textRunCache = new Map<number, TextRun[]>();
  private markSeq = 0;
  private dragStart: { x: number; y: number } | null = null;
  private imagePoint: { page: number; x: number; y: number } | null = null;
  private pdfBytes: Uint8Array | null = null;
  private sourceUrl = '';
  private downloadName = 'document.pdf';
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
  }

  override disconnectedCallback(): void {
    this.removeEventListener('keydown', this.onKeydown);
    window.removeEventListener('pointerdown', this.onWindowPointerDown);
    this.resizeObserver?.disconnect();
    window.clearTimeout(this.searchTimer);
    this.cancelThumbnails();
    this.renderToken += 1;
    this.renderTask?.cancel();
    void this.pdf?.loadingTask.destroy();
    this.pdf = null;
    super.disconnectedCallback();
  }

  protected override firstUpdated(): void {
    this.resizeObserver = new ResizeObserver(() => {
      if (!this.pdf || this.zoomMode === 'percent') return;
      this.scheduleRender();
    });
    this.resizeObserver.observe(this.viewerEl);
  }

  protected override updated(changed: PropertyValues): void {
    if (changed.has('src') && this.src) {
      void this.load(this.src);
    }
    if (changed.has('pagesPressed') || changed.has('totalPages')) {
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
        const field = this.renderRoot.querySelector('.note-box textarea');
        if (field instanceof HTMLTextAreaElement) field.focus();
      });
    }
    if (changed.has('formDraft') && this.formDraft && !changed.get('formDraft')) {
      void this.updateComplete.then(() => {
        const field = this.renderRoot.querySelector('.form-box textarea');
        if (field instanceof HTMLTextAreaElement) field.focus();
      });
    }
  }

  /** Open a PDF from a URL, a File, or binary data. */
  async load(source: PdfSource): Promise<void> {
    const id = ++this.loadId;
    this.loading = true;
    this.errorMessage = '';
    this.pageReady = false;
    this.clearSearch();
    this.marks = [];
    this.selectedMarkId = null;
    this.markBoxes = [];
    this.noteDraft = null;
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

  private async save(): Promise<void> {
    if (!this.pdf || this.saving) return;
    this.saving = true;
    try {
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
        });
      }
      const bytes = await applyMarks(await this.ensureBytes(), drawn);
      this.downloadFile(bytes, this.downloadName);
    } catch (error) {
      this.reportError('Could not save this PDF.', error);
    } finally {
      this.saving = false;
    }
  }

  private downloadFile(bytes: Uint8Array, name: string): void {
    const copy = new Uint8Array(bytes);
    const url = URL.createObjectURL(new Blob([copy.buffer], { type: 'application/pdf' }));
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
    this.thumbsPainted.clear();
    this.thumbsPending.clear();
    this.scheduleRender();
    if (this.findQuery.trim()) this.queueSearch();
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
      const size = await pageSizeAtScale1(pdf, this.currentPage);
      if (token !== this.renderToken) return;
      const scale = this.scaleFor(size.width, size.height, viewer);
      this.scale = scale;
      const task = await paintPage(pdf, this.currentPage, canvas, scale);
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
    const availableWidth = Math.max(viewer.clientWidth - 72, 120);
    const availableHeight = Math.max(viewer.clientHeight - 68, 120);
    if (this.zoomMode === 'fit-width') return availableWidth / width;
    return Math.min(availableWidth / width, availableHeight / height);
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
    if (next === this.currentPage) return;
    this.highlights = [];
    this.noteDraft = null;
    this.stampMenu = null;
    this.imagePoint = null;
    this.dragStart = null;
    this.draftBox = null;
    this.draftStroke = [];
    this.draftArrow = 'none';
    this.polygonHover = null;
    this.currentPage = next;
    this.pageReady = false;
    this.resetScroll = true;
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
    const viewport = page.getViewport({ scale: this.scale });
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
    const root = this.pagesEl;
    if (!this.pagesPressed || !root || !this.pdf) return;
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
      const size = await pageSizeAtScale1(pdf, pageNumber);
      if (token !== this.thumbToken || this.pdf !== pdf) return;
      const task = await paintPage(pdf, pageNumber, canvas, 120 / size.width);
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
    this.pagesEl?.querySelector('.thumb[aria-current="page"]')?.scrollIntoView({ block: 'nearest' });
  }

  private chooseFile(): void {
    this.menuOpen = false;
    this.fileInput.click();
  }

  private onFileChosen(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) void this.load(file);
  }

  private onKeydown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      if (this.noteDraft || this.editDraft || this.formDraft || this.draftStroke.length > 0 || this.stampMenu) {
        this.cancelDraft();
        event.preventDefault();
        return;
      }
      this.menuOpen = false;
      this.zoomOpen = false;
      this.searchOpen = false;
      return;
    }
    const target = event.composedPath()[0];
    if (target instanceof HTMLElement && target.closest('input, textarea, select')) return;
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
    if (this.menuOpen && this.menuRoot && !path.includes(this.menuRoot)) this.menuOpen = false;
    if (this.zoomOpen && this.zoomRoot && !path.includes(this.zoomRoot)) this.zoomOpen = false;
    if (this.searchOpen && this.searchRoot && !path.includes(this.searchRoot)) this.searchOpen = false;
  };

  private onPointerDown(event: PointerEvent): void {
    if (event.button !== 0) return;
    if (this.eventIn(event, '.note-box') || this.eventIn(event, '.stamp-menu') || this.eventIn(event, '.text-line') || this.eventIn(event, '.field-target')) return;
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
    if (this.canDraw && (kind === 'note' || kind === 'freetext' || kind === 'text' || kind === 'image' || kind === 'stamp' || kind === 'initials' || kind === 'date')) {
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
    const markId = this.markIdFromEvent(event);
    this.selectedMarkId = markId;
    void this.refreshMarkBoxes();
  }

  private onPointerMove(event: PointerEvent): void {
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
      this.drawKind !== 'note' &&
      this.drawKind !== 'freetext' &&
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
    if (this.dragStart && this.canDraw && this.drawKind) {
      const start = this.dragStart;
      const point = this.canvasPoint(event) ?? start;
      const kind = this.drawKind;
      this.dragStart = null;
      const distance = Math.hypot(point.x - start.x, point.y - start.y);
      if (kind === 'note' || kind === 'freetext' || kind === 'text' || kind === 'initials') {
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
    const viewport = page.getViewport({ scale: this.scale });
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
    const viewport = page.getViewport({ scale });
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
    const from = viewportToPdf(page.getViewport({ scale: this.scale }), start.x, start.y);
    const to = viewportToPdf(page.getViewport({ scale: this.scale }), end.x, end.y);
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
    const viewport = page.getViewport({ scale: this.scale });
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

  private openText(kind: 'note' | 'freetext' | 'text' | 'initials', point: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber) return;
      if (kind === 'text' && this.insertKind !== 'text') return;
      if (kind === 'initials' && this.signKind !== 'initials') return;
      const pdfPoint = viewportToPdf(page.getViewport({ scale: this.scale }), point.x, point.y);
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
      };
    });
  }

  private openCallout(start: { x: number; y: number }, end: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber) return;
      const viewport = page.getViewport({ scale: this.scale });
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
      };
    });
  }

  private async createPen(kind: 'pen' | 'signature', points: Array<{ x: number; y: number }>): Promise<void> {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf || points.length < 2) return;
    const page = await pdf.getPage(pageNumber);
    if (this.pdf !== pdf) return;
    const viewport = page.getViewport({ scale: this.scale });
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
      const pdfPoint = viewportToPdf(page.getViewport({ scale: this.scale }), point.x, point.y);
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
    const viewport = page.getViewport({ scale: this.scale });
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
    const viewport = page.getViewport({ scale: this.scale });
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
    const mark =
      draft.kind === 'callout' && draft.tip
        ? this.makeMark('callout', draft.page, [anchor], text, [draft.tip, { x: draft.pdfX, y: draft.pdfY }])
        : draft.kind === 'link' && draft.rect
          ? this.makeMark('link', draft.page, [draft.rect], text)
          : draft.kind === 'text'
            ? this.makeMark('text', draft.page, [anchor], text)
            : draft.kind === 'initials'
              ? this.makeMark('initials', draft.page, [anchor], text)
              : this.makeMark(draft.kind === 'freetext' ? 'freetext' : 'note', draft.page, [anchor], text);
    this.cancelDraft();
    this.marks = [...this.marks, mark];
    this.selectedMarkId = mark.id;
    void this.refreshMarkBoxes();
  }

  private cancelDraft(): void {
    this.noteDraft = null;
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
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber || this.insertKind !== 'link') return;
      const viewport = page.getViewport({ scale: this.scale });
      const startPdf = viewportToPdf(viewport, start.x, start.y);
      const endPdf = viewportToPdf(viewport, end.x, end.y);
      const rect = pdfRectFromPoints(startPdf.x, startPdf.y, endPdf.x, endPdf.y);
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
      };
    });
  }

  private openStampMenu(point: { x: number; y: number }): void {
    const pdf = this.pdf;
    const pageNumber = this.currentPage;
    if (!pdf) return;
    void pdf.getPage(pageNumber).then((page) => {
      if (this.pdf !== pdf || this.currentPage !== pageNumber || this.insertKind !== 'stamp') return;
      const pdfPoint = viewportToPdf(page.getViewport({ scale: this.scale }), point.x, point.y);
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
      const pdfPoint = viewportToPdf(page.getViewport({ scale: this.scale }), point.x, point.y);
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
    const viewport = page.getViewport({ scale: this.scale });
    const noteNumbers = new Map<string, number>();
    let noteNumber = 0;
    for (const mark of this.marks) {
      if (mark.kind === 'note') noteNumbers.set(mark.id, ++noteNumber);
    }
    const boxes: MarkBox[] = [];
    for (const mark of this.marks) {
      if (mark.page !== pageNumber) continue;
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
        });
      }
    }
    this.markBoxes = boxes;
  }

  private focusMark(id: string): void {
    const mark = this.marks.find((item) => item.id === id);
    if (!mark) return;
    this.selectedMarkId = id;
    if (mark.page !== this.currentPage) this.goToPage(mark.page);
    else void this.refreshMarkBoxes();
  }

  private deleteMark(id: string): void {
    this.marks = this.marks.filter((mark) => mark.id !== id);
    if (this.selectedMarkId === id) this.selectedMarkId = null;
    void this.refreshMarkBoxes();
  }

  private get listedMarks(): Mark[] {
    const query = this.commentQuery.trim().toLowerCase();
    const sorted = [...this.marks].sort((left, right) => left.page - right.page || right.sortY - left.sortY);
    if (!query) return sorted;
    return sorted.filter((mark) => mark.text.toLowerCase().includes(query));
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
    return 'Note text';
  }

  private markFace(box: MarkBox) {
    if (box.kind === 'pen' || box.kind === 'signature' || box.kind === 'line' || box.kind === 'arrow' || box.kind === 'rect' || box.kind === 'ellipse' || box.kind === 'polygon') {
      return nothing;
    }
    if (box.kind === 'note') {
      return html`
        <button
          class="note-pin ${box.selected ? 'selected' : ''}"
          type="button"
          data-mark=${box.id}
          aria-label=${`Note ${box.label}`}
          style="left:${box.left}px;top:${box.top - 22}px"
        >
          ${box.label}
        </button>
      `;
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
    return html`
      <div
        class="mark ${box.kind} ${box.selected ? 'selected' : ''}"
        data-mark=${box.id}
        style="left:${box.left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px"
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

  private get pagesPanel() {
    if (!this.pagesPressed) return null;
    const pages = Array.from({ length: this.totalPages }, (_, index) => index + 1);
    return html`
      <aside class="pages" aria-label="Pages">
        ${pages.length === 0 ? html`<p class="pages-empty">Open a PDF to see pages.</p>` : null}
        ${pages.map(
          (page) => html`
            <button
              class="thumb"
              type="button"
              data-page=${page}
              aria-label=${`Page ${page}`}
              aria-current=${page === this.currentPage ? 'page' : nothing}
              @click=${() => this.goToPage(page)}
            >
              <span class="thumb-frame"><canvas></canvas></span>
              <span>${page}</span>
            </button>
          `,
        )}
      </aside>
    `;
  }

  private zoomChecked(step: number): boolean {
    return this.zoomMode === 'percent' && Math.abs(this.zoomPercent - step) < 0.001;
  }

  override render() {
    const pageLabel = this.totalPages > 0 ? `Page ${this.currentPage} of ${this.totalPages}` : 'No document';
    return html`
      <div
        class="root"
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
                }}
              >
                ${iconMenu}
              </button>
              ${this.menuOpen
                ? html`
                    <div class="popup" role="menu">
                      <button type="button" role="menuitem" @click=${() => this.chooseFile()}>Open…</button>
                      <div class="rule"></div>
                      <button
                        type="button"
                        role="menuitem"
                        ?disabled=${!this.pdf || this.currentPage <= 1}
                        @click=${() => this.goToPage(this.currentPage - 1)}
                      >
                        Previous page
                      </button>
                      <button
                        type="button"
                        role="menuitem"
                        ?disabled=${!this.pdf || this.currentPage >= this.totalPages}
                        @click=${() => this.goToPage(this.currentPage + 1)}
                      >
                        Next page
                      </button>
                      <div class="rule"></div>
                      <div class="menu-label">${pageLabel}</div>
                    </div>
                  `
                : null}
            </div>
            <button
              class="icon"
              type="button"
              aria-label="Page panel"
              aria-pressed=${this.pagesPressed}
              @click=${() => {
                this.pagesPressed = !this.pagesPressed;
              }}
            >
              ${iconPages}
            </button>
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
          ${this.pagesPanel}
          <div
            class="viewer ${this.tool === 'hand' ? 'hand' : ''} ${this.panning ? 'panning' : ''} ${this.canDraw ? 'markup' : ''}"
            tabindex="0"
            @pointerdown=${this.onPointerDown}
            @pointermove=${this.onPointerMove}
            @pointerup=${this.onPointerUp}
            @pointercancel=${this.onPointerCancel}
          >
            <div class="stage">
              <div class="page-frame ${this.pageReady ? '' : 'is-hidden'}">
                <canvas class="page-canvas"></canvas>
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
                              class="field-target"
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
                ${this.noteDraft && this.noteDraft.page === this.currentPage
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
              </div>
            </div>
            ${this.pageReady
              ? null
              : html`
                  <div class="status ${this.errorMessage ? 'error' : ''}">
                    ${this.errorMessage || (this.loading ? 'Opening PDF…' : 'Open a PDF to start viewing.')}
                  </div>
                `}
            ${this.dragOver ? html`<div class="drop-hint">Drop a PDF to open it</div>` : null}
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
                    <select
                      aria-label="Sort comments"
                      .value=${this.sortBy}
                      @change=${(event: Event) => {
                        this.sortBy = (event.target as HTMLSelectElement).value;
                      }}
                    >
                      <option value="position">Position</option>
                    </select>
                    <div class="sort-actions">
                      <button class="icon" type="button" aria-label="Comment list options">${iconOpen}</button>
                      <button class="icon" type="button" aria-label="Filter comments">${iconFilter}</button>
                    </div>
                  </div>
                  ${this.marks.length === 0
                    ? html`
                        <div class="empty">
                          ${iconEmptyDoc}
                          <span>Start making annotations to leave a comment.</span>
                        </div>
                      `
                    : html`
                        <div class="comment-list">
                          ${this.listedMarks.length === 0
                            ? html`<div class="empty"><span>No comments match that search.</span></div>`
                            : this.listedMarks.map(
                                (mark) => html`
                                  <div class="comment-row ${mark.id === this.selectedMarkId ? 'selected' : ''}">
                                    <button class="comment" type="button" @click=${() => this.focusMark(mark.id)}>
                                      <span class="comment-kind">${this.markLabel(mark.kind)}</span>
                                      <span class="comment-page">Page ${mark.page}</span>
                                      <span class="comment-text">${mark.text}</span>
                                    </button>
                                    ${mark.id === this.selectedMarkId
                                      ? html`
                                          <button
                                            class="comment-delete"
                                            type="button"
                                            aria-label="Delete comment"
                                            @click=${() => this.deleteMark(mark.id)}
                                          >
                                            Delete
                                          </button>
                                        `
                                      : null}
                                  </div>
                                `,
                              )}
                        </div>
                      `}
                </aside>
              `
            : null}
        </div>
        <input class="file-input" type="file" accept="application/pdf,.pdf" @change=${this.onFileChosen} />
        <input
          class="image-input file-input"
          type="file"
          accept="image/png,image/jpeg,image/gif,image/webp"
          @change=${this.onImageChosen}
        />
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    'pdf-editor': PdfEditor;
  }
}
