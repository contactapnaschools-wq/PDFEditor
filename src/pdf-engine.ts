import { PDFDocument, StandardFonts, degrees, rgb, type PDFFont, type PDFPage, type RGB } from 'pdf-lib';
import {
  getDocument,
  GlobalWorkerOptions,
  RenderingCancelledException,
  type PDFDocumentProxy,
  type PageViewport,
  type RenderTask,
} from 'pdfjs-dist';
import { pdfjsAsset } from './pdfjs-assets';

let workerReady = false;

function ensureWorker(): void {
  if (workerReady) return;
  GlobalWorkerOptions.workerSrc = pdfjsAsset('pdf.worker.min.mjs');
  workerReady = true;
}

function documentOptions() {
  return {
    cMapUrl: pdfjsAsset('cmaps/'),
    cMapPacked: true,
    standardFontDataUrl: pdfjsAsset('standard_fonts/'),
    wasmUrl: pdfjsAsset('wasm/'),
  };
}

export function isRenderCancel(error: unknown): boolean {
  return (
    error instanceof RenderingCancelledException ||
    (error instanceof Error && error.name === 'RenderingCancelledException')
  );
}

export async function openDocument(source: Uint8Array | string): Promise<PDFDocumentProxy> {
  ensureWorker();
  const shared = documentOptions();
  const task =
    typeof source === 'string'
      ? getDocument({ url: source, ...shared })
      : getDocument({ data: source, ...shared });
  return task.promise;
}

export interface OutlineEntry {
  id: string;
  title: string;
  page: number | null;
  url: string | null;
  items: OutlineEntry[];
}

interface RawOutline {
  title?: string;
  dest?: string | unknown[] | null;
  url?: string | null;
  items?: RawOutline[];
}

/** The document's built-in contents, with each entry resolved to a page when it has one. */
export async function readOutline(pdf: PDFDocumentProxy): Promise<OutlineEntry[]> {
  const root = await pdf.getOutline();
  if (!root?.length) return [];
  const walk = async (nodes: RawOutline[], prefix: string): Promise<OutlineEntry[]> => {
    const entries: OutlineEntry[] = [];
    for (let index = 0; index < nodes.length; index += 1) {
      const node = nodes[index];
      const id = `${prefix}${index}`;
      entries.push({
        id,
        title: node.title?.trim() || 'Untitled',
        page: await outlinePage(pdf, node.dest),
        url: typeof node.url === 'string' ? node.url : null,
        items: await walk(node.items ?? [], `${id}-`),
      });
    }
    return entries;
  };
  return walk(root as RawOutline[], '');
}

async function outlinePage(pdf: PDFDocumentProxy, dest: string | unknown[] | null | undefined): Promise<number | null> {
  try {
    const resolved = typeof dest === 'string' ? await pdf.getDestination(dest) : dest;
    if (!Array.isArray(resolved) || resolved[0] == null) return null;
    const index = await pdf.getPageIndex(resolved[0] as never);
    return Number.isFinite(index) ? index + 1 : null;
  } catch {
    return null;
  }
}

function turnedRotation(pageRotate: number, quarterTurns: number): number {
  return (((pageRotate || 0) + quarterTurns * 90) % 360 + 360) % 360;
}

type LayerConfig = Awaited<ReturnType<PDFDocumentProxy['getOptionalContentConfig']>>;

export async function paintPage(
  pdf: PDFDocumentProxy,
  pageNumber: number,
  canvas: HTMLCanvasElement,
  cssScale: number,
  quarterTurns = 0,
  optionalContent?: Promise<LayerConfig> | null,
): Promise<RenderTask> {
  const page = await pdf.getPage(pageNumber);
  const viewport = page.getViewport({ scale: cssScale, rotation: turnedRotation(page.rotate, quarterTurns) });
  const pixelRatio = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.floor(viewport.width * pixelRatio));
  canvas.height = Math.max(1, Math.floor(viewport.height * pixelRatio));
  canvas.style.width = `${Math.floor(viewport.width)}px`;
  canvas.style.height = `${Math.floor(viewport.height)}px`;
  return page.render({
    canvas,
    viewport,
    transform: pixelRatio === 1 ? undefined : [pixelRatio, 0, 0, pixelRatio, 0, 0],
    optionalContentConfigPromise: optionalContent ?? undefined,
  });
}

export async function pageSizeAtScale1(
  pdf: PDFDocumentProxy,
  pageNumber: number,
  quarterTurns = 0,
): Promise<{ width: number; height: number }> {
  const page = await pdf.getPage(pageNumber);
  const viewport = page.getViewport({ scale: 1, rotation: turnedRotation(page.rotate, quarterTurns) });
  return { width: viewport.width, height: viewport.height };
}

/** A rectangle in PDF user space, origin at the bottom left. */
export interface PdfRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** One find result. A phrase that crosses text items has several rectangles. */
export interface TextMatch {
  page: number;
  rects: PdfRect[];
}

export interface ViewportBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

export function projectRect(viewport: PageViewport, rect: PdfRect): ViewportBox {
  const [x1, y1] = viewport.convertToViewportPoint(rect.x, rect.y);
  const [x2, y2] = viewport.convertToViewportPoint(rect.x + rect.width, rect.y + rect.height);
  return {
    left: Math.min(x1, x2),
    top: Math.min(y1, y2),
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1),
  };
}

interface PdfTextItem {
  str: string;
  transform: number[];
  width: number;
  height: number;
  hasEOL: boolean;
}

function isTextItem(item: unknown): item is PdfTextItem {
  return (
    typeof item === 'object' &&
    item !== null &&
    'str' in item &&
    typeof (item as PdfTextItem).str === 'string' &&
    'transform' in item &&
    'width' in item
  );
}

function joinBetween(previous: PdfTextItem, next: PdfTextItem): string {
  if (previous.hasEOL) return '\n';
  const lineShift = Math.abs(previous.transform[5] - next.transform[5]);
  if (lineShift > Math.max(previous.height, next.height) * 0.5) return '\n';
  if (previous.str.endsWith(' ') || previous.str.endsWith('\n') || next.str.startsWith(' ')) return '';
  return ' ';
}

function rectForSlice(item: PdfTextItem, start: number, end: number): PdfRect {
  const length = Math.max(item.str.toLowerCase().length, 1);
  const from = start / length;
  const to = end / length;
  const horizontal = Math.abs(item.transform[1]) < 0.01 && Math.abs(item.transform[2]) < 0.01;
  if (!horizontal) {
    return {
      x: item.transform[4],
      y: item.transform[5] - item.height * 0.25,
      width: item.width,
      height: item.height * 1.1,
    };
  }
  return {
    x: item.transform[4] + item.width * from,
    y: item.transform[5] - item.height * 0.25,
    width: Math.max(item.width * (to - from), 0.5),
    height: item.height * 1.1,
  };
}

/** Case-insensitive search. Each match keeps the PDF-space boxes to highlight. */
export async function findText(pdf: PDFDocumentProxy, query: string): Promise<TextMatch[]> {
  const needle = query.trim().toLowerCase();
  if (!needle) return [];
  const matches: TextMatch[] = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const items: PdfTextItem[] = [];
    for (const item of content.items) {
      if (isTextItem(item) && item.str.length > 0) items.push(item);
    }
    let text = '';
    const spans: Array<{ start: number; end: number; item: PdfTextItem }> = [];
    items.forEach((item, index) => {
      if (index > 0) text += joinBetween(items[index - 1], item);
      const start = text.length;
      text += item.str.toLowerCase();
      spans.push({ start, end: text.length, item });
    });

    let from = 0;
    while (from <= text.length - needle.length) {
      const at = text.indexOf(needle, from);
      if (at < 0) break;
      const end = at + needle.length;
      const rects: PdfRect[] = [];
      for (const span of spans) {
        const overlapStart = Math.max(at, span.start);
        const overlapEnd = Math.min(end, span.end);
        if (overlapStart >= overlapEnd) continue;
        rects.push(rectForSlice(span.item, overlapStart - span.start, overlapEnd - span.start));
      }
      if (rects.length > 0) matches.push({ page: pageNumber, rects });
      from = end;
    }
  }

  return matches;
}

/** One word on a page, with its box in PDF user space. */
export interface TextRun {
  text: string;
  rect: PdfRect;
}

export async function pageWordRects(pdf: PDFDocumentProxy, pageNumber: number): Promise<TextRun[]> {
  const page = await pdf.getPage(pageNumber);
  const content = await page.getTextContent();
  const runs: TextRun[] = [];
  for (const item of content.items) {
    if (!isTextItem(item) || item.str.length === 0) continue;
    const expression = /\S+/g;
    let match: RegExpExecArray | null;
    while ((match = expression.exec(item.str))) {
      runs.push({
        text: match[0],
        rect: rectForSlice(item, match.index, match.index + match[0].length),
      });
    }
  }
  return runs;
}

export function viewportToPdf(viewport: PageViewport, x: number, y: number): { x: number; y: number } {
  const [px, py] = viewport.convertToPdfPoint(x, y);
  return { x: Number(px), y: Number(py) };
}

export function pdfRectFromPoints(x1: number, y1: number, x2: number, y2: number): PdfRect {
  return {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    width: Math.max(Math.abs(x2 - x1), 0.5),
    height: Math.max(Math.abs(y2 - y1), 0.5),
  };
}

export function rectsIntersect(a: PdfRect, b: PdfRect): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
}

/** One horizontal line of PDF text, in PDF user space. */
export interface TextLine {
  text: string;
  rect: PdfRect;
  baseline: number;
  fontSize: number;
}

function isHorizontal(item: PdfTextItem): boolean {
  return Math.abs(item.transform[1]) < 0.01 && Math.abs(item.transform[2]) < 0.01;
}

/** Group horizontal text items that share a baseline into lines. */
export async function pageTextLines(pdf: PDFDocumentProxy, pageNumber: number): Promise<TextLine[]> {
  const page = await pdf.getPage(pageNumber);
  const content = await page.getTextContent();
  const items: PdfTextItem[] = [];
  for (const item of content.items) {
    if (isTextItem(item) && item.str.length > 0 && isHorizontal(item)) items.push(item);
  }
  const groups: PdfTextItem[][] = [];
  for (const item of items) {
    const group = groups.find((entry) => {
      const sample = entry[0];
      return Math.abs(sample.transform[5] - item.transform[5]) <= Math.max(sample.height, item.height) * 0.5;
    });
    if (group) group.push(item);
    else groups.push([item]);
  }
  return groups.map((group) => {
    group.sort((left, right) => left.transform[4] - right.transform[4]);
    let text = '';
    const rects: PdfRect[] = [];
    group.forEach((item, index) => {
      if (index > 0) {
        const gap = joinBetween(group[index - 1], item);
        text += gap === '\n' ? ' ' : gap;
      }
      text += item.str;
      rects.push(rectForSlice(item, 0, item.str.length));
    });
    const left = Math.min(...rects.map((rect) => rect.x));
    const bottom = Math.min(...rects.map((rect) => rect.y));
    const right = Math.max(...rects.map((rect) => rect.x + rect.width));
    const top = Math.max(...rects.map((rect) => rect.y + rect.height));
    return {
      text,
      rect: { x: left, y: bottom, width: Math.max(right - left, 0.5), height: Math.max(top - bottom, 0.5) },
      baseline: group[0].transform[5],
      fontSize: Math.max(...group.map((item) => Math.abs(item.transform[3]) || item.height), 1),
    };
  });
}

export interface FormField {
  kind: 'text' | 'check';
  name: string;
  value: string;
  checked: boolean;
  rect: PdfRect;
}

interface WidgetData {
  fieldType?: string;
  fieldName?: string;
  fieldValue?: unknown;
  checkBox?: boolean;
  radioButton?: boolean;
  rect?: number[];
}

function widgetRect(rect: number[]): PdfRect | null {
  if (rect.length < 4) return null;
  const [x1, y1, x2, y2] = rect;
  const width = Math.abs(x2 - x1);
  const height = Math.abs(y2 - y1);
  if (width < 0.5 || height < 0.5) return null;
  return { x: Math.min(x1, x2), y: Math.min(y1, y2), width, height };
}

function checkboxOn(value: unknown): boolean {
  return typeof value === 'string' && value.length > 0 && value !== 'Off';
}

/** Text fields and checkboxes on one page, in PDF user space. */
export async function pageFormFields(pdf: PDFDocumentProxy, pageNumber: number): Promise<FormField[]> {
  const page = await pdf.getPage(pageNumber);
  const annotations: WidgetData[] = await page.getAnnotations();
  const fields: FormField[] = [];
  for (const item of annotations) {
    const rect = item.rect ? widgetRect(item.rect) : null;
    const name = item.fieldName ?? '';
    if (!rect || !name) continue;
    if (item.fieldType === 'Tx') {
      fields.push({
        kind: 'text',
        name,
        value: typeof item.fieldValue === 'string' ? item.fieldValue : '',
        checked: false,
        rect,
      });
    } else if (item.fieldType === 'Btn' && item.checkBox && !item.radioButton) {
      fields.push({ kind: 'check', name, value: '', checked: checkboxOn(item.fieldValue), rect });
    }
  }
  return fields;
}

function nextFieldName(names: Set<string>, prefix: string): string {
  let index = 1;
  while (names.has(`${prefix} ${index}`)) index += 1;
  return `${prefix} ${index}`;
}

/** Write a text field value, including an empty string, and refresh its appearance. */
export async function setTextField(bytes: Uint8Array, name: string, value: string): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const form = doc.getForm();
  form.getTextField(name).setText(value);
  form.updateFieldAppearances();
  return doc.save();
}

/** Flip a checkbox and refresh its appearance. */
export async function toggleCheckBox(bytes: Uint8Array, name: string): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const form = doc.getForm();
  const box = form.getCheckBox(name);
  if (box.isChecked()) box.uncheck();
  else box.check();
  form.updateFieldAppearances();
  return doc.save();
}

/** Add an empty text field with the next free name, such as `Text 1`. */
export async function addTextField(bytes: Uint8Array, pageNumber: number, rect: PdfRect): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const form = doc.getForm();
  const name = nextFieldName(new Set(form.getFields().map((field) => field.getName())), 'Text');
  const field = form.createTextField(name);
  field.addToPage(doc.getPages()[pageNumber - 1], { x: rect.x, y: rect.y, width: rect.width, height: rect.height });
  form.updateFieldAppearances();
  return doc.save();
}

/** Add an unchecked checkbox with the next free name, such as `Check 1`. */
export async function addCheckBox(bytes: Uint8Array, pageNumber: number, rect: PdfRect): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const form = doc.getForm();
  const name = nextFieldName(new Set(form.getFields().map((field) => field.getName())), 'Check');
  const field = form.createCheckBox(name);
  field.addToPage(doc.getPages()[pageNumber - 1], { x: rect.x, y: rect.y, width: rect.width, height: rect.height });
  form.updateFieldAppearances();
  return doc.save();
}

/** Cover one line with white and draw the replacement in Helvetica. Returns new PDF bytes. */
export async function replaceLineText(
  bytes: Uint8Array,
  pageNumber: number,
  rect: PdfRect,
  baseline: number,
  fontSize: number,
  text: string,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[pageNumber - 1];
  const pad = 1;
  page.drawRectangle({
    x: rect.x - pad,
    y: rect.y - pad,
    width: rect.width + pad * 2,
    height: rect.height + pad * 2,
    color: rgb(1, 1, 1),
  });
  const font = await doc.embedFont(StandardFonts.Helvetica);
  page.drawText(text, {
    x: rect.x,
    y: baseline,
    size: Math.max(fontSize, 1),
    font,
    color: rgb(0, 0, 0),
  });
  return doc.save();
}

export interface DrawnMark {
  kind: string;
  page: number;
  rects: PdfRect[];
  text: string;
  points: Array<{ x: number; y: number }>;
  image?: Uint8Array;
  imageType?: 'png' | 'jpg';
  file?: Uint8Array;
  fileName?: string;
  color?: string;
  opacity?: number;
  weight?: number;
}

export type StampShape = 'badge' | 'arrow' | 'check' | 'cross';

export interface StampStyle {
  id: string;
  label: string;
  shape: StampShape;
  ink: string;
  fill: string;
  text: string;
}

export const STANDARD_STAMPS: readonly StampStyle[] = [
  { id: 'approved', label: 'APPROVED', shape: 'badge', ink: '#188038', fill: '#e6f4ea', text: '#188038' },
  { id: 'as-is', label: 'AS IS', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'completed', label: 'COMPLETED', shape: 'badge', ink: '#188038', fill: '#e6f4ea', text: '#188038' },
  { id: 'confidential', label: 'CONFIDENTIAL', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'departmental', label: 'DEPARTMENTAL', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'draft', label: 'DRAFT', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'experimental', label: 'EXPERIMENTAL', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'expired', label: 'EXPIRED', shape: 'badge', ink: '#d93025', fill: '#fce8e6', text: '#d93025' },
  { id: 'final', label: 'FINAL', shape: 'badge', ink: '#188038', fill: '#e6f4ea', text: '#188038' },
  { id: 'for-comment', label: 'FOR COMMENT', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'for-public-release', label: 'FOR PUBLIC RELEASE', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'information-only', label: 'INFORMATION ONLY', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'not-approved', label: 'NOT APPROVED', shape: 'badge', ink: '#d93025', fill: '#fce8e6', text: '#d93025' },
  { id: 'not-for-public-release', label: 'NOT FOR PUBLIC RELEASE', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'preliminary-results', label: 'PRELIMINARY RESULTS', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'sold', label: 'SOLD', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'top-secret', label: 'TOP SECRET', shape: 'badge', ink: '#1a73e8', fill: '#e8f0fe', text: '#1a73e8' },
  { id: 'void', label: 'VOID', shape: 'badge', ink: '#d93025', fill: '#fce8e6', text: '#d93025' },
  { id: 'sign-here', label: 'SIGN HERE', shape: 'arrow', ink: '#ffffff', fill: '#e25c4c', text: '#ffffff' },
  { id: 'witness', label: 'WITNESS', shape: 'arrow', ink: '#202124', fill: '#f6d96a', text: '#202124' },
  { id: 'initial-here', label: 'INITIAL HERE', shape: 'arrow', ink: '#ffffff', fill: '#7e57c2', text: '#ffffff' },
  { id: 'check', label: 'Check', shape: 'check', ink: '#188038', fill: 'transparent', text: '#188038' },
  { id: 'cross', label: 'Cross', shape: 'cross', ink: '#d93025', fill: 'transparent', text: '#d93025' },
];

export function stampFill(ink: string): string {
  const match = /^#([0-9a-f]{6})$/i.exec(ink);
  if (!match) return '#e8f0fe';
  const value = Number.parseInt(match[1], 16);
  const mix = (channel: number) => Math.round(channel + (255 - channel) * 0.86);
  const channels = [mix((value >> 16) & 255), mix((value >> 8) & 255), mix(value & 255)];
  return `#${channels.map((channel) => channel.toString(16).padStart(2, '0')).join('')}`;
}

export function stampStyle(label: string, ink = '#1a73e8'): StampStyle {
  const found = STANDARD_STAMPS.find((item) => item.label.toUpperCase() === label.toUpperCase());
  if (found) return found;
  return { id: 'custom', label, shape: 'badge', ink, fill: stampFill(ink), text: ink };
}

export function stampSize(style: StampStyle): { width: number; height: number } {
  if (style.shape === 'check' || style.shape === 'cross') return { width: 28, height: 28 };
  const letters = Math.max(style.label.length, 2);
  if (style.shape === 'arrow') return { width: Math.max(112, letters * 6.4 + 36), height: 26 };
  return { width: Math.max(78, letters * 6.6 + 22), height: 24 };
}

function paintOf(hex: string | undefined, fallback: RGB): RGB {
  const match = /^#([0-9a-f]{6})$/i.exec(hex ?? '');
  if (!match) return fallback;
  const value = Number.parseInt(match[1], 16);
  return rgb(((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255);
}

const ink = rgb(0.125, 0.129, 0.141);
const blue = rgb(0.102, 0.451, 0.91);
const yellow = rgb(1, 0.839, 0);
const orange = rgb(0.89, 0.455, 0);
const red = rgb(0.851, 0.188, 0.145);
const pin = rgb(0.976, 0.671, 0);
const hairline = rgb(0.855, 0.863, 0.878);

/** Draw in-memory marks onto a copy of the working PDF. The input bytes are not changed. */
export async function applyMarks(bytes: Uint8Array, marks: DrawnMark[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const usedNames = new Set<string>();
  for (const mark of marks) {
    if (mark.kind === 'file' && mark.file) await doc.attach(mark.file, uniqueFileName(usedNames, mark.fileName || mark.text || 'attachment'));
    const page = doc.getPages()[mark.page - 1];
    if (!page) continue;
    await drawSavedMark(doc, page, font, bold, mark);
  }
  return doc.save();
}

function uniqueFileName(used: Set<string>, name: string): string {
  const clean = name.replace(/[\\/:*?"<>|]/g, ' ').trim() || 'attachment';
  if (!used.has(clean)) {
    used.add(clean);
    return clean;
  }
  let count = 2;
  let next = `${count} ${clean}`;
  while (used.has(next)) {
    count += 1;
    next = `${count} ${clean}`;
  }
  used.add(next);
  return next;
}

async function drawSavedMark(doc: PDFDocument, page: PDFPage, font: PDFFont, bold: PDFFont, mark: DrawnMark): Promise<void> {
  const rect = mark.rects[0];
  const stroke = 1.5;
  if (mark.kind === 'highlight') {
    const color = paintOf(mark.color, yellow);
    const opacity = mark.opacity ?? 0.45;
    for (const box of mark.rects) page.drawRectangle({ x: box.x, y: box.y, width: box.width, height: box.height, color, opacity });
    return;
  }
  if (mark.kind === 'underline') {
    const color = paintOf(mark.color, orange);
    const thickness = mark.weight ?? 2;
    const opacity = mark.opacity ?? 1;
    for (const box of mark.rects) page.drawLine({ start: { x: box.x, y: box.y }, end: { x: box.x + box.width, y: box.y }, thickness, color, opacity });
    return;
  }
  if (mark.kind === 'squiggly') {
    const color = paintOf(mark.color, red);
    const thickness = mark.weight ?? 1.25;
    const opacity = mark.opacity ?? 1;
    for (const box of mark.rects) drawSquiggle(page, box, color, thickness, opacity);
    return;
  }
  if (mark.kind === 'strike') {
    const color = paintOf(mark.color, red);
    const thickness = mark.weight ?? 2;
    const opacity = mark.opacity ?? 1;
    for (const box of mark.rects) {
      const y = box.y + box.height / 2;
      page.drawLine({ start: { x: box.x, y }, end: { x: box.x + box.width, y }, thickness, color, opacity });
    }
    return;
  }
  if (mark.kind === 'note') {
    for (const box of mark.rects) {
      page.drawRectangle({ x: box.x, y: box.y, width: box.width, height: box.height, color: pin, opacity: 0.35 });
    }
    return;
  }
  if (mark.kind === 'freehand-highlight') {
    drawPolyline(page, mark.points, paintOf(mark.color, yellow), mark.weight ?? 12, mark.opacity ?? 0.45);
    return;
  }
  if (mark.kind === 'arc' || mark.kind === 'polyline') {
    drawPolyline(page, mark.points, ink, stroke);
    return;
  }
  if (mark.kind === 'cloud' && mark.points.length >= 3) {
    const height = page.getHeight();
    const path = svgPath(cloudOutline(mark.points, 8), height, true);
    page.drawSvgPath(path, { x: 0, y: height, borderColor: ink, borderWidth: stroke });
    return;
  }
  if (mark.kind === 'pen' || mark.kind === 'signature' || mark.kind === 'line' || mark.kind === 'distance' || mark.kind === 'perimeter') {
    const color = mark.kind === 'signature' || mark.kind === 'distance' || mark.kind === 'perimeter' ? blue : ink;
    const thickness = mark.kind === 'signature' ? 2.25 : stroke;
    drawPolyline(page, mark.points, color, thickness);
    if (rect && (mark.kind === 'distance' || mark.kind === 'perimeter')) drawCenteredText(page, font, mark.text, rect.x, rect.y, 12, blue);
    return;
  }
  if (mark.kind === 'arrow' && mark.points.length >= 2) {
    drawArrow(page, mark.points[0], mark.points[mark.points.length - 1], ink, stroke);
    return;
  }
  if (mark.kind === 'callout' && mark.points.length >= 2) {
    drawArrow(page, mark.points[1], mark.points[0], ink, stroke);
    drawBoxedText(page, font, mark.text, mark.points[1].x, mark.points[1].y, 12, ink, hairline);
    return;
  }
  if (mark.kind === 'rect' && rect) {
    page.drawRectangle({ x: rect.x, y: rect.y, width: rect.width, height: rect.height, borderColor: ink, borderWidth: stroke });
    return;
  }
  if (mark.kind === 'ellipse' && rect) {
    page.drawEllipse({
      x: rect.x + rect.width / 2,
      y: rect.y + rect.height / 2,
      xScale: rect.width / 2,
      yScale: rect.height / 2,
      borderColor: ink,
      borderWidth: stroke,
    });
    return;
  }
  if ((mark.kind === 'polygon' || mark.kind === 'area') && mark.points.length >= 3) {
    const height = page.getHeight();
    const path = svgPath(mark.points, height, true);
    if (mark.kind === 'area') {
      page.drawSvgPath(path, { x: 0, y: height, color: blue, opacity: 0.12, borderColor: blue, borderWidth: stroke, borderOpacity: 1 });
      if (rect) drawCenteredText(page, font, mark.text, rect.x, rect.y, 12, blue);
    } else {
      page.drawSvgPath(path, { x: 0, y: height, borderColor: ink, borderWidth: stroke });
    }
    return;
  }
  if (mark.kind === 'link' && rect) {
    page.drawRectangle({ x: rect.x, y: rect.y, width: rect.width, height: rect.height, borderColor: blue, borderWidth: 2 });
    return;
  }
  if (mark.kind === 'image' && rect && mark.image && mark.imageType) {
    const image = mark.imageType === 'png' ? await doc.embedPng(mark.image) : await doc.embedJpg(mark.image);
    page.drawImage(image, { x: rect.x, y: rect.y, width: rect.width, height: rect.height });
    return;
  }
  if (mark.kind === 'signature' && rect && mark.image && mark.imageType) {
    const image = mark.imageType === 'png' ? await doc.embedPng(mark.image) : await doc.embedJpg(mark.image);
    page.drawImage(image, { x: rect.x, y: rect.y, width: rect.width, height: rect.height });
    return;
  }
  if (mark.kind === 'stamp' && rect) {
    drawStamp(page, bold, mark, page.getHeight());
    return;
  }
  if (mark.kind === 'file' && rect) {
    drawPaperclip(page, rect, page.getHeight());
    return;
  }
  if (rect && (mark.kind === 'freetext' || mark.kind === 'text' || mark.kind === 'initials' || mark.kind === 'date')) {
    const size = mark.kind === 'initials' ? 18 : 12;
    const color = mark.kind === 'initials' || mark.kind === 'date' ? blue : ink;
    drawTopLeftText(page, font, mark.text, rect.x, rect.y, size, color);
  }
}

function drawSquiggle(
  page: PDFPage,
  box: { x: number; y: number; width: number },
  color: RGB,
  thickness: number,
  opacity: number,
): void {
  const points: Array<{ x: number; y: number }> = [];
  const end = box.x + box.width;
  const base = box.y + 1;
  const step = 1.6;
  for (let x = box.x; x <= end + 0.01; x += step) {
    const along = Math.min(x, end) - box.x;
    points.push({ x: Math.min(x, end), y: base + Math.sin(along / 1.7) * 1.15 });
  }
  drawPolyline(page, points, color, thickness, opacity);
}

function drawPolyline(page: PDFPage, points: Array<{ x: number; y: number }>, color: RGB, thickness: number, opacity = 1): void {
  for (let index = 1; index < points.length; index += 1) {
    page.drawLine({ start: points[index - 1], end: points[index], thickness, color, opacity });
  }
}

/** Scallops around a closed shape. The radius is in the same space as the corners. */
export function cloudOutline(corners: Array<{ x: number; y: number }>, radius: number): Array<{ x: number; y: number }> {
  if (corners.length < 3) return corners.map((point) => ({ ...point }));
  const bumpRadius = Math.max(radius, 1);
  const center = {
    x: corners.reduce((sum, point) => sum + point.x, 0) / corners.length,
    y: corners.reduce((sum, point) => sum + point.y, 0) / corners.length,
  };
  const outline: Array<{ x: number; y: number }> = [];
  for (let index = 0; index < corners.length; index += 1) {
    const start = corners[index];
    const end = corners[(index + 1) % corners.length];
    const dx = end.x - start.x;
    const dy = end.y - start.y;
    const length = Math.hypot(dx, dy);
    if (length < 0.5) continue;
    const bumps = Math.max(1, Math.round(length / (bumpRadius * 2)));
    const ux = dx / length;
    const uy = dy / length;
    let nx = -uy;
    let ny = ux;
    const midX = (start.x + end.x) / 2;
    const midY = (start.y + end.y) / 2;
    if (nx * (center.x - midX) + ny * (center.y - midY) > 0) {
      nx = -nx;
      ny = -ny;
    }
    const step = length / bumps;
    const curve = step / 2;
    for (let bump = 0; bump < bumps; bump += 1) {
      const cx = start.x + ux * (bump + 0.5) * step;
      const cy = start.y + uy * (bump + 0.5) * step;
      const samples = 6;
      for (let sample = 0; sample <= samples; sample += 1) {
        const theta = Math.PI - (Math.PI * sample) / samples;
        const along = Math.cos(theta);
        const side = Math.sin(theta);
        outline.push({ x: cx + along * curve * ux + side * curve * nx, y: cy + along * curve * uy + side * curve * ny });
      }
    }
  }
  return outline;
}

function drawArrow(page: PDFPage, from: { x: number; y: number }, to: { x: number; y: number }, color: RGB, thickness: number): void {
  page.drawLine({ start: from, end: to, thickness, color });
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const length = Math.hypot(dx, dy) || 1;
  const ux = dx / length;
  const uy = dy / length;
  const size = 9;
  const wing = 3.5;
  const base = { x: to.x - ux * size, y: to.y - uy * size };
  page.drawLine({ start: to, end: { x: base.x - uy * wing, y: base.y + ux * wing }, thickness, color });
  page.drawLine({ start: to, end: { x: base.x + uy * wing, y: base.y - ux * wing }, thickness, color });
}

function svgPath(points: Array<{ x: number; y: number }>, height: number, close: boolean): string {
  const commands = points.map((point, index) => `${index === 0 ? 'M' : 'L'} ${point.x.toFixed(2)} ${(height - point.y).toFixed(2)}`);
  if (close) commands.push('Z');
  return commands.join(' ');
}

function drawTopLeftText(page: PDFPage, font: PDFFont, text: string, x: number, top: number, size: number, color: RGB): void {
  if (!text) return;
  page.drawText(text, { x, y: top - size, size, font, color, lineHeight: size * 1.2 });
}

function drawCenteredText(page: PDFPage, font: PDFFont, text: string, x: number, y: number, size: number, color: RGB): void {
  if (!text) return;
  const width = font.widthOfTextAtSize(text.split('\n')[0] ?? '', size);
  page.drawText(text, { x: x - width / 2, y: y - size / 2, size, font, color, lineHeight: size * 1.2 });
}

/** Quarter turns are clockwise. Rotation is stored on each page. */
export async function rotateDocumentPages(bytes: Uint8Array, pages: number[], quarterTurns: number): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const delta = (((quarterTurns % 4) + 4) % 4) * 90;
  const count = doc.getPageCount();
  for (const pageNumber of pages) {
    if (pageNumber < 1 || pageNumber > count) continue;
    const page = doc.getPage(pageNumber - 1);
    page.setRotation(degrees((page.getRotation().angle + delta) % 360));
  }
  return doc.save();
}

/** `order` is the 1-based page numbers that should remain, in their new sequence. */
export async function rebuildPages(bytes: Uint8Array, order: number[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const count = doc.getPageCount();
  if (order.length < 1) throw new Error('Keep at least one page in the document.');
  if (order.some((page) => page < 1 || page > count)) throw new Error('That page is not in this document.');
  const next = await PDFDocument.create();
  const copied = await next.copyPages(doc, order.map((page) => page - 1));
  copied.forEach((page) => next.addPage(page));
  return next.save();
}

export async function insertBlankPage(bytes: Uint8Array, afterPage: number): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const count = doc.getPageCount();
  const after = Math.min(Math.max(afterPage, 0), count);
  const reference = doc.getPage(Math.min(Math.max(after, 1), count) - 1);
  const { width, height } = reference.getSize();
  doc.insertPage(after, [width, height]);
  return doc.save();
}

export async function deleteDocumentPages(bytes: Uint8Array, pages: number[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const count = doc.getPageCount();
  const drop = [...new Set(pages)].filter((page) => page >= 1 && page <= count).sort((a, b) => b - a);
  if (drop.length < 1) return doc.save();
  if (drop.length >= count) throw new Error('Keep at least one page in the document.');
  for (const page of drop) doc.removePage(page - 1);
  return doc.save();
}

/** Crops one page to a rectangle in PDF user space. */
export async function cropDocumentPage(bytes: Uint8Array, pageNumber: number, rect: PdfRect): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[pageNumber - 1];
  if (!page) throw new Error('Could not crop this page.');
  const x = Math.max(0, Math.min(rect.x, page.getWidth() - 8));
  const y = Math.max(0, Math.min(rect.y, page.getHeight() - 8));
  const width = Math.max(8, Math.min(rect.width, page.getWidth() - x));
  const height = Math.max(8, Math.min(rect.height, page.getHeight() - y));
  page.setCropBox(x, y, width, height);
  page.setMediaBox(x, y, width, height);
  return doc.save();
}

/** Replaces one page with the first page of another PDF. */
export async function replaceDocumentPage(bytes: Uint8Array, pageNumber: number, source: Uint8Array): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const incoming = await PDFDocument.load(source);
  if (incoming.getPageCount() < 1) throw new Error('That PDF has no pages.');
  if (pageNumber < 1 || pageNumber > doc.getPageCount()) throw new Error('That page is not in this document.');
  const [page] = await doc.copyPages(incoming, [0]);
  const index = pageNumber - 1;
  doc.removePage(index);
  doc.insertPage(index, page);
  return doc.save();
}

/** A new PDF containing the requested pages, in the order given. */
export async function extractDocumentPages(bytes: Uint8Array, pages: number[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const count = doc.getPageCount();
  const order = pages.filter((page) => page >= 1 && page <= count);
  if (!order.length) throw new Error('Choose a page to save.');
  const next = await PDFDocument.create();
  const copied = await next.copyPages(doc, order.map((page) => page - 1));
  copied.forEach((page) => next.addPage(page));
  return next.save();
}

function drawStamp(page: PDFPage, font: PDFFont, mark: DrawnMark, pageHeight: number): void {
  const rect = mark.rects[0];
  if (!rect) return;
  const style = stampStyle(mark.text, mark.color);
  const border = paintOf(style.ink, blue);
  const fill = paintOf(style.fill, rgb(1, 1, 1));
  const text = paintOf(style.text, border);
  if (style.shape === 'badge' || style.shape === 'arrow') {
    const path = style.shape === 'arrow' ? arrowPath(rect, pageHeight) : roundedRectPath(rect, 4, pageHeight);
    page.drawSvgPath(path, { x: 0, y: pageHeight, color: fill, borderColor: style.shape === 'arrow' ? fill : border, borderWidth: style.shape === 'arrow' ? 0.5 : 1.4 });
    const size = 8;
    const width = font.widthOfTextAtSize(style.label, size);
    const textX = style.shape === 'arrow' ? rect.x + (rect.width + 12 - width) / 2 : rect.x + (rect.width - width) / 2;
    page.drawText(style.label, { x: textX, y: rect.y + (rect.height - size) / 2, size, font, color: text });
    return;
  }
  const path = style.shape === 'check' ? checkPath(rect, pageHeight) : crossPath(rect, pageHeight);
  page.drawSvgPath(path, { x: 0, y: pageHeight, borderColor: border, borderWidth: 2.2 });
}

function roundedRectPath(rect: PdfRect, radius: number, pageHeight: number): string {
  const r = Math.min(radius, rect.width / 2, rect.height / 2);
  const left = rect.x;
  const right = rect.x + rect.width;
  const top = pageHeight - (rect.y + rect.height);
  const bottom = pageHeight - rect.y;
  return `M ${left + r} ${top} H ${right - r} Q ${right} ${top} ${right} ${top + r} V ${bottom - r} Q ${right} ${bottom} ${right - r} ${bottom} H ${left + r} Q ${left} ${bottom} ${left} ${bottom - r} V ${top + r} Q ${left} ${top} ${left + r} ${top} Z`;
}

function arrowPath(rect: PdfRect, pageHeight: number): string {
  const notch = Math.min(14, rect.width * 0.22);
  const points = [
    { x: rect.x + notch, y: rect.y + rect.height },
    { x: rect.x + rect.width, y: rect.y + rect.height },
    { x: rect.x + rect.width, y: rect.y },
    { x: rect.x + notch, y: rect.y },
    { x: rect.x, y: rect.y + rect.height / 2 },
  ];
  return svgPath(points, pageHeight, true);
}

function checkPath(rect: PdfRect, pageHeight: number): string {
  const points = [
    { x: rect.x + rect.width * 0.18, y: rect.y + rect.height * 0.48 },
    { x: rect.x + rect.width * 0.4, y: rect.y + rect.height * 0.24 },
    { x: rect.x + rect.width * 0.84, y: rect.y + rect.height * 0.76 },
  ];
  return svgPath(points, pageHeight, false);
}

function crossPath(rect: PdfRect, pageHeight: number): string {
  const inset = rect.width * 0.22;
  const first = svgPath(
    [
      { x: rect.x + inset, y: rect.y + rect.height - inset },
      { x: rect.x + rect.width - inset, y: rect.y + inset },
    ],
    pageHeight,
    false,
  );
  const second = svgPath(
    [
      { x: rect.x + inset, y: rect.y + inset },
      { x: rect.x + rect.width - inset, y: rect.y + rect.height - inset },
    ],
    pageHeight,
    false,
  );
  return `${first} ${second}`;
}

function drawPaperclip(page: PDFPage, rect: PdfRect, pageHeight: number): void {
  const left = rect.x + rect.width * 0.32;
  const right = rect.x + rect.width * 0.7;
  const top = rect.y + rect.height * 0.82;
  const bottom = rect.y + rect.height * 0.16;
  const mid = rect.x + rect.width * 0.5;
  const path = svgPath(
    [
      { x: mid, y: bottom },
      { x: left, y: bottom + rect.height * 0.12 },
      { x: left, y: top },
      { x: right, y: top },
      { x: right, y: bottom + rect.height * 0.28 },
      { x: mid, y: bottom + rect.height * 0.28 },
      { x: mid, y: top - rect.height * 0.16 },
    ],
    pageHeight,
    false,
  );
  page.drawSvgPath(path, { x: 0, y: pageHeight, borderColor: ink, borderWidth: 1.4 });
}

function drawBoxedText(page: PDFPage, font: PDFFont, text: string, x: number, top: number, size: number, color: RGB, border: RGB): void {
  if (!text) return;
  const lines = text.split('\n');
  const width = Math.max(...lines.map((line) => font.widthOfTextAtSize(line, size)), 8) + 12;
  const height = size * 1.2 * lines.length + 8;
  page.drawRectangle({ x, y: top - height, width, height, borderColor: border, borderWidth: 1.5, color: rgb(1, 1, 1) });
  page.drawText(text, { x: x + 6, y: top - size - 2, size, font, color, lineHeight: size * 1.2 });
}
