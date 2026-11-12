import { read, utils, type CellObject, type Range, type WorkSheet } from 'xlsx';

const MAX_ROWS = 500;
const MAX_COLS = 50;

/**
 * Renders a workbook into `body` as one `.preview-sheet` per worksheet.
 * The caller shows a single sheet at a time. Macros are not run.
 */
export function renderSheetPages(bytes: Uint8Array, body: HTMLElement): number {
  body.replaceChildren();
  let names: string[] = [];
  let sheets: Record<string, WorkSheet> = {};
  try {
    const workbook = read(bytes, { type: 'array', cellDates: true });
    names = workbook.SheetNames;
    sheets = workbook.Sheets;
  } catch {
    throw new Error('Could not open this spreadsheet.');
  }
  if (names.length < 1) return 0;
  const host = document.createElement('div');
  host.className = 'sheet-preview';
  for (const name of names) {
    host.append(sheetPage(name, sheets[name]));
  }
  body.append(host);
  return names.length;
}

function sheetPage(name: string, sheet: WorkSheet | undefined): HTMLElement {
  const page = document.createElement('section');
  page.className = 'preview-sheet';
  const title = document.createElement('h2');
  title.className = 'sheet-name';
  title.textContent = name;
  page.append(title);
  const built = sheet ? tableFromSheet(sheet) : null;
  if (!built) {
    const empty = document.createElement('p');
    empty.className = 'sheet-empty';
    empty.textContent = 'This sheet is empty.';
    page.append(empty);
    return page;
  }
  page.append(built.table);
  if (built.truncated) {
    const note = document.createElement('p');
    note.className = 'sheet-note';
    note.textContent = `Showing the first ${MAX_ROWS} rows and ${MAX_COLS} columns of this sheet.`;
    page.append(note);
  }
  return page;
}

function tableFromSheet(sheet: WorkSheet): { table: HTMLTableElement; truncated: boolean } | null {
  const ref = sheet['!ref'];
  if (!ref) return null;
  const range = utils.decode_range(ref);
  const rowCount = range.e.r - range.s.r + 1;
  const colCount = range.e.c - range.s.c + 1;
  if (rowCount < 1 || colCount < 1) return null;
  const rows = Math.min(rowCount, MAX_ROWS);
  const cols = Math.min(colCount, MAX_COLS);
  const endRow = range.s.r + rows - 1;
  const endCol = range.s.c + cols - 1;
  const { covered, spans } = mergeMap(sheet['!merges'] ?? [], range.s.r, range.s.c, endRow, endCol);

  const table = document.createElement('table');
  table.className = 'sheet-table';
  const head = document.createElement('tr');
  head.append(document.createElement('th'));
  for (let column = range.s.c; column <= endCol; column += 1) {
    const cell = document.createElement('th');
    cell.textContent = utils.encode_col(column);
    head.append(cell);
  }
  const thead = document.createElement('thead');
  thead.append(head);
  table.append(thead);

  const body = document.createElement('tbody');
  for (let row = range.s.r; row <= endRow; row += 1) {
    const line = document.createElement('tr');
    const label = document.createElement('th');
    label.textContent = String(row + 1);
    line.append(label);
    for (let column = range.s.c; column <= endCol; column += 1) {
      const key = `${row}:${column}`;
      if (covered.has(key)) continue;
      const cell = document.createElement('td');
      const span = spans.get(key);
      if (span) {
        if (span.cols > 1) cell.colSpan = span.cols;
        if (span.rows > 1) cell.rowSpan = span.rows;
      }
      cell.textContent = cellText(sheet[utils.encode_cell({ r: row, c: column })]);
      line.append(cell);
    }
    body.append(line);
  }
  table.append(body);
  return { table, truncated: rows < rowCount || cols < colCount };
}

function mergeMap(
  merges: Range[],
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): { covered: Set<string>; spans: Map<string, { rows: number; cols: number }> } {
  const covered = new Set<string>();
  const spans = new Map<string, { rows: number; cols: number }>();
  for (const merge of merges) {
    if (merge.s.r > endRow || merge.s.c > endCol || merge.e.r < startRow || merge.e.c < startCol) continue;
    const row = Math.max(merge.s.r, startRow);
    const column = Math.max(merge.s.c, startCol);
    const lastRow = Math.min(merge.e.r, endRow);
    const lastCol = Math.min(merge.e.c, endCol);
    if (merge.s.r >= startRow && merge.s.c >= startCol) {
      spans.set(`${merge.s.r}:${merge.s.c}`, { rows: lastRow - merge.s.r + 1, cols: lastCol - merge.s.c + 1 });
    }
    for (let r = row; r <= lastRow; r += 1) {
      for (let c = column; c <= lastCol; c += 1) {
        if (r === merge.s.r && c === merge.s.c) continue;
        covered.add(`${r}:${c}`);
      }
    }
  }
  return { covered, spans };
}

function cellText(value: unknown): string {
  const cell = asCell(value);
  if (!cell) return '';
  if (typeof cell.w === 'string' && cell.w.length > 0) return cell.w;
  if (cell.v != null && cell.t !== 'z') {
    try {
      return utils.format_cell(cell);
    } catch {
      return cell.v instanceof Date ? cell.v.toLocaleDateString() : String(cell.v);
    }
  }
  if (typeof cell.f === 'string' && cell.f) return `=${cell.f}`;
  return '';
}

function asCell(value: unknown): CellObject | null {
  if (!value || typeof value !== 'object' || !('t' in value)) return null;
  return value as CellObject;
}
