/**
 * A rectangle of cells as clipboard text, and back.
 *
 * Tab-separated, newline between rows — which is not a house format but the one
 * Excel, Google Sheets, Numbers and every database console already speak. It is
 * what makes a copy out of this grid paste into a spreadsheet, and a copy out of
 * a spreadsheet paste into this grid, without either end knowing about the
 * other.
 *
 * The quoting is Excel's: a cell containing a tab, a newline or a quote is
 * wrapped in quotes, and a quote inside is doubled. Anything else is written
 * bare, so the common case is exactly the text the author sees.
 */
import {
    cellToInput, columnToIndex, indexToColumn, parseRange, type SheetDto,
} from './sheet-document.model';

const NEEDS_QUOTING = /["\t\r\n]/;

/** One cell as clipboard text. */
function quote(value: string): string {
    return NEEDS_QUOTING.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/**
 * A range as the text a spreadsheet expects on the clipboard.
 *
 * The cells' RAW text, formulas included: within this grid a formula pasted
 * elsewhere is offset and stays a formula, and pasted into Excel it arrives as
 * one too. Copying the computed values instead would quietly turn a table of
 * formulas into a table of numbers, which is a different document.
 */
export function toClipboardText(sheet: SheetDto, range: string): string {
    const box = parseRange(range.includes(':') ? range : `${range}:${range}`);
    if (!box) return '';

    const lines: string[] = [];
    for (let row = box.top; row <= box.bottom; row++) {
        const line: string[] = [];
        for (let column = box.left; column <= box.right; column++) {
            line.push(quote(cellToInput(sheet.cells[indexToColumn(column) + row])));
        }
        lines.push(line.join('\t'));
    }

    return lines.join('\n');
}

/**
 * Clipboard text as a rectangle.
 *
 * Quoted cells may contain the very characters the format separates on, so this
 * cannot be `split('\n').map(split('\t'))` — a single-cell paste of a note with
 * a line break in it would arrive as two rows. Scanned character by character
 * instead, which is the only way to know whether a newline is data.
 *
 * A trailing newline is dropped: every spreadsheet ends its clipboard text with
 * one, and honouring it would paste an empty row that clears whatever was
 * underneath.
 */
export function parseClipboardText(text: string): string[][] {
    const rows: string[][] = [];
    let line: string[] = [];
    let cell = '';
    let quoted = false;

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];

        if (quoted) {
            if (ch === '"') {
                if (text[i + 1] === '"') { cell += '"'; i += 1; continue; }
                quoted = false;
                continue;
            }
            cell += ch;
            continue;
        }

        if (ch === '"' && cell === '') { quoted = true; continue; }
        if (ch === '\t') { line.push(cell); cell = ''; continue; }
        if (ch === '\r') continue;
        if (ch === '\n') { line.push(cell); rows.push(line); line = []; cell = ''; continue; }

        cell += ch;
    }

    // Whatever is left is the last cell -- unless the text ended on a newline,
    // in which case the row was already closed and there is nothing pending.
    if (cell !== '' || line.length > 0) {
        line.push(cell);
        rows.push(line);
    }

    return rows;
}

/** The range a block of this shape fills when pasted at `at`. */
export function pastedRange(at: string, block: readonly (readonly string[])[]): string {
    const corner = parseRange(`${at}:${at}`);
    if (!corner || block.length === 0) return at;

    const width = Math.max(...block.map(line => line.length));
    const right = indexToColumn(corner.left + Math.max(0, width - 1));

    return `${indexToColumn(corner.left)}${corner.top}:${right}${corner.top + block.length - 1}`;
}

/** Where a paste at `at` moves a formula copied from `from`. */
export function pasteOffset(from: string, at: string): { rows: number; columns: number } | null {
    const source = parseRange(`${from}:${from}`);
    const target = parseRange(`${at}:${at}`);
    if (!source || !target) return null;

    return { rows: target.top - source.top, columns: target.left - source.left };
}

/** The top-left corner of a range, which is where a copy is relative to. */
export function rangeCorner(range: string): string {
    const box = parseRange(range.includes(':') ? range : `${range}:${range}`);

    return box ? indexToColumn(box.left) + box.top : range;
}
