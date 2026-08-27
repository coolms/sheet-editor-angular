/**
 * The `.dsheet` document (ADR-155), and the A1 arithmetic a grid needs.
 *
 * Mirrors the backend's `SheetDocument` / `SheetCell` exactly — same key names,
 * same sparseness, same "value XOR formula" rule — because the file this reads
 * is the same file the renderer reads. Anything invented here would survive the
 * editor and fail at generation time, where nobody is watching.
 *
 * Everything in this file is PURE so it can be tested without a browser; the
 * component is the only part that needs one.
 */

import type { ConditionalDto } from './conditional';
import { parseDateInput } from './number-format';

export interface SheetCellDto {
    /** Literal content. May carry DTMPL tokens. Mutually exclusive with `formula`. */
    value?: string;
    /** An A1 formula WITHOUT its leading `=`. */
    formula?: string;
    /** An OOXML number-format code — `@` declares text (#1977). */
    numberFormat?: string;
    bold?: boolean;
    italic?: boolean;
    /** Font family NAME as Excel stores it. Absent inherits the workbook default. */
    fontFamily?: string;
    /** POINTS, as Excel and the toolbar both speak. */
    fontSize?: number;
    /** Text colour, `#RRGGBB` upper-case — the backend normalises the spelling. */
    color?: string;
    /** Solid fill, `#RRGGBB`. Absent means NO fill, which is not the same as white. */
    background?: string;
    align?: CellAlign;
    /**
     * Hyperlink target. May carry DTMPL tokens — `{var:order.trackingUrl}` is
     * the point of putting one in a template — so it is NOT validated here.
     * The backend's writer is the gate, at the last moment before the URL
     * enters a workbook someone will click (#2102).
     */
    link?: string;
    /**
     * Break the text inside the cell instead of letting it run past the edge.
     *
     * The switch between Excel's two behaviours: unwrapped, text SPILLS across
     * empty neighbours and clips at the first occupied one; wrapped, it breaks
     * inside the cell and the row grows to hold it. Absent is off — `wrapText`
     * is a boolean in OOXML with a real default, so there is no third state.
     */
    wrap?: boolean;
    /** Where content sits in a row TALLER than itself. Absent inherits bottom. */
    valign?: CellVAlign;
    /**
     * Ruled edges: `top` | `right` | `bottom` | `left` => `"thin"` or
     * `"thin #FF0000"`.
     *
     * Per EDGE because Excel rules each side separately and so does every
     * instruction an author gives — "underline the header", "box the table". A
     * box round a RANGE is therefore stored as an edge on each cell of its
     * perimeter, which is how OOXML stores it too; the editor offers "outer"
     * and "all" as gestures over a selection rather than as a stored shape.
     */
    borders?: Record<string, string>;
}

/** The edges a cell can rule, in CSS order. */
export const CELL_EDGES = ['top', 'right', 'bottom', 'left'] as const;

export type CellEdge = (typeof CELL_EDGES)[number];

/**
 * The line styles the writer can express, with what each looks like in CSS.
 *
 * OOXML's vocabulary trimmed to the ones an author asks for by name — `hair`
 * and the slanted/medium-dashed family exist in the format, nobody draws them
 * on purpose, and every one offered is one more thing the control must explain.
 */
export const BORDER_STYLES: Readonly<Record<string, string>> = {
    thin: '1px solid',
    medium: '2px solid',
    thick: '3px solid',
    double: '3px double',
    dashed: '1px dashed',
    dotted: '1px dotted',
};

/** What a border control can do to a selection. */
export type BorderPreset = 'all' | 'outer' | 'top' | 'right' | 'bottom' | 'left' | 'none';

/** What the writer can express; anything else is dropped server-side. */
export type CellAlign = 'left' | 'center' | 'right';
export type CellVAlign = 'top' | 'middle' | 'bottom';

export interface SheetDto {
    cells: Record<string, SheetCellDto>;
    /**
     * `B2:B9` range => the rules that repaint it, in the author's order.
     *
     * A LIST per range, and the ORDER is the priority: Excel applies the first
     * rule that matches and stops, so "red over 90" before "amber over 60" is
     * red at 95 -- and the other way round, nothing is ever red. Typed loosely
     * here and validated by `conditional.ts`, which owns the vocabulary.
     */
    conditionals?: Record<string, ConditionalDto[]>;
    /**
     * The cell BELOW AND RIGHT of the frozen panes, or absent for none.
     *
     * `B2` freezes the first row and the first column; `A3` freezes two rows
     * and no column. One ref rather than two counts because it is what OOXML
     * stores -- `<pane topLeftCell="B2" state="frozen">` -- and a shape that
     * mirrors the file cannot disagree with it. `A1` freezes nothing and is
     * therefore never stored: absence is how the format says "no panes".
     */
    freeze?: string;
    columnWidths?: Record<string, number>;
    /**
     * Row number => height in POINTS. Keyed by number, so it must survive as a
     * JSON object — the backend casts it for exactly that reason (ADR-157).
     */
    rowHeights?: Record<string, number>;
    merges?: string[];
    /**
     * The ONE range this sheet's filter covers, `A1:D9`, header row included.
     *
     * Singular where `merges` is a list, and that mirrors OOXML rather than
     * simplifying it: a worksheet carries at most one `<autoFilter>`, so a list
     * here would let the editor express a workbook Excel refuses to open.
     *
     * It is a DECLARATION, not a view state. What it buys is dropdowns in the
     * generated workbook over the right rows — including rows a `{loop:}` band
     * has not created yet, which the backend grows the range to cover. Which
     * values are hidden while editing is not stored, because a template has no
     * data to hide yet.
     */
    autoFilter?: string;
    /**
     * Range => the rule for what may be typed there — the sheet's form
     * controls.
     *
     * Keyed by RANGE because OOXML is: a `<dataValidation>` carries an `sqref`
     * and covers every cell in it. A single cell is a legal key, unlike a merge
     * or a filter — "this one cell is a dropdown" is the ordinary case.
     */
    validations?: Record<string, SheetValidationDto>;
}

/** What the backend's `SheetValidation` can express. */
export type ValidationType = 'list' | 'checkbox' | 'date' | 'time' | 'whole' | 'decimal' | 'textLength';

export interface SheetValidationDto {
    type: ValidationType;
    /** A `list`'s options, in the order they are shown. May carry DTMPL tokens. */
    values?: string[];
    /** OOXML operator; only meaningful for the comparing types. */
    operator?: string;
    min?: string;
    max?: string;
    /** Whether an empty cell passes. Absent means yes, which is Excel's default. */
    allowBlank?: boolean;
    prompt?: string;
    error?: string;
}

/**
 * What a checkbox writes into the cell.
 *
 * Excel has no checkbox validation type — a real one is a form control living
 * in a drawing part — so the backend emits a TRUE/FALSE list and this editor
 * draws the tick. These are the two values that dropdown offers, stated in one
 * place so the grid and the writer cannot drift.
 */
export const CHECKBOX_VALUES: readonly string[] = ['TRUE', 'FALSE'];

export interface SheetDocumentDto {
    version: number;
    sheets: Record<string, SheetDto>;
    /**
     * Name → the range it stands for: `items_amount` → `Sheet1!$B$2:$B$3`.
     *
     * Document-level, because a defined name is: the backend keeps it on the
     * workbook and the range names its own sheet.
     *
     * ⚠️ The grid needs these to PREVIEW honestly. A formula is written through
     * to the `.xlsx` verbatim, so `SUM(items_amount)` renders correctly whether
     * or not this editor understands it — and an editor that showed `#NAME?`
     * for a formula the document computes would be lying about the document.
     * Optional so a `.dsheet` written before names existed still loads.
     */
    definedNames?: Record<string, string>;
}

/** What the backend mints for a new native template. */
export const EMPTY_SHEET_DOCUMENT: SheetDocumentDto = {
    version: 1,
    sheets: { Sheet1: { cells: {} } },
};

/** `A` → 1, `Z` → 26, `AA` → 27. */
export function columnToIndex(letters: string): number {
    let index = 0;
    for (const char of letters) {
        index = index * 26 + (char.charCodeAt(0) - 64);
    }

    return index;
}

/** 1 → `A`, 26 → `Z`, 27 → `AA`. */
export function indexToColumn(index: number): string {
    let out = '';
    let n = index;
    while (n > 0) {
        const remainder = (n - 1) % 26;
        out = String.fromCharCode(65 + remainder) + out;
        n = Math.floor((n - 1) / 26);
    }

    return out;
}

/** `B4` → `{ column: 'B', row: 4 }`; null when it is not an A1 reference. */
export function parseRef(ref: string): { column: string; row: number } | null {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(ref.toUpperCase());

    return match ? { column: match[1], row: Number(match[2]) } : null;
}

/**
 * How far the grid must reach to show every cell the document defines, plus
 * room to type into.
 *
 * The padding is what makes the grid an EDITOR rather than a viewer: a document
 * whose last cell is B4 would otherwise render four rows with nowhere to add a
 * fifth. Minimums cover the empty document, which is what a new native template
 * is.
 */
/**
 * How much grid EXISTS. Since #2067 that is no longer how much is rendered.
 *
 * 1,000 rows x 26 columns is what Google Sheets opens on, and the row count is
 * now free: {@link SheetEditorDialogComponent} renders only the rows in the
 * viewport and stands spacer boxes in for the rest, so a thousand rows costs
 * the same DOM as thirty.
 *
 * ⚠️ COLUMNS are still all rendered, so their floor is a real cost and they
 * grow on demand instead. Column virtualisation is harder than row
 * virtualisation here — a `<colgroup>` sizes the table and every row would have
 * to agree on which columns it skips — and 26 columns is what an author expects
 * to find, so the trade is left where it is rather than half-made.
 */
export function gridExtent(
    sheet: SheetDto,
    { minRows = 1000, minCols = 26, padRows = 6, padCols = 3 } = {},
): { rows: number; cols: number } {
    let maxRow = 0;
    let maxCol = 0;

    for (const ref of Object.keys(sheet.cells)) {
        const parsed = parseRef(ref);
        if (!parsed) continue;
        maxRow = Math.max(maxRow, parsed.row);
        maxCol = Math.max(maxCol, columnToIndex(parsed.column));
    }

    return {
        rows: Math.max(minRows, maxRow + padRows),
        cols: Math.max(minCols, maxCol + padCols),
    };
}

/**
 * What the operator types for a cell: a formula shows with its `=`, everything
 * else shows its value.
 */
export function cellToInput(cell: SheetCellDto | undefined): string {
    if (!cell) return '';

    return cell.formula !== undefined ? '=' + cell.formula : (cell.value ?? '');
}

/**
 * The reverse, preserving everything the grid does not edit.
 *
 * `numberFormat` and `bold` are carried through untouched — the number format
 * is the author's TYPE DECLARATION (#1977) and losing it on an unrelated edit
 * would turn an order number back into arithmetic. Returning `null` for an
 * emptied cell keeps the document sparse, which is how the backend writes it.
 */
export function inputToCell(raw: string, previous: SheetCellDto | undefined): SheetCellDto | null {
    const text = raw;

    // ⚠️ Everything the cell had EXCEPT what this edit replaces. Naming the
    // fields to KEEP is what let ten of them go missing: the list held
    // `numberFormat` and `bold`, so retyping the words in a cell silently threw
    // away its colour, its fill, its alignment, its link and its borders. Every
    // field added to the DTO since joined the document, the toolbar and the
    // writer -- and quietly failed to join that list. Subtracting what is
    // REPLACED cannot go stale in the same way.
    const kept: SheetCellDto = { ...(previous ?? {}) };
    delete kept.value;
    delete kept.formula;

    if (text === '') {
        // Nothing typed: keep the cell ONLY if it still carries formatting the
        // author set deliberately, otherwise drop it from the document.
        return Object.keys(kept).length > 0 ? kept : null;
    }

    if (text.startsWith('=')) return { formula: text.slice(1), ...kept };

    // A date cell EDITS as a date and STORES as a serial -- see `editForm`.
    // Without this half the round trip is one-way: the author is shown
    // `21/08/2026`, types over it, and the cell becomes the text of a date
    // that no longer computes.
    const serial = parseDateInput(text, kept.numberFormat);

    return { value: serial === null ? text : String(serial), ...kept };
}

/** A number format the toolbar offers. `code` undefined means General. */
export interface NumberFormatOption {
    label: string;
    code?: string;
}

/**
 * The formats worth a menu entry.
 *
 * `Text` is first among the non-default ones because it is the one that
 * CHANGES MEANING rather than appearance: `@` is the author's declaration that
 * a value is not arithmetic, and it is what keeps `00412` an order number
 * rather than the integer 412 (#1977). The rest are presentation.
 *
 * Deliberately short. A curated list an operator can read beats a complete one
 * they cannot, and a code outside it is still preserved — see
 * {@link isKnownFormat}.
 */
export const NUMBER_FORMATS: readonly NumberFormatOption[] = [
    { label: 'General' },
    { label: 'Text', code: '@' },
    { label: 'Number', code: '#,##0.00' },
    { label: 'Integer', code: '0' },
    { label: 'Percent', code: '0.00%' },
    { label: 'Date', code: 'dd/mm/yyyy' },
];

/**
 * Whether the toolbar can represent this code.
 *
 * A `.dsheet` may carry any OOXML format code — an author who hand-wrote
 * `#,##0.00\ [$€-407]` in the JSON has one the menu does not list. The editor
 * must SHOW that rather than display "General" and overwrite it on the next
 * unrelated change, which is how a silent loss happens.
 */
export function isKnownFormat(code: string | undefined): boolean {
    return NUMBER_FORMATS.some(option => option.code === code);
}

/**
 * True when nothing is left worth storing — mirrors the backend's isEmpty().
 *
 * ⚠️ EVERY styling field counts. A cell holding only a background colour has no
 * text and is still something the author made on purpose; treating it as blank
 * would delete a shaded row the moment anything else on it changed. The two
 * sides of this rule have to agree, or the editor and the file disagree about
 * which cells exist.
 */
function isBlank(cell: SheetCellDto): boolean {
    return cell.value === undefined
        && cell.formula === undefined
        && cell.numberFormat === undefined
        && cell.bold !== true
        && cell.italic !== true
        && cell.fontFamily === undefined
        && cell.fontSize === undefined
        && cell.color === undefined
        && cell.background === undefined
        && cell.align === undefined
        && cell.wrap !== true
        && cell.valign === undefined
        && cell.link === undefined;
}

/**
 * Apply a number format, keeping everything else. `undefined` clears it back to
 * General. Returns null when the cell has nothing left to store.
 *
 * A format may be applied to a cell that does not exist yet — an author marks a
 * column as Text BEFORE typing into it, and that intent has to survive.
 */
export function withNumberFormat(cell: SheetCellDto | undefined, code: string | undefined): SheetCellDto | null {
    const next: SheetCellDto = { ...(cell ?? {}) };
    if (code === undefined) {
        delete next.numberFormat;
    } else {
        next.numberFormat = code;
    }

    return isBlank(next) ? null : next;
}

/** As {@link withNumberFormat}, for weight. */
export function withBold(cell: SheetCellDto | undefined, bold: boolean): SheetCellDto | null {
    const next: SheetCellDto = { ...(cell ?? {}) };
    if (bold) {
        next.bold = true;
    } else {
        delete next.bold;
    }

    return isBlank(next) ? null : next;
}

/** As {@link withBold}, for slant. */
export function withItalic(cell: SheetCellDto | undefined, italic: boolean): SheetCellDto | null {
    const next: SheetCellDto = { ...(cell ?? {}) };
    if (italic) {
        next.italic = true;
    } else {
        delete next.italic;
    }

    return isBlank(next) ? null : next;
}

/** As {@link withBold}, for wrapping. */
export function withWrap(cell: SheetCellDto | undefined, wrap: boolean): SheetCellDto | null {
    const next: SheetCellDto = { ...(cell ?? {}) };
    if (wrap) {
        next.wrap = true;
    } else {
        delete next.wrap;
    }

    return isBlank(next) ? null : next;
}

/**
 * Set or clear one of the optional style fields (#2060).
 *
 * `undefined` CLEARS rather than storing an empty value, which is the same rule
 * every other `with*` here follows: absent means "inherit the workbook
 * default", and storing `''` or `0` would mean "this cell states nothing" in a
 * way the writer has to special-case and a reader cannot distinguish from a
 * deliberate choice.
 */
export function withStyle<K extends 'fontFamily' | 'fontSize' | 'color' | 'background' | 'align' | 'valign' | 'link'>(
    cell: SheetCellDto | undefined,
    key: K,
    value: SheetCellDto[K] | undefined,
): SheetCellDto | null {
    const next: SheetCellDto = { ...(cell ?? {}) };
    if (value === undefined || '' === value) {
        delete next[key];
    } else {
        next[key] = ('color' === key || 'background' === key)
            ? (normaliseColour(value as string) as SheetCellDto[K])
            : value;
    }

    return isBlank(next) ? null : next;
}

/**
 * `#RRGGBB` upper-case — the same canonical form the backend's `SheetCell`
 * produces when it parses a `.dsheet` (#2076).
 *
 * A browser colour input emits LOWER case, so without this the editor wrote
 * `#ffee00` into a file the backend would rewrite as `#FFEE00`. Both work —
 * nothing compares colours as strings — but a re-save then shows a case-only
 * diff on a line nobody touched, which is noise in a source file an author
 * reads. One canonical form, decided by whoever writes it first.
 */
export function normaliseColour(value: string): string {
    return /^#?[0-9A-Fa-f]{6}$/.test(value)
        ? '#' + value.replace('#', '').toUpperCase()
        : value;
}

/**
 * Store a row's height, or clear it. Blank clears.
 *
 * Refuses a non-positive height for the same reason {@link withColumnWidth}
 * refuses a non-positive width: Excel reads 0 as a HIDDEN row, so a stray zero
 * would remove the row from the generated workbook with nothing in the editor
 * to explain it.
 */
export function withRowHeight(sheet: SheetDto, row: number, height: number | undefined): SheetDto {
    const key = String(row);
    const heights = { ...(sheet.rowHeights ?? {}) };

    if (height === undefined) {
        delete heights[key];
    } else {
        if (!Number.isFinite(height) || height <= 0) return sheet;
        heights[key] = height;
    }

    const next: SheetDto = { ...sheet };
    if (Object.keys(heights).length > 0) {
        next.rowHeights = heights;
    } else {
        delete next.rowHeights;
    }

    return next;
}

export function rowHeightOf(sheet: SheetDto, row: number): number | undefined {
    return sheet.rowHeights?.[String(row)];
}

/**
 * Approximate pixels for a stored row height, so the grid SHOWS what the author
 * set. Heights are in POINTS, and a CSS pixel is 1/96in against a point's
 * 1/72in — the one conversion, unlike column widths, that is exact.
 */
export function rowHeightToPx(points: number): number {
    return Math.round(points * (96 / 72));
}

/** The inverse, for the row-header drag handle. */
export function rowHeightFromPx(px: number): number {
    return Math.max(MIN_ROW_HEIGHT, Math.round(px * (72 / 96) * 100) / 100);
}

/** As {@link MIN_COLUMN_WIDTH}: a drag must not reach Excel's "hidden". */
export const MIN_ROW_HEIGHT = 4;

/** Excel's cap, mirrored from `SheetDocumentWriter::safeSheetName()`. */
export const SHEET_NAME_MAX = 31;

/**
 * The name a sheet will actually carry in the generated workbook.
 *
 * Mirrors the backend's `safeSheetName()` exactly — same forbidden characters,
 * same 31-character cap — so the tab an author names is the tab they get. The
 * writer applies these rules whatever the editor does; normalising HERE is what
 * stops a name silently changing between the grid and the xlsx, which is the
 * same "what you see is what renders" rule that decides merge clearing.
 */
export function safeSheetName(name: string, fallbackIndex = 0): string {
    const clean = name.replace(/[\\/*?:[\]]/g, '-').slice(0, SHEET_NAME_MAX).trim();

    return '' !== clean ? clean : 'Sheet' + (fallbackIndex + 1);
}

/** A name not already taken, with a numeric suffix if needed. */
export function uniqueSheetName(doc: SheetDocumentDto, desired: string): string {
    const base = safeSheetName(desired, Object.keys(doc.sheets).length);
    if (!(base in doc.sheets)) return base;

    for (let n = 2; n < 1000; n++) {
        // Re-trim: the suffix must not push the name past Excel's cap.
        const candidate = safeSheetName(base.slice(0, SHEET_NAME_MAX - String(n).length - 1) + ' ' + n);
        if (!(candidate in doc.sheets)) return candidate;
    }

    return base;
}

/** Append a sheet. A name already in use is suffixed rather than overwriting. */
export function withNewSheet(doc: SheetDocumentDto, desired: string): { doc: SheetDocumentDto; name: string } {
    const name = uniqueSheetName(doc, desired);

    return { doc: { ...doc, sheets: { ...doc.sheets, [name]: { cells: {} } } }, name };
}

/**
 * Rename a sheet IN PLACE, keeping its position.
 *
 * Rebuilding the map in order is the whole point: JS object keys iterate in
 * insertion order, the backend writes sheets in that same order, and
 * `SheetDocumentWriter` makes index 0 the ACTIVE sheet. Deleting and re-adding
 * would silently move a renamed sheet to the end of the workbook — and renaming
 * the first sheet would hand the operator a different opening tab.
 *
 * A no-op rename, an unknown source, or a name already taken all return the
 * document untouched.
 */
export function withRenamedSheet(doc: SheetDocumentDto, from: string, to: string): SheetDocumentDto {
    if (!(from in doc.sheets)) return doc;
    const name = safeSheetName(to);
    if ('' === name || name === from || name in doc.sheets) return doc;

    const sheets: Record<string, SheetDto> = {};
    for (const [key, sheet] of Object.entries(doc.sheets)) {
        sheets[key === from ? name : key] = sheet;
    }

    return { ...doc, sheets };
}

/**
 * Remove a sheet — never the last one.
 *
 * `SheetDocumentWriter::write()` throws on a document with no sheets, so an
 * editor that allowed it would produce a template that only fails at generation
 * time, where nobody is watching.
 */
export function withoutSheet(doc: SheetDocumentDto, name: string): SheetDocumentDto {
    if (!(name in doc.sheets) || Object.keys(doc.sheets).length <= 1) return doc;

    const sheets = { ...doc.sheets };
    delete sheets[name];

    return { ...doc, sheets };
}

/** A merge range resolved to numbers. Columns are INDEXES, not letters. */
export interface MergeBox {
    top: number;
    left: number;
    bottom: number;
    right: number;
}

/** `A1:D2` → its box; null when it is not a range. */
export function parseRange(range: string): MergeBox | null {
    const [from, to] = range.toUpperCase().split(':');
    const a = from ? parseRef(from) : null;
    const b = to ? parseRef(to) : null;
    if (!a || !b) return null;

    return {
        top: Math.min(a.row, b.row),
        bottom: Math.max(a.row, b.row),
        left: Math.min(columnToIndex(a.column), columnToIndex(b.column)),
        right: Math.max(columnToIndex(a.column), columnToIndex(b.column)),
    };
}

/**
 * The range spanning two cells, normalised so the top-left comes first.
 *
 * Normalising here rather than at the call site is what lets an author drag or
 * shift-click in ANY direction — up-left to bottom-right is the same merge as
 * bottom-right to up-left, and a `D4:A1` written into the document would not
 * match the OOXML the renderer emits.
 */
export function rangeBetween(a: string, b: string): string | null {
    const box = parseRange(a + ':' + b);
    if (!box) return null;

    return indexToColumn(box.left) + box.top + ':' + indexToColumn(box.right) + box.bottom;
}

/**
 * Whether a range covers a cell, without building the range first.
 *
 * The same answer as `refsInRange(range).includes(ref)` and the reason that call
 * is no longer made per-cell: the grid asks this once for EVERY rendered cell on
 * every change detection, so the list version is quadratic in the selection. A
 * whole-column selection — which the column headers now make a one-click gesture
 * — turns a 500-row sheet into millions of string comparisons per keystroke.
 */
export function rangeContains(range: string, ref: string): boolean {
    const box = parseRange(range);
    const cell = parseRef(ref);
    if (!box || !cell) return false;
    const column = columnToIndex(cell.column);

    return cell.row >= box.top && cell.row <= box.bottom
        && column >= box.left && column <= box.right;
}

/** Every A1 reference inside a range, the anchor included. */
export function refsInRange(range: string): string[] {
    const box = parseRange(range);
    if (!box) return [];

    const out: string[] = [];
    for (let row = box.top; row <= box.bottom; row++) {
        for (let col = box.left; col <= box.right; col++) {
            out.push(indexToColumn(col) + row);
        }
    }

    return out;
}

/** The merge covering a cell, if any. */
export function mergeCovering(sheet: SheetDto, ref: string): string | null {
    const cell = parseRef(ref);
    if (!cell) return null;
    const column = columnToIndex(cell.column);

    for (const range of sheet.merges ?? []) {
        const box = parseRange(range);
        if (!box) continue;
        if (cell.row >= box.top && cell.row <= box.bottom && column >= box.left && column <= box.right) {
            return range;
        }
    }

    return null;
}

/** True when `ref` is a merge's TOP-LEFT — the only cell of it that renders. */
export function isMergeAnchor(range: string, ref: string): boolean {
    const box = parseRange(range);
    const cell = parseRef(ref);

    return !!box && !!cell && cell.row === box.top && columnToIndex(cell.column) === box.left;
}


/**
 * A sheet's merges and validation rules, answered per cell in constant time.
 *
 * WHY this exists, when {@link mergeCovering} and {@link validationAt} already
 * answer the same questions: the grid asks for the merge FOUR times and the
 * rule TWICE for every rendered cell, on every change detection -- and both
 * are linear scans that re-parse every range on every call. The cost is
 * therefore cells x ranges, and ranges are ordinary: a report with a merged
 * banner over each section has dozens.
 *
 * MEASURED in the browser, not reasoned about. 384 rendered cells over 25
 * merges and 10 validations cost 25.5ms a change-detection pass, against 7.1ms
 * for the same grid over one merge and one rule -- slower with FEWER cells,
 * and well past a frame, so the sheet stuttered on content nobody would call
 * unusual. This is the same trap {@link rangeContains} records one screen up,
 * met again from the other side.
 *
 * The LIFETIME is the whole design. Build one per document-and-sheet and drop
 * it when either changes: a lookup that outlives the sheet it read is a grid
 * showing merges the document no longer has. The dialog holds it in a
 * `computed` over the document signal, which is what makes that automatic.
 *
 * First match wins, and the ranges are visited in their original order,
 * because that is the rule {@link validationRangeAt} already documents.
 */
export interface SheetLookup {
    /** The merge covering a cell, if any. */
    mergeCovering(ref: string): string | null;

    /** The validation rule covering a cell, if any. */
    validationAt(ref: string): SheetValidationDto | null;
}

/** Build a {@link SheetLookup} over one sheet. An absent sheet answers nothing. */
export function sheetLookup(sheet: SheetDto | undefined): SheetLookup {
    const merges: { range: string; box: MergeBox }[] = [];
    for (const range of sheet?.merges ?? []) {
        const box = parseRange(range);
        if (null !== box) merges.push({ range, box });
    }

    const rules: { rule: SheetValidationDto; box: MergeBox }[] = [];
    for (const [range, rule] of Object.entries(sheet?.validations ?? {})) {
        // A single-cell key is a legal range, widened as validationRangeAt does.
        const box = parseRange(range.includes(':') ? range : range + ':' + range);
        if (null !== box) rules.push({ rule, box });
    }

    const covering = new Map<string, string | null>();
    const validating = new Map<string, SheetValidationDto | null>();

    const inside = (box: MergeBox, row: number, column: number): boolean =>
        row >= box.top && row <= box.bottom && column >= box.left && column <= box.right;

    return {
        mergeCovering(ref: string): string | null {
            const known = covering.get(ref);
            if (undefined !== known) return known;

            const cell = parseRef(ref);
            let found: string | null = null;
            if (null !== cell) {
                const column = columnToIndex(cell.column);
                for (const merge of merges) {
                    if (inside(merge.box, cell.row, column)) {
                        found = merge.range;
                        break;
                    }
                }
            }
            covering.set(ref, found);

            return found;
        },

        validationAt(ref: string): SheetValidationDto | null {
            const known = validating.get(ref);
            if (undefined !== known) return known;

            const cell = parseRef(ref);
            let found: SheetValidationDto | null = null;
            if (null !== cell) {
                const column = columnToIndex(cell.column);
                for (const entry of rules) {
                    if (inside(entry.box, cell.row, column)) {
                        found = entry.rule;
                        break;
                    }
                }
            }
            validating.set(ref, found);

            return found;
        },
    };
}

/** How many grid cells a merge's anchor must span. */
export function mergeSpan(range: string): { colspan: number; rowspan: number } {
    const box = parseRange(range);
    if (!box) return { colspan: 1, rowspan: 1 };

    return { colspan: box.right - box.left + 1, rowspan: box.bottom - box.top + 1 };
}

/** Two ranges sharing at least one cell. */
function overlaps(a: MergeBox, b: MergeBox): boolean {
    return a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;
}

/**
 * Add a merge, and CLEAR the cells it covers.
 *
 * Clearing is deliberate and is the honest choice here. A merged range keeps
 * only its top-left value — that is what `SheetDocumentWriter` does when it
 * renders, because PhpSpreadsheet's `mergeCells()` empties the rest — so a
 * document that quietly held values under a merge would render differently from
 * what the grid shows, which is the exact divergence [#1998] existed to close.
 * Excel behaves the same way and warns before it does.
 *
 * Any merge OVERLAPPING the new one is dropped: two merges sharing a cell is
 * not a state Excel will open, and silently keeping both would produce a
 * workbook that prompts for repair.
 *
 * A single-cell "range" is not a merge and is refused, so an accidental
 * click-without-drag cannot write `A1:A1` into the document.
 */
export function withMerge(sheet: SheetDto, range: string): SheetDto {
    const box = parseRange(range);
    if (!box || (box.top === box.bottom && box.left === box.right)) return sheet;

    const kept = (sheet.merges ?? []).filter(existing => {
        const other = parseRange(existing);

        return !other || !overlaps(box, other);
    });

    const anchor = indexToColumn(box.left) + box.top;
    const cells = { ...sheet.cells };
    for (const ref of refsInRange(range)) {
        if (ref !== anchor) delete cells[ref];
    }

    return { ...sheet, cells, merges: [...kept, range] };
}

/** Remove a merge. Cleared cells do NOT come back — there is nothing to restore. */
export function withoutMerge(sheet: SheetDto, range: string): SheetDto {
    const merges = (sheet.merges ?? []).filter(existing => existing !== range);

    const next: SheetDto = { ...sheet };
    if (merges.length > 0) {
        next.merges = merges;
    } else {
        delete next.merges;
    }

    return next;
}

/** The range this sheet's filter covers, or null when it declares none. */
export function autoFilterOf(sheet: SheetDto): string | null {
    const range = sheet.autoFilter;

    return undefined !== range && null !== parseRange(range) ? range : null;
}

/**
 * Declare a filter over a range.
 *
 * A single cell is refused for the same reason {@link withMerge} refuses one: a
 * click without a drag would otherwise write `A1:A1` into the document, and a
 * filter over one cell is not a thing an author meant to ask for.
 *
 * The TOP row of the range is its header — that is Excel's rule, not ours, and
 * it is why a range must be given with the header included. A filter declared
 * over the body alone puts dropdowns on the first line of data.
 */
export function withAutoFilter(sheet: SheetDto, range: string): SheetDto {
    const box = parseRange(range);
    if (!box || (box.top === box.bottom && box.left === box.right)) return sheet;

    return { ...sheet, autoFilter: range.toUpperCase() };
}

/** Remove the filter, dropping the key rather than leaving an empty string. */
export function withoutAutoFilter(sheet: SheetDto): SheetDto {
    const next: SheetDto = { ...sheet };
    delete next.autoFilter;

    return next;
}

/** True when `ref` is one of the filter's HEADER cells — the ones that get a button. */
export function isFilterHeader(sheet: SheetDto, ref: string): boolean {
    const range = autoFilterOf(sheet);
    const box = range ? parseRange(range) : null;
    const cell = parseRef(ref);
    if (!box || !cell) return false;
    const column = columnToIndex(cell.column);

    return cell.row === box.top && column >= box.left && column <= box.right;
}

/** The columns a filter covers, left to right. */
export function filterColumns(range: string): string[] {
    const box = parseRange(range);
    if (!box) return [];

    const out: string[] = [];
    for (let col = box.left; col <= box.right; col++) out.push(indexToColumn(col));

    return out;
}

/**
 * The rows a filter's dropdowns act on: everything under the header.
 *
 * Empty when the range is one row tall, which is a filter whose table has no
 * body yet — the ordinary state of a template before a `{loop:}` band expands.
 */
export function filterBodyRows(range: string): number[] {
    const box = parseRange(range);
    if (!box) return [];

    const out: number[] = [];
    for (let row = box.top + 1; row <= box.bottom; row++) out.push(row);

    return out;
}

// ---- Inserting and deleting rows and columns --------------------------------
//
// ## The arithmetic mirrors the backend's `RowExpansionMap`, deliberately
//
// A `{loop:}` expansion and an inserted row are the same question — where does
// everything below end up, and what does that do to the references pointing at
// it. The backend keeps those rules in one class precisely so its two carriers
// cannot drift; this is the third carrier, and it follows the same rules for
// the same reason. A total that sums the wrong range still looks like a number.
//
// ## Where this DIFFERS from a loop expansion, and why it is simpler
//
// An expansion asks a range's two ends different questions — a start pins the
// table, an end grows with it. An insert does not: every reference at or below
// the line moves by exactly one, both ends alike. So each A1 reference is
// shifted independently and there is no start/end asymmetry to get wrong.
//
// ## A deleted row takes its references with it
//
// Excel answers `#REF!` for a formula pointing at a row that no longer exists,
// and so does this. Shifting such a reference to whatever moved INTO that row
// would be silently wrong: the formula would keep computing, on the wrong cell.

/** What a reference to something deleted becomes — Excel's own answer. */
/** The marker that pins part of a reference against a move. */
const ABSOLUTE = '$';

export const REF_ERROR = '#REF!';

/**
 * A RANGE or a lone A1 reference, in ONE alternation.
 *
 * They must be matched together, and the backend's `RowExpansionMap` says the
 * same for its own reason. Here the reason is delete: matching only single
 * refs treats `SUM(B2:B3)` as two independent references, so deleting row 2
 * turns the START into `#REF!` and yields `SUM(#REF!:B2)`. Excel SHRINKS such
 * a range — `SUM(B2:B2)` — and reserves `#REF!` for a range deleted entirely.
 *
 * The trailing lookahead keeps `LOG10(` from reading as column LOG, row 10.
 */
const A1_REFERENCE = /(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)|\$?[A-Z]{1,3}\$?\d+(?![\d(])/g;

/** Apply a rewrite to the parts of a formula that are NOT inside quotes. */
function outsideQuotes(formula: string, rewrite: (chunk: string) => string): string {
    return formula
        .split(/("(?:[^"]|"")*")/)
        .map(piece => (piece.startsWith('"') ? piece : rewrite(piece)))
        .join('');
}

/** Where an index lands, or null when the thing it named was deleted. */
function shiftIndex(index: number, at: number, delta: number): number | null {
    if (delta > 0) return index >= at ? index + delta : index;
    if (index === at) return null;

    return index > at ? index + delta : index;
}

/**
 * Move one A1 reference across an insert or delete.
 *
 * `null` means the row or column it named is gone.
 */
export function shiftRef(ref: string, axis: 'row' | 'column', at: number, delta: number): string | null {
    return mapRef(ref, axis, index => shiftIndex(index, at, delta));
}

/** Rebuild one reference with a different index on `axis`, keeping its `$` markers. */
function mapRef(ref: string, axis: 'row' | 'column', map: (index: number) => number | null): string | null {
    const match = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(ref.toUpperCase());
    if (!match) return ref;
    const [, colAbs, letters, rowAbs, digits] = match;

    if (axis === 'row') {
        const row = map(Number(digits));

        return null === row ? null : `${colAbs}${letters}${rowAbs}${row}`;
    }

    const column = map(columnToIndex(letters));

    return null === column ? null : `${colAbs}${indexToColumn(column)}${rowAbs}${digits}`;
}

/** One reference's position on `axis`, or null when it is not a reference. */
function axisIndex(ref: string, axis: 'row' | 'column'): number | null {
    const match = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref.toUpperCase());
    if (!match) return null;

    return axis === 'row' ? Number(match[2]) : columnToIndex(match[1]);
}

/**
 * A range INSIDE a formula, whose ends follow the range rules rather than the
 * single-reference ones.
 *
 * On delete a start sitting ON the deleted line stays put — whatever moved up
 * into that line is now the range's first member — while the end shrinks. Only
 * when the end falls BELOW the start has the whole range gone, and that is the
 * one case Excel answers `#REF!` for.
 */
function shiftFormulaRange(
    from: string, to: string, axis: 'row' | 'column', at: number, delta: number,
): string | null {
    const first = axisIndex(from, axis);
    const last = axisIndex(to, axis);
    if (null === first || null === last) return `${from}:${to}`;

    const start = delta > 0 ? (first >= at ? first + delta : first) : (first > at ? first + delta : first);
    const end = delta > 0 ? (last >= at ? last + delta : last) : (last >= at ? last + delta : last);
    if (end < start) return null;

    const shiftedFrom = mapRef(from, axis, () => start);
    const shiftedTo = mapRef(to, axis, () => end);

    return null === shiftedFrom || null === shiftedTo ? null : `${shiftedFrom}:${shiftedTo}`;
}

/** Rewrite every reference in a formula. A reference to a deleted line becomes `#REF!`. */
export function shiftFormula(formula: string, axis: 'row' | 'column', at: number, delta: number): string {
    return outsideQuotes(formula, chunk =>
        chunk.replace(A1_REFERENCE, (whole: string, from?: string, to?: string) =>
            (undefined !== from && undefined !== to
                ? shiftFormulaRange(from, to, axis, at, delta)
                : shiftRef(whole, axis, at, delta)) ?? REF_ERROR));
}

/**
 * Move a RANGE across an insert or delete.
 *
 * A range whose lines are entirely deleted returns null; one that is partly
 * deleted SHRINKS, which is what Excel does and what keeps a merge or a filter
 * describing the rows that are still there.
 */
export function shiftRange(range: string, axis: 'row' | 'column', at: number, delta: number): string | null {
    const single = !range.includes(':');
    const box = parseRange(single ? `${range}:${range}` : range);
    if (!box) return range;

    const isRow = axis === 'row';
    let first = isRow ? box.top : box.left;
    let last = isRow ? box.bottom : box.right;

    if (delta > 0) {
        // An insert INSIDE a range grows it; one at its start pushes it down.
        if (first >= at) first += delta;
        if (last >= at) last += delta;
    } else {
        if (first > at) first += delta;
        if (last >= at) last += delta;
        if (last < first) return null;
    }

    const top = isRow ? first : box.top;
    const bottom = isRow ? last : box.bottom;
    const left = isRow ? box.left : first;
    const right = isRow ? box.right : last;

    const from = indexToColumn(left) + top;
    const to = indexToColumn(right) + bottom;

    return single && from === to ? from : `${from}:${to}`;
}

/** Rebuild every part of a sheet across one insert or delete. */
function shiftSheet(sheet: SheetDto, axis: 'row' | 'column', at: number, delta: number): SheetDto {
    const isRow = axis === 'row';

    const cells: Record<string, SheetCellDto> = {};
    for (const [ref, cell] of Object.entries(sheet.cells)) {
        const moved = shiftRef(ref, axis, at, delta);
        // The cell itself was on the deleted line — it goes with it.
        if (null === moved) continue;
        cells[moved] = undefined === cell.formula
            ? cell
            : { ...cell, formula: shiftFormula(cell.formula, axis, at, delta) };
    }

    const next: SheetDto = { ...sheet, cells };

    // Sizes are keyed by the line itself, so they simply move or vanish.
    if (isRow && sheet.rowHeights) {
        const heights: Record<string, number> = {};
        for (const [row, px] of Object.entries(sheet.rowHeights)) {
            const moved = shiftIndex(Number(row), at, delta);
            if (null !== moved) heights[String(moved)] = px;
        }
        if (Object.keys(heights).length > 0) next.rowHeights = heights; else delete next.rowHeights;
    }
    if (!isRow && sheet.columnWidths) {
        const widths: Record<string, number> = {};
        for (const [letter, width] of Object.entries(sheet.columnWidths)) {
            const moved = shiftIndex(columnToIndex(letter), at, delta);
            if (null !== moved) widths[indexToColumn(moved)] = width;
        }
        if (Object.keys(widths).length > 0) next.columnWidths = widths; else delete next.columnWidths;
    }

    if (sheet.merges) {
        const merges = sheet.merges
            .map(range => shiftRange(range, axis, at, delta))
            // A merge reduced to ONE cell is no longer a merge. Tested on the
            // ends rather than on the presence of a colon: `A1:B1` losing
            // column B comes back as the range `A1:A1`, which has a colon and
            // is still not a merge.
            .filter((range): range is string => {
                const box = null === range ? null : parseRange(range);

                return !!box && !(box.top === box.bottom && box.left === box.right);
            });
        if (merges.length > 0) next.merges = merges; else delete next.merges;
    }

    if (sheet.autoFilter) {
        const moved = shiftRange(sheet.autoFilter, axis, at, delta);
        if (null !== moved) next.autoFilter = moved; else delete next.autoFilter;
    }

    if (sheet.validations) {
        const validations: Record<string, SheetValidationDto> = {};
        for (const [range, rule] of Object.entries(sheet.validations)) {
            const moved = shiftRange(range, axis, at, delta);
            if (null !== moved) validations[moved] = rule;
        }
        if (Object.keys(validations).length > 0) next.validations = validations; else delete next.validations;
    }

    return next;
}

/** Insert a blank row, pushing `row` and everything under it down. */
export function withInsertedRow(sheet: SheetDto, row: number): SheetDto {
    return row < 1 ? sheet : shiftSheet(sheet, 'row', row, 1);
}

/** Delete a row, pulling everything under it up. */
export function withDeletedRow(sheet: SheetDto, row: number): SheetDto {
    return row < 1 ? sheet : shiftSheet(sheet, 'row', row, -1);
}

/** Insert a blank column, pushing `column` and everything right of it along. */
export function withInsertedColumn(sheet: SheetDto, column: string): SheetDto {
    const index = columnToIndex(column.toUpperCase());

    return index < 1 ? sheet : shiftSheet(sheet, 'column', index, 1);
}

/** Delete a column, pulling everything right of it back. */
export function withDeletedColumn(sheet: SheetDto, column: string): SheetDto {
    const index = columnToIndex(column.toUpperCase());

    return index < 1 ? sheet : shiftSheet(sheet, 'column', index, -1);
}

/**
 * Empty every cell in a range, leaving the layout alone.
 *
 * Contents only: merges, sizes, filters and form elements describe the SHAPE
 * of the sheet, and an author clearing three cells has not asked to dismantle
 * the table around them.
 */
export function withClearedRange(sheet: SheetDto, range: string): SheetDto {
    const cells = { ...sheet.cells };
    for (const ref of refsInRange(range.includes(':') ? range : `${range}:${range}`)) {
        delete cells[ref];
    }

    return { ...sheet, cells };
}

/** One edge's CSS shorthand, or null when that edge is not ruled. */
export function borderCss(cell: SheetCellDto | undefined, edge: CellEdge): string | null {
    const spec = cell?.borders?.[edge];
    if (undefined === spec) return null;

    const [style, colour] = spec.split(' ');
    const width = BORDER_STYLES[style];

    return undefined === width ? null : `${width} ${colour ?? DEFAULT_BORDER_COLOUR}`;
}

/** What a border IS unless it is told otherwise -- see {@link withBorderPreset}. */
export const DEFAULT_BORDER_COLOUR = '#000000';

/** The edge on the far side of a shared line. */
const OPPOSITE_EDGE: Readonly<Record<CellEdge, CellEdge>> = {
    top: 'bottom',
    right: 'left',
    bottom: 'top',
    left: 'right',
};

/** The cell one step over, or null where the grid ends. */
function neighbourRef(ref: string, edge: CellEdge): string | null {
    const cell = parseRef(ref);
    if (!cell) return null;

    const column = columnToIndex(cell.column) + ('left' === edge ? -1 : 'right' === edge ? 1 : 0);
    const row = cell.row + ('top' === edge ? -1 : 'bottom' === edge ? 1 : 0);

    return column < 1 || row < 1 ? null : indexToColumn(column) + row;
}

/**
 * One cell's rendered borders, as the CSS the grid binds.
 *
 * An edge the cell does not rule is taken from the NEIGHBOUR across it, and
 * that is not a nicety. The grid is `border-collapse: collapse`, so the line
 * between two cells is ONE line and the browser picks a winner: at equal width
 * and style, the cell further up or further left. Every cell already carries a
 * grey gridline, so a black `top` lost to the grey `bottom` of the cell above
 * and a black `left` lost to the grey `right` of the cell beside -- "All
 * borders" on one cell drew its right and bottom ONLY. Declaring the same line
 * on both sides settles the conflict the same way whichever side wins, and it
 * is also how a spreadsheet thinks of it: the line between two cells belongs
 * to the boundary, not to one of them.
 */
export function borderCssAt(sheet: SheetDto, ref: string): Record<string, string> {
    const cell = sheet.cells[ref];
    const out: Record<string, string> = {};

    for (const edge of CELL_EDGES) {
        const across = neighbourRef(ref, edge);
        const css = borderCss(cell, edge)
            ?? (null === across ? null : borderCss(sheet.cells[across], OPPOSITE_EDGE[edge]));

        // An author's line REPLACES the grid rule on that edge rather than
        // sitting inside it, which is what makes a thin black border look like
        // a border and not like a slightly darker gridline.
        if (null !== css) out[`border-${edge}`] = css;
    }

    return out;
}

/**
 * Which edges of ONE cell a gesture rules, given where the cell sits in the
 * selection.
 *
 * Shared by the gesture below and by the readout after it, so the two cannot
 * drift: a control that reports a state its own gesture would not produce is
 * worse than one that reports nothing at all.
 */
function edgesForPreset(preset: BorderPreset, box: MergeBox, row: number, col: number): CellEdge[] {
    if ('none' === preset) return [];
    if ('all' === preset) return [...CELL_EDGES];

    const edges: CellEdge[] = [];
    // Only the cells ON each side of the selection carry that side.
    if (('outer' === preset || 'top' === preset) && row === box.top) edges.push('top');
    if (('outer' === preset || 'right' === preset) && col === box.right) edges.push('right');
    if (('outer' === preset || 'bottom' === preset) && row === box.bottom) edges.push('bottom');
    if (('outer' === preset || 'left' === preset) && col === box.left) edges.push('left');

    return edges;
}

/**
 * Apply a border gesture to a RANGE.
 *
 * The presets are Google Sheets' vocabulary because that is what an author
 * already means by them: `top` rules the top of the SELECTION, not the top of
 * every cell in it -- the difference between underlining a table and
 * underlining each of its rows, and getting it backwards is the kind of thing
 * nobody reports as a bug, they just stop using the control.
 *
 * `all` is the exception and is per-cell by definition. `none` clears every
 * edge in the range, including the perimeter another gesture put there.
 */
export function withBorderPreset(
    sheet: SheetDto,
    range: string,
    preset: BorderPreset,
    style = 'thin',
    colour: string | null = null,
): SheetDto {
    const box = parseRange(range.includes(':') ? range : `${range}:${range}`);
    if (!box) return sheet;

    // A bare style MEANS black to the writer and to the importer both, so the
    // default colour is left unwritten rather than stamped onto every edge.
    const spec = null === colour || DEFAULT_BORDER_COLOUR === colour.toLowerCase()
        ? style
        : `${style} ${colour.toUpperCase()}`;

    const cells = { ...sheet.cells };

    for (let row = box.top; row <= box.bottom; row++) {
        for (let col = box.left; col <= box.right; col++) {
            const ref = indexToColumn(col) + row;
            const current = cells[ref] ?? {};
            const borders = { ...(current.borders ?? {}) };

            if ('none' === preset) {
                CELL_EDGES.forEach(edge => delete borders[edge]);
            } else {
                edgesForPreset(preset, box, row, col).forEach(edge => (borders[edge] = spec));
            }

            const next: SheetCellDto = { ...current };
            if (Object.keys(borders).length > 0) {
                next.borders = borders;
            } else {
                delete next.borders;
            }

            // A cell that now holds NOTHING -- no text, no style, no edge -- is
            // dropped rather than written as `{}`: clearing borders off empty
            // cells must not leave the document littered with them.
            if (Object.keys(next).length > 0) {
                cells[ref] = next;
            } else {
                delete cells[ref];
            }
        }
    }

    return { ...sheet, cells };
}

/** What a selection's borders currently are, in the control's own vocabulary. */
export interface BorderState {
    /** The gesture that would reproduce what is there, or null when none does. */
    readonly preset: BorderPreset | null;
    /** The line style every ruled edge shares, or null when they differ. */
    readonly style: string | null;
    /** The colour every ruled edge shares, or null when they differ. */
    readonly colour: string | null;
}

/**
 * `all` before `outer` because on a SINGLE cell the two gestures produce the
 * same four edges, and `all` is the one the author reached for.
 */
const BORDER_PRESET_ORDER: readonly BorderPreset[] = ['all', 'outer', 'top', 'right', 'bottom', 'left'];

/**
 * Read a selection's borders back.
 *
 * The control has to show what the cells ARE, not merely offer what they could
 * become: a cell with every edge ruled that reports nothing tells the author
 * their last gesture did not land. Reads the cells' OWN edges -- the mirrored
 * ones {@link borderCssAt} renders belong to the neighbour, and reporting them
 * would claim a border this selection does not hold.
 */
export function borderStateIn(sheet: SheetDto, range: string): BorderState {
    const nothing: BorderState = { preset: null, style: null, colour: null };
    const box = parseRange(range.includes(':') ? range : `${range}:${range}`);
    if (!box) return nothing;

    const actual = new Map<string, Set<CellEdge>>();
    const specs = new Set<string>();

    for (let row = box.top; row <= box.bottom; row++) {
        for (let col = box.left; col <= box.right; col++) {
            const ref = indexToColumn(col) + row;
            const borders = sheet.cells[ref]?.borders ?? {};
            const edges = new Set<CellEdge>();

            for (const edge of CELL_EDGES) {
                const spec = borders[edge];
                if (undefined === spec) continue;
                edges.add(edge);
                specs.add(spec);
            }

            actual.set(ref, edges);
        }
    }

    if (0 === specs.size) return nothing;

    const preset = BORDER_PRESET_ORDER.find(candidate => {
        for (let row = box.top; row <= box.bottom; row++) {
            for (let col = box.left; col <= box.right; col++) {
                const want = edgesForPreset(candidate, box, row, col);
                const have = actual.get(indexToColumn(col) + row);
                if (undefined === have) return false;
                if (want.length !== have.size || !want.every(edge => have.has(edge))) return false;
            }
        }

        return true;
    }) ?? null;

    // Mixed lines have no single answer, and guessing one would let the next
    // gesture quietly rewrite the edges it did not describe.
    if (1 !== specs.size) return { preset, style: null, colour: null };

    // A bare style means black, exactly as the renderer and the writer read it.
    const [style, colour = DEFAULT_BORDER_COLOUR] = [...specs][0].split(' ');

    return { preset, style, colour };
}

/**
 * Move every RELATIVE reference in a formula by a row and column offset.
 *
 * What a PASTE does, and the reason copying a total down a column is the
 * commonest thing anybody does in a spreadsheet: `=SUM(B2:B4)` pasted one
 * column right has to become `=SUM(C2:C4)` or it is not a copy of anything.
 *
 * `$` is what makes a part of a reference absolute, and an absolute part does
 * NOT move -- that is the entire point of writing one. A reference pushed off
 * the top or the left of the sheet becomes `#REF!`, as Excel's does, rather
 * than wrapping round to a cell nobody meant.
 */
export function offsetFormula(formula: string, rows: number, columns: number): string {
    const move = (ref: string): string => {
        const match = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(ref.toUpperCase());
        if (!match) return ref;
        const [, colAbs, letters, rowAbs, digits] = match;

        const column = ABSOLUTE === colAbs ? columnToIndex(letters) : columnToIndex(letters) + columns;
        const row = ABSOLUTE === rowAbs ? Number(digits) : Number(digits) + rows;
        if (column < 1 || row < 1) return REF_ERROR;

        return `${colAbs}${indexToColumn(column)}${rowAbs}${row}`;
    };

    return outsideQuotes(formula, chunk =>
        chunk.replace(A1_REFERENCE, (whole: string, from?: string, to?: string) =>
            (undefined !== from && undefined !== to ? `${move(from)}:${move(to)}` : move(whole))));
}

/**
 * Write a rectangle of text into the sheet, its top-left corner at `at`.
 *
 * Each cell goes through {@link inputToCell}, so a pasted `=A1+1` is a formula,
 * a pasted date lands in a date-formatted cell as a serial, and the formatting
 * already on the target cells survives -- a paste replaces CONTENT, not the
 * look of the table it is pasted into.
 *
 * `offset` moves the references in any pasted formula; null for text that came
 * from OUTSIDE, where a formula is just the characters somebody typed and has
 * no origin to be relative to.
 */
export function withPastedBlock(
    sheet: SheetDto,
    at: string,
    block: readonly (readonly string[])[],
    offset: { rows: number; columns: number } | null = null,
): SheetDto {
    const corner = parseRef(at);
    if (!corner) return sheet;

    const cells = { ...sheet.cells };
    const left = columnToIndex(corner.column);

    block.forEach((line, rowOffset) => {
        line.forEach((raw, columnOffset) => {
            const ref = indexToColumn(left + columnOffset) + (corner.row + rowOffset);
            const text = null !== offset && raw.startsWith('=')
                ? '=' + offsetFormula(raw.slice(1), offset.rows, offset.columns)
                : raw;
            const next = inputToCell(text, cells[ref]);
            if (null === next) {
                delete cells[ref];
            } else {
                cells[ref] = next;
            }
        });
    });

    return { ...sheet, cells };
}

/** How many rows and columns a sheet's freeze holds in place. */
export function frozenAt(sheet: SheetDto | undefined): { rows: number; columns: number } {
    const cell = sheet?.freeze === undefined ? null : parseRef(sheet.freeze);

    return null === cell
        ? { rows: 0, columns: 0 }
        : { rows: cell.row - 1, columns: columnToIndex(cell.column) - 1 };
}

/**
 * Freeze above and left of a cell, or clear it.
 *
 * `A1` clears, because freezing above and left of the first cell freezes
 * nothing -- the same answer the backend and OOXML give, so a round trip
 * through any of the three says the same thing.
 */
export function withFreeze(sheet: SheetDto, at: string | null): SheetDto {
    const next = { ...sheet };
    const cell = null === at ? null : parseRef(at);

    if (null === cell || (1 === cell.row && 1 === columnToIndex(cell.column))) {
        delete next.freeze;
    } else {
        next.freeze = cell.column + cell.row;
    }

    return next;
}

/** The range whose rule covers `ref`, or null when nothing does. */
export function validationRangeAt(sheet: SheetDto, ref: string): string | null {
    for (const range of Object.keys(sheet.validations ?? {})) {
        // A single-cell key is a legal range here and `rangeContains` wants a
        // pair, so it is widened to one whose ends coincide.
        if (rangeContains(range.includes(':') ? range : `${range}:${range}`, ref)) return range;
    }

    return null;
}

/** The rule covering `ref`, or null. */
export function validationAt(sheet: SheetDto, ref: string): SheetValidationDto | null {
    const range = validationRangeAt(sheet, ref);

    return null === range ? (null) : (sheet.validations?.[range] ?? null);
}

/**
 * The options a rule offers, or empty when it is not a list-shaped one.
 *
 * A checkbox's options are its TYPE and are never read from the document —
 * {@link CHECKBOX_VALUES} says why.
 */
export function validationOptions(rule: SheetValidationDto): string[] {
    if (rule.type === 'checkbox') return [...CHECKBOX_VALUES];

    return rule.type === 'list' ? [...(rule.values ?? [])] : [];
}

/** Whether a rule draws as a control in the grid — the two list-shaped types. */
export function isControl(rule: SheetValidationDto): boolean {
    return rule.type === 'list' || rule.type === 'checkbox';
}

/**
 * Attach a rule to a range, replacing any it OVERLAPS.
 *
 * Overlap is resolved rather than allowed for the reason {@link withMerge}
 * gives about merges: two rules covering one cell is a state whose behaviour
 * depends on which the reader happens to apply, and Excel keeps only one. The
 * author sees the rule they just made, which is the one they were thinking
 * about.
 *
 * A `list` with no options is refused — that is not a dropdown, it is a cell
 * nobody can type into, and the backend refuses the same shape.
 */
export function withValidation(sheet: SheetDto, range: string, rule: SheetValidationDto): SheetDto {
    const box = parseRange(range.includes(':') ? range : `${range}:${range}`);
    if (!box) return sheet;
    if (rule.type === 'list' && (rule.values ?? []).length === 0) return sheet;

    const validations: Record<string, SheetValidationDto> = {};
    for (const [existing, kept] of Object.entries(sheet.validations ?? {})) {
        const other = parseRange(existing.includes(':') ? existing : `${existing}:${existing}`);
        if (!other || !overlaps(box, other)) validations[existing] = kept;
    }
    validations[range.toUpperCase()] = rule;

    return { ...sheet, validations };
}

/** Remove a rule, dropping the key entirely with the last one. */
export function withoutValidation(sheet: SheetDto, range: string): SheetDto {
    const validations = { ...(sheet.validations ?? {}) };
    delete validations[range];

    const next: SheetDto = { ...sheet };
    if (Object.keys(validations).length > 0) {
        next.validations = validations;
    } else {
        delete next.validations;
    }

    return next;
}

/**
 * The stored width of a column, if the author set one.
 *
 * The unit is OOXML's, not pixels: a column's width is measured in CHARACTERS
 * of the workbook's default font. It is stored exactly as the backend's
 * `SheetDocumentWriter` passes it to `setWidth()`, so nothing here converts on
 * the way in or out — see {@link columnWidthToPx} for the display side.
 */
export function columnWidthOf(sheet: SheetDto, column: string): number | undefined {
    return sheet.columnWidths?.[column.toUpperCase()];
}

/**
 * Set or clear a column's width, keeping the document sparse.
 *
 * `undefined` clears, and clearing the LAST width drops the `columnWidths` key
 * entirely rather than leaving `{}` behind — the backend writes the object only
 * when it has entries, and a `.dsheet` is a source file an author may read.
 *
 * A non-finite or non-positive width is refused rather than stored: Excel reads
 * width 0 as a HIDDEN column, so letting a stray `0` through would make a column
 * vanish from the generated workbook with nothing in the editor to explain it.
 */
export function withColumnWidth(sheet: SheetDto, column: string, width: number | undefined): SheetDto {
    const key = column.toUpperCase();
    const widths = { ...(sheet.columnWidths ?? {}) };

    if (width === undefined) {
        delete widths[key];
    } else {
        if (!Number.isFinite(width) || width <= 0) return sheet;
        widths[key] = width;
    }

    const next: SheetDto = { ...sheet };
    if (Object.keys(widths).length > 0) {
        next.columnWidths = widths;
    } else {
        delete next.columnWidths;
    }

    return next;
}

/**
 * Approximate pixels for a stored character width, so the grid SHOWS what the
 * author set.
 *
 * Excel's own conversion is `px = round(width * maxDigitWidth) + padding`, where
 * `maxDigitWidth` depends on the workbook font — 7px for the 11pt Calibri that
 * is the usual default, with 5px of cell padding. Both are reproduced here as
 * named constants rather than a magic `w * 7 + 5`.
 *
 * This is DISPLAY ONLY and deliberately approximate: the editor renders in the
 * browser's font, not the workbook's, so the two cannot agree exactly. What must
 * be exact is the stored number, which is why nothing converts on save.
 */
const PX_PER_CHARACTER = 7;
const CELL_PADDING_PX = 5;

export function columnWidthToPx(width: number): number {
    return Math.round(width * PX_PER_CHARACTER) + CELL_PADDING_PX;
}

/**
 * The floor a drag may reach. Excel reads a width of 0 as a HIDDEN column, so a
 * drag that ran to the left edge would not make a thin column — it would make
 * one that does not appear in the generated workbook at all, with nothing on
 * screen to say so.
 */
export const MIN_COLUMN_WIDTH = 0.5;

/**
 * Pixels back to a stored character width — the inverse of
 * {@link columnWidthToPx}, for the drag-resize handle on the column headers.
 *
 * Rounded to two decimals because the drag produces a new value on every
 * mousemove and the raw division yields things like `12.714285714285714`, which
 * would be written into the document and shown in the toolbar's width box. The
 * conversion is approximate in both directions anyway (see above); what matters
 * is that the number an author ends up with is one they could have typed.
 */
export function columnWidthFromPx(px: number): number {
    const width = (px - CELL_PADDING_PX) / PX_PER_CHARACTER;

    return Math.max(MIN_COLUMN_WIDTH, Math.round(width * 100) / 100);
}

/**
 * Parse `.dsheet` bytes, falling back to an empty document.
 *
 * A file the editor cannot understand must NOT be silently replaced with a
 * blank one — the caller is told, and decides. Returning the fallback plus a
 * flag rather than throwing keeps the dialog openable on a damaged file, which
 * is the only state from which an operator can repair it.
 */
export function parseSheetDocument(content: string): { doc: SheetDocumentDto; ok: boolean } {
    if (content.trim() === '') {
        return { doc: structuredClone(EMPTY_SHEET_DOCUMENT), ok: true };
    }

    try {
        const parsed: unknown = JSON.parse(content);
        if (
            typeof parsed !== 'object' || parsed === null ||
            typeof (parsed as SheetDocumentDto).sheets !== 'object' ||
            (parsed as SheetDocumentDto).sheets === null
        ) {
            return { doc: structuredClone(EMPTY_SHEET_DOCUMENT), ok: false };
        }

        const doc = parsed as SheetDocumentDto;

        // `sheets` is a map keyed by name and can arrive as an ARRAY for the same
        // reason `cells` could: PHP coerces numeric string keys, so a file whose
        // sheets were named "0" and "1" encoded as `"sheets": [ … ]`. `Object.keys`
        // still yields "0"/"1" so the tabs LOOK right, but adding a sheet then sets
        // a string key on an array and `JSON.stringify` drops it on save — the
        // #2002 loss again, one level up. Rebuild it as a plain object, which keeps
        // the positional names rather than discarding them.
        if (Array.isArray(doc.sheets)) {
            doc.sheets = { ...(doc.sheets as unknown as Record<string, SheetDto>) };
        }
        if (Object.keys(doc.sheets).length === 0) {
            return { doc: structuredClone(EMPTY_SHEET_DOCUMENT), ok: false };
        }

        for (const sheet of Object.values(doc.sheets)) {
            // `cells` must be a plain OBJECT, and `??=` is not enough to make it
            // one. A template minted by the backend carried `"cells": []` — PHP
            // cannot distinguish an empty map from an empty list — and `[]` is
            // neither null nor undefined, so the old `??=` left it as a JS ARRAY.
            // `cells['A1'] = …` then sets a STRING KEY on an array, which
            // `JSON.stringify` DISCARDS: every cell typed into a brand-new
            // native template was silently lost on save, behind a green "Saved"
            // toast, with the stored blob coming back byte-identical.
            //
            // Normalising on READ is what fixes the templates already on disk;
            // the backend now emits `{}` so new ones never carry the array.
            // Nothing is lost by discarding an array here — a JSON array cannot
            // hold the A1 keys this format is made of.
            if (typeof sheet.cells !== 'object' || sheet.cells === null || Array.isArray(sheet.cells)) {
                sheet.cells = {};
            }
        }
        doc.version ??= 1;

        return { doc, ok: true };
    } catch {
        return { doc: structuredClone(EMPTY_SHEET_DOCUMENT), ok: false };
    }
}

/** Serialise for the VFS write — indented, because a `.dsheet` is a SOURCE file. */
export function serialiseSheetDocument(doc: SheetDocumentDto): string {
    return JSON.stringify(doc, null, 4);
}
