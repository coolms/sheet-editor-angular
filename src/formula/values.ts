/**
 * What a formula can evaluate TO, and how those values coerce.
 *
 * Six outcomes, and the sixth is the one a spreadsheet library would not have.
 *
 * **`unresolved` is not an error.** This grid edits a TEMPLATE: a cell may hold
 * a DTMPL token like `{var:order.total}`, whose value is not known until the
 * backend fills the document. A formula reaching such a cell has no answer yet
 * — and answering `0` would be a lie, while answering `#VALUE!` would report a
 * mistake the author has not made. So it propagates as its own outcome, and the
 * grid can say "depends on the data" rather than showing a number nobody should
 * trust.
 *
 * `blank` is likewise distinct from `0` and from `""`: an empty cell is empty,
 * and only coercion decides what that means in context. `COUNT` must not count
 * it; `SUM` must treat it as nothing rather than as zero-the-number, which
 * happens to be the same result but for a reason worth keeping straight.
 */

/** The error codes this engine emits, spelled as a spreadsheet spells them. */
export type FormulaError =
    | '#DIV/0!'
    | '#VALUE!'
    | '#REF!'
    | '#NAME?'
    | '#NUM!'
    | '#N/A'
    | '#CYCLE!';

export const FORMULA_ERRORS: readonly FormulaError[] = [
    '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#CYCLE!',
];

export type CellValue =
    | { readonly kind: 'number'; readonly value: number }
    | { readonly kind: 'text'; readonly value: string }
    | { readonly kind: 'boolean'; readonly value: boolean }
    | { readonly kind: 'blank' }
    | { readonly kind: 'error'; readonly code: FormulaError }
    /** Known to be unknowable here: a template token stands where data will. */
    | { readonly kind: 'unresolved'; readonly because: string };

export const BLANK: CellValue = { kind: 'blank' };

export const num = (value: number): CellValue => ({ kind: 'number', value });
export const text = (value: string): CellValue => ({ kind: 'text', value });
export const bool = (value: boolean): CellValue => ({ kind: 'boolean', value });
export const err = (code: FormulaError): CellValue => ({ kind: 'error', code });
export const unresolved = (because: string): CellValue => ({ kind: 'unresolved', because });

export const isError = (v: CellValue): boolean => v.kind === 'error';
export const isUnresolved = (v: CellValue): boolean => v.kind === 'unresolved';

/**
 * Errors and unresolved values travel OUTWARD: an operand that has no answer
 * makes the whole expression have no answer, and an error names itself all the
 * way up rather than being swallowed into a number.
 *
 * Checked before every operation, which is why it is one function: a single
 * place that decides propagation cannot disagree with itself.
 */
export function propagate(...values: readonly CellValue[]): CellValue | null {
    for (const v of values) {
        if (v.kind === 'error') return v;
    }
    // An error outranks an unresolved: a mistake is worth reporting even when
    // part of the expression is merely unknown.
    for (const v of values) {
        if (v.kind === 'unresolved') return v;
    }

    return null;
}

/**
 * A number, or the error that explains why not.
 *
 * Text coerces when it LOOKS like a number, which is what a spreadsheet does
 * with `="5"+1`. Anything else is `#VALUE!` rather than `NaN`, because `NaN`
 * spreads silently and an error does not.
 */
export function toNumber(v: CellValue): { ok: true; value: number } | { ok: false; value: CellValue } {
    switch (v.kind) {
        case 'number':
            return { ok: true, value: v.value };
        case 'boolean':
            return { ok: true, value: v.value ? 1 : 0 };
        case 'blank':
            return { ok: true, value: 0 };
        case 'text': {
            const trimmed = v.value.trim();
            if (trimmed === '') return { ok: true, value: 0 };
            const n = Number(trimmed);

            return Number.isFinite(n)
                ? { ok: true, value: n }
                : { ok: false, value: err('#VALUE!') };
        }
        default:
            return { ok: false, value: v };
    }
}

/** Text, as a spreadsheet renders a value inside `&` or `CONCAT`. */
export function toText(v: CellValue): string {
    switch (v.kind) {
        case 'text':
            return v.value;
        case 'number':
            return formatNumber(v.value);
        case 'boolean':
            return v.value ? 'TRUE' : 'FALSE';
        case 'blank':
            return '';
        case 'error':
            return v.code;
        default:
            return '';
    }
}

/**
 * Truthiness for `IF`, `AND`, `OR`.
 *
 * A number is true when non-zero; text is NOT coerced, because `IF("yes", …)`
 * is a mistake worth reporting rather than a convention worth guessing at.
 */
export function toBoolean(v: CellValue): { ok: true; value: boolean } | { ok: false; value: CellValue } {
    switch (v.kind) {
        case 'boolean':
            return { ok: true, value: v.value };
        case 'number':
            return { ok: true, value: v.value !== 0 };
        case 'blank':
            return { ok: true, value: false };
        case 'text': {
            const upper = v.value.trim().toUpperCase();
            if (upper === 'TRUE') return { ok: true, value: true };
            if (upper === 'FALSE') return { ok: true, value: false };

            return { ok: false, value: err('#VALUE!') };
        }
        default:
            return { ok: false, value: v };
    }
}

/**
 * Comparison, with a spreadsheet's ordering rather than JavaScript's.
 *
 * Text compares case-INSENSITIVELY, so `"a" = "A"` is true — which is what a
 * spreadsheet says and what `===` does not. A number never equals text; they
 * order as number < text < boolean instead of coercing, so `1 < "a"` is true
 * without `"a"` becoming `NaN`.
 *
 * Lives HERE, beside the coercions, because two callers need it: the `=` and
 * `<` operators, and the criterion in `SUMIF`/`COUNTIF` — which is a
 * comparison spelled as text. A second implementation of the same ordering is
 * a disagreement waiting to happen, and the one that would drift is the one
 * nobody reads.
 */
export function compareValues(op: string, left: CellValue, right: CellValue): CellValue {
    //  A BLANK takes the type of whatever it is compared WITH. An empty cell
    // equals 0 and it also equals "" -- both are true in a spreadsheet, and
    // ranking blank as a number made the second one false.
    if (left.kind === 'blank' && right.kind === 'text') left = text('');
    if (right.kind === 'blank' && left.kind === 'text') right = text('');

    const rank = (v: CellValue): number =>
        (v.kind === 'number' || v.kind === 'blank' ? 0 : v.kind === 'text' ? 1 : 2);

    let cmp: number;
    if (rank(left) !== rank(right)) {
        cmp = rank(left) < rank(right) ? -1 : 1;
    } else if (rank(left) === 1) {
        const a = toText(left).toUpperCase();
        const b = toText(right).toUpperCase();
        cmp = a === b ? 0 : (a < b ? -1 : 1);
    } else {
        const a = toNumber(left);
        const b = toNumber(right);
        if (!a.ok) return a.value;
        if (!b.ok) return b.value;
        cmp = a.value === b.value ? 0 : (a.value < b.value ? -1 : 1);
    }

    switch (op) {
        case '=': return bool(cmp === 0);
        case '<>': return bool(cmp !== 0);
        case '<': return bool(cmp < 0);
        case '>': return bool(cmp > 0);
        case '<=': return bool(cmp <= 0);
        case '>=': return bool(cmp >= 0);
        default: return err('#VALUE!');
    }
}

/** True when the comparison holds; false when it does not OR cannot be made. */
export function comparesTrue(op: string, left: CellValue, right: CellValue): boolean {
    const result = compareValues(op, left, right);

    return result.kind === 'boolean' && result.value;
}

/**
 * How a computed number reaches the grid.
 *
 * Rounded at the twelfth significant digit before display, because binary
 * floating point makes `0.1 + 0.2` end in `0000000000000004` and a spreadsheet
 * that shows that is reporting its arithmetic rather than the author's. The
 * stored VALUE is untouched; this is presentation only.
 */
export function formatNumber(n: number): string {
    if (!Number.isFinite(n)) return '#NUM!';
    if (Number.isInteger(n)) return String(n);

    return String(Number.parseFloat(n.toPrecision(12)));
}

/** How a finished value reaches the grid, errors and all. */
export function displayValue(v: CellValue): string {
    return v.kind === 'unresolved' ? '' : toText(v);
}
