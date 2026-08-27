/**
 * A formula, against a document, to a value.
 *
 * Two things here are not obvious from a spreadsheet's behaviour, and both come
 * from this being a TEMPLATE editor rather than a workbook:
 *
 * 1. **A DTMPL token makes a cell unresolved, not zero.** `{var:order.total}`
 *    is a placeholder the backend fills; a formula reaching it has no answer
 *    while authoring. Reporting `0` would show a total nobody should trust.
 *
 * 2. **Nothing is cached across calls.** A template is small and edited by
 *    hand, and a stale cached value is a far worse failure than a recomputed
 *    one. Cells already visited within ONE evaluation are remembered, which is
 *    what makes a diamond dependency linear instead of exponential -- but that
 *    memo dies with the call.
 *
 * Cycles are detected rather than run: a formula that reaches itself returns
 * `#CYCLE!` for every cell on the loop.
 */
import { columnToIndex, indexToColumn, parseRef, type SheetDocumentDto } from '../sheet-document.model';
import type { BinaryOperator, FormulaNode } from './ast';
import { lookupFunction, type FunctionContext } from './functions';
import { parseFormula } from './parse';
import {
    BLANK, bool, compareValues, err, num, propagate, text, toNumber, toText,
    type CellValue,
} from './values';

/** A DTMPL token: `{var:…}`, `{t:…}` and friends. Anything in braces counts. */
const DTMPL_TOKEN = /\{[a-z]+:[^}]*\}/iu;

export interface EvaluationOptions {
    /** The sheet an unqualified reference belongs to. */
    readonly sheet: string;
    /**
     * How deep a chain of references may go before it is called a cycle.
     * A guard against pathological documents, not against ordinary ones.
     */
    readonly maxDepth?: number;
}

class Evaluator implements FunctionContext {
    private readonly memo = new Map<string, CellValue>();
    private readonly visiting = new Set<string>();

    constructor(
        private readonly doc: SheetDocumentDto,
        private readonly options: EvaluationOptions,
    ) {}

    /** The value of one cell, following a formula if it holds one. */
    cellValue(sheetName: string, ref: string): CellValue {
        const key = `${sheetName}!${ref}`;
        const cached = this.memo.get(key);
        if (cached) return cached;

        if (this.visiting.has(key)) return err('#CYCLE!');
        if (this.visiting.size > (this.options.maxDepth ?? 256)) return err('#CYCLE!');

        const sheet = this.doc.sheets[sheetName];
        if (!sheet) return err('#REF!');
        const cell = sheet.cells[ref];
        if (!cell) return BLANK;

        if (cell.formula !== undefined) {
            this.visiting.add(key);
            const parsed = parseFormula(cell.formula);
            const value = parsed.ok
                ? this.node(parsed.node, sheetName)
                : err('#NAME?');
            this.visiting.delete(key);
            this.memo.set(key, value);

            return value;
        }

        const literal = literalValue(cell.value, cell.numberFormat);
        this.memo.set(key, literal);

        return literal;
    }

    /** One value from a node; a range collapses to its first cell. */
    evaluate(node: FormulaNode): CellValue {
        return this.node(node, this.options.sheet);
    }

    spread(node: FormulaNode): readonly CellValue[] {
        return this.spreadIn(node, this.options.sheet);
    }

    /**
     * A range as a flat list of values, evaluated in a named sheet.
     *
     * Takes the sheet explicitly for the same reason {@see gridIn} does: a
     * formula that crosses sheets is evaluated in the sheet the CALL sits in,
     * not the evaluator's own. One implementation, so the scoped context handed
     * to a function cannot answer differently from `spread` itself.
     */
    private spreadIn(node: FormulaNode, sheetName: string): readonly CellValue[] {
        node = this.deref(node);
        if (node.kind !== 'range') return [this.node(node, sheetName)];

        return rangeRefs(node.from, node.to)
            .map((r) => this.cellValue(node.sheet ?? sheetName, r));
    }

    grid(node: FormulaNode): readonly (readonly CellValue[])[] {
        return this.gridIn(node, this.options.sheet);
    }

    /**
     * A defined name, replaced by the range it stands for.
     *
     * Done as a NODE REWRITE rather than a third branch inside every consumer:
     * once a name has become a `range`, `spread` and `grid` need no knowledge
     * of names at all, and a name can never behave differently from the range
     * it denotes.
     *
     * ⚠️ Case-INSENSITIVE, because a spreadsheet's names are: an author who
     * declares `items_amount` and types `Items_Amount` gets the range, not an
     * error. An unknown name is `#NAME?`, which is what Excel answers and what
     * the parser used to refuse to produce at all.
     */
    private deref(node: FormulaNode): FormulaNode {
        if (node.kind !== 'name') return node;

        const declared = this.declaredRange(node.name);
        if (declared === undefined) return { kind: 'error', code: '#NAME?' };

        return parseScopedRange(declared) ?? { kind: 'error', code: '#REF!' };
    }

    private declaredRange(name: string): string | undefined {
        const names = this.doc.definedNames;
        if (!names) return undefined;

        const exact = names[name];
        if (exact !== undefined) return exact;

        const wanted = name.toUpperCase();
        for (const [declared, range] of Object.entries(names)) {
            if (declared.toUpperCase() === wanted) return range;
        }

        return undefined;
    }

    /**
     * A range as ROWS of values.
     *
     * `spread` flattens, which is all `SUM` ever needed. `VLOOKUP` and
     * `INDEX` are about position -- "the third column of this block" -- and a
     * flat list cannot answer that: 12 values are a 3x4 block or a 4x3 one and
     * the difference is the answer. Anything that is not a range is a 1x1,
     * because a lookup over one cell is a legal, if pointless, lookup.
     */
    private gridIn(node: FormulaNode, sheetName: string): readonly (readonly CellValue[])[] {
        node = this.deref(node);
        if (node.kind !== 'range') return [[this.node(node, sheetName)]];

        const scope = node.sheet ?? sheetName;
        const box = rangeBox(node.from, node.to);
        if (!box) return [[err('#REF!')]];

        const rows: CellValue[][] = [];
        for (let row = box.top; row <= box.bottom; row++) {
            const line: CellValue[] = [];
            for (let col = box.left; col <= box.right; col++) {
                line.push(this.cellValue(scope, indexToColumn(col) + row));
            }
            rows.push(line);
        }

        return rows;
    }

    private node(node: FormulaNode, sheetName: string): CellValue {
        switch (node.kind) {
            case 'number':
                return num(node.value);
            case 'text':
                return text(node.value);
            case 'boolean':
                return bool(node.value);
            case 'error':
                return err(node.code);

            case 'name':
                return this.node(this.deref(node), sheetName);

            case 'ref': {
                if (!parseRef(node.ref)) return err('#REF!');

                return this.cellValue(node.sheet ?? sheetName, node.ref);
            }

            case 'range': {
                // A range in a scalar position is its first cell, which is what
                // a spreadsheet does outside an array formula.
                const refs = rangeRefs(node.from, node.to);

                return refs.length === 0
                    ? err('#REF!')
                    : this.cellValue(node.sheet ?? sheetName, refs[0]);
            }

            case 'unary': {
                const operand = this.node(node.operand, sheetName);
                const stop = propagate(operand);
                if (stop) return stop;
                const n = toNumber(operand);
                if (!n.ok) return n.value;

                return num(node.op === '-' ? -n.value : n.value);
            }

            case 'percent': {
                const operand = this.node(node.operand, sheetName);
                const stop = propagate(operand);
                if (stop) return stop;
                const n = toNumber(operand);

                return n.ok ? num(n.value / 100) : n.value;
            }

            case 'binary':
                return this.binary(node.op, node.left, node.right, sheetName);

            case 'call': {
                const fn = lookupFunction(node.name);
                if (!fn) return err('#NAME?');
                if (node.args.length < fn.minArgs || node.args.length > fn.maxArgs) {
                    return err('#VALUE!');
                }
                // The context evaluates in the sheet the CALL sits in, which is
                // not always the evaluator's own when a formula crosses sheets.
                // ⚠️ Every one of these DELEGATES. This object used to carry its
                // own copy of `spread`, and the copy is what functions actually
                // receive -- so a defined name resolved everywhere except inside
                // a function call, which is the only place `SUM(items_amount)`
                // ever appears. `grid` delegated and worked; `spread` did not
                // and silently summed the range's first cell (#2384).
                const scoped: FunctionContext = {
                    evaluate: (n2) => this.node(n2, sheetName),
                    spread: (n2) => this.spreadIn(n2, sheetName),
                    grid: (n2) => this.gridIn(n2, sheetName),
                };

                return fn.call(node.args, scoped);
            }

            default:
                return err('#VALUE!');
        }
    }

    private binary(
        op: BinaryOperator,
        leftNode: FormulaNode,
        rightNode: FormulaNode,
        sheetName: string,
    ): CellValue {
        const left = this.node(leftNode, sheetName);
        const right = this.node(rightNode, sheetName);
        const stop = propagate(left, right);
        if (stop) return stop;

        if (op === '&') return text(toText(left) + toText(right));

        if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
            return compareValues(op, left, right);
        }

        const a = toNumber(left);
        const b = toNumber(right);
        if (!a.ok) return a.value;
        if (!b.ok) return b.value;

        switch (op) {
            case '+': return num(a.value + b.value);
            case '-': return num(a.value - b.value);
            case '*': return num(a.value * b.value);
            case '/': return b.value === 0 ? err('#DIV/0!') : num(a.value / b.value);
            case '^': {
                const result = a.value ** b.value;

                return Number.isFinite(result) ? num(result) : err('#NUM!');
            }
            default: return err('#VALUE!');
        }
    }
}

/**
 * What a literal cell holds, once template tokens are accounted for.
 *
 * ## ⚠️ The DOCUMENT decides the type, not the spelling
 *
 * A cell whose text happens to parse as a number is not necessarily a number:
 * `42` stored as text is `t="s"` in the file, and the model says the same thing
 * with `numberFormat: '@'` (#1977). Inferring the type from the string made
 * `ISNUMBER(G1)` answer TRUE and `G1 = 42` answer TRUE for a cell holding TEXT
 * -- both of which LibreOffice and Excel answer FALSE (#2349).
 *
 * Arithmetic still coerces: `G1 + 1` is 43 and `G1 & "!"` is `42!`, because
 * that is what a spreadsheet does with text that looks like a number. What
 * changes is the cell's TYPE, which is what `=` and `ISNUMBER` ask about.
 */
function literalValue(raw: string | undefined, numberFormat?: string): CellValue {
    if (raw === undefined || raw === '') return BLANK;
    if (DTMPL_TOKEN.test(raw)) {
        return {
            kind: 'unresolved',
            because: 'This cell holds a template token, so its value is not known until the document is generated.',
        };
    }
    if (numberFormat === '@') return text(raw);
    const trimmed = raw.trim();
    const n = Number(trimmed);
    if (trimmed !== '' && Number.isFinite(n)) return num(n);
    const upper = trimmed.toUpperCase();
    if (upper === 'TRUE') return bool(true);
    if (upper === 'FALSE') return bool(false);

    return text(raw);
}

/** A range as numbers, normalised however the author wrote its corners. */
function rangeBox(
    from: string,
    to: string,
): { top: number; bottom: number; left: number; right: number } | null {
    const a = parseRef(from);
    const b = parseRef(to);
    if (!a || !b) return null;

    return {
        top: Math.min(a.row, b.row),
        bottom: Math.max(a.row, b.row),
        left: Math.min(columnToIndex(a.column), columnToIndex(b.column)),
        right: Math.max(columnToIndex(a.column), columnToIndex(b.column)),
    };
}

/** Every reference in a range, in reading order. */
/**
 * `Sheet1!$B$2:$B$3` → the node it denotes. Null when it is not a range at all.
 *
 * ⚠️ The `$` signs are stripped rather than honoured. They mean "do not move me
 * when this formula is COPIED", which no evaluation here performs — and a
 * defined name is written with them by every editor, so treating `$B$2` as a
 * different reference from `B2` would make every imported name unresolvable.
 *
 * A quoted sheet (`'My Sheet'!A1:B2`) is accepted because that is how a name
 * with a space in its sheet is written.
 */
function parseScopedRange(declared: string): FormulaNode | null {
    const scoped = /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!(.+)$/.exec(declared);
    const sheet = scoped ? (scoped[1] ?? '').replace(/''/g, "'") || scoped[2] : undefined;
    const body = (scoped ? scoped[3] : declared).replace(/\$/g, '').toUpperCase();

    const [from, to] = body.split(':');
    if (!from || !parseRef(from)) return null;

    if (to === undefined) return { kind: 'ref', ref: from, sheet };

    return parseRef(to) ? { kind: 'range', from, to, sheet } : null;
}

function rangeRefs(from: string, to: string): string[] {
    const box = rangeBox(from, to);
    if (!box) return [];

    const out: string[] = [];
    for (let row = box.top; row <= box.bottom; row++) {
        for (let col = box.left; col <= box.right; col++) {
            out.push(indexToColumn(col) + row);
        }
    }

    return out;
}

/**
 * The value of a formula written in a given sheet.
 *
 * `source` is the formula WITHOUT its leading `=`, matching how the document
 * stores it, so no caller strips it twice.
 */
export function evaluateFormula(
    source: string,
    doc: SheetDocumentDto,
    options: EvaluationOptions,
): CellValue {
    const parsed = parseFormula(source);
    if (!parsed.ok) return err('#NAME?');

    return new Evaluator(doc, options).evaluate(parsed.node);
}

/** The value of a cell, formula or literal. */
export function evaluateCell(
    doc: SheetDocumentDto,
    sheet: string,
    ref: string,
): CellValue {
    return new Evaluator(doc, { sheet }).cellValue(sheet, ref);
}

/**
 * Every formula cell on a sheet, in ONE pass.
 *
 * The grid needs all of them at once, and calling `evaluateCell` per cell would
 * build a fresh evaluator each time -- so a chain `A3 = A2 + 1`, `A2 = A1 + 1`
 * would walk `A1` once per dependent instead of once in total. Sharing the memo
 * across the sheet turns that from quadratic into linear on the common shape,
 * a column of running totals.
 *
 * Only formula cells are returned: a literal already displays itself, and
 * putting every cell in the map would make it the size of the sheet for no
 * reason.
 */
export function evaluateSheet(
    doc: SheetDocumentDto,
    sheet: string,
): ReadonlyMap<string, CellValue> {
    const out = new Map<string, CellValue>();
    const cells = doc.sheets[sheet]?.cells;
    if (!cells) return out;

    const evaluator = new Evaluator(doc, { sheet });
    for (const [ref, cell] of Object.entries(cells)) {
        if (cell.formula !== undefined) {
            out.set(ref, evaluator.cellValue(sheet, ref));
        }
    }

    return out;
}
