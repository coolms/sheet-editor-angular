/**
 * The functions this engine knows, and what they tell the author about
 * themselves.
 *
 * Every entry carries a `signature`, a `summary` and a `category` because the
 * two surfaces that show functions — the helper popup that appears as you type
 * `=SUM(`, and the browsable list behind the toolbar's Σ — both read THIS.
 * Putting the description here rather than in either popup means there is ONE
 * list: a function cannot exist without being describable, and neither surface
 * can offer one that does not exist.
 *
 * **Functions receive unevaluated NODES**, not values, so a function can decide
 * what to evaluate. `IF(A1=0, 0, 1/A1)` must not compute `1/A1` when `A1` is
 * zero; eager arguments would make that `#DIV/0!` and be wrong.
 *
 * **No volatile functions.** `TODAY`, `NOW` and `RAND` are deliberately absent:
 * this grid edits a TEMPLATE, and a date captured while authoring is not the
 * date the document will be generated on. Offering them would invite an author
 * to bake in an answer the backend is supposed to compute.
 *
 * **Where the semantics are Excel's rather than the safer choice, they are
 * Excel's on purpose.** A formula is written through to the `.xlsx` verbatim,
 * so anything computed differently here would be a preview that disagrees with
 * the document it previews. `VLOOKUP`'s fourth argument defaulting to an
 * approximate match is the standing example: it is a footgun, and changing it
 * would be a worse one.
 */
import { dateToSerial, formatCellValue, serialToDate, type DateParts } from '../number-format';
import type { FormulaNode } from './ast';
import {
    BLANK, bool, comparesTrue, err, isError, num, propagate, text, toBoolean, toNumber, toText,
    type CellValue,
} from './values';

/**
 * The shelves the browsable list puts functions on.
 *
 * A spreadsheet author already knows this vocabulary from every other
 * spreadsheet, so it is theirs and not ours. Kept short for the same reason the
 * number-format menu is: a category nobody can predict the contents of is worse
 * than one more entry in a category they can.
 */
export type FunctionCategory =
    'Math' | 'Statistical' | 'Logical' | 'Text' | 'Date' | 'Lookup' | 'Info';

export const FUNCTION_CATEGORIES: readonly FunctionCategory[] =
    ['Math', 'Statistical', 'Logical', 'Text', 'Date', 'Lookup', 'Info'];

export interface FunctionContext {
    /** One value: a range collapses to its first cell, as a scalar context does. */
    evaluate(node: FormulaNode): CellValue;
    /** Every value: a range spreads, a scalar yields one. */
    spread(node: FormulaNode): readonly CellValue[];
    /**
     * Every value, in ROWS.
     *
     * `spread` flattens, which is all `SUM` ever needed. `VLOOKUP` and `INDEX`
     * are about POSITION — "the third column of this block" — and a flat list
     * cannot answer that: twelve values are a 3x4 block or a 4x3 one, and the
     * difference is the answer.
     */
    grid(node: FormulaNode): readonly (readonly CellValue[])[];
}

export interface FormulaFunction {
    readonly name: string;
    readonly category: FunctionCategory;
    readonly minArgs: number;
    readonly maxArgs: number;
    /** Shown in the helper, e.g. `SUM(number1, [number2, …])`. */
    readonly signature: string;
    /** One sentence, shown under the signature. */
    readonly summary: string;
    call(args: readonly FormulaNode[], ctx: FunctionContext): CellValue;
}

/** An entry before it is shelved — see {@link inCategory}. */
type FunctionSpec = Omit<FormulaFunction, 'category'>;

/**
 * Stamp a group of entries with the category they are declared under.
 *
 * The alternative is a `category` field on all forty-odd entries, where a
 * wrong one reads exactly like a right one. Declaring the group once puts the
 * shelf in the SHAPE of the file, so a function in the wrong place is visible.
 */
const inCategory = (category: FunctionCategory, fns: readonly FunctionSpec[]): FormulaFunction[] =>
    fns.map((fn) => ({ ...fn, category }));

/** Every argument's values, flattened, with the first error or unresolved won. */
function collect(args: readonly FormulaNode[], ctx: FunctionContext): CellValue[] | CellValue {
    const out: CellValue[] = [];
    for (const arg of args) {
        for (const v of ctx.spread(arg)) {
            const stop = propagate(v);
            if (stop) return stop;
            out.push(v);
        }
    }

    return out;
}

/** The numbers among some values — text and blanks ignored, as `SUM` does. */
function numbersOf(values: readonly CellValue[]): number[] | CellValue {
    const out: number[] = [];
    for (const v of values) {
        if (v.kind === 'blank' || v.kind === 'text') continue;
        const n = toNumber(v);
        if (!n.ok) return n.value;
        out.push(n.value);
    }

    return out;
}

/** A function over the numbers in its arguments. */
function numeric(
    name: string,
    signature: string,
    summary: string,
    minArgs: number,
    reduce: (numbers: readonly number[]) => CellValue,
): FunctionSpec {
    return {
        name,
        minArgs,
        maxArgs: Number.POSITIVE_INFINITY,
        signature,
        summary,
        call(args, ctx) {
            const values = collect(args, ctx);
            if (!Array.isArray(values)) return values;
            const numbers = numbersOf(values);
            if (!Array.isArray(numbers)) return numbers;

            return reduce(numbers);
        },
    };
}

/** A function of exactly one number. */
function unary(
    name: string,
    signature: string,
    summary: string,
    apply: (n: number) => CellValue,
): FunctionSpec {
    return {
        name,
        minArgs: 1,
        maxArgs: 1,
        signature,
        summary,
        call(args, ctx) {
            const v = ctx.evaluate(args[0]);
            const stop = propagate(v);
            if (stop) return stop;
            const n = toNumber(v);

            return n.ok ? apply(n.value) : n.value;
        },
    };
}

/** A function of exactly one string. */
function textual(
    name: string,
    signature: string,
    summary: string,
    apply: (s: string) => CellValue,
): FunctionSpec {
    return {
        name,
        minArgs: 1,
        maxArgs: 1,
        signature,
        summary,
        call(args, ctx) {
            const v = ctx.evaluate(args[0]);
            const stop = propagate(v);
            if (stop) return stop;

            return apply(toText(v));
        },
    };
}

/** A function that asks one question about one value's kind. */
function predicate(
    name: string,
    signature: string,
    summary: string,
    holds: (v: CellValue) => boolean,
): FunctionSpec {
    return {
        name,
        minArgs: 1,
        maxArgs: 1,
        signature,
        summary,
        // NOT propagated: asking "is this an error?" about an error must answer
        // the question rather than become the error.
        call: (args, ctx) => bool(holds(ctx.evaluate(args[0]))),
    };
}

const round = (n: number, digits: number, mode: 'half' | 'up' | 'down'): number => {
    const factor = 10 ** digits;
    // ⚠️ `1.005 * 100` lands at 100.49999999999999, so rounding the scaled
    // double gives 1 where every spreadsheet gives 1.01. Excel rounds the
    // DECIMAL the author sees, so normalise to the 15 significant digits a
    // double actually carries before deciding. The same artefact bites
    // ROUNDDOWN: `4.35 * 100` is 434.99999999999994.
    const scaled = Number((n * factor).toPrecision(15));
    const rounded = mode === 'half'
        // Away from zero at the halfway point, which is what a spreadsheet does
        // and what `Math.round` does NOT for negatives.
        ? Math.sign(scaled) * Math.round(Math.abs(scaled))
        : mode === 'up' ? Math.sign(scaled) * Math.ceil(Math.abs(scaled))
            : Math.sign(scaled) * Math.floor(Math.abs(scaled));

    return rounded / factor;
};

const digitsOf = (args: readonly FormulaNode[], ctx: FunctionContext): number | CellValue => {
    if (args.length < 2) return 0;
    const v = ctx.evaluate(args[1]);
    const stop = propagate(v);
    if (stop) return stop;
    const n = toNumber(v);

    return n.ok ? Math.trunc(n.value) : n.value;
};

function rounder(name: string, summary: string, mode: 'half' | 'up' | 'down'): FunctionSpec {
    return {
        name,
        minArgs: 1,
        maxArgs: 2,
        signature: `${name}(number, [digits])`,
        summary,
        call(args, ctx) {
            const v = ctx.evaluate(args[0]);
            const stop = propagate(v);
            if (stop) return stop;
            const n = toNumber(v);
            if (!n.ok) return n.value;
            const digits = digitsOf(args, ctx);
            if (typeof digits !== 'number') return digits;

            return num(round(n.value, digits, mode));
        },
    };
}

// ── Criteria, wildcards and lookup ──────────────────────────────────────────

/** `">10"` splits into an operator and an operand; a bare value is equality. */
const CRITERION = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/u;

/** Every character a regex would read as syntax. */
const escapeRegex = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&');

/**
 * `*` matches any run, `?` exactly one, `~` escapes either — as a spreadsheet
 * spells it. Case-insensitive, because every other text comparison here is.
 */
function wildcardPattern(source: string): RegExp {
    let out = '';
    for (let i = 0; i < source.length; i++) {
        const ch = source[i];
        const next = source[i + 1];
        if (ch === '~' && (next === '*' || next === '?' || next === '~')) {
            out += escapeRegex(next);
            i += 1;
            continue;
        }
        if (ch === '*') { out += '[\\s\\S]*'; continue; }
        if (ch === '?') { out += '[\\s\\S]'; continue; }
        out += escapeRegex(ch);
    }

    return new RegExp(`^${out}$`, 'iu');
}

/**
 * A `SUMIF`/`COUNTIF` criterion, as the test it stands for.
 *
 * The criterion is a comparison SPELLED AS TEXT — `">10"` is not the text
 * `">10"`, it is "greater than ten" — which is the one thing about these
 * functions everybody gets wrong the first time. The ordering it compares with
 * is `compareValues`, the same one `>` uses, so a criterion and an operator can
 * never disagree about what is bigger.
 */
function matcherFor(criterion: CellValue): (candidate: CellValue) => boolean {
    if (criterion.kind !== 'text') {
        return (candidate) => comparesTrue('=', candidate, criterion);
    }

    const parsed = CRITERION.exec(criterion.value);
    const op = parsed?.[1] ?? '=';
    const rest = (parsed?.[2] ?? '').trim();
    const asNumber = Number(rest);
    const operand: CellValue = rest !== '' && Number.isFinite(asNumber) ? num(asNumber) : text(rest);

    // Wildcards are an EQUALITY test and nothing else: `">a*"` has no meaning a
    // spreadsheet defines, so a pattern under an ordering operator is treated
    // as the literal text it looks like.
    if ((op === '=' || op === '<>') && operand.kind === 'text' && /[*?]/u.test(operand.value)) {
        const pattern = wildcardPattern(operand.value);

        return (candidate) => pattern.test(toText(candidate)) === ('=' === op);
    }

    return (candidate) => comparesTrue(op, candidate, operand);
}

/**
 * Equality as a LOOKUP means it.
 *
 * Deliberately not {@link matcherFor}: a lookup key is a value, not a
 * criterion, so a key that happens to begin `<` is that text and not a
 * comparison. Wildcards still apply, because `MATCH("wid*", …)` is an idiom.
 */
function looksUp(key: CellValue): (candidate: CellValue) => boolean {
    if (key.kind === 'text' && /[*?]/u.test(key.value)) {
        const pattern = wildcardPattern(key.value);

        return (candidate) => pattern.test(toText(candidate));
    }

    return (candidate) => comparesTrue('=', candidate, key);
}

/**
 * Where a key sits in a vector, or -1.
 *
 * `type` 0 is an exact match. 1 wants the vector ASCENDING and answers with the
 * last value not greater than the key; -1 wants it descending. Both give a
 * wrong answer on unsorted data rather than an error — which is what every
 * spreadsheet does, and this engine has to agree with the workbook the formula
 * is written into.
 */
function matchIndex(key: CellValue, values: readonly CellValue[], type: number): number {
    if (0 === type) {
        const matches = looksUp(key);

        return values.findIndex((v) => matches(v));
    }

    let found = -1;
    for (let i = 0; i < values.length; i++) {
        const within = type > 0
            ? comparesTrue('<=', values[i], key)
            : comparesTrue('>=', values[i], key);
        if (!within) break;
        found = i;
    }

    return found;
}

/**
 * The values a criterion picks out of a range, paired with a second range.
 *
 * ⚠️ The two ranges must hold the same NUMBER of cells. Excel resizes a short
 * `sum_range` from its top-left corner to match; this engine cannot see past
 * the range it was handed, so it says `#VALUE!` rather than quietly summing a
 * column that stops early. The mismatch is a mistake far more often than an
 * intention.
 */
function pickedBy(
    args: readonly FormulaNode[],
    ctx: FunctionContext,
    targetIndex: number,
): CellValue[] | CellValue {
    const tested = ctx.spread(args[0]);
    const criterion = ctx.evaluate(args[1]);
    const stop = propagate(criterion, ...tested);
    if (stop) return stop;

    const target = args[targetIndex] === undefined ? tested : ctx.spread(args[targetIndex]);
    if (target.length !== tested.length) return err('#VALUE!');
    const stopTarget = propagate(...target);
    if (stopTarget) return stopTarget;

    const matches = matcherFor(criterion);
    const out: CellValue[] = [];
    for (let i = 0; i < tested.length; i++) {
        if (matches(tested[i])) out.push(target[i]);
    }

    return out;
}

/** A run of arguments as plain numbers, or the value that stopped them. */
function numberArgs(
    args: readonly FormulaNode[],
    ctx: FunctionContext,
    from: number,
    count: number,
): number[] | CellValue {
    const out: number[] = [];
    for (let i = from; i < from + count; i++) {
        const v = ctx.evaluate(args[i]);
        const stop = propagate(v);
        if (stop) return stop;
        const n = toNumber(v);
        if (!n.ok) return n.value;
        out.push(n.value);
    }

    return out;
}

/** One argument as a calendar date, or the value that stopped it. */
function dateArg(
    args: readonly FormulaNode[],
    ctx: FunctionContext,
    index: number,
): DateParts | CellValue {
    const v = ctx.evaluate(args[index]);
    const stop = propagate(v);
    if (stop) return stop;
    const n = toNumber(v);
    if (!n.ok) return n.value;

    return serialToDate(n.value) ?? err('#NUM!');
}

/** `YEAR`, `MONTH`, `DAY` — one reading of a date, three times over. */
function datePart(name: string, summary: string, read: (d: DateParts) => number): FunctionSpec {
    return {
        name,
        minArgs: 1,
        maxArgs: 1,
        signature: `${name}(date)`,
        summary,
        call(args, ctx) {
            const date = dateArg(args, ctx, 0);

            return 'year' in date ? num(read(date)) : date;
        },
    };
}

/**
 * `EDATE` and `EOMONTH`, which differ only in which day of the month they land on.
 *
 * The 31st of a month shifted onto a 30-day month CLAMPS to the 30th rather
 * than rolling into the next month: an invoice dated the 31st of January is due
 * on the 28th of February, not the 3rd of March, and every spreadsheet agrees.
 */
function shiftMonths(
    args: readonly FormulaNode[],
    ctx: FunctionContext,
    endOfMonth: boolean,
): CellValue {
    const date = dateArg(args, ctx, 0);
    if (!('year' in date)) return date;
    const months = toNumber(ctx.evaluate(args[1]));
    if (!months.ok) return months.value;

    const total = date.year * 12 + (date.month - 1) + Math.trunc(months.value);
    const year = Math.floor(total / 12);
    const month = (total % 12) + 1;
    // Day 0 of the NEXT month is the last day of this one, which is how the
    // length of February gets answered without a table of month lengths.
    const last = serialToDate(dateToSerial(year, month + 1, 0))?.day ?? 28;
    const day = endOfMonth ? last : Math.min(date.day, last);
    const serial = dateToSerial(year, month, day);

    return Number.isFinite(serial) && serial >= 1 ? num(serial) : err('#NUM!');
}

/** The first column of a grid, which is what `VLOOKUP` searches. */
const firstColumn = (rows: readonly (readonly CellValue[])[]): CellValue[] =>
    rows.map((row) => row[0] ?? BLANK);

const FUNCTIONS: readonly FormulaFunction[] = [
    ...inCategory('Math', [
        numeric('SUM', 'SUM(number1, [number2, …])', 'Adds its arguments.', 1,
            (ns) => num(ns.reduce((a, b) => a + b, 0))),
        numeric('PRODUCT', 'PRODUCT(number1, [number2, …])', 'Multiplies its arguments.', 1,
            (ns) => num(ns.reduce((a, b) => a * b, 1))),
        unary('ABS', 'ABS(number)', 'The number without its sign.', (n) => num(Math.abs(n))),
        unary('INT', 'INT(number)', 'Rounds down to a whole number.', (n) => num(Math.floor(n))),
        unary('SQRT', 'SQRT(number)', 'The square root.',
            (n) => (n < 0 ? err('#NUM!') : num(Math.sqrt(n)))),
        unary('SIGN', 'SIGN(number)', 'Minus one, zero or one.', (n) => num(Math.sign(n))),
        rounder('ROUND', 'Rounds to a number of digits.', 'half'),
        rounder('ROUNDUP', 'Rounds away from zero.', 'up'),
        rounder('ROUNDDOWN', 'Rounds towards zero.', 'down'),
        {
            name: 'MOD',
            minArgs: 2,
            maxArgs: 2,
            signature: 'MOD(number, divisor)',
            summary: 'The remainder after division.',
            call(args, ctx) {
                const a = toNumber(ctx.evaluate(args[0]));
                const b = toNumber(ctx.evaluate(args[1]));
                if (!a.ok) return a.value;
                if (!b.ok) return b.value;
                if (b.value === 0) return err('#DIV/0!');

                // Sign follows the DIVISOR, as a spreadsheet defines it -- which is
                // not what `%` does in JavaScript for negative operands.
                return num(a.value - b.value * Math.floor(a.value / b.value));
            },
        },
        {
            name: 'POWER',
            minArgs: 2,
            maxArgs: 2,
            signature: 'POWER(number, exponent)',
            summary: 'The number raised to a power.',
            call(args, ctx) {
                const a = toNumber(ctx.evaluate(args[0]));
                const b = toNumber(ctx.evaluate(args[1]));
                if (!a.ok) return a.value;
                if (!b.ok) return b.value;
                const result = a.value ** b.value;

                return Number.isFinite(result) ? num(result) : err('#NUM!');
            },
        },
        {
            name: 'SUMIF',
            minArgs: 2,
            maxArgs: 3,
            signature: 'SUMIF(range, criterion, [sum_range])',
            summary: 'Adds the cells that meet a condition, e.g. SUMIF(A2:A9, ">10").',
            call(args, ctx) {
                const values = pickedBy(args, ctx, 2);
                if (!Array.isArray(values)) return values;
                const numbers = numbersOf(values);
                if (!Array.isArray(numbers)) return numbers;

                return num(numbers.reduce((a, b) => a + b, 0));
            },
        },
    ]),

    ...inCategory('Statistical', [
        numeric('MIN', 'MIN(number1, [number2, …])', 'The smallest number.', 1,
            (ns) => num(ns.length === 0 ? 0 : Math.min(...ns))),
        numeric('MAX', 'MAX(number1, [number2, …])', 'The largest number.', 1,
            (ns) => num(ns.length === 0 ? 0 : Math.max(...ns))),
        numeric('AVERAGE', 'AVERAGE(number1, [number2, …])', 'The mean of the numbers.', 1,
            // Dividing by zero here is `#DIV/0!` and not `0`: the average of nothing
            // is not a number, and a spreadsheet says so.
            (ns) => (ns.length === 0 ? err('#DIV/0!') : num(ns.reduce((a, b) => a + b, 0) / ns.length))),
        {
            name: 'COUNT',
            minArgs: 1,
            maxArgs: Number.POSITIVE_INFINITY,
            signature: 'COUNT(value1, [value2, …])',
            summary: 'How many arguments are numbers.',
            // ⚠️ COUNT IGNORES errors rather than propagating them -- it is
            // asking "how many of these are numbers?", and an error is simply
            // not one. Built on `numeric()` it answered `#DIV/0!` where a
            // spreadsheet answers 0 (#2349).
            call(args, ctx) {
                // ⚠️ Not `collect()`, which propagates the first error it
                // meets. COUNT has to walk past one.
                let count = 0;
                for (const arg of args) {
                    for (const value of ctx.spread(arg)) {
                        // ⚠️ UNRESOLVED still stops it: a template token means
                        // the answer is not known until the document is
                        // generated, which is not the same as "not a number".
                        if (value.kind === 'unresolved') return value;
                        if (value.kind === 'error' || value.kind === 'blank' || value.kind === 'text') continue;
                        if (toNumber(value).ok) ++count;
                    }
                }

                return num(count);
            },
        },
        {
            name: 'COUNTA',
            minArgs: 1,
            maxArgs: Number.POSITIVE_INFINITY,
            signature: 'COUNTA(value1, [value2, …])',
            summary: 'How many arguments are not empty.',
            call(args, ctx) {
                const values = collect(args, ctx);
                if (!Array.isArray(values)) return values;

                return num(values.filter((v) => v.kind !== 'blank').length);
            },
        },
        {
            name: 'COUNTBLANK',
            minArgs: 1,
            maxArgs: Number.POSITIVE_INFINITY,
            signature: 'COUNTBLANK(range)',
            summary: 'How many cells in the range are empty.',
            call(args, ctx) {
                const values = collect(args, ctx);
                if (!Array.isArray(values)) return values;

                return num(values.filter((v) => v.kind === 'blank').length);
            },
        },
        {
            name: 'COUNTIF',
            minArgs: 2,
            maxArgs: 2,
            signature: 'COUNTIF(range, criterion)',
            summary: 'How many cells meet a condition, e.g. COUNTIF(A2:A9, "Widget").',
            call(args, ctx) {
                const values = pickedBy(args, ctx, 2);

                return Array.isArray(values) ? num(values.length) : values;
            },
        },
        {
            name: 'AVERAGEIF',
            minArgs: 2,
            maxArgs: 3,
            signature: 'AVERAGEIF(range, criterion, [average_range])',
            summary: 'The mean of the cells that meet a condition.',
            call(args, ctx) {
                const values = pickedBy(args, ctx, 2);
                if (!Array.isArray(values)) return values;
                const numbers = numbersOf(values);
                if (!Array.isArray(numbers)) return numbers;

                return numbers.length === 0
                    ? err('#DIV/0!')
                    : num(numbers.reduce((a, b) => a + b, 0) / numbers.length);
            },
        },
    ]),

    ...inCategory('Logical', [
        {
            name: 'IF',
            minArgs: 2,
            maxArgs: 3,
            signature: 'IF(condition, then, [otherwise])',
            summary: 'One value when the condition holds, another when it does not.',
            call(args, ctx) {
                const condition = ctx.evaluate(args[0]);
                const stop = propagate(condition);
                if (stop) return stop;
                const b = toBoolean(condition);
                if (!b.ok) return b.value;
                // The branch NOT taken is never evaluated -- which is the whole
                // reason a function receives nodes instead of values.
                if (b.value) return ctx.evaluate(args[1]);

                return args[2] === undefined ? bool(false) : ctx.evaluate(args[2]);
            },
        },
        {
            name: 'IFERROR',
            minArgs: 2,
            maxArgs: 2,
            signature: 'IFERROR(value, fallback)',
            summary: 'The value, or the fallback when it is an error.',
            call(args, ctx) {
                const value = ctx.evaluate(args[0]);
                // An UNRESOLVED value is not an error and must not be swallowed: the
                // author would see the fallback and believe the formula settled,
                // when in truth it is still waiting for the data.
                if (isError(value)) return ctx.evaluate(args[1]);

                return value;
            },
        },
        {
            name: 'AND',
            minArgs: 1,
            maxArgs: Number.POSITIVE_INFINITY,
            signature: 'AND(condition1, [condition2, …])',
            summary: 'True when every condition holds.',
            call(args, ctx) {
                const values = collect(args, ctx);
                if (!Array.isArray(values)) return values;
                for (const v of values) {
                    const b = toBoolean(v);
                    if (!b.ok) return b.value;
                    if (!b.value) return bool(false);
                }

                return bool(true);
            },
        },
        {
            name: 'OR',
            minArgs: 1,
            maxArgs: Number.POSITIVE_INFINITY,
            signature: 'OR(condition1, [condition2, …])',
            summary: 'True when any condition holds.',
            call(args, ctx) {
                const values = collect(args, ctx);
                if (!Array.isArray(values)) return values;
                for (const v of values) {
                    const b = toBoolean(v);
                    if (!b.ok) return b.value;
                    if (b.value) return bool(true);
                }

                return bool(false);
            },
        },
        {
            name: 'NOT',
            minArgs: 1,
            maxArgs: 1,
            signature: 'NOT(condition)',
            summary: 'Reverses a condition.',
            call(args, ctx) {
                const b = toBoolean(ctx.evaluate(args[0]));

                return b.ok ? bool(!b.value) : b.value;
            },
        },
        {
            name: 'TRUE',
            minArgs: 0,
            maxArgs: 0,
            signature: 'TRUE()',
            summary: 'The value TRUE.',
            call: () => bool(true),
        },
        {
            name: 'FALSE',
            minArgs: 0,
            maxArgs: 0,
            signature: 'FALSE()',
            summary: 'The value FALSE.',
            call: () => bool(false),
        },
    ]),

    ...inCategory('Text', [
        {
            name: 'CONCAT',
            minArgs: 1,
            maxArgs: Number.POSITIVE_INFINITY,
            signature: 'CONCAT(text1, [text2, …])',
            summary: 'Joins text together.',
            call(args, ctx) {
                const values = collect(args, ctx);
                if (!Array.isArray(values)) return values;

                return text(values.map(toText).join(''));
            },
        },
        textual('LEN', 'LEN(text)', 'How many characters the text has.',
            (s) => num([...s].length)),
        textual('UPPER', 'UPPER(text)', 'The text in capitals.', (s) => text(s.toUpperCase())),
        textual('LOWER', 'LOWER(text)', 'The text in lower case.', (s) => text(s.toLowerCase())),
        // ⚠️ Excel's TRIM also collapses INTERNAL runs of spaces to one --
        // `TRIM("a  b")` is `"a b"`. Trimming only the ends is what
        // `String.trim()` does, and it is not what the function means.
        textual('TRIM', 'TRIM(text)', 'The text without leading, trailing or repeated spaces.',
            (s) => text(s.replace(/ +/gu, ' ').trim())),
        {
            name: 'LEFT',
            minArgs: 1,
            maxArgs: 2,
            signature: 'LEFT(text, [count])',
            summary: 'The first characters of the text.',
            call(args, ctx) {
                const v = ctx.evaluate(args[0]);
                const stop = propagate(v);
                if (stop) return stop;
                const count = digitsOf([args[0], ...(args[1] ? [args[1]] : [])], ctx);
                if (typeof count !== 'number') return count;

                return text([...toText(v)].slice(0, args[1] === undefined ? 1 : count).join(''));
            },
        },
        {
            name: 'RIGHT',
            minArgs: 1,
            maxArgs: 2,
            signature: 'RIGHT(text, [count])',
            summary: 'The last characters of the text.',
            call(args, ctx) {
                const v = ctx.evaluate(args[0]);
                const stop = propagate(v);
                if (stop) return stop;
                const count = digitsOf([args[0], ...(args[1] ? [args[1]] : [])], ctx);
                if (typeof count !== 'number') return count;
                const chars = [...toText(v)];
                const n = args[1] === undefined ? 1 : count;

                return text(n <= 0 ? '' : chars.slice(Math.max(0, chars.length - n)).join(''));
            },
        },
        {
            name: 'MID',
            minArgs: 3,
            maxArgs: 3,
            signature: 'MID(text, start, count)',
            summary: 'Characters from the middle of the text, counting from one.',
            call(args, ctx) {
                const v = ctx.evaluate(args[0]);
                const stop = propagate(v);
                if (stop) return stop;
                const start = toNumber(ctx.evaluate(args[1]));
                const count = toNumber(ctx.evaluate(args[2]));
                if (!start.ok) return start.value;
                if (!count.ok) return count.value;
                if (start.value < 1) return err('#VALUE!');
                const chars = [...toText(v)];

                return text(chars.slice(start.value - 1, start.value - 1 + Math.max(0, count.value)).join(''));
            },
        },
        {
            name: 'TEXT',
            minArgs: 2,
            maxArgs: 2,
            signature: 'TEXT(value, format)',
            summary: 'A number written out the way a format code says, e.g. TEXT(A1, "dd/mm/yyyy").',
            call(args, ctx) {
                const value = ctx.evaluate(args[0]);
                const code = ctx.evaluate(args[1]);
                const stop = propagate(value, code);
                if (stop) return stop;

                // The same renderer the GRID uses, which is the point: what
                // TEXT produces is what the cell beside it shows.
                return text(formatCellValue(toText(value), toText(code)));
            },
        },
        {
            name: 'SUBSTITUTE',
            minArgs: 3,
            maxArgs: 4,
            signature: 'SUBSTITUTE(text, find, replace, [occurrence])',
            summary: 'Replaces text, every time or only the nth.',
            call(args, ctx) {
                const source = ctx.evaluate(args[0]);
                const find = ctx.evaluate(args[1]);
                const replace = ctx.evaluate(args[2]);
                const stop = propagate(source, find, replace);
                if (stop) return stop;

                const haystack = toText(source);
                const needle = toText(find);
                const with_ = toText(replace);
                // Replacing nothing would loop forever on an empty needle, and a
                // spreadsheet answers with the text untouched.
                if (needle === '') return text(haystack);

                if (args[3] === undefined) return text(haystack.split(needle).join(with_));

                const nth = toNumber(ctx.evaluate(args[3]));
                if (!nth.ok) return nth.value;
                const wanted = Math.trunc(nth.value);
                if (wanted < 1) return err('#VALUE!');

                let seen = 0;
                let at = haystack.indexOf(needle);
                while (at >= 0) {
                    seen += 1;
                    if (seen === wanted) {
                        return text(haystack.slice(0, at) + with_ + haystack.slice(at + needle.length));
                    }
                    at = haystack.indexOf(needle, at + needle.length);
                }

                return text(haystack);
            },
        },
    ]),

    ...inCategory('Date', [
        {
            name: 'DATE',
            minArgs: 3,
            maxArgs: 3,
            signature: 'DATE(year, month, day)',
            summary: 'A date, as the serial number a spreadsheet stores.',
            call(args, ctx) {
                const parts = numberArgs(args, ctx, 0, 3);
                if (!Array.isArray(parts)) return parts;
                // Out-of-range months and days ROLL: DATE(2026, 13, 1) is
                // January 2027 in every spreadsheet, and it is how an author
                // writes "a year on" without arithmetic.
                const serial = dateToSerial(parts[0], parts[1], parts[2]);

                return Number.isFinite(serial) && serial >= 1 ? num(serial) : err('#NUM!');
            },
        },
        datePart('YEAR', 'The year of a date.', (d) => d.year),
        datePart('MONTH', 'The month of a date, 1 to 12.', (d) => d.month),
        datePart('DAY', 'The day of the month.', (d) => d.day),
        {
            name: 'WEEKDAY',
            minArgs: 1,
            maxArgs: 2,
            signature: 'WEEKDAY(date, [type])',
            summary: 'Which day of the week: 1 counts from Sunday, 2 from Monday, 3 from Monday at zero.',
            call(args, ctx) {
                const date = dateArg(args, ctx, 0);
                if (!('year' in date)) return date;

                let type = 1;
                if (args[1] !== undefined) {
                    const t = toNumber(ctx.evaluate(args[1]));
                    if (!t.ok) return t.value;
                    type = Math.trunc(t.value);
                }

                // `weekday` is 1..7 from Sunday, which IS type 1.
                switch (type) {
                    case 1: return num(date.weekday);
                    case 2: return num(date.weekday === 1 ? 7 : date.weekday - 1);
                    case 3: return num(date.weekday === 1 ? 6 : date.weekday - 2);
                    default: return err('#NUM!');
                }
            },
        },
        {
            name: 'EDATE',
            minArgs: 2,
            maxArgs: 2,
            signature: 'EDATE(date, months)',
            summary: 'The same day a number of months away — the usual way to date an invoice.',
            call: (args, ctx) => shiftMonths(args, ctx, false),
        },
        {
            name: 'EOMONTH',
            minArgs: 2,
            maxArgs: 2,
            signature: 'EOMONTH(date, months)',
            summary: 'The last day of the month a number of months away.',
            call: (args, ctx) => shiftMonths(args, ctx, true),
        },
        {
            name: 'DAYS',
            minArgs: 2,
            maxArgs: 2,
            signature: 'DAYS(end, start)',
            summary: 'How many days between two dates.',
            call(args, ctx) {
                const parts = numberArgs(args, ctx, 0, 2);

                return Array.isArray(parts)
                    ? num(Math.trunc(parts[0]) - Math.trunc(parts[1]))
                    : parts;
            },
        },
    ]),

    ...inCategory('Lookup', [
        {
            name: 'VLOOKUP',
            minArgs: 3,
            maxArgs: 4,
            signature: 'VLOOKUP(key, range, index, [sorted])',
            summary: 'Searches the first column and returns a value from the same row.',
            call(args, ctx) {
                const key = ctx.evaluate(args[0]);
                const rows = ctx.grid(args[1]);
                const index = toNumber(ctx.evaluate(args[2]));
                if (!index.ok) return index.value;

                // ⚠️ Absent, `sorted` is TRUE -- an approximate match. That is a
                // footgun, and it is Excel's, and this formula is written into an
                // Excel file verbatim: an editor that defaulted to an exact match
                // would preview a different answer than the document gives.
                let sorted = true;
                if (args[3] !== undefined) {
                    const flag = toBoolean(ctx.evaluate(args[3]));
                    if (!flag.ok) return flag.value;
                    sorted = flag.value;
                }

                const column = Math.trunc(index.value);
                if (column < 1) return err('#VALUE!');
                const first = firstColumn(rows);
                const stop = propagate(key, ...first);
                if (stop) return stop;

                const at = matchIndex(key, first, sorted ? 1 : 0);
                if (at < 0) return err('#N/A');
                const row = rows[at];
                if (column > row.length) return err('#REF!');

                return row[column - 1];
            },
        },
        {
            name: 'HLOOKUP',
            minArgs: 3,
            maxArgs: 4,
            signature: 'HLOOKUP(key, range, index, [sorted])',
            summary: 'Searches the first row and returns a value from the same column.',
            call(args, ctx) {
                const key = ctx.evaluate(args[0]);
                const rows = ctx.grid(args[1]);
                const index = toNumber(ctx.evaluate(args[2]));
                if (!index.ok) return index.value;

                let sorted = true;
                if (args[3] !== undefined) {
                    const flag = toBoolean(ctx.evaluate(args[3]));
                    if (!flag.ok) return flag.value;
                    sorted = flag.value;
                }

                const rowIndex = Math.trunc(index.value);
                if (rowIndex < 1) return err('#VALUE!');
                const header = rows[0] ?? [];
                const stop = propagate(key, ...header);
                if (stop) return stop;

                const at = matchIndex(key, header, sorted ? 1 : 0);
                if (at < 0) return err('#N/A');
                if (rowIndex > rows.length) return err('#REF!');
                const row = rows[rowIndex - 1];

                return row[at] ?? err('#REF!');
            },
        },
        {
            name: 'MATCH',
            minArgs: 2,
            maxArgs: 3,
            signature: 'MATCH(key, range, [type])',
            summary: 'The position of a value: 0 matches exactly, 1 wants ascending order, -1 descending.',
            call(args, ctx) {
                const key = ctx.evaluate(args[0]);
                const values = ctx.spread(args[1]);
                const stop = propagate(key, ...values);
                if (stop) return stop;

                let type = 1;
                if (args[2] !== undefined) {
                    const t = toNumber(ctx.evaluate(args[2]));
                    if (!t.ok) return t.value;
                    type = Math.sign(Math.trunc(t.value));
                }

                const at = matchIndex(key, values, type);

                return at < 0 ? err('#N/A') : num(at + 1);
            },
        },
        {
            name: 'INDEX',
            minArgs: 2,
            maxArgs: 3,
            signature: 'INDEX(range, row, [column])',
            summary: 'The cell at a position in the range, counting from one.',
            call(args, ctx) {
                const rows = ctx.grid(args[0]);
                const first = toNumber(ctx.evaluate(args[1]));
                if (!first.ok) return first.value;

                let row = Math.trunc(first.value);
                let column = 1;
                if (args[2] === undefined) {
                    // One row across: the single index is the COLUMN, which is
                    // what `INDEX(A1:E1, 3)` plainly means and what a strict
                    // row-then-column reading would call out of range.
                    if (rows.length === 1) {
                        column = row;
                        row = 1;
                    }
                } else {
                    const second = toNumber(ctx.evaluate(args[2]));
                    if (!second.ok) return second.value;
                    column = Math.trunc(second.value);
                }

                // Zero means "the whole row or column" in a spreadsheet's array
                // context. This engine has no arrays, so it is out of range
                // rather than a silent first cell.
                if (row < 1 || row > rows.length) return err('#REF!');
                const line = rows[row - 1];
                if (column < 1 || column > line.length) return err('#REF!');

                return line[column - 1];
            },
        },
    ]),

    ...inCategory('Info', [
        predicate('ISBLANK', 'ISBLANK(value)', 'True when the cell is empty.',
            (v) => v.kind === 'blank'),
        predicate('ISNUMBER', 'ISNUMBER(value)', 'True when the value is a number.',
            (v) => v.kind === 'number'),
        predicate('ISTEXT', 'ISTEXT(value)', 'True when the value is text.',
            (v) => v.kind === 'text'),
        // Deliberately false for an UNRESOLVED value: a template token is not a
        // mistake, and answering TRUE would let IFERROR-shaped logic swallow a
        // cell that is merely waiting for its data.
        predicate('ISERROR', 'ISERROR(value)', 'True when the value is an error.',
            (v) => isError(v)),
        {
            name: 'VALUE',
            minArgs: 1,
            maxArgs: 1,
            signature: 'VALUE(text)',
            summary: 'Text read as a number.',
            call(args, ctx) {
                const v = ctx.evaluate(args[0]);
                const stop = propagate(v);
                if (stop) return stop;
                const n = toNumber(v);

                return n.ok ? num(n.value) : n.value;
            },
        },
    ]),
];

const BY_NAME = new Map(FUNCTIONS.map((f) => [f.name, f]));

/**
 * OOXML stores a function NEWER than the format's baseline under this prefix.
 *
 * `CONCAT` written in Excel is saved as `_xlfn.CONCAT`, and every reader is
 * expected to strip it -- Excel and LibreOffice both display and evaluate the
 * bare name. It is a storage artefact of the format, not part of the
 * function's identity.
 *
 * ⚠️ Found by comparing this engine against LibreOffice on the same workbook
 * (#2347): `_xlfn.CONCAT` and `_xlfn.DAYS` came back `#NAME?` here and computed
 * there. Since a formula is written through to the `.xlsx` VERBATIM, an author
 * uploading a real Excel file saw the grid disagree with the document it was
 * previewing -- the exact failure this file's own header warns about.
 *
 * ⚠️ `_xludf.` is deliberately NOT stripped. That prefix marks a USER-DEFINED
 * function, which this engine genuinely cannot run, and `#NAME?` is the honest
 * answer for it.
 */
const FUTURE_FUNCTION_PREFIX = '_XLFN.';

export const lookupFunction = (name: string): FormulaFunction | undefined => {
    const upper = name.toUpperCase();

    return BY_NAME.get(
        upper.startsWith(FUTURE_FUNCTION_PREFIX) ? upper.slice(FUTURE_FUNCTION_PREFIX.length) : upper,
    );
};

/** Every function, for the helper's list and for `=SU` style completion. */
export const allFunctions = (): readonly FormulaFunction[] => FUNCTIONS;

/**
 * Every function on its shelf, for the browsable list.
 *
 * Derived rather than declared a second time: a category that held its own copy
 * of the names would be a list to forget to update, which is the failure this
 * whole file is arranged to avoid. An empty category is dropped, so a shelf
 * cannot appear with nothing on it.
 */
export function functionsByCategory(): ReadonlyArray<{
    readonly category: FunctionCategory;
    readonly functions: readonly FormulaFunction[];
}> {
    return FUNCTION_CATEGORIES
        .map((category) => ({
            category,
            functions: FUNCTIONS.filter((f) => f.category === category),
        }))
        .filter((shelf) => shelf.functions.length > 0);
}

export const BLANK_VALUE = BLANK;
