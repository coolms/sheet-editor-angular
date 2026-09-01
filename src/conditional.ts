/**
 * Rules that change how a range LOOKS, according to what is in it.
 *
 * The grid's half of `SheetConditional` on the backend: the same vocabulary,
 * the same range keying, and the evaluation that decides which rule a cell
 * answers to. Kept beside the model rather than inside it because this is the
 * only part of the document that is a QUESTION about a value rather than a fact
 * about it.
 *
 * ## Evaluated against what the cell SHOWS
 *
 * A formula's result, not its text -- "colour it red when it is over 90" means
 * the computed 95, and a rule that read `=B2*C2` as text would never match a
 * number at all. That is the opposite of what find-and-replace does, and both
 * are right: one is about the document's content, this is about its appearance
 * once the content has been worked out.
 */
import { columnToIndex, parseRange, parseRef, type SheetDto } from './sheet-document.model';

export type ConditionalWhen =
    | 'greaterThan' | 'lessThan' | 'equal' | 'notEqual' | 'between' | 'contains' | 'empty';

/** In the order the editor offers them, which is roughly how often they are wanted. */
export const CONDITIONAL_WHENS: ReadonlyArray<{ value: ConditionalWhen; label: string }> = [
    { value: 'greaterThan', label: 'Greater than' },
    { value: 'lessThan', label: 'Less than' },
    { value: 'equal', label: 'Equal to' },
    { value: 'notEqual', label: 'Not equal to' },
    { value: 'between', label: 'Between' },
    { value: 'contains', label: 'Text contains' },
    { value: 'empty', label: 'Is empty' },
];

/** The conditions that compare against a value the author supplies. */
const NEEDS_VALUE: ReadonlySet<ConditionalWhen> =
    new Set(['greaterThan', 'lessThan', 'equal', 'notEqual', 'between', 'contains']);

export interface ConditionalDto {
    readonly when: ConditionalWhen;
    readonly value?: string;
    readonly value2?: string;
    readonly background?: string;
    readonly color?: string;
    readonly bold?: boolean;
    readonly italic?: boolean;
}

/** The look a matching cell takes on, as the grid binds it. */
export interface ConditionalLook {
    background?: string;
    color?: string;
    bold?: boolean;
    italic?: boolean;
}

export const needsValue = (when: ConditionalWhen): boolean => NEEDS_VALUE.has(when);

export const needsSecondValue = (when: ConditionalWhen): boolean => 'between' === when;

/** Whether a rule would leave a matching cell looking exactly as it did. */
export function formatsNothing(rule: ConditionalDto): boolean {
    return undefined === rule.background && undefined === rule.color
        && true !== rule.bold && true !== rule.italic;
}

/**
 * Whether one displayed value answers a rule.
 *
 * Numbers compare as numbers and everything else as text, which is what makes
 * "greater than 9" put 10 above 9 rather than below it -- the trap of comparing
 * `"10" > "9"` as strings, and the reason this is a function and not an
 * expression written twice.
 */
export function matchesRule(shown: string, rule: ConditionalDto): boolean {
    const text = shown.trim();

    if ('empty' === rule.when) return '' === text;
    if (undefined === rule.value) return false;
    if ('contains' === rule.when) {
        return text.toLowerCase().includes(rule.value.toLowerCase());
    }

    const left = Number(text);
    const right = Number(rule.value);
    const numeric = '' !== text && Number.isFinite(left) && Number.isFinite(right);

    switch (rule.when) {
        case 'greaterThan':
            return numeric ? left > right : text > rule.value;
        case 'lessThan':
            return numeric ? left < right : text < rule.value;
        case 'equal':
            return numeric ? left === right : text === rule.value;
        case 'notEqual':
            return numeric ? left !== right : text !== rule.value;
        case 'between': {
            if (undefined === rule.value2) return false;
            const far = Number(rule.value2);
            if (!numeric || !Number.isFinite(far)) {
                const [low, high] = rule.value <= rule.value2 ? [rule.value, rule.value2] : [rule.value2, rule.value];

                return text >= low && text <= high;
            }
            // Written in either order: "between 90 and 10" is what somebody
            // means when they type the bigger number first, not an empty range.
            return left >= Math.min(right, far) && left <= Math.max(right, far);
        }
        default:
            return false;
    }
}

/** Whether a range covers a cell, by numbers rather than by string surgery. */
function covers(range: string, ref: string): boolean {
    const box = parseRange(range.includes(':') ? range : `${range}:${range}`);
    const cell = parseRef(ref);
    if (!box || !cell) return false;

    const column = columnToIndex(cell.column);

    return cell.row >= box.top && cell.row <= box.bottom
        && column >= box.left && column <= box.right;
}

/**
 * The look a cell takes on, or null when no rule claims it.
 *
 *  FIRST match wins, and the rules are in the author's order. Excel applies
 * the first `<cfRule>` that matches and stops, so "red over 90" written before
 * "amber over 60" is red at 95 -- and written the other way round, everything
 * above 60 is amber and nothing is ever red. The editor keeps the order the
 * author put them in for exactly that reason, and the writer emits it.
 */
export function lookFor(
    sheet: SheetDto | undefined,
    ref: string,
    shown: string,
): ConditionalLook | null {
    const all = sheet?.conditionals;
    if (undefined === all) return null;

    for (const [range, rules] of Object.entries(all)) {
        if (!covers(range, ref)) continue;

        for (const rule of rules) {
            if (!matchesRule(shown, rule)) continue;

            return {
                background: rule.background,
                color: rule.color,
                bold: rule.bold,
                italic: rule.italic,
            };
        }
    }

    return null;
}

/** Add a rule to a range, keeping the author's order. */
export function withConditional(sheet: SheetDto, range: string, rule: ConditionalDto): SheetDto {
    if (formatsNothing(rule)) return sheet;

    const conditionals = { ...(sheet.conditionals ?? {}) };
    conditionals[range.toUpperCase()] = [...(conditionals[range.toUpperCase()] ?? []), rule];

    return { ...sheet, conditionals };
}

/**
 * Drop one rule, and the range with it when that was its last.
 *
 * An empty list left behind would be a key in the source file that says
 * nothing, which is the same reason no other optional field is written empty.
 */
export function withoutConditional(sheet: SheetDto, range: string, index: number): SheetDto {
    const all = sheet.conditionals;
    if (undefined === all || undefined === all[range]) return sheet;

    const rules = all[range].filter((_, i) => i !== index);
    const conditionals = { ...all };
    if (0 === rules.length) {
        delete conditionals[range];
    } else {
        conditionals[range] = rules;
    }

    const next = { ...sheet };
    if (0 === Object.keys(conditionals).length) {
        delete next.conditionals;
    } else {
        next.conditionals = conditionals;
    }

    return next;
}

/** Every rule that covers a cell, with the range it came from. */
export function conditionalsAt(
    sheet: SheetDto | undefined,
    ref: string,
): ReadonlyArray<{ range: string; index: number; rule: ConditionalDto }> {
    const all = sheet?.conditionals;
    if (undefined === all || '' === ref) return [];

    const out: { range: string; index: number; rule: ConditionalDto }[] = [];
    for (const [range, rules] of Object.entries(all)) {
        if (!covers(range, ref)) continue;
        rules.forEach((rule, index) => out.push({ range, index, rule }));
    }

    return out;
}
