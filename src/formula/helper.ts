/**
 * What the formula helper should show, given a formula and a caret.
 *
 * A pure function of (text, caret), deliberately: the popup is the easy half,
 * and the hard half — "what is the caret sitting inside?" — is a question about
 * tokens that can be answered and TESTED without a DOM. Every case below is a
 * string and a number.
 *
 * Two states, matching what a spreadsheet does as you type:
 *
 *   `=SU`            -> COMPLETIONS: the functions beginning `SU`
 *   `=SUM(A1, `      -> SIGNATURE:   SUM, with its second argument highlighted
 *
 * A formula being typed is almost never parseable — `=SUM(A1,` is not a tree —
 * so this reads the token stream rather than a parse. That is why the tokeniser
 * keeps whitespace: a caret offset has to land on a token, and a stream with
 * holes in it cannot say where the caret is.
 */
import { allFunctions, lookupFunction, type FormulaFunction } from './functions';
import { tokenise, type Token } from './tokenise';

export type HelperState =
    | { readonly kind: 'none' }
    | {
        readonly kind: 'completions';
        /** What the author has typed so far, e.g. `SU`. */
        readonly prefix: string;
        readonly matches: readonly FormulaFunction[];
        /** The span the accepted name replaces, as offsets into the raw text. */
        readonly from: number;
        readonly to: number;
    }
    | {
        readonly kind: 'signature';
        readonly fn: FormulaFunction;
        /** Which argument the caret is in, counting from zero. */
        readonly argumentIndex: number;
    };

const NONE: HelperState = { kind: 'none' };

/**
 * The argument names in a signature, so the popup can bold the current one.
 *
 * Split at commas OUTSIDE the square brackets: `SUM(number1, [number2, …])`
 * has two arguments, not three, and a plain `split(',')` reports the optional
 * group as two — which would then highlight the wrong one from the second
 * argument onwards.
 */
export function signatureParts(fn: FormulaFunction): readonly string[] {
    const inside = /\((.*)\)/su.exec(fn.signature)?.[1] ?? '';
    if (inside === '') return [];

    const parts: string[] = [];
    let depth = 0;
    let current = '';
    for (const ch of inside) {
        if (ch === '[') depth += 1;
        if (ch === ']') depth -= 1;
        if (ch === ',' && depth === 0) {
            parts.push(current.trim());
            current = '';
            continue;
        }
        current += ch;
    }
    parts.push(current.trim());

    return parts.filter((p) => p !== '');
}

/**
 * Which argument a signature is describing at `index`.
 *
 * The last argument of a variadic signature repeats — `SUM(number1,
 * [number2, …])` describes the fifth argument with its second part — so the
 * index is clamped rather than running off the end.
 */
export function argumentLabel(fn: FormulaFunction, index: number): string | null {
    const parts = signatureParts(fn);
    if (parts.length === 0) return null;

    return parts[Math.min(index, parts.length - 1)] ?? null;
}

/** The next significant token at or after `i`, skipping whitespace. */
function nextSignificant(tokens: readonly Token[], i: number): Token | undefined {
    for (let j = i; j < tokens.length; j++) {
        const t = tokens[j];
        if (t && t.kind !== 'whitespace') return t;
    }

    return undefined;
}

/** One reference a formula makes, and where it sits in the source text. */
export interface FormulaReference {
    /** `B3` or `B3:B5`, exactly as written. */
    readonly range: string;
    /** The sheet it was qualified with, when it was. */
    readonly sheet?: string;
    readonly from: number;
    readonly to: number;
}

/**
 * Every range a formula points at, in source order.
 *
 * ## Why this reads TOKENS rather than a regex over the text
 *
 * Because `"B4"` inside a string is not a reference and `LOG10(` is not column
 * LOG row 10 — the tokeniser already knows both, and a second implementation
 * that half-knew them would highlight cells the formula never touches, which
 * is worse than highlighting none: it teaches the author to distrust the
 * highlight.
 *
 * A `ref colon ref` triple is ONE reference. Returning `B3` and `B5` as two
 * would outline the ends of a range and not its middle, which is exactly the
 * wrong picture of `SUM(B3:B5)`.
 */
export function referencesIn(text: string): FormulaReference[] {
    if (!text.startsWith('=')) return [];

    const tokens = tokenise(text);
    const out: FormulaReference[] = [];

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (token.kind !== 'ref') continue;

        const after = nextSignificant(tokens, i + 1);
        const beyond = after?.kind === 'colon'
            ? nextSignificant(tokens, tokens.indexOf(after) + 1)
            : undefined;

        if (beyond?.kind === 'ref') {
            out.push({
                range: `${token.text}:${beyond.text}`,
                sheet: token.sheet,
                from: token.start,
                to: beyond.start + beyond.text.length,
            });
            // Skip the colon and the far end -- they are part of THIS reference.
            i = tokens.indexOf(beyond);

            continue;
        }

        out.push({
            range: token.text,
            sheet: token.sheet,
            from: token.start,
            to: token.start + token.text.length,
        });
    }

    return out;
}

/**
 * Where a CLICKED cell reference should go, or null when a click means
 * something else.
 *
 * ## Point mode, and why it has to be conditional
 *
 * In every spreadsheet, clicking a cell while a formula is being typed inserts
 * that cell's reference instead of moving the cursor. That is only true in the
 * places a reference could legally go — after `=`, an operator, `(`, `,` or `:`
 * — and everywhere else a click has to keep its ordinary meaning. Getting that
 * wrong is worse than not having the feature: an author who clicks away from a
 * half-typed formula would find their formula edited instead of left alone.
 *
 * ## Clicking again REPLACES rather than appends
 *
 * When the caret sits anywhere in a reference the author just placed, the next
 * click swaps the whole of it. Without that rule a second click produces
 * `A1B2`, which is not a formula anybody meant, and the author has to reach
 * for backspace between every two clicks.
 *
 * ##  The token the caret is INSIDE is not the token before it
 *
 * This read only the last token that ENDED before the caret, so a caret parked
 * in the middle of something was treated as though it sat just after whatever
 * came before that thing. Both halves were reported from the editor:
 *
 *   `=SUM(B|3`      -> "just after the (" -> inserted between B and 3: `=SUM(BD93`
 *   `=SU|M(B3:B5)`  -> "just after the =" -> inserted at 3: `=SUB4M(B3:B5)`
 *
 * A caret TOUCHING a reference replaces the whole reference; a caret touching a
 * name or a literal declines, because a cell reference cannot go inside either.
 *
 * Returns the span to overwrite with the new reference.
 */
/** Tokens a reference cannot be spliced into: they are values, not gaps. */
const VALUE_KINDS: readonly string[] = ['name', 'number', 'string', 'boolean', 'error'];

export function pointInsertAt(text: string, caret: number): { from: number; to: number } | null {
    if (!text.startsWith('=')) return null;

    const at = Math.max(0, Math.min(caret, text.length));
    const tokens = tokenise(text);

    // Every token the caret TOUCHES -- at a boundary it touches two, and the
    // one that matters is whichever of them a reference cannot be spliced into.
    const touching = tokens.filter(
        (t) => t.kind !== 'whitespace' && at >= t.start && at <= t.start + t.text.length,
    );

    //  A reference is replaced WHOLE, and a RANGE is one reference. Read
    // from the tokens, `B3:B5` is three of them, so a caret at its end replaced
    // only the `B5`: dragging the end of a range wrote the new range INSIDE the
    // old one, `=SUM(B3:B3:B5`. `referencesIn` already knows a `ref colon ref`
    // triple is one thing -- it is what outlines the range -- so the two agree
    // about what the author is pointing at by construction.
    const reference = referencesIn(text).find((r) => at >= r.from && at <= r.to);
    if (reference) return { from: reference.from, to: reference.to };

    // A name or a literal is something being typed, and a cell reference has no
    // place inside it. `=SU|M(` is a function, not a gap.
    if (touching.some((t) => VALUE_KINDS.includes(t.kind))) return null;

    let previous: Token | undefined;
    for (const token of tokens) {
        if (token.kind === 'whitespace') continue;
        const end = token.start + token.text.length;
        if (end > at) break;
        previous = token;
    }

    // Nothing but the leading `=` yet.
    if (undefined === previous || (previous.kind === 'operator' && previous.text === '=' && previous.start === 0)) {
        return { from: at, to: at };
    }

    return ['operator', 'open-paren', 'comma', 'colon'].includes(previous.kind)
        ? { from: at, to: at }
        : null;
}

/**
 * The state for a raw cell input and a caret offset into it.
 *
 * `text` is what the input holds, INCLUDING the leading `=` — this reads what
 * the author sees, not what the document stores, and a cell that is not a
 * formula gets no helper at all.
 */
export function helperAt(text: string, caret: number): HelperState {
    if (!text.startsWith('=')) return NONE;
    const at = Math.max(0, Math.min(caret, text.length));
    const tokens = tokenise(text);

    // --- Completions: the caret is inside or just after a bare name ---------
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (token.kind !== 'name') continue;
        const end = token.start + token.text.length;
        if (at < token.start || at > end) continue;
        // A name already followed by `(` is a CALL, and the author has finished
        // choosing it -- describing it beats offering to replace it.
        if (nextSignificant(tokens, i + 1)?.kind === 'open-paren') break;

        const prefix = text.slice(token.start, at);
        const matches = allFunctions()
            .filter((f) => f.name.startsWith(prefix.toUpperCase()))
            // SHORTEST first, then alphabetical -- the closest name to what has
            // actually been typed. Alphabetically, `=SU` offers SUBSTITUTE
            // before SUM and Enter takes it: the commonest function in the
            // sheet loses to a longer one that merely sorts earlier. That is
            // what Google Sheets does, and it is a papercut every time.
            .sort((a, b) => a.name.length - b.name.length || a.name.localeCompare(b.name));

        return matches.length === 0
            ? NONE
            : { kind: 'completions', prefix, matches, from: token.start, to: end };
    }

    // --- Signature: the innermost call the caret is inside ------------------
    interface Frame { readonly fn: FormulaFunction | null; argumentIndex: number }
    const stack: Frame[] = [];

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (token.start >= at) break;
        if (token.kind === 'whitespace') continue;

        if (token.kind === 'open-paren') {
            // A `(` is a call only when a name sits immediately before it;
            // otherwise it is ordinary grouping and describes nothing.
            const before = tokens.slice(0, i).reverse()
                .find((t) => t.kind !== 'whitespace');
            const named = before && (before.kind === 'name' || before.kind === 'boolean')
                ? lookupFunction(before.text) ?? null
                : null;
            stack.push({ fn: named, argumentIndex: 0 });
            continue;
        }
        if (token.kind === 'close-paren') {
            stack.pop();
            continue;
        }
        if (token.kind === 'comma' && stack.length > 0) {
            stack[stack.length - 1].argumentIndex += 1;
        }
    }

    for (let i = stack.length - 1; i >= 0; i--) {
        const frame = stack[i];
        if (frame.fn) {
            return { kind: 'signature', fn: frame.fn, argumentIndex: frame.argumentIndex };
        }
    }

    return NONE;
}

/**
 * Accept a completion: the text to put in the cell, and where the caret lands.
 *
 * The `(` comes with it, and the caret goes INSIDE — the author's next keystroke
 * is an argument, never the bracket they just asked for. A function taking no
 * arguments closes itself, so `TRUE()` needs no second thought.
 */
export function applyCompletion(
    text: string,
    state: Extract<HelperState, { kind: 'completions' }>,
    fn: FormulaFunction,
): { readonly text: string; readonly caret: number } {
    const closes = fn.maxArgs === 0;
    const insert = `${fn.name}(${closes ? ')' : ''}`;
    const next = text.slice(0, state.from) + insert + text.slice(state.to);

    return { text: next, caret: state.from + fn.name.length + 1 };
}
