/**
 * A formula's text, split into the pieces the parser reads.
 *
 * Kept separate from the parser because the FORMULA HELPER needs it too: the
 * popup that offers a function signature has to know what the caret is sitting
 * inside — a function name, an argument, a half-typed reference — and that is a
 * question about tokens, not about a finished parse tree. A formula being typed
 * is usually not parseable yet; it is always tokenisable.
 */
import { FORMULA_ERRORS, type FormulaError } from './values';

export type TokenKind =
    | 'number' | 'string' | 'boolean' | 'error'
    | 'ref' | 'name'
    | 'operator' | 'open-paren' | 'close-paren' | 'comma' | 'colon'
    | 'whitespace';

export interface Token {
    readonly kind: TokenKind;
    readonly text: string;
    /** Index of the token's first character in the source. */
    readonly start: number;
    /** For `ref`: the sheet part, when the reference was qualified. */
    readonly sheet?: string;
}

const OPERATORS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%'];

/**
 * `A1`, `$A$1`, `AA100` — the `$` is accepted and dropped, since this engine
 * does not move formulas, so absolute and relative mean the same thing here.
 *
 * ⚠️ The lookahead excludes `(` as well as word characters, so a NAME that
 * happens to look like a reference — `LOG10(` is the standing example — is a
 * call and not column LOG row 10. The backend's `RowExpansionMap::REFERENCE`
 * has carried that guard from the start and this had not; no function with
 * digits in its name is registered yet, so nothing was broken, which is
 * exactly why it would have been found the hard way. Surfaced by
 * `referencesIn()`, which highlights whatever this calls a reference.
 */
const REF = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)(?![\w.(])/;
/** A bare name: a function, or `TRUE`/`FALSE`. */
const NAME = /^[A-Za-z_][\w.]*/;
/** `Sheet1!` or `'My Sheet'!` in front of a reference. */
const SHEET_PREFIX = /^(?:'((?:[^']|'')+)'|([A-Za-z_][\w. ]*))!/;

/**
 * Tokens, in order, including whitespace.
 *
 * Whitespace is KEPT rather than skipped: the helper needs to map a caret
 * offset back to a token, and a stream with holes in it cannot do that. The
 * parser drops it.
 */
export function tokenise(input: string): Token[] {
    const out: Token[] = [];
    let i = 0;

    while (i < input.length) {
        const rest = input.slice(i);
        const ch = input[i];

        if (/\s/u.test(ch)) {
            const ws = /^\s+/u.exec(rest)![0];
            out.push({ kind: 'whitespace', text: ws, start: i });
            i += ws.length;
            continue;
        }

        // A quoted string. `""` inside is a literal quote, as in a spreadsheet.
        if (ch === '"') {
            let j = i + 1;
            let value = '';
            while (j < input.length) {
                if (input[j] === '"') {
                    if (input[j + 1] === '"') {
                        value += '"';
                        j += 2;
                        continue;
                    }
                    break;
                }
                value += input[j];
                j += 1;
            }
            const closed = input[j] === '"';
            out.push({ kind: 'string', text: value, start: i });
            i = closed ? j + 1 : input.length;
            continue;
        }

        // An error literal typed by hand, e.g. `=IFERROR(x, #N/A)`.
        const errorLiteral = FORMULA_ERRORS.find((e) => rest.toUpperCase().startsWith(e));
        if (errorLiteral !== undefined) {
            out.push({ kind: 'error', text: errorLiteral satisfies FormulaError, start: i });
            i += errorLiteral.length;
            continue;
        }

        // A number BEFORE a reference, so `1` is not read as part of a name.
        const number = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|^\.\d+/.exec(rest);
        if (number) {
            out.push({ kind: 'number', text: number[0], start: i });
            i += number[0].length;
            continue;
        }

        // A sheet-qualified reference: the prefix only counts when a real
        // reference follows it, so `SUM!` is not silently accepted.
        const prefix = SHEET_PREFIX.exec(rest);
        if (prefix) {
            const after = rest.slice(prefix[0].length);
            const ref = REF.exec(after);
            if (ref) {
                const sheet = (prefix[1] ?? prefix[2] ?? '').replace(/''/gu, "'");
                out.push({
                    kind: 'ref',
                    text: (ref[1] + ref[2]).toUpperCase(),
                    start: i,
                    sheet,
                });
                i += prefix[0].length + ref[0].length;
                continue;
            }
        }

        const ref = REF.exec(rest);
        if (ref) {
            out.push({ kind: 'ref', text: (ref[1] + ref[2]).toUpperCase(), start: i });
            i += ref[0].length;
            continue;
        }

        const name = NAME.exec(rest);
        if (name) {
            const upper = name[0].toUpperCase();
            out.push({
                kind: upper === 'TRUE' || upper === 'FALSE' ? 'boolean' : 'name',
                text: name[0],
                start: i,
            });
            i += name[0].length;
            continue;
        }

        if (ch === '(') { out.push({ kind: 'open-paren', text: ch, start: i }); i += 1; continue; }
        if (ch === ')') { out.push({ kind: 'close-paren', text: ch, start: i }); i += 1; continue; }
        if (ch === ',' || ch === ';') { out.push({ kind: 'comma', text: ch, start: i }); i += 1; continue; }
        if (ch === ':') { out.push({ kind: 'colon', text: ch, start: i }); i += 1; continue; }

        const op = OPERATORS.find((o) => rest.startsWith(o));
        if (op !== undefined) {
            out.push({ kind: 'operator', text: op, start: i });
            i += op.length;
            continue;
        }

        // Anything else is a character the grammar has no place for. Emitted as
        // an operator token so the parser reports WHERE, rather than the
        // tokeniser throwing and losing the position.
        out.push({ kind: 'operator', text: ch, start: i });
        i += 1;
    }

    return out;
}

/** Tokens with whitespace removed — what the parser consumes. */
export const significant = (tokens: readonly Token[]): Token[] =>
    tokens.filter((t) => t.kind !== 'whitespace');
