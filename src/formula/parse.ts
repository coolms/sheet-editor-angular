/**
 * Tokens to a tree, by precedence climbing.
 *
 * The precedence table is the spreadsheet's, not TypeScript's, and the two
 * differ in ways that matter: `^` is RIGHT-associative, `%` is a postfix
 * operator binding tighter than `^`, `&` sits between arithmetic and
 * comparison, and unary minus binds tighter than `^` so `-2^2` is `4` rather
 * than `-4`. Getting any of those wrong produces a formula that computes
 * something plausible and wrong, which is the worst kind.
 *
 * Never throws. A formula being typed is usually incomplete, and the helper
 * needs a POSITION to point at rather than an exception to catch.
 */
import type { BinaryOperator, FormulaNode, ParseResult } from './ast';
import { significant, tokenise, type Token } from './tokenise';
import type { FormulaError } from './values';

/** Binding power. Higher binds tighter. */
const BINARY_PRECEDENCE: Readonly<Record<string, number>> = {
    '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
    '&': 2,
    '+': 3, '-': 3,
    '*': 4, '/': 4,
    '^': 6,
};
/**
 * ABOVE `^`, which is the whole point and easy to get wrong.
 *
 * A spreadsheet reads `-2^2` as `(-2)^2` = 4, where almost every programming
 * language reads it as `-(2^2)` = -4. Put this below `^` and the engine
 * computes the second one -- a plausible number, silently wrong, in the kind of
 * formula nobody re-checks.
 */
const UNARY_PRECEDENCE = 7;

class Parser {
    private index = 0;

    constructor(private readonly tokens: readonly Token[], private readonly source: string) {}

    private peek(): Token | undefined {
        return this.tokens[this.index];
    }

    private next(): Token | undefined {
        const t = this.tokens[this.index];
        this.index += 1;

        return t;
    }

    private get position(): number {
        return this.peek()?.start ?? this.source.length;
    }

    parse(): ParseResult {
        if (this.tokens.length === 0) {
            return { ok: false, error: { message: 'A formula cannot be empty.', at: 0 } };
        }
        const node = this.expression(0);
        if (!node.ok) return node;
        const leftover = this.peek();
        if (leftover) {
            return {
                ok: false,
                error: {
                    message: `Unexpected ${describe(leftover)} after the end of the formula.`,
                    at: leftover.start,
                },
            };
        }

        return node;
    }

    /** Everything binding at least as tightly as `minimum`. */
    private expression(minimum: number): ParseResult {
        let left = this.unary();
        if (!left.ok) return left;

        for (;;) {
            const token = this.peek();
            if (!token || token.kind !== 'operator') break;
            const precedence = BINARY_PRECEDENCE[token.text];
            if (precedence === undefined || precedence < minimum) break;
            this.next();

            // `^` is right-associative, so it recurses at its OWN precedence;
            // everything else recurses one higher to stay left-associative.
            // ⚠️ `^` is LEFT-associative here, as in every spreadsheet:
            // `2^3^2` is `(2^3)^2` = 64, not 512. TypeScript's `**` is
            // right-associative and this used to follow it -- which is exactly
            // what this file exists NOT to do. Measured against LibreOffice
            // (#2348), which agrees with Excel at 64.
            const right = this.expression(precedence + 1);
            if (!right.ok) return right;
            left = {
                ok: true,
                node: {
                    kind: 'binary',
                    op: token.text as BinaryOperator,
                    left: left.node,
                    right: right.node,
                },
            };
        }

        return left;
    }

    private unary(): ParseResult {
        const token = this.peek();
        if (token?.kind === 'operator' && (token.text === '-' || token.text === '+')) {
            this.next();
            const operand = this.expression(UNARY_PRECEDENCE);
            if (!operand.ok) return operand;

            return { ok: true, node: { kind: 'unary', op: token.text, operand: operand.node } };
        }

        return this.postfix();
    }

    /** `%` after a value, and it may repeat: `50%%` is a half-percent. */
    private postfix(): ParseResult {
        let node = this.primary();
        if (!node.ok) return node;
        for (;;) {
            const token = this.peek();
            if (token?.kind === 'operator' && token.text === '%') {
                this.next();
                node = { ok: true, node: { kind: 'percent', operand: node.node } };
                continue;
            }
            break;
        }

        return node;
    }

    private primary(): ParseResult {
        const token = this.next();
        if (!token) {
            return {
                ok: false,
                error: { message: 'The formula ends before it is finished.', at: this.source.length },
            };
        }

        switch (token.kind) {
            case 'number':
                return { ok: true, node: { kind: 'number', value: Number(token.text) } };
            case 'string':
                return { ok: true, node: { kind: 'text', value: token.text } };
            case 'boolean':
                // `TRUE` is a literal and `TRUE()` is a call, and a spreadsheet
                // accepts both. The tokeniser cannot tell them apart -- only
                // what FOLLOWS decides -- so the parser looks.
                return this.peek()?.kind === 'open-paren'
                    ? this.call(token)
                    : {
                        ok: true,
                        node: { kind: 'boolean', value: token.text.toUpperCase() === 'TRUE' },
                    };
            case 'error':
                return { ok: true, node: { kind: 'error', code: token.text as FormulaError } };
            case 'ref':
                return this.maybeRange(token);
            case 'name':
                return this.call(token);
            case 'open-paren': {
                const inner = this.expression(0);
                if (!inner.ok) return inner;
                const close = this.next();
                if (close?.kind !== 'close-paren') {
                    return {
                        ok: false,
                        error: { message: 'This bracket is never closed.', at: token.start },
                    };
                }

                return inner;
            }
            default:
                return {
                    ok: false,
                    error: { message: `${describe(token)} cannot start a value.`, at: token.start },
                };
        }
    }

    /** A reference, or the left half of `A1:B7`. */
    private maybeRange(first: Token): ParseResult {
        if (this.peek()?.kind !== 'colon') {
            return {
                ok: true,
                node: first.sheet !== undefined
                    ? { kind: 'ref', ref: first.text, sheet: first.sheet }
                    : { kind: 'ref', ref: first.text },
            };
        }
        this.next();
        const second = this.next();
        if (second?.kind !== 'ref') {
            return {
                ok: false,
                error: {
                    message: 'A range needs a cell on both sides of the colon.',
                    at: second?.start ?? this.source.length,
                },
            };
        }
        // The sheet of the LEFT side wins: `Sheet2!A1:B7` is all of Sheet2, and
        // a range spanning two sheets is not a thing this grammar admits.
        const sheet = first.sheet ?? second.sheet;

        return {
            ok: true,
            node: sheet !== undefined
                ? { kind: 'range', from: first.text, to: second.text, sheet }
                : { kind: 'range', from: first.text, to: second.text },
        };
    }

    private call(name: Token): ParseResult {
        if (this.peek()?.kind !== 'open-paren') {
            // A name with no `(` after it is a DEFINED NAME -- `SUM(items_amount)`.
            // Whether the workbook actually declares it is the evaluator's
            // question, answered with `#NAME?` exactly as a spreadsheet does;
            // refusing to parse here would make the editor reject a formula the
            // rendered document computes perfectly well.
            return { ok: true, node: { kind: 'name', name: name.text } };
        }
        this.next();
        const args: FormulaNode[] = [];

        if (this.peek()?.kind === 'close-paren') {
            this.next();

            return { ok: true, node: { kind: 'call', name: name.text.toUpperCase(), args } };
        }

        for (;;) {
            const arg = this.expression(0);
            if (!arg.ok) {
                // Running out of input INSIDE a call is almost always an
                // unclosed bracket, and saying which function is unclosed beats
                // "the formula ends before it is finished" -- the author is
                // usually mid-typing and needs to know what to close.
                return this.peek() === undefined
                    ? {
                        ok: false,
                        error: {
                            message: `\`${name.text}\` is missing its closing bracket.`,
                            at: this.source.length,
                        },
                    }
                    : arg;
            }
            args.push(arg.node);
            const separator = this.next();
            if (separator?.kind === 'comma') continue;
            if (separator?.kind === 'close-paren') break;

            return {
                ok: false,
                error: {
                    message: `\`${name.text}\` is missing its closing bracket.`,
                    at: separator?.start ?? this.source.length,
                },
            };
        }

        return { ok: true, node: { kind: 'call', name: name.text.toUpperCase(), args } };
    }
}

function describe(token: Token): string {
    switch (token.kind) {
        case 'operator':
            return `\`${token.text}\``;
        case 'close-paren':
            return 'a closing bracket';
        case 'comma':
            return 'a comma';
        case 'colon':
            return 'a colon';
        default:
            return `\`${token.text}\``;
    }
}

/**
 * Parse a formula WITHOUT its leading `=` — which is how the document stores
 * one, so callers never strip it twice or forget to.
 */
export function parseFormula(source: string): ParseResult {
    return new Parser(significant(tokenise(source)), source).parse();
}
