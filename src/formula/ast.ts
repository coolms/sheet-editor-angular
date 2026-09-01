/** The shapes a parsed formula takes. Deliberately small: what the grammar
 * admits is what the evaluator and the formula helper both read, so a node
 * added here is a feature both must answer for. */
import type { FormulaError } from './values';

export type BinaryOperator =
    | '+' | '-' | '*' | '/' | '^'
    | '&'
    | '=' | '<>' | '<' | '>' | '<=' | '>=';

export type FormulaNode =
    | { readonly kind: 'number'; readonly value: number }
    | { readonly kind: 'text'; readonly value: string }
    | { readonly kind: 'boolean'; readonly value: boolean }
    | { readonly kind: 'error'; readonly code: FormulaError }
    /** `B4`, or `Sheet2!B4`. `sheet` absent means the formula's own sheet. */
    | { readonly kind: 'ref'; readonly ref: string; readonly sheet?: string }
    /** `A1:B7`, normalised so `from` is the top-left corner. */
    | { readonly kind: 'range'; readonly from: string; readonly to: string; readonly sheet?: string }
    | { readonly kind: 'unary'; readonly op: '-' | '+'; readonly operand: FormulaNode }
    /** `50%` — postfix, and it binds tighter than `^`. */
    | { readonly kind: 'percent'; readonly operand: FormulaNode }
    | {
        readonly kind: 'binary';
        readonly op: BinaryOperator;
        readonly left: FormulaNode;
        readonly right: FormulaNode;
    }
    | { readonly kind: 'call'; readonly name: string; readonly args: readonly FormulaNode[] }
    /**
     * A defined name used as a value: `SUM(items_amount)`.
     *
     *  Distinct from `call`, which is a name FOLLOWED BY `(`. This one stands
     * for a range the workbook declares elsewhere, and resolving it needs the
     * document — so the parser records the name and the evaluator looks it up.
     */
    | { readonly kind: 'name'; readonly name: string };

/** Where a parse gave up, and on what. */
export interface FormulaParseError {
    readonly message: string;
    /** Index into the formula text, so a helper can put a caret under it. */
    readonly at: number;
}

export type ParseResult =
    | { readonly ok: true; readonly node: FormulaNode }
    | { readonly ok: false; readonly error: FormulaParseError };
