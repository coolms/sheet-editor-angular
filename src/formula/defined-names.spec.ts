/**
 * Defined names in the grid's preview (#2384).
 *
 * ## Why the editor has to understand these
 *
 * A formula is written through to the `.xlsx` verbatim, so `SUM(items_amount)`
 * renders correctly whether or not this editor understands it. Before this the
 * PARSER refused it outright -- "`items_amount` is not a value" -- so the grid
 * showed an error for a formula the delivered document computes perfectly well.
 * An editor that disagrees with its own output is worse than one that says
 * nothing.
 *
 * A name is also the only reference into a loop band that survives every
 * editor: `SUM(B2:B2)` is collapsed to `SUM(B2)` by LibreOffice on save, while
 * a name's range is rewritten when the band grows.
 */
import type { SheetDocumentDto } from '../sheet-document.model';
import { evaluateFormula } from './evaluate';
import { parseFormula } from './parse';
import { displayValue } from './values';

const doc = (
    cells: Record<string, { value?: string; formula?: string }>,
    definedNames?: Record<string, string>,
): SheetDocumentDto => ({ version: 1, sheets: { Sheet1: { cells } }, definedNames });

const shown = (
    formula: string,
    cells: Record<string, { value?: string; formula?: string }> = {},
    definedNames?: Record<string, string>,
): string => displayValue(evaluateFormula(formula, doc(cells, definedNames), { sheet: 'Sheet1' }));

const LINES = { B2: { value: '10' }, B3: { value: '20' }, B4: { value: '30' } };

describe('defined names', () => {
    it('parses a bare name as a value rather than refusing it', () => {
        const parsed = parseFormula('SUM(items_amount)');
        expect(parsed.ok).toBe(true);
    });

    /**
     * ⚠️ A name resolved everywhere EXCEPT inside a function call, which is the
     * only place `SUM(items_amount)` ever appears. The call path built its own
     * copy of `spread` instead of delegating, and the copy never dereferenced --
     * so this summed the range's FIRST CELL and looked like a working feature.
     *
     * Pinned as a trio because the scalar answer alone cannot tell the two
     * apart: a range in scalar position IS its first cell, so `items_amount`
     * answering 10 is correct and identical to the broken behaviour.
     */
    it('resolves a name in a call exactly as it does outside one', () => {
        const names = { items_amount: 'Sheet1!$B$2:$B$4' };
        expect(shown('SUM(B2:B4)', LINES, names)).toBe('60');
        expect(shown('items_amount', LINES, names)).toBe('10');
        expect(shown('SUM(items_amount)', LINES, names)).toBe('60');
    });

    it('is case-insensitive, as a spreadsheet is', () => {
        expect(shown('SUM(Items_Amount)', LINES, { items_amount: 'Sheet1!$B$2:$B$4' })).toBe('60');
    });

    it('answers #NAME? for a name the workbook does not declare', () => {
        expect(shown('SUM(mystery)', LINES, { items_amount: 'Sheet1!$B$2:$B$4' })).toBe('#NAME?');
    });

    it('answers #NAME? when the workbook declares nothing at all', () => {
        expect(shown('SUM(items_amount)', LINES)).toBe('#NAME?');
    });

    /**
     * ⚠️ `$` means "do not move when COPIED", which no evaluation here performs
     * -- and every editor writes a name's range with them, so treating `$B$2`
     * as different from `B2` would make every imported name unresolvable.
     */
    it('reads a range whether or not it is written with dollars', () => {
        expect(shown('SUM(items_amount)', LINES, { items_amount: 'Sheet1!B2:B4' })).toBe('60');
    });

    it('resolves a name that stands for a single cell', () => {
        expect(shown('total*2', LINES, { total: 'Sheet1!$B$3' })).toBe('40');
    });

    it('reads a quoted sheet, which is how a name with a space is written', () => {
        const two: SheetDocumentDto = {
            version: 1,
            sheets: { Sheet1: { cells: {} }, 'Line items': { cells: { B2: { value: '7' } } } },
            definedNames: { first_line: "'Line items'!$B$2" },
        };
        expect(displayValue(evaluateFormula('first_line*3', two, { sheet: 'Sheet1' }))).toBe('21');
    });

    it('answers #REF! when a name stands for something that is not a range', () => {
        expect(shown('SUM(items_amount)', LINES, { items_amount: 'Sheet1!not-a-range' })).toBe('#REF!');
    });

    /** A name in a position that wants a GRID, not a flat list. */
    it('serves a name to a function that needs rows and columns', () => {
        const cells = {
            A2: { value: 'north' }, B2: { value: '10' },
            A3: { value: 'south' }, B3: { value: '20' },
        };
        expect(shown('VLOOKUP("south",lookup_block,2,FALSE)', cells, { lookup_block: 'Sheet1!$A$2:$B$3' }))
            .toBe('20');
    });

    /** A function CALL is still a call -- the name branch must not swallow it. */
    it('still treats a name followed by a bracket as a function', () => {
        expect(shown('SUM(1,2)')).toBe('3');
        expect(shown('NOTAFUNCTION(1)')).toBe('#NAME?');
    });
});
