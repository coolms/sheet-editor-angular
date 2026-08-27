/**
 * Declaring a name, and the rules that stop one shadowing a cell (#2385).
 */
import {
    definedNamesOf, nameProblem, scopedRange, withDefinedName, withoutDefinedName,
} from './defined-names';
import type { SheetDocumentDto } from './sheet-document.model';

const doc = (definedNames?: Record<string, string>): SheetDocumentDto =>
    ({ version: 1, sheets: { Sheet1: { cells: {} } }, definedNames });

describe('declaring a name', () => {
    it('accepts an ordinary one', () => {
        expect(nameProblem('items_amount')).toBeNull();
        expect(nameProblem('Total.Net')).toBeNull();
        expect(nameProblem('_private')).toBeNull();
    });

    /**
     * ⚠️ The rule that actually bites. `Q4` is a natural name for a quarter's
     * figures and it is also a cell, so a formula could not tell them apart.
     */
    it('refuses a name that is a cell reference', () => {
        expect(nameProblem('Q4')).toContain('cell reference');
        expect(nameProblem('B2')).toContain('cell reference');
        expect(nameProblem('xfd1')).toContain('cell reference');
    });

    it('refuses the single letters a spreadsheet reserves', () => {
        expect(nameProblem('R')).toContain('reserved');
        expect(nameProblem('c')).toContain('reserved');
        // Only alone -- these are ordinary names.
        expect(nameProblem('Rate')).toBeNull();
        expect(nameProblem('Cost')).toBeNull();
    });

    it('refuses a shape a spreadsheet cannot store', () => {
        expect(nameProblem('')).toContain('needed');
        expect(nameProblem('   ')).toContain('needed');
        expect(nameProblem('2fast')).toContain('starting with a letter');
        expect(nameProblem('has space')).toContain('starting with a letter');
        expect(nameProblem('has-dash')).toContain('starting with a letter');
        expect(nameProblem('a'.repeat(256))).toContain('at most');
    });

    it('refuses a name already taken, whatever its case', () => {
        expect(nameProblem('items_amount', ['items_amount'])).toContain('already used');
        expect(nameProblem('ITEMS_AMOUNT', ['items_amount'])).toContain('already used');
        expect(nameProblem('other', ['items_amount'])).toBeNull();
    });
});

describe('the range a name stands for', () => {
    it('writes it absolute, as every editor does', () => {
        expect(scopedRange('Sheet1', 'B2:B4')).toBe('Sheet1!$B$2:$B$4');
        expect(scopedRange('Sheet1', 'B2')).toBe('Sheet1!$B$2');
    });

    /** A sheet name with a space is quoted, and an embedded quote is doubled. */
    it('quotes a sheet name that needs it', () => {
        expect(scopedRange('Line items', 'B2:B4')).toBe("'Line items'!$B$2:$B$4");
        expect(scopedRange("Ann's", 'A1')).toBe("'Ann''s'!$A$1");
    });
});

describe('the document', () => {
    it('gains a name', () => {
        const next = withDefinedName(doc(), 'items_amount', 'Sheet1!$B$2:$B$4');
        expect(next.definedNames).toEqual({ items_amount: 'Sheet1!$B$2:$B$4' });
    });

    it('trims the name it stores, so a stray space cannot make two of them', () => {
        expect(withDefinedName(doc(), '  spaced  ', 'Sheet1!$A$1').definedNames).toEqual({
            spaced: 'Sheet1!$A$1',
        });
    });

    it('is not mutated in place, because undo holds the previous one', () => {
        const before = doc({ a: 'Sheet1!$A$1' });
        const after = withDefinedName(before, 'b', 'Sheet1!$B$1');

        expect(before.definedNames).toEqual({ a: 'Sheet1!$A$1' });
        expect(after.definedNames).toEqual({ a: 'Sheet1!$A$1', b: 'Sheet1!$B$1' });
    });

    it('loses one', () => {
        const next = withoutDefinedName(doc({ a: 'Sheet1!$A$1', b: 'Sheet1!$B$1' }), 'a');
        expect(next.definedNames).toEqual({ b: 'Sheet1!$B$1' });
    });

    /**
     * ⚠️ The last one takes the whole map with it. A `.dsheet` is a source file
     * an operator reads in a diff, and `"definedNames": {}` on every template is
     * noise -- the backend omits it on the same rule, so a document that has
     * never had a name and one that lost its last are byte-identical.
     */
    it('drops the map entirely when the last name goes', () => {
        const next = withoutDefinedName(doc({ only: 'Sheet1!$A$1' }), 'only');
        expect('definedNames' in next).toBe(false);
    });

    it('lists them sorted, so the panel does not reorder itself on edit', () => {
        const listed = definedNamesOf(doc({ zeta: 'Sheet1!$Z$1', alpha: 'Sheet1!$A$1' }));
        expect(listed.map((n) => n.name)).toEqual(['alpha', 'zeta']);
    });

    it('lists nothing for a document that declares none', () => {
        expect(definedNamesOf(doc())).toEqual([]);
        expect(definedNamesOf(null)).toEqual([]);
    });
});
