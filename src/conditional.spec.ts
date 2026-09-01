/**
 * Rules that repaint a range, pinned against the traps in comparing values.
 *
 * The cases that earn their place are where the obvious implementation is
 * quietly wrong: `"10" > "9"` as strings, a range typed with its ends the wrong
 * way round, and which of two overlapping rules a cell answers to.
 */
import {
    conditionalsAt, formatsNothing, lookFor, matchesRule, withConditional, withoutConditional,
    type ConditionalDto,
} from './conditional';
import type { SheetDto } from './sheet-document.model';

const RED: Pick<ConditionalDto, 'background'> = { background: '#FFCCCC' };

describe('answering one rule', () => {
    /**
     *  The trap: as STRINGS, "10" sorts below "9". A grid that compared text
     * would put every two-digit number below every one-digit one and colour
     * exactly the wrong rows.
     */
    it('compares numbers as numbers', () => {
        expect(matchesRule('10', { when: 'greaterThan', value: '9', ...RED })).toBe(true);
        expect(matchesRule('9', { when: 'greaterThan', value: '10', ...RED })).toBe(false);
    });

    it('compares anything else as text', () => {
        expect(matchesRule('banana', { when: 'greaterThan', value: 'apple', ...RED })).toBe(true);
        expect(matchesRule('apple', { when: 'equal', value: 'apple', ...RED })).toBe(true);
    });

    it('finds text inside a cell, ignoring case', () => {
        expect(matchesRule('Overdue payment', { when: 'contains', value: 'overdue', ...RED })).toBe(true);
        expect(matchesRule('Paid', { when: 'contains', value: 'overdue', ...RED })).toBe(false);
    });

    it('knows an empty cell', () => {
        expect(matchesRule('', { when: 'empty', ...RED })).toBe(true);
        expect(matchesRule('   ', { when: 'empty', ...RED })).withContext('spaces are empty').toBe(true);
        expect(matchesRule('0', { when: 'empty', ...RED })).toBe(false);
    });

    it('takes a range inclusively at both ends', () => {
        const rule: ConditionalDto = { when: 'between', value: '10', value2: '20', ...RED };

        expect(matchesRule('10', rule)).toBe(true);
        expect(matchesRule('20', rule)).toBe(true);
        expect(matchesRule('21', rule)).toBe(false);
    });

    /** Typing the bigger number first is a range, not an empty one. */
    it('takes a range written either way round', () => {
        expect(matchesRule('15', { when: 'between', value: '20', value2: '10', ...RED })).toBe(true);
    });

    it('matches nothing when the rule has no value to compare', () => {
        expect(matchesRule('5', { when: 'greaterThan', ...RED })).toBe(false);
    });
});

describe('which rule a cell wears', () => {
    const SHEET: SheetDto = {
        cells: { B2: { value: '95' } },
        conditionals: {
            'B1:B9': [
                { when: 'greaterThan', value: '90', background: '#FFCCCC' },
                { when: 'greaterThan', value: '60', background: '#FFF3CD' },
            ],
        },
    };

    /**
     *  FIRST match wins, in the author's order -- as Excel does. Written the
     * other way round, everything over 60 is amber and nothing is ever red.
     */
    it('takes the first rule that matches, not the last', () => {
        expect(lookFor(SHEET, 'B2', '95')?.background).toBe('#FFCCCC');
    });

    it('takes the later rule when the first does not match', () => {
        expect(lookFor(SHEET, 'B2', '75')?.background).toBe('#FFF3CD');
    });

    it('leaves a cell alone when no rule claims it', () => {
        expect(lookFor(SHEET, 'B2', '5')).toBeNull();
        expect(lookFor(SHEET, 'D4', '95')).withContext('outside the range').toBeNull();
    });

    it('leaves a sheet with no rules alone', () => {
        expect(lookFor({ cells: {} }, 'A1', 'x')).toBeNull();
    });

    it('covers a single-cell range', () => {
        const one: SheetDto = { cells: {}, conditionals: { B2: [{ when: 'empty', background: '#EEE' }] } };

        expect(lookFor(one, 'B2', '')?.background).toBe('#EEE');
        expect(lookFor(one, 'B3', '')).toBeNull();
    });
});

describe('keeping the rules', () => {
    const SHEET: SheetDto = { cells: {} };

    it('adds a rule to a range, in order', () => {
        let next = withConditional(SHEET, 'B1:B9', { when: 'greaterThan', value: '90', background: '#F00' });
        next = withConditional(next, 'B1:B9', { when: 'greaterThan', value: '60', background: '#FF0' });

        expect(next.conditionals!['B1:B9'].map(r => r.value)).toEqual(['90', '60']);
    });

    /** A rule that changes nothing would be an element Excel does nothing with. */
    it('refuses a rule that formats nothing', () => {
        expect(formatsNothing({ when: 'empty' })).toBe(true);
        expect(withConditional(SHEET, 'B1', { when: 'empty' })).toBe(SHEET);
    });

    it('removes one rule and keeps the rest', () => {
        let next = withConditional(SHEET, 'B1:B9', { when: 'greaterThan', value: '90', background: '#F00' });
        next = withConditional(next, 'B1:B9', { when: 'greaterThan', value: '60', background: '#FF0' });

        const after = withoutConditional(next, 'B1:B9', 0);

        expect(after.conditionals!['B1:B9'].map(r => r.value)).toEqual(['60']);
    });

    /** An empty list left behind would be a key in the file that says nothing. */
    it('drops the range when its last rule goes', () => {
        const one = withConditional(SHEET, 'B1', { when: 'empty', background: '#EEE' });

        expect(withoutConditional(one, 'B1', 0).conditionals).toBeUndefined();
    });

    it('lists the rules covering a cell, so they can be removed', () => {
        const next = withConditional(SHEET, 'B1:B9', { when: 'empty', background: '#EEE' });

        expect(conditionalsAt(next, 'B4').map(r => r.range)).toEqual(['B1:B9']);
        expect(conditionalsAt(next, 'D4')).toEqual([]);
    });
});
