/**
 * Finding and replacing, pinned against what a TEMPLATE needs.
 *
 * The job this exists for is renaming a `{var:…}` token across a document, so
 * the cases that earn their place are the ones where searching computed values
 * would answer differently: a token has no value, and a formula's text is not
 * its result.
 */
import {
    findMatches, matchesQuery, nextMatch, replaceIn, withReplacedAll, withReplacedIn,
} from './find-replace';
import type { SheetDto } from './sheet-document.model';

const SHEET: SheetDto = {
    cells: {
        A1: { value: 'Order number' },
        B1: { value: '{var:order.number}' },
        A2: { value: 'order total' },
        B2: { formula: 'SUM(C1:C9)' },
        C1: { value: '5', numberFormat: '@', bold: true },
    },
};

describe('matching one cell', () => {
    it('finds text anywhere in the cell, ignoring case', () => {
        expect(matchesQuery('Order number', 'order')).toBe(true);
        expect(matchesQuery('Order number', 'ORDER')).toBe(true);
    });

    it('respects case when asked', () => {
        expect(matchesQuery('Order number', 'order', { matchCase: true })).toBe(false);
        expect(matchesQuery('Order number', 'Order', { matchCase: true })).toBe(true);
    });

    it('can require the whole cell', () => {
        expect(matchesQuery('Order number', 'Order', { wholeCell: true })).toBe(false);
        expect(matchesQuery('Order number', 'order number', { wholeCell: true })).toBe(true);
    });

    it('matches nothing for an empty query', () => {
        expect(matchesQuery('anything', '')).toBe(false);
    });
});

describe('finding across a sheet', () => {
    /**
     * A token has no computed value -- it is filled when the document is
     * generated -- so a search over results could not find one at all.
     */
    it('searches what the cell holds, tokens included', () => {
        expect(findMatches(SHEET, '{var:order')).toEqual(['B1']);
    });

    /** A formula is its TEXT here: renaming a column means the one inside SUM. */
    it('searches a formula as its source', () => {
        expect(findMatches(SHEET, 'SUM(C1')).toEqual(['B2']);
    });

    /**
     * READING order, not the order the cells were written into the map -- "find
     * next" has to walk the sheet the way an author reads it.
     */
    it('returns matches in reading order', () => {
        expect(findMatches(SHEET, 'order')).toEqual(['A1', 'B1', 'A2']);
    });

    it('finds nothing for an empty query, rather than everything', () => {
        expect(findMatches(SHEET, '')).toEqual([]);
    });
});

describe('walking the matches', () => {
    const matches = ['A1', 'B1', 'A2'];

    it('goes to the next one after the cell', () => {
        expect(nextMatch(matches, 'A1')).toBe('B1');
        expect(nextMatch(matches, 'B1')).toBe('A2');
    });

    /** Every find box that has ever existed wraps; one that stops looks broken. */
    it('wraps round at the end', () => {
        expect(nextMatch(matches, 'A2')).toBe('A1');
        expect(nextMatch(matches, 'Z99')).toBe('A1');
    });

    it('starts from a cell that is not itself a match', () => {
        expect(nextMatch(matches, 'A1')).toBe('B1');
        expect(nextMatch([], 'A1')).toBeNull();
    });
});

describe('replacing text', () => {
    it('replaces every occurrence in a cell', () => {
        expect(replaceIn('a-b-a', 'a', 'x')).toBe('x-b-x');
    });

    /**
     * Case-insensitively the match and the text differ, so this cannot be a
     * split-and-join: the needle is found in the folded text and cut out of the
     * original one, which is what keeps the untouched characters as they were.
     */
    it('keeps the original case of what it does not replace', () => {
        expect(replaceIn('Order ORDER order', 'order', 'x')).toBe('x x x');
        expect(replaceIn('Order ORDER order', 'order', 'x', { matchCase: true })).toBe('Order ORDER x');
    });

    it('replaces the whole cell when asked', () => {
        expect(replaceIn('Order number', 'order number', 'Ref', { wholeCell: true })).toBe('Ref');
        expect(replaceIn('Order number', 'order', 'Ref', { wholeCell: true }))
            .withContext('not the whole cell, so nothing')
            .toBe('Order number');
    });
});

describe('replacing across a sheet', () => {
    it('renames a token everywhere it appears', () => {
        const { sheet, count } = withReplacedAll(SHEET, '{var:order.number}', '{var:invoice.number}');

        expect(sheet.cells['B1'].value).toBe('{var:invoice.number}');
        expect(count).toBe(1);
    });

    /** CELLS changed, which is a number an author can check by looking. */
    it('counts the cells it changed', () => {
        const { count } = withReplacedAll(SHEET, 'order', 'invoice');

        expect(count).toBe(3);
    });

    /** A replaced formula is still a formula, not the text of one. */
    it('keeps a formula a formula', () => {
        const { sheet } = withReplacedAll(SHEET, 'C1:C9', 'D1:D9');

        expect(sheet.cells['B2']).toEqual({ formula: 'SUM(D1:D9)' });
    });

    /** A replacement changes TEXT, not the look of the table it runs through. */
    it('keeps the formatting on the cells it rewrites', () => {
        const { sheet } = withReplacedAll(SHEET, '5', '6');

        expect(sheet.cells['C1']).toEqual({ value: '6', numberFormat: '@', bold: true });
    });

    it('leaves the sheet alone when nothing matches', () => {
        const { sheet, count } = withReplacedAll(SHEET, 'nothing here', 'x');

        expect(sheet).toBe(SHEET);
        expect(count).toBe(0);
    });

    it('replaces in one cell only when asked for one', () => {
        const next = withReplacedIn(SHEET, 'A1', 'order', 'invoice');

        expect(next.cells['A1'].value).toBe('invoice number');
        expect(next.cells['A2'].value).withContext('untouched').toBe('order total');
    });

    /** Emptying a cell drops it, which is how the document stays sparse. */
    it('drops a cell a replacement empties', () => {
        const next = withReplacedIn({ cells: { A1: { value: 'gone' } } }, 'A1', 'gone', '', { wholeCell: true });

        expect(next.cells['A1']).toBeUndefined();
    });
});
