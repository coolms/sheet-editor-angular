/**
 * The clipboard shape, pinned against what a SPREADSHEET puts there.
 *
 * Tab-separated with newlines between rows is not a house format: it is what
 * Excel, Google Sheets and every database console already speak, and getting it
 * right is the difference between a grid you can move data into and one you
 * have to retype into.
 */
import { pasteOffset, pastedRange, parseClipboardText, rangeCorner, toClipboardText } from './clipboard';
import { offsetFormula, withPastedBlock, type SheetDto } from './sheet-document.model';

const SHEET: SheetDto = {
    cells: {
        A1: { value: 'Item' }, B1: { value: 'Qty' },
        A2: { value: 'Bolt' }, B2: { value: '4' },
        A3: { value: 'Gadget' }, B3: { formula: 'B2*2' },
    },
};

describe('copying cells out', () => {
    it('writes a rectangle as tabs and newlines', () => {
        expect(toClipboardText(SHEET, 'A1:B2')).toBe('Item\tQty\nBolt\t4');
    });

    it('writes an empty cell as an empty field, keeping the shape', () => {
        expect(toClipboardText(SHEET, 'A1:C1')).toBe('Item\tQty\t');
    });

    /**
     * The RAW text, formulas included. Copying the computed values instead
     * would quietly turn a table of formulas into a table of numbers, which is
     * a different document.
     */
    it('copies a formula as a formula', () => {
        expect(toClipboardText(SHEET, 'B3')).toBe('=B2*2');
    });

    it('quotes a cell that contains the separators', () => {
        const awkward: SheetDto = { cells: { A1: { value: 'one\ttwo' }, B1: { value: 'say "hi"' } } };

        expect(toClipboardText(awkward, 'A1:B1')).toBe('"one\ttwo"\t"say ""hi"""');
    });
});

describe('reading a paste', () => {
    it('reads a rectangle back', () => {
        expect(parseClipboardText('a\tb\nc\td')).toEqual([['a', 'b'], ['c', 'd']]);
    });

    it('reads the line endings a Windows spreadsheet writes', () => {
        expect(parseClipboardText('a\tb\r\nc\td')).toEqual([['a', 'b'], ['c', 'd']]);
    });

    /**
     * Every spreadsheet ends its clipboard text with a newline. Honouring it
     * would paste an empty row and clear whatever was underneath.
     */
    it('drops the trailing newline rather than pasting an empty row', () => {
        expect(parseClipboardText('a\tb\n')).toEqual([['a', 'b']]);
    });

    /**
     * A quoted cell may contain the very characters the format separates on, so
     * this cannot be `split('\n').map(split('\t'))`: a one-cell note with a line
     * break in it would arrive as two rows.
     */
    it('keeps a separator that is inside a quoted cell', () => {
        expect(parseClipboardText('"one\ttwo"\tb')).toEqual([['one\ttwo', 'b']]);
        expect(parseClipboardText('"line\nbreak"')).toEqual([['line\nbreak']]);
        expect(parseClipboardText('"say ""hi"""')).toEqual([['say "hi"']]);
    });

    it('round-trips whatever was copied', () => {
        const text = toClipboardText(SHEET, 'A1:B3');

        expect(parseClipboardText(text)).toEqual([
            ['Item', 'Qty'],
            ['Bolt', '4'],
            ['Gadget', '=B2*2'],
        ]);
    });
});

describe('where a paste lands', () => {
    it('reports the range a block fills', () => {
        expect(pastedRange('C3', [['a', 'b'], ['c', 'd']])).toBe('C3:D4');
    });

    it('measures the offset from one corner to another', () => {
        expect(pasteOffset('B2', 'C5')).toEqual({ rows: 3, columns: 1 });
    });

    it('takes the top-left corner of a range', () => {
        expect(rangeCorner('B2:D9')).toBe('B2');
        expect(rangeCorner('D9')).toBe('D9');
    });
});

/**
 * Copying a total down a column is the commonest thing anybody does in a
 * spreadsheet, and it only works if the references move with it.
 */
describe('moving a formula with its paste', () => {
    it('moves a relative reference by the offset', () => {
        expect(offsetFormula('SUM(B2:B4)', 0, 1)).toBe('SUM(C2:C4)');
        expect(offsetFormula('A1+1', 2, 0)).toBe('A3+1');
    });

    /** `$` is what makes a part absolute, and an absolute part does not move. */
    it('leaves an anchored part exactly where it is', () => {
        expect(offsetFormula('$B$2', 3, 3)).toBe('$B$2');
        expect(offsetFormula('$B2', 3, 3)).toBe('$B5');
        expect(offsetFormula('B$2', 3, 3)).toBe('E$2');
    });

    it('does not touch a reference-shaped thing inside a string', () => {
        expect(offsetFormula('IF(A1="B2","yes","no")', 1, 0)).toBe('IF(A2="B2","yes","no")');
    });

    it('says #REF! rather than wrapping off the top of the sheet', () => {
        expect(offsetFormula('A1', -5, 0)).toBe('#REF!');
    });

    it('pastes a block, offsetting the formulas in it', () => {
        const next = withPastedBlock(SHEET, 'D1', [['=A1+1']], { rows: 0, columns: 3 });

        expect(next.cells['D1']).toEqual({ formula: 'D1+1' });
    });

    /** Text from elsewhere has no origin, so a formula in it is what it says. */
    it('leaves a formula alone when the text came from outside', () => {
        const next = withPastedBlock(SHEET, 'D1', [['=A1+1']], null);

        expect(next.cells['D1']).toEqual({ formula: 'A1+1' });
    });

    /** A paste replaces CONTENT, not the look of the table it lands in. */
    it('keeps the formatting already on the cells it fills', () => {
        const formatted: SheetDto = { cells: { D1: { value: 'old', numberFormat: '@', bold: true } } };

        const next = withPastedBlock(formatted, 'D1', [['new']]);

        expect(next.cells['D1']).toEqual({ value: 'new', numberFormat: '@', bold: true });
    });

    it('clears a cell the block pastes nothing into', () => {
        const next = withPastedBlock(SHEET, 'A2', [['']]);

        expect(next.cells['A2']).toBeUndefined();
    });
});
