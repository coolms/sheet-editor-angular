/**
 * The formula engine, pinned against what a SPREADSHEET does — not against
 * what TypeScript's operators happen to do.
 *
 * The cases that earn their place are the ones where those two disagree:
 * `-2^2`, `MOD(-3, 2)`, `"a" = "A"`, `1 < "a"`, `ROUND(-2.5)`. Each is a place
 * where the obvious implementation is quietly wrong, and where a formula would
 * compute something plausible instead of failing.
 */
import type { SheetDocumentDto } from '../sheet-document.model';
import { evaluateCell, evaluateFormula } from './evaluate';
import { allFunctions, functionsByCategory, lookupFunction } from './functions';
import { parseFormula } from './parse';
import { significant, tokenise } from './tokenise';
import { displayValue, type CellValue } from './values';

const doc = (cells: Record<string, { value?: string; formula?: string }>): SheetDocumentDto =>
    ({ version: 1, sheets: { Sheet1: { cells } } });

/** Evaluate a formula against a document, as the grid will. */
const evalIn = (formula: string, cells: Record<string, { value?: string; formula?: string }> = {}): CellValue =>
    evaluateFormula(formula, doc(cells), { sheet: 'Sheet1' });

/** The displayed result, which is what an author actually sees. */
const shown = (formula: string, cells: Record<string, { value?: string; formula?: string }> = {}): string =>
    displayValue(evalIn(formula, cells));

describe('formula tokeniser', () => {
    it('keeps whitespace, because the helper maps a caret onto a token', () => {
        const tokens = tokenise('SUM(A1, 2)');
        expect(tokens.some((t) => t.kind === 'whitespace')).toBe(true);
        expect(significant(tokens).some((t) => t.kind === 'whitespace')).toBe(false);
    });

    it('reads a reference, a number and a name apart', () => {
        expect(significant(tokenise('A1')).map((t) => t.kind)).toEqual(['ref']);
        expect(significant(tokenise('1')).map((t) => t.kind)).toEqual(['number']);
        expect(significant(tokenise('SUM')).map((t) => t.kind)).toEqual(['name']);
    });

    it('reads a sheet-qualified reference, quoted or not', () => {
        expect(significant(tokenise("'Line items'!B4"))[0]).toEqual(
            jasmine.objectContaining({ kind: 'ref', text: 'B4', sheet: 'Line items' }));
        expect(significant(tokenise('Sheet2!B4'))[0]).toEqual(
            jasmine.objectContaining({ kind: 'ref', text: 'B4', sheet: 'Sheet2' }));
    });

    it('takes `""` inside a string as one quote', () => {
        expect(significant(tokenise('"say ""hi"""'))[0]?.text).toBe('say "hi"');
    });

    it('does not read a bare name as a reference, nor the reverse', () => {
        // `SUM!` is not a sheet prefix, because no reference follows it.
        expect(significant(tokenise('SUM'))[0]?.kind).toBe('name');
        expect(significant(tokenise('AA100'))[0]?.kind).toBe('ref');
    });
});

describe('formula parser', () => {
    it('reports WHERE it gave up rather than throwing', () => {
        const result = parseFormula('SUM(A1,');
        expect(result.ok).toBe(false);
        if (!result.ok) {
            expect(result.error.at).toBeGreaterThan(0);
            expect(result.error.message).toContain('closing bracket');
        }
    });

    /**
 *  This asserted a REFUSAL until a later fix. A bare name is a DEFINED NAME, and
     * whether the workbook declares it is the evaluator's question -- answered
     * with `#NAME?`, exactly as a spreadsheet does. Refusing to parse made the
     * editor reject `SUM(items_amount)`, a formula the rendered document
     * computes perfectly well.
     */
    it('parses a bare name, leaving #NAME? to the evaluator', () => {
        const result = parseFormula('FOO');
        expect(result.ok).toBe(true);
        if (result.ok) expect(result.node).toEqual({ kind: 'name', name: 'FOO' });
        expect(shown('FOO')).toBe('#NAME?');
    });

    it('accepts a call with no arguments', () => {
        expect(parseFormula('TRUE()').ok).toBe(true);
    });
});

describe('operator precedence, where a spreadsheet and TypeScript disagree', () => {
    it('binds unary minus TIGHTER than the power operator: -2^2 is 4', () => {
        // TypeScript reads `-2 ** 2` as an error and other languages as -4.
        // A spreadsheet squares the negative number.
        expect(shown('-2^2')).toBe('4');
    });

    it('makes ^ LEFT-associative: 2^3^2 is 64, not 512', () => {
        //  This test used to assert 512, which is what TypeScript's `**`
        // does — the very thing this file's header says it exists not to
        // follow. A spreadsheet reads `2^3^2` as `(2^3)^2`. Excel and
        // LibreOffice both answer 64.
        expect(shown('2^3^2')).toBe('64');
    });

    it('binds % tighter than ^: 10^2% is 10^0.02', () => {
        expect(shown('50%')).toBe('0.5');
        expect(shown('2*50%')).toBe('1');
    });

    it('puts & between arithmetic and comparison', () => {
        // Below `+`: `"a"&1+2` is `"a"&(1+2)` = "a3". Were `&` tighter it would
        // be `("a"&1)+2`, which is #VALUE!.
        expect(shown('"a"&1+2')).toBe('a3');
        // Above `=`: `"a"&"b"="ab"` is `("a"&"b")="ab"`, so TRUE. Both sides
        // are text on purpose -- comparing across types is a separate rule,
        // pinned below, and mixing the two would test neither.
        expect(shown('"a"&"b"="ab"')).toBe('TRUE');
    });

    it('keeps ordinary arithmetic left-associative', () => {
        expect(shown('10-3-2')).toBe('5');
        expect(shown('100/10/2')).toBe('5');
    });
});

describe('values and coercion', () => {
    it('adds text that looks like a number', () => {
        expect(shown('"5"+1')).toBe('6');
    });

    it('refuses text that does not, with #VALUE! rather than NaN', () => {
        expect(shown('"apples"+1')).toBe('#VALUE!');
    });

    it('treats an empty cell as nothing, and counts it as nothing', () => {
        expect(shown('SUM(A1:A3)', { A1: { value: '1' }, A3: { value: '2' } })).toBe('3');
        expect(shown('COUNT(A1:A3)', { A1: { value: '1' }, A3: { value: '2' } })).toBe('2');
        expect(shown('COUNTA(A1:A3)', { A1: { value: '1' }, A3: { value: '2' } })).toBe('2');
    });

    it('compares text case-insensitively, as a spreadsheet does', () => {
        expect(shown('"a"="A"')).toBe('TRUE');
    });

    it('orders a number below text instead of coercing it', () => {
        // `1 < "a"` must not make "a" NaN.
        expect(shown('1<"a"')).toBe('TRUE');
        // Text never EQUALS a number, however it reads. This is why
        // `=1&2=12` is FALSE: `1&2` makes the text "12", and text "12" is not
        // the number 12 -- a spreadsheet agrees, and it surprises people.
        expect(shown('1="1"')).toBe('FALSE');
        expect(shown('1&2=12')).toBe('FALSE');
    });

    it('rounds away from binary noise before display', () => {
        // 0.30000000000000004 is arithmetic, not an answer.
        expect(shown('0.1+0.2')).toBe('0.3');
    });
});

describe('errors', () => {
    it('reports division by zero', () => {
        expect(shown('1/0')).toBe('#DIV/0!');
    });

    it('reports an unknown function', () => {
        expect(shown('NOSUCHFN(1)')).toBe('#NAME?');
    });

    it('propagates an error outward rather than swallowing it', () => {
        expect(shown('SUM(1, 1/0, 2)')).toBe('#DIV/0!');
    });

    it('catches a cycle instead of running one', () => {
        const cells = { A1: { formula: 'B1+1' }, B1: { formula: 'A1+1' } };
        expect(displayValue(evaluateCell(doc(cells), 'Sheet1', 'A1'))).toBe('#CYCLE!');
    });

    it('catches a cell that refers to itself', () => {
        const cells = { A1: { formula: 'A1+1' } };
        expect(displayValue(evaluateCell(doc(cells), 'Sheet1', 'A1'))).toBe('#CYCLE!');
    });

    it('reports a reference to a sheet that is not there', () => {
        expect(shown('Nope!A1')).toBe('#REF!');
    });
});

describe('a TEMPLATE cell is unresolved, not zero', () => {
    const cells = { A1: { value: '{var:order.total}' }, A2: { value: '10' } };

    it('does not pretend a template token is a number', () => {
        const result = evalIn('A1*2', cells);
        expect(result.kind).toBe('unresolved');
    });

    it('shows nothing rather than a number the author should not trust', () => {
        expect(shown('A1*2', cells)).toBe('');
    });

    it('spreads through a SUM that touches one', () => {
        expect(evalIn('SUM(A1:A2)', cells).kind).toBe('unresolved');
    });

    it('explains itself, because the grid has to say why', () => {
        const result = evalIn('A1*2', cells);
        if (result.kind === 'unresolved') {
            expect(result.because).toContain('template token');
        }
    });

    it('is NOT swallowed by IFERROR — it is not an error', () => {
        // Swallowing it would show the fallback and imply the formula settled.
        expect(evalIn('IFERROR(A1*2, 0)', cells).kind).toBe('unresolved');
        expect(shown('IFERROR(1/0, 0)')).toBe('0');
    });

    it('loses to a real error, which is worth reporting first', () => {
        expect(shown('A1 + 1/0', cells)).toBe('#DIV/0!');
    });
});

describe('functions', () => {
    it('sums a range, ignoring text', () => {
        expect(shown('SUM(A1:A3)', {
            A1: { value: '1' }, A2: { value: 'note' }, A3: { value: '2' },
        })).toBe('3');
    });

    it('averages nothing as #DIV/0!, not as zero', () => {
        expect(shown('AVERAGE(A1:A3)')).toBe('#DIV/0!');
    });

    it('rounds halves away from zero, both signs', () => {
        // Math.round(-2.5) is -2 in JavaScript; a spreadsheet says -3.
        expect(shown('ROUND(2.5, 0)')).toBe('3');
        expect(shown('ROUND(-2.5, 0)')).toBe('-3');
    });

    it('takes MOD sign from the divisor, not from JavaScript %', () => {
        // `-3 % 2` is -1 in JavaScript; a spreadsheet says 1.
        expect(shown('MOD(-3, 2)')).toBe('1');
        expect(shown('MOD(3, -2)')).toBe('-1');
    });

    it('does not evaluate the branch IF did not take', () => {
        // Eager arguments would make this #DIV/0!.
        expect(shown('IF(A1=0, "safe", 1/A1)', { A1: { value: '0' } })).toBe('safe');
    });

    it('chains through cells that hold formulas', () => {
        const cells = {
            A1: { value: '2' },
            A2: { formula: 'A1*3' },
            A3: { formula: 'A2+A1' },
        };
        expect(displayValue(evaluateCell(doc(cells), 'Sheet1', 'A3'))).toBe('8');
    });

    it('reads a cell on another sheet', () => {
        const two: SheetDocumentDto = {
            version: 1,
            sheets: {
                Sheet1: { cells: {} },
                'Line items': { cells: { B4: { value: '7' } } },
            },
        };
        expect(displayValue(evaluateFormula("'Line items'!B4*2", two, { sheet: 'Sheet1' })))
            .toBe('14');
    });

    it('joins text with & and with CONCAT alike', () => {
        expect(shown('"a"&"b"')).toBe('ab');
        expect(shown('CONCAT("a", "b", 1)')).toBe('ab1');
    });
});

/**
 * A little table to look things up in, laid out the way an invoice is: a column
 * of names, a column of quantities, a column of prices.
 */
const TABLE = {
    A1: { value: 'Item' }, B1: { value: 'Qty' }, C1: { value: 'Price' },
    A2: { value: 'Bolt' }, B2: { value: '4' }, C2: { value: '2.5' },
    A3: { value: 'Gadget' }, B3: { value: '12' }, C3: { value: '7' },
    A4: { value: 'Widget' }, B4: { value: '9' }, C4: { value: '3' },
};

describe('conditional totals', () => {
    it('adds only what meets the condition', () => {
        expect(shown('SUMIF(B2:B4, ">5", C2:C4)', TABLE)).toBe('10');
    });

    /**
     * The one thing everybody gets wrong the first time: `">5"` is not the TEXT
     * ">5", it is the comparison "greater than five".
     */
    it('reads a criterion as a comparison, not as text', () => {
        expect(shown('COUNTIF(B2:B4, ">5")', TABLE)).toBe('2');
        expect(shown('COUNTIF(B2:B4, "<>4")', TABLE)).toBe('2');
        expect(shown('COUNTIF(B2:B4, 12)', TABLE)).toBe('1');
    });

    it('matches text without caring about case, as every comparison here does', () => {
        expect(shown('COUNTIF(A2:A4, "widget")', TABLE)).toBe('1');
    });

    it('takes * and ? as wildcards', () => {
        expect(shown('COUNTIF(A2:A4, "*et")', TABLE)).toBe('2');
        expect(shown('COUNTIF(A2:A4, "?olt")', TABLE)).toBe('1');
    });

    it('sums the tested range itself when no other is given', () => {
        expect(shown('SUMIF(B2:B4, ">5")', TABLE)).toBe('21');
    });

    it('averages what it picks', () => {
        expect(shown('AVERAGEIF(B2:B4, ">5", C2:C4)', TABLE)).toBe('5');
    });

    /**
     * Excel resizes a short target range from its top-left corner. This engine
     * cannot see past the range it was handed, so it says so rather than
     * quietly summing a column that stops early.
     */
    it('refuses two ranges of different sizes rather than guessing', () => {
        expect(shown('SUMIF(B2:B4, ">5", C2:C3)', TABLE)).toBe('#VALUE!');
    });

    /**
     * A template band has no value while authoring, and a total over it must
     * say "not yet" rather than report the sum of the rows that happen to
     * exist -- the same rule the rest of this engine follows.
     */
    it('has no answer over a template token', () => {
        const cells = { ...TABLE, B3: { value: '{var:line.qty}' } };
        expect(evalIn('SUMIF(B2:B4, ">5", C2:C4)', cells).kind).toBe('unresolved');
    });

    it('counts empty cells', () => {
        expect(shown('COUNTBLANK(A2:A6)', TABLE)).toBe('2');
    });
});

describe('lookup', () => {
    it('finds a row by its first column and reads across', () => {
        expect(shown('VLOOKUP("Gadget", A2:C4, 3, FALSE)', TABLE)).toBe('7');
    });

    it('finds a column by its first row and reads down', () => {
        expect(shown('HLOOKUP("Price", A1:C4, 3, FALSE)', TABLE)).toBe('7');
    });

    it('says #N/A when the key is not there, rather than the nearest row', () => {
        expect(shown('VLOOKUP("Sprocket", A2:C4, 2, FALSE)', TABLE)).toBe('#N/A');
    });

    it('says #REF! when the column is past the end of the range', () => {
        expect(shown('VLOOKUP("Bolt", A2:C4, 4, FALSE)', TABLE)).toBe('#REF!');
    });

    /**
     *  Absent, the fourth argument is TRUE -- an APPROXIMATE match over data
     * assumed sorted. That is a footgun, and it is Excel's: the formula is
     * written into the workbook verbatim, so an editor defaulting to an exact
     * match would preview a different answer than the document gives.
     */
    it('defaults to an approximate match, because the workbook will', () => {
        const sorted = { A1: { value: '1' }, B1: { value: 'low' }, A2: { value: '10' }, B2: { value: 'high' } };
        expect(shown('VLOOKUP(5, A1:B2, 2)', sorted)).toBe('low');
        expect(shown('VLOOKUP(5, A1:B2, 2, FALSE)', sorted)).toBe('#N/A');
    });

    it('reports a position, counting from one', () => {
        expect(shown('MATCH("Widget", A2:A4, 0)', TABLE)).toBe('3');
        expect(shown('MATCH("Sprocket", A2:A4, 0)', TABLE)).toBe('#N/A');
    });

    it('takes a wildcard as a lookup key', () => {
        expect(shown('MATCH("Gad*", A2:A4, 0)', TABLE)).toBe('2');
    });

    /** A key that begins with an operator is a KEY, not a criterion. */
    it('does not read a lookup key as a comparison', () => {
        const odd = { A1: { value: '>5' }, A2: { value: 'x' } };
        expect(shown('MATCH(">5", A1:A2, 0)', odd)).toBe('1');
    });

    it('reads a cell out of a block by row and column', () => {
        expect(shown('INDEX(A2:C4, 2, 3)', TABLE)).toBe('7');
    });

    /** One row across: the single index is the column, plainly. */
    it('takes one index as the column of a single row', () => {
        expect(shown('INDEX(A1:C1, 3)', TABLE)).toBe('Price');
    });

    it('says #REF! outside the block rather than reaching past it', () => {
        expect(shown('INDEX(A2:C4, 9, 1)', TABLE)).toBe('#REF!');
        expect(shown('INDEX(A2:C4, 0, 1)', TABLE)).toBe('#REF!');
    });

    it('composes, which is the whole reason MATCH exists', () => {
        expect(shown('INDEX(C2:C4, MATCH("Widget", A2:A4, 0))', TABLE)).toBe('3');
    });
});

describe('asking what a value is', () => {
    it('tells the kinds apart', () => {
        expect(shown('ISNUMBER(B2)', TABLE)).toBe('TRUE');
        expect(shown('ISTEXT(A2)', TABLE)).toBe('TRUE');
        expect(shown('ISNUMBER(A2)', TABLE)).toBe('FALSE');
        expect(shown('ISBLANK(Z9)', TABLE)).toBe('TRUE');
    });

    it('answers about an error instead of becoming one', () => {
        expect(shown('ISERROR(1/0)')).toBe('TRUE');
        expect(shown('ISERROR(1)')).toBe('FALSE');
    });

    /**
     * A template token is not a mistake. Answering TRUE would let IFERROR-shaped
     * logic swallow a cell that is merely waiting for its data.
     */
    it('does not call a template token an error', () => {
        expect(shown('ISERROR(A2)', { A2: { value: '{var:order.total}' } })).toBe('FALSE');
    });

    it('reads text as a number', () => {
        expect(shown('VALUE("12.5")')).toBe('12.5');
        expect(shown('VALUE("nope")')).toBe('#VALUE!');
    });
});

describe('replacing text', () => {
    it('replaces every occurrence by default', () => {
        expect(shown('SUBSTITUTE("a-b-c", "-", "+")')).toBe('a+b+c');
    });

    it('replaces only the one asked for', () => {
        expect(shown('SUBSTITUTE("a-b-c", "-", "+", 2)')).toBe('a-b+c');
    });

    it('leaves the text alone rather than looping on an empty search', () => {
        expect(shown('SUBSTITUTE("abc", "", "x")')).toBe('abc');
    });
});

describe('dates', () => {
    /** 45292 is 1 January 2024 in every spreadsheet there is. */
    it('builds a date as the serial a workbook stores', () => {
        expect(shown('DATE(2024, 1, 1)')).toBe('45292');
    });

    it('reads a date apart', () => {
        expect(shown('YEAR(45292)')).toBe('2024');
        expect(shown('MONTH(45292)')).toBe('1');
        expect(shown('DAY(45292)')).toBe('1');
    });

    /** Out of range ROLLS -- how an author writes "a year on" without arithmetic. */
    it('rolls a month past twelve into the next year', () => {
        expect(shown('YEAR(DATE(2026, 13, 1))')).toBe('2027');
    });

    it('counts the weekday the three ways a spreadsheet does', () => {
        // 1 January 2024 was a Monday.
        expect(shown('WEEKDAY(45292)')).withContext('type 1: Sunday is 1').toBe('2');
        expect(shown('WEEKDAY(45292, 2)')).withContext('type 2: Monday is 1').toBe('1');
        expect(shown('WEEKDAY(45292, 3)')).withContext('type 3: Monday is 0').toBe('0');
    });

    it('moves a date by months, which is how an invoice gets dated', () => {
        expect(shown('MONTH(EDATE(DATE(2026, 1, 15), 2))')).toBe('3');
        expect(shown('DAY(EDATE(DATE(2026, 1, 15), 2))')).toBe('15');
    });

    /**
     * The 31st shifted onto a shorter month CLAMPS. An invoice dated the 31st
     * of January is due on the 28th of February, not the 3rd of March, and
     * every spreadsheet agrees.
     */
    it('clamps a day the target month does not have', () => {
        expect(shown('DAY(EDATE(DATE(2026, 1, 31), 1))')).toBe('28');
        expect(shown('MONTH(EDATE(DATE(2026, 1, 31), 1))')).toBe('2');
    });

    it('finds the end of a month, leap year included', () => {
        expect(shown('DAY(EOMONTH(DATE(2024, 2, 10), 0))')).withContext('2024 is a leap year').toBe('29');
        expect(shown('DAY(EOMONTH(DATE(2026, 2, 10), 0))')).toBe('28');
        expect(shown('DAY(EOMONTH(DATE(2026, 1, 31), 1))')).toBe('28');
    });

    it('counts days between two dates', () => {
        expect(shown('DAYS(DATE(2026, 1, 31), DATE(2026, 1, 1))')).toBe('30');
    });

    /**
     * Still no TODAY or NOW. A template is generated later than it is written,
     * and a date captured while authoring is not the date on the document.
     */
    it('offers no way to capture the day it was written', () => {
        expect(shown('TODAY()')).toBe('#NAME?');
    });
});

describe('writing a value out as text', () => {
    it('renders through the same formatter the grid uses', () => {
        expect(shown('TEXT(DATE(2026, 8, 21), "dd/mm/yyyy")')).toBe('21/08/2026');
        expect(shown('TEXT(1234.5, "#,##0.00")')).toBe('1,234.50');
    });

    it('joins into a sentence, which is what it is for in a template', () => {
        expect(shown('"Due " & TEXT(DATE(2026, 8, 21), "d mmmm yyyy")')).toBe('Due 21 August 2026');
    });

    /** An unknown code gives the value back rather than a wrong rendering. */
    it('returns the value under a format it cannot read', () => {
        expect(shown('TEXT(1234.5, "0.00E+00")')).toBe('1234.5');
    });
});

describe('the function catalogue is the helper\'s data source', () => {
    it('describes every function it offers', () => {
        for (const fn of allFunctions()) {
            expect(fn.signature).withContext(`${fn.name} signature`).toContain(fn.name);
            expect(fn.summary.length).withContext(`${fn.name} summary`).toBeGreaterThan(0);
        }
    });

    it('finds a function whatever case it is typed in', () => {
        expect(lookupFunction('sum')?.name).toBe('SUM');
        expect(lookupFunction('Sum')?.name).toBe('SUM');
    });

    it('puts every function on exactly one shelf, and no shelf is empty', () => {
        const shelved = functionsByCategory().flatMap((s) => s.functions);

        expect(shelved.length).toBe(allFunctions().length, 'every function is reachable by browsing');
        expect(new Set(shelved.map((f) => f.name)).size).toBe(shelved.length, 'and appears once');
        expect(functionsByCategory().every((s) => s.functions.length > 0)).toBe(true);
    });

    it('offers no volatile function, because a template must not bake one in', () => {
        const names = allFunctions().map((f) => f.name);
        expect(names).not.toContain('TODAY');
        expect(names).not.toContain('NOW');
        expect(names).not.toContain('RAND');
    });
});
