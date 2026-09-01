import { evaluateSheet } from './evaluate';
import { FORMULA_ORACLE, ORACLE_DATA } from './formula-oracle.fixture';
import type { CellValue } from './values';
import type { SheetCellDto, SheetDocumentDto } from '../sheet-document.model';

/**
 * This engine against an INDEPENDENT one, on the same formulas.
 *
 * ##  Why an oracle and not more hand-written expectations
 *
 * `formula.spec.ts` says what this engine should do, and it was written by the
 * same hand that wrote the engine — so the two share their assumptions, and a
 * function that is confidently wrong is confidently asserted. The header of
 * `functions.ts` names the stake: a formula is written through to the `.xlsx`
 * VERBATIM, so "anything computed differently here would be a preview that
 * disagrees with the document it previews."
 *
 * The expectations in `formula-oracle.fixture.ts` were computed by LibreOffice
 * from a workbook carrying no cached values. Nobody here chose them.
 *
 * It earned its place immediately: two formulas came back `#NAME?` here and
 * computed there, because Excel stores a post-2007 function as `_xlfn.CONCAT`
 * and this engine had never been shown one.
 */
describe('the formula engine against LibreOffice', () => {
    /**
     * The literal cells, taken from the fixture rather than restated here.
     *
     *  They used to be written out in this file as well, which is one
     * definition too many: a cell changed in the workbook and not here would
     * silently change what every formula was compared against, and the suite
     * would go red somewhere unrelated. The generator now emits both from the
     * same computed workbook.
     */
    const dataCells = (): Record<string, SheetCellDto> => {
        const cells: Record<string, SheetCellDto> = {};
        for (const [ref, cell] of Object.entries(ORACLE_DATA)) {
            cells[ref] = cell.numberFormat === undefined
                ? { value: cell.value }
                : { value: cell.value, numberFormat: cell.numberFormat };
        }

        return cells;
    };

    const render = (value: CellValue | undefined): string => {
        if (!value) return '(no value)';
        switch (value.kind) {
            case 'number': return String(value.value);
            case 'text': return value.value;
            case 'boolean': return value.value ? 'TRUE' : 'FALSE';
            case 'error': return value.code;
            default: return value.kind;
        }
    };

    /**
     *  Compared numerically when both sides are numbers. LibreOffice writes
     * `4.33333333333333` for a third; asserting string equality would fail on
     * the fifteenth digit and say nothing about the arithmetic.
     */
    const agrees = (ours: string, theirs: string): boolean => {
        if (ours === theirs) return true;
        const a = Number(ours);
        const b = Number(theirs);
        if (Number.isNaN(a) || Number.isNaN(b)) return false;

        return Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-9);
    };

    it('agrees with it on every formula in the oracle', () => {
        // The control: a fixture that shrank to nothing would pass silently.
        expect(FORMULA_ORACLE.length).toBeGreaterThan(50);
        expect(Object.keys(ORACLE_DATA).length).toBeGreaterThan(10);

        const cells = dataCells();
        FORMULA_ORACLE.forEach((testCase, index) => {
            cells['H' + (index + 1)] = { formula: testCase.formula };
        });

        const doc: SheetDocumentDto = { version: 1, sheets: { Sheet1: { cells } } };
        const evaluated = evaluateSheet(doc, 'Sheet1');

        const disagreements = FORMULA_ORACLE
            .map((testCase, index) => ({ testCase, ours: render(evaluated.get('H' + (index + 1))) }))
            .filter(({ testCase, ours }) => !agrees(ours, testCase.expected))
            .map(({ testCase, ours }) => `${testCase.formula}: ours=${ours} LibreOffice=${testCase.expected}`);

        expect(disagreements).toEqual([]);
    });

    /**
     *  The prefix that started this. Excel STORES `CONCAT` as `_xlfn.CONCAT`,
     * so an uploaded workbook carries it and every reader is expected to strip
     * it.
     */
    it('evaluates a function stored under the future-function prefix', () => {
        const doc: SheetDocumentDto = {
            version: 1,
            sheets: { Sheet1: { cells: { A1: { value: '2' }, A2: { value: '3' }, B1: { formula: '_xlfn.SUM(A1:A2)' } } } },
        };

        expect(render(evaluateSheet(doc, 'Sheet1').get('B1'))).toBe('5');
    });

    /**
     *  Where the oracle is the one that is wrong.
     *
     * The oracle is a second opinion, not an authority. Where LibreOffice and
     * Excel disagree, this engine follows EXCEL on purpose: the formula is
     * written through to the `.xlsx` verbatim, and Excel is what opens it.
     *
     * Such a case cannot BE an oracle case, because the oracle is wrong. Each
     * is listed in `DIVERGENCES` in `tools/formula-oracle-fixture.py` — which
     * prints what it excluded, so the exclusion is visible rather than silent —
     * and asserted here, where the reason can be written down.
     */
    describe('follows Excel rather than LibreOffice', () => {
        const cases: readonly [string, string, string][] = [
            ['POWER(-8,1/3)', '#NUM!', 'a fractional power of a negative is #NUM!; LibreOffice returns -2'],
            ['TRUE()>1', 'TRUE', 'a boolean ranks ABOVE every number; LibreOffice coerces TRUE to 1'],
            ['DAY(59)', '28', 'serial 59 is 28 Feb 1900 because of the phantom 29 Feb; LibreOffice says 27'],
            ['DATE(1900,2,28)', '59', 'the same phantom day, from the other side; LibreOffice says 60'],
        ];

        cases.forEach(([formula, expected, why]) => {
            it(`${formula} is ${expected} — ${why}`, () => {
                const doc: SheetDocumentDto = {
                    version: 1,
                    sheets: { Sheet1: { cells: { A1: { formula } } } },
                };

                expect(render(evaluateSheet(doc, 'Sheet1').get('A1'))).toBe(expected);
            });
        });

        it('answers #REF! for a lookup column past the end of the range', () => {
            const cells = dataCells();
            cells['H1'] = { formula: 'VLOOKUP(1,A2:D5,5,FALSE())' };
            const doc: SheetDocumentDto = { version: 1, sheets: { Sheet1: { cells } } };

            //  #REF!, not #VALUE!: there is no fifth column, which is a
            // reference problem. LibreOffice calls it a value problem.
            expect(render(evaluateSheet(doc, 'Sheet1').get('H1'))).toBe('#REF!');
        });
    });

    /**
     *  And the other direction: `_xludf.` marks a USER-DEFINED function,
     * which this engine genuinely cannot run. Stripping every prefix alike
     * would turn "I cannot do this" into a wrong answer.
     */
    it('still refuses a user-defined function', () => {
        const doc: SheetDocumentDto = {
            version: 1,
            sheets: { Sheet1: { cells: { A1: { value: '2' }, B1: { formula: '_xludf.MYMACRO(A1)' } } } },
        };

        expect(render(evaluateSheet(doc, 'Sheet1').get('B1'))).toBe('#NAME?');
    });
});
