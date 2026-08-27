/**
 * Number formats, pinned against what a WORKBOOK shows.
 *
 * The cases that earn their place are the ones where the obvious reading is
 * wrong: `m` meaning minutes beside an hour and months beside a day, Excel's
 * phantom 29 February 1900, and every code outside the supported subset having
 * to come back as the raw value rather than as an approximation of itself.
 */
import {
    dateToSerial, editForm, formatCellValue, isDateFormat, parseDateInput, serialToDate,
} from './number-format';

describe('date serials', () => {
    /**
     * Excel believes 1900 was a leap year. Serial 60 is a day that did not
     * happen, and every date after it is offset by one from a naive count —
     * which is why a converter that quietly fixed the bug would be a day out
     * from the workbook for every date since February 1900.
     */
    it('agrees with the workbook on both sides of the phantom leap day', () => {
        expect(serialToDate(1)).toEqual(jasmine.objectContaining({ year: 1900, month: 1, day: 1 }));
        expect(serialToDate(59)).toEqual(jasmine.objectContaining({ year: 1900, month: 2, day: 28 }));
        expect(serialToDate(60)).toEqual(jasmine.objectContaining({ year: 1900, month: 2, day: 29 }));
        expect(serialToDate(61)).toEqual(jasmine.objectContaining({ year: 1900, month: 3, day: 1 }));
    });

    /** 45292 is 1 January 2024 in every spreadsheet there is. */
    it('agrees with the workbook on a modern date', () => {
        expect(serialToDate(45292)).toEqual(jasmine.objectContaining({ year: 2024, month: 1, day: 1 }));
        expect(dateToSerial(2024, 1, 1)).toBe(45292);
    });

    it('round-trips', () => {
        for (const [y, m, d] of [[1900, 1, 1], [1999, 12, 31], [2026, 8, 21], [2100, 2, 28]]) {
            const back = serialToDate(dateToSerial(y, m, d));
            expect([back?.year, back?.month, back?.day]).withContext(`${y}-${m}-${d}`).toEqual([y, m, d]);
        }
    });

    it('carries the time of day in the fraction', () => {
        const noon = serialToDate(45292.5);
        expect([noon?.hour, noon?.minute]).toEqual([12, 0]);
    });

    it('reports the weekday, counting Sunday as one', () => {
        // 1 January 2024 was a Monday.
        expect(serialToDate(45292)?.weekday).toBe(2);
    });
});

describe('formatting a value', () => {
    it('leaves a value alone when nothing says otherwise', () => {
        expect(formatCellValue('1234.5', undefined)).toBe('1234.5');
        expect(formatCellValue('1234.5', 'General')).toBe('1234.5');
    });

    /**
     * `@` is the author's declaration that the content is not arithmetic
     * (#1977). Formatting it as a number would undo exactly that.
     */
    it('leaves a text-declared cell alone', () => {
        expect(formatCellValue('00412', '@')).toBe('00412');
    });

    it('leaves text alone under a numeric format', () => {
        expect(formatCellValue('Widget', '#,##0.00')).toBe('Widget');
    });

    it('groups thousands and fixes the decimals', () => {
        expect(formatCellValue('1234.5', '#,##0.00')).toBe('1,234.50');
        expect(formatCellValue('1234567', '#,##0')).toBe('1,234,567');
        expect(formatCellValue('5', '0')).toBe('5');
        expect(formatCellValue('5.678', '0.00')).toBe('5.68');
    });

    it('pads to the zeros the code asks for', () => {
        expect(formatCellValue('42', '00000')).toBe('00042');
    });

    it('shows a trailing # only when there is a digit for it', () => {
        expect(formatCellValue('1.5', '0.##')).toBe('1.5');
        expect(formatCellValue('1.5', '0.00')).toBe('1.50');
    });

    it('multiplies a percentage by a hundred, as the format implies', () => {
        expect(formatCellValue('0.256', '0.00%')).toBe('25.60%');
    });

    it('keeps the literals around the number', () => {
        expect(formatCellValue('9.5', '"$"#,##0.00')).toBe('$9.50');
        expect(formatCellValue('9.5', '[$€-407]#,##0.00')).toBe('€9.50');
    });

    /** A negative section carries its own sign — brackets instead of a minus. */
    it('uses the negative section when there is one', () => {
        expect(formatCellValue('-1234.5', '#,##0.00;(#,##0.00)')).toBe('(1,234.50)');
        expect(formatCellValue('-1234.5', '#,##0.00')).toBe('-1,234.50');
    });

    it('uses the zero section when there is one', () => {
        expect(formatCellValue('0', '0.00;-0.00;"—"')).toBe('—');
    });

    it('drops a colour modifier, which the grid owns', () => {
        expect(formatCellValue('-5', '0.00;[Red]-0.00')).toBe('-5.00');
    });

    describe('dates', () => {
        // 21 August 2026.
        const SERIAL = String(dateToSerial(2026, 8, 21));

        it('renders the format the toolbar offers', () => {
            expect(formatCellValue(SERIAL, 'dd/mm/yyyy')).toBe('21/08/2026');
        });

        it('renders the orderings and widths', () => {
            expect(formatCellValue(SERIAL, 'yyyy-mm-dd')).toBe('2026-08-21');
            expect(formatCellValue(SERIAL, 'd/m/yy')).toBe('21/8/26');
            expect(formatCellValue(SERIAL, 'mmm yyyy')).toBe('Aug 2026');
            expect(formatCellValue(SERIAL, 'dddd')).toBe('Friday');
        });

        /**
         * The one genuine ambiguity in the grammar: `m` is a MINUTE beside an
         * hour and a MONTH beside a day, and nothing else tells them apart.
         */
        it('reads m as minutes beside an hour and months beside a day', () => {
            const noon = String(Number(SERIAL) + 0.5);
            expect(formatCellValue(noon, 'hh:mm')).toBe('12:00');
            expect(formatCellValue(noon, 'mm/dd')).toBe('08/21');
            expect(formatCellValue(noon, 'dd/mm/yyyy hh:mm')).toBe('21/08/2026 12:00');
        });

        it('shows a twelve-hour clock when the code asks for one', () => {
            const evening = String(Number(SERIAL) + 13 / 24);
            expect(formatCellValue(evening, 'h:mm AM/PM')).toBe('1:00 PM');
        });
    });

    /**
     * ⚠️ The safety property of the whole file. OOXML's grammar is far larger
     * than this subset, and a rendered guess is a number the document does not
     * agree with — shown in the one place an author cannot check it.
     */
    describe('a code it does not understand', () => {
        it('returns the raw value rather than an approximation', () => {
            expect(formatCellValue('1234.5', '# ?/?')).withContext('a fraction').toBe('1234.5');
            expect(formatCellValue('1234.5', '0.00E+00')).withContext('scientific').toBe('1234.5');
            expect(formatCellValue('1234.5', '[h]:mm')).withContext('elapsed time').toBe('1234.5');
            expect(formatCellValue('1234.5', '#,##0,')).withContext('scaled by thousands').toBe('1234.5');
            expect(formatCellValue('45000', 'yyyy 0.00')).withContext('two grammars at once').toBe('45000');
        });
    });
});

describe('telling a date format from any other', () => {
    it('knows one when it sees one', () => {
        expect(isDateFormat('dd/mm/yyyy')).toBe(true);
        expect(isDateFormat('hh:mm')).toBe(true);
        expect(isDateFormat('#,##0.00')).toBe(false);
        expect(isDateFormat('@')).toBe(false);
        expect(isDateFormat(undefined)).toBe(false);
    });
});

describe('editing a formatted cell', () => {
    /**
     * A spreadsheet shows `21/08/2026` in the formula bar of a date cell, not
     * `46255`. Nobody wants to type a serial.
     */
    it('edits a date as a date and everything else as it is stored', () => {
        const serial = String(dateToSerial(2026, 8, 21));
        expect(editForm(serial, 'dd/mm/yyyy')).toBe('21/08/2026');
        expect(editForm('1234.5', '#,##0.00')).withContext('Excel shows the number, not 1,234.50').toBe('1234.5');
    });

    it('reads a typed date back to the serial it came from', () => {
        const serial = dateToSerial(2026, 8, 21);
        expect(parseDateInput('21/08/2026', 'dd/mm/yyyy')).toBe(serial);
        expect(parseDateInput('2026-08-21', 'dd/mm/yyyy')).toBe(serial);
        expect(parseDateInput('21.08.2026', 'dd/mm/yyyy')).toBe(serial);
    });

    /**
     * Day-first, because that is what the toolbar's format says and what most
     * of the world writes. Guessing the other way turns the fifth of March into
     * the third of May without saying so.
     */
    it('reads the day first unless a four-digit year comes first', () => {
        expect(serialToDate(parseDateInput('03/05/2026', 'dd/mm/yyyy')!))
            .toEqual(jasmine.objectContaining({ day: 3, month: 5 }));
        expect(serialToDate(parseDateInput('2026/05/03', 'dd/mm/yyyy')!))
            .toEqual(jasmine.objectContaining({ day: 3, month: 5 }));
    });

    /** A date that does not exist was a typing mistake, not a date. */
    it('refuses a date the calendar does not have', () => {
        expect(parseDateInput('31/02/2026', 'dd/mm/yyyy')).toBeNull();
        expect(parseDateInput('21/13/2026', 'dd/mm/yyyy')).toBeNull();
    });

    it('leaves anything that is not a date alone', () => {
        expect(parseDateInput('Widget', 'dd/mm/yyyy')).toBeNull();
        expect(parseDateInput('21/08/2026', '#,##0.00')).withContext('not a date cell').toBeNull();
    });
});
