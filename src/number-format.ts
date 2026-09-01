/**
 * Rendering a stored value the way its number format says it looks.
 *
 * Two callers need this and they are the reason it exists at all:
 *
 * 1. **The grid.** A `.xlsx` stores a date as the SERIAL NUMBER its format
 *    describes, and the importer keeps it that way on purpose — converting to a
 *    date string would discard the format and make the value unarithmetic. So
 *    until this existed, an imported invoice showed `46255` where the generated
 *    document showed `21/08/2026`. The editor was not showing the document.
 * 2. **`TEXT(value, format)`**, which is this function with the format written
 *    out by the author instead of stored on the cell.
 *
 *  **An unrecognised format returns the RAW value.** OOXML's format grammar
 * is far larger than this, and a `.dsheet` may carry any code an author
 * hand-wrote. Rendering a code we half-understand would put a number on screen
 * that the document does not agree with — silently, and in the one place an
 * author has no way to check. Showing the underlying value is honest: it is
 * plainly not the final look, and it is never a wrong answer.
 */

/**
 * Excel's day zero, and its famous mistake.
 *
 * Serial 1 is 1 January 1900. Excel also believes 1900 was a leap year, so
 * serial 60 is "29 February 1900" — a day that did not happen. Every date after
 * it is therefore offset by one from a naive count, which is why the two
 * branches below exist rather than one addition. The bug is thirty years old,
 * every spreadsheet reproduces it deliberately, and a converter that quietly
 * fixed it would be one day out from the workbook for every date after
 * February 1900.
 */
const PHANTOM_LEAP_DAY = 60;
const MS_PER_DAY = 86_400_000;
/** Serial 61 (1 March 1900) onwards counts from here. */
const EPOCH_AFTER = Date.UTC(1899, 11, 30);
/** Serials 1..59 count from here instead — the phantom day is not yet in play. */
const EPOCH_BEFORE = Date.UTC(1899, 11, 31);

export interface DateParts {
    readonly year: number;
    /** 1-12, as an author counts them. */
    readonly month: number;
    readonly day: number;
    readonly hour: number;
    readonly minute: number;
    readonly second: number;
    /** 1 = Sunday, as `WEEKDAY`'s default type reports it. */
    readonly weekday: number;
}

/** A serial number as a calendar date, or null when it is not one. */
export function serialToDate(serial: number): DateParts | null {
    if (!Number.isFinite(serial) || serial < 0) return null;

    const whole = Math.floor(serial);
    const fraction = serial - whole;
    const seconds = Math.round(fraction * 86_400);

    if (whole === PHANTOM_LEAP_DAY) {
        // The day that never was. Reported as Excel reports it, because a
        // document containing it must round-trip unchanged.
        return { year: 1900, month: 2, day: 29, ...clock(seconds), weekday: 4 };
    }

    const ms = (whole < PHANTOM_LEAP_DAY ? EPOCH_BEFORE : EPOCH_AFTER) + whole * MS_PER_DAY;
    const date = new Date(ms);
    if (Number.isNaN(date.getTime())) return null;

    return {
        year: date.getUTCFullYear(),
        month: date.getUTCMonth() + 1,
        day: date.getUTCDate(),
        weekday: date.getUTCDay() + 1,
        ...clock(seconds),
    };
}

function clock(seconds: number): { hour: number; minute: number; second: number } {
    return {
        hour: Math.floor(seconds / 3600) % 24,
        minute: Math.floor(seconds / 60) % 60,
        second: seconds % 60,
    };
}

/**
 * A calendar date as its serial number.
 *
 * Out-of-range months and days ROLL, which is not sloppiness: `DATE(2026, 13, 1)`
 * is January 2027 in every spreadsheet, and it is how an author writes "a year
 * from now" without arithmetic.
 */
export function dateToSerial(year: number, month: number, day: number): number {
    const ms = Date.UTC(year, month - 1, day);
    if (Number.isNaN(ms)) return Number.NaN;

    const days = Math.round((ms - EPOCH_AFTER) / MS_PER_DAY);

    return days > PHANTOM_LEAP_DAY ? days : days - 1;
}

/** The pieces a format code is made of. */
type Piece =
    | { kind: 'literal'; text: string }
    | { kind: 'token'; text: string };

const TOKEN_CHARS = /[ymdhs0#?]/i;
/** The colour names OOXML allows in a format code, and nothing else. */
const COLOUR = /^(black|blue|cyan|green|magenta|red|white|yellow|color\s?[1-9][0-9]?)$/i;
const DATE_CHARS = /[ymdhs]/i;
const NUMBER_CHARS = /[0#?]/;

/**
 * Split one section of a format code into literals and tokens.
 *
 * Returns null for anything this does not understand — an elapsed-time
 * `[h]`, a fraction `# ?/?`, a scientific `E+00` — so the caller can fall back
 * to the raw value rather than render an approximation of it.
 */
function pieces(section: string): Piece[] | null {
    const out: Piece[] = [];
    let i = 0;

    const literal = (text: string): void => {
        const last = out[out.length - 1];
        if (last?.kind === 'literal') out[out.length - 1] = { kind: 'literal', text: last.text + text };
        else out.push({ kind: 'literal', text });
    };

    while (i < section.length) {
        const ch = section[i];

        if (ch === '"') {
            const end = section.indexOf('"', i + 1);
            if (end < 0) return null;
            literal(section.slice(i + 1, end));
            i = end + 1;
            continue;
        }

        if (ch === '\\') {
            if (i + 1 >= section.length) return null;
            literal(section[i + 1]);
            i += 2;
            continue;
        }

        if (ch === '[') {
            const end = section.indexOf(']', i);
            if (end < 0) return null;
            const inside = section.slice(i + 1, end);
            i = end + 1;
            // `[$€-407]` is a currency symbol and a locale; the symbol is the
            // half that shows. `[Red]` is a colour the grid already owns, so it
            // is dropped rather than obeyed. `[h]` is elapsed time and is a
            // different clock entirely -- not understood, so: raw value.
            if (inside.startsWith('$')) {
                literal(inside.slice(1).split('-')[0]);
                continue;
            }
            if (COLOUR.test(inside)) continue;

            //  NOT "any word is a colour". `[h]` is elapsed time — a clock
            // that counts past 24 hours — and treating it as an unknown colour
            // dropped the bracket and rendered `[h]:mm` as minutes alone, which
            // is a plausible wrong time. An unrecognised bracket is
            // UNSUPPORTED, and unsupported means the raw value.
            return null;
        }

        // `_x` reserves the width of x; a space is what it looks like.
        if (ch === '_') { literal(' '); i += 2; continue; }
        // `*x` repeats x to fill the column, which a grid cell does not have.
        if (ch === '*') { i += 2; continue; }

        if (/[eE]/.test(ch) && /[+-]/.test(section[i + 1] ?? '')) return null;

        // A digit placeholder takes the whole numeric pattern with it —
        // `#,##0.00` is ONE token, because the point and the commas are part of
        // the shape of the number and not decoration around it. Splitting them
        // off as literals loses where the decimal point was, and the number
        // then renders beside its own punctuation: `1235,.`
        if (NUMBER_CHARS.test(ch)) {
            let j = i;
            while (j < section.length && /[0#?.,]/.test(section[j])) j += 1;
            out.push({ kind: 'token', text: section.slice(i, j) });
            i = j;
            continue;
        }

        // A date token is a run of ONE character: `mm` is a month, `mmm` its
        // name, and `dd/mm` is two tokens with a separator between them.
        if (TOKEN_CHARS.test(ch)) {
            let j = i;
            while (j < section.length && section[j].toLowerCase() === ch.toLowerCase()) j += 1;
            out.push({ kind: 'token', text: section.slice(i, j) });
            i = j;
            continue;
        }

        const ampm = /^(AM\/PM|A\/P)/i.exec(section.slice(i));
        if (ampm) {
            out.push({ kind: 'token', text: ampm[0] });
            i += ampm[0].length;
            continue;
        }

        literal(ch);
        i += 1;
    }

    return out;
}

/**
 * Split a code into its sections: positive, negative, zero, text.
 *
 * Split on `;` OUTSIDE quotes and brackets, because `[$-409]` and `"a;b"` both
 * legitimately contain one and a plain `split(';')` would cut a code in half.
 */
function sections(code: string): string[] {
    const out: string[] = [];
    let current = '';
    let quoted = false;
    let bracketed = false;

    for (let i = 0; i < code.length; i++) {
        const ch = code[i];
        if (ch === '\\') { current += ch + (code[i + 1] ?? ''); i += 1; continue; }
        if (ch === '"') { quoted = !quoted; current += ch; continue; }
        if (!quoted && ch === '[') bracketed = true;
        if (!quoted && ch === ']') bracketed = false;
        if (ch === ';' && !quoted && !bracketed) { out.push(current); current = ''; continue; }
        current += ch;
    }
    out.push(current);

    return out;
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const pad = (n: number, width: number): string => String(Math.abs(n)).padStart(width, '0');

/**
 * Render a date, resolving the one genuine ambiguity in the grammar.
 *
 * `m` is MONTH or MINUTE depending on its neighbours: minute when it follows an
 * hour or precedes a second, month otherwise. `hh:mm` is hours and minutes,
 * `mm/dd` is a month and a day, and the only difference is what sits beside it.
 */
function renderDate(list: readonly Piece[], parts: DateParts): string | null {
    const tokens = list.map((p, i) => ({ piece: p, index: i }))
        .filter((e) => e.piece.kind === 'token');

    let out = '';
    for (let i = 0; i < list.length; i++) {
        const piece = list[i];
        if (piece.kind === 'literal') { out += piece.text; continue; }

        const lower = piece.text.toLowerCase();
        const here = tokens.findIndex((t) => t.index === i);
        const before = tokens[here - 1]?.piece.text.toLowerCase() ?? '';
        const after = tokens[here + 1]?.piece.text.toLowerCase() ?? '';

        if (/^y+$/.test(lower)) {
            out += lower.length <= 2 ? pad(parts.year % 100, 2) : pad(parts.year, 4);
            continue;
        }
        if (/^m+$/.test(lower)) {
            const isMinute = before.startsWith('h') || after.startsWith('s');
            if (isMinute) { out += lower.length >= 2 ? pad(parts.minute, 2) : String(parts.minute); continue; }
            if (lower.length >= 4) { out += MONTHS[parts.month - 1]; continue; }
            if (lower.length === 3) { out += MONTHS[parts.month - 1].slice(0, 3); continue; }
            out += lower.length === 2 ? pad(parts.month, 2) : String(parts.month);
            continue;
        }
        if (/^d+$/.test(lower)) {
            if (lower.length >= 4) { out += DAYS[parts.weekday - 1]; continue; }
            if (lower.length === 3) { out += DAYS[parts.weekday - 1].slice(0, 3); continue; }
            out += lower.length === 2 ? pad(parts.day, 2) : String(parts.day);
            continue;
        }
        if (/^h+$/.test(lower)) {
            const twelve = list.some((p) => p.kind === 'token' && /^(am\/pm|a\/p)$/i.test(p.text));
            const hour = twelve ? (parts.hour % 12 === 0 ? 12 : parts.hour % 12) : parts.hour;
            out += lower.length >= 2 ? pad(hour, 2) : String(hour);
            continue;
        }
        if (/^s+$/.test(lower)) {
            out += lower.length >= 2 ? pad(parts.second, 2) : String(parts.second);
            continue;
        }
        if (/^(am\/pm|a\/p)$/.test(lower)) {
            const pm = parts.hour >= 12;
            out += lower === 'a/p' ? (pm ? 'P' : 'A') : (pm ? 'PM' : 'AM');
            continue;
        }

        // A digit placeholder inside a date code: two grammars at once, and not
        // one this understands.
        return null;
    }

    return out;
}

/** Render a number against the `0`, `#`, `.` and `,` placeholders. */
function renderNumber(list: readonly Piece[], value: number): string | null {
    const patterns = list.filter((p) => p.kind === 'token');

    // Exactly ONE numeric pattern. Two of them means a second grammar inside
    // the code -- a fraction, `# ?/?`, is the standing example -- and rendering
    // the first while dropping the rest would be a plausible wrong number.
    if (patterns.length !== 1) return null;
    const pattern = patterns[0].text;
    if (!NUMBER_CHARS.test(pattern)) return null;

    const dot = pattern.indexOf('.');
    const integerPart = (dot < 0 ? pattern : pattern.slice(0, dot)).replace(/[^0#?]/g, '');
    const decimalPart = (dot < 0 ? '' : pattern.slice(dot + 1)).replace(/[^0#?]/g, '');

    // `,` between digit placeholders groups thousands; one at the END of them
    // divides by a thousand per comma, which is a different meaning and not one
    // this renders. Raw value rather than a number a thousand times wrong.
    const grouped = /[0#?],[0#?]/.test(pattern);
    if (/,(?![0#?])/.test(pattern)) return null;

    const percent = list.some((p) => p.kind === 'literal' && p.text.includes('%'));
    const scaled = percent ? value * 100 : value;

    const fixed = Math.abs(scaled).toFixed(decimalPart.length);
    let [whole, fraction = ''] = fixed.split('.');

    // A trailing `#` shows a digit only when there is one; a `0` always shows.
    while (fraction.length > 0 && decimalPart[fraction.length - 1] !== '0' && fraction.endsWith('0')) {
        fraction = fraction.slice(0, -1);
    }

    const minimumDigits = (integerPart.match(/0/g) ?? []).length;
    if (whole.length < minimumDigits) whole = whole.padStart(minimumDigits, '0');
    if (whole === '0' && minimumDigits === 0) whole = '';
    if (grouped) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

    const body = whole + (fraction === '' ? '' : '.' + fraction);
    const sign = scaled < 0 ? '-' : '';

    // Put the rendered number where the placeholders were, keeping the literals
    // around it: `$#,##0.00` is a currency symbol and a number, in that order.
    let out = '';
    let placed = false;
    for (const piece of list) {
        if (piece.kind === 'literal') { out += piece.text; continue; }
        if (!placed) { out += sign + body; placed = true; }
    }

    return out;
}

/** Whether a stored value is a number this can format. */
const asNumber = (value: string): number | null => {
    const trimmed = value.trim();
    if (trimmed === '') return null;
    const n = Number(trimmed);

    return Number.isFinite(n) ? n : null;
};

/** A format code that shows a date rather than a number. */
export function isDateFormat(code: string | undefined): boolean {
    if (code === undefined || code === '' || code === '@') return false;
    const list = pieces(sections(code)[0] ?? '');
    if (!list) return false;

    const tokens = list.filter((p) => p.kind === 'token').map((p) => p.text).join('');

    return DATE_CHARS.test(tokens) && !NUMBER_CHARS.test(tokens);
}

/**
 * A stored value as its number format says it looks.
 *
 * Returns the value UNCHANGED whenever it cannot do better: no format, a text
 * format, a value that is not a number, or a code outside the subset above.
 */
export function formatCellValue(value: string, code: string | undefined): string {
    if (code === undefined || code === '' || code === 'General' || code === '@') return value;

    const number = asNumber(value);
    if (number === null) return value;

    // Positive; negative; zero; text — Excel's four, and a code may give one,
    // two or three of them. A negative value with no section of its own uses
    // the positive one with a minus in front, which `renderNumber` does.
    const parts = sections(code);
    const chosen = number < 0 && parts.length > 1
        ? parts[1]
        : (number === 0 && parts.length > 2 ? parts[2] : parts[0]);

    const list = pieces(chosen ?? '');
    if (!list) return value;

    const tokens = list.filter((p) => p.kind === 'token').map((p) => p.text).join('');
    const hasDate = DATE_CHARS.test(tokens);
    const hasNumber = NUMBER_CHARS.test(tokens);

    // Both at once is a grammar this does not read; neither means the section is
    // pure literal text, which IS the answer — `;;;"paid"` renders "paid".
    if (hasDate && hasNumber) return value;
    if (!hasDate && !hasNumber) return list.map((p) => (p.kind === 'literal' ? p.text : '')).join('');

    if (hasDate) {
        const date = serialToDate(number);

        return date === null ? value : (renderDate(list, date) ?? value);
    }

    // A negative section renders the magnitude: the section carries its own
    // sign, as `#,##0.00;(#,##0.00)` does with brackets.
    const magnitude = number < 0 && parts.length > 1 ? Math.abs(number) : number;

    return renderNumber(list, magnitude) ?? value;
}

/**
 * How a date-formatted cell reads while it is being EDITED.
 *
 * A spreadsheet shows `21/08/2026` in the formula bar of a date cell, not
 * `46255`. The serial is the truth on disk and nobody wants to type one, so the
 * edit form of a date is the date — and {@link parseDateInput} turns it back.
 * Every other format edits as the value it stores, which is what Excel does
 * with `1234.5` under `#,##0.00`.
 */
export function editForm(value: string, code: string | undefined): string {
    return isDateFormat(code) ? formatCellValue(value, code) : value;
}

/** `21/08/2026`, `2026-08-21`, `21.08.2026` — the ways a date is typed. */
const TYPED_DATE = /^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})$/;

/**
 * A typed date as its serial, or null when the text is not one.
 *
 * Day-first when the format is: `dd/mm/yyyy` is what the toolbar offers and
 * what most of the world writes, and guessing the other way round turns the
 * fifth of March into the third of May without saying so. A four-digit FIRST
 * number is unambiguous and read as a year, since nobody means day 2026.
 */
export function parseDateInput(text: string, code: string | undefined): number | null {
    if (!isDateFormat(code)) return null;

    const match = TYPED_DATE.exec(text.trim());
    if (!match) return null;

    const [, a, b, c] = match;
    const yearFirst = a.length === 4;
    const year = Number(yearFirst ? a : c);
    const month = Number(b);
    const day = Number(yearFirst ? c : a);

    if (month < 1 || month > 12 || day < 1 || day > 31) return null;

    const full = year < 100 ? 2000 + year : year;
    const serial = dateToSerial(full, month, day);
    if (!Number.isFinite(serial) || serial < 1) return null;

    // Rolled rather than rejected would accept the 31st of February as the 3rd
    // of March. A date that does not exist was a typing mistake, not a date.
    const back = serialToDate(serial);

    return back?.year === full && back.month === month && back.day === day ? serial : null;
}
