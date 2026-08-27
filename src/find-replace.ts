/**
 * Finding and replacing text across a sheet.
 *
 * Searches what the cell HOLDS -- its literal text or its formula source --
 * rather than what it computes. In a template that is the whole point: the job
 * this exists for is renaming `{var:order.number}` to `{var:invoice.number}`
 * across a document, and a search over computed values could not see a token at
 * all, because a token has no value until the document is generated.
 *
 * It follows that a formula is searchable and replaceable as its text. That can
 * break a formula, and it is still right: an author renaming a column heading
 * means the one inside `SUM(Orders!B2:B9)` too, and a tool that silently
 * skipped it would leave exactly the occurrence they cannot see.
 *
 * Literal matching, not regular expressions. The thing being searched for here
 * is usually a token full of braces and colons, and a regex over that is a
 * mistake waiting to be made by someone who did not know they were writing one.
 */
import {
    cellToInput, columnToIndex, inputToCell, parseRef, type SheetDto,
} from './sheet-document.model';

export interface FindOptions {
    /** Off by default, as every editor's find box is. */
    readonly matchCase?: boolean;
    /** The whole cell must equal the query, not merely contain it. */
    readonly wholeCell?: boolean;
}

const fold = (text: string, options: FindOptions): string =>
    (options.matchCase === true ? text : text.toLowerCase());

/** Whether one cell's text answers the query. */
export function matchesQuery(text: string, query: string, options: FindOptions = {}): boolean {
    if ('' === query) return false;

    const haystack = fold(text, options);
    const needle = fold(query, options);

    return options.wholeCell === true ? haystack === needle : haystack.includes(needle);
}

/**
 * Every matching cell, in READING order.
 *
 * Reading order and not the order the cells happen to sit in the map: "find
 * next" has to walk the sheet the way an author reads it, and a JSON object's
 * key order is the order things were written, which is nobody's idea of next.
 */
export function findMatches(
    sheet: SheetDto | undefined,
    query: string,
    options: FindOptions = {},
): string[] {
    if (!sheet || '' === query) return [];

    const found: { ref: string; row: number; column: number }[] = [];
    for (const [ref, cell] of Object.entries(sheet.cells)) {
        if (!matchesQuery(cellToInput(cell), query, options)) continue;
        const at = parseRef(ref);
        if (at) found.push({ ref, row: at.row, column: columnToIndex(at.column) });
    }

    found.sort((a, b) => a.row - b.row || a.column - b.column);

    return found.map(f => f.ref);
}

/**
 * The match after a cell, wrapping round to the first.
 *
 * Wrapping rather than stopping: a search that runs out at the bottom and does
 * nothing looks broken, and every editor that has ever had a find box wraps.
 * Returns null only when there is nothing to find at all.
 */
export function nextMatch(matches: readonly string[], after: string): string | null {
    if (0 === matches.length) return null;

    const at = parseRef(after);
    if (!at) return matches[0];

    const column = columnToIndex(at.column);
    const beyond = matches.find(ref => {
        const cell = parseRef(ref);
        if (!cell) return false;

        return cell.row > at.row || (cell.row === at.row && columnToIndex(cell.column) > column);
    });

    return beyond ?? matches[0];
}

/** One cell's text with the query replaced -- every occurrence, or the whole cell. */
export function replaceIn(
    text: string,
    query: string,
    replacement: string,
    options: FindOptions = {},
): string {
    if (!matchesQuery(text, query, options)) return text;
    if (options.wholeCell === true) return replacement;

    // Case-insensitively, the match and the text differ, so this cannot be
    // `split(query).join(replacement)`: the needle has to be found in the
    // FOLDED text and cut out of the original one.
    const haystack = fold(text, options);
    const needle = fold(query, options);

    let out = '';
    let from = 0;
    for (;;) {
        const at = haystack.indexOf(needle, from);
        if (at < 0) break;
        out += text.slice(from, at) + replacement;
        from = at + needle.length;
    }

    return out + text.slice(from);
}

/**
 * Replace in ONE cell, leaving the rest of the sheet alone.
 *
 * What the Replace button does, as against Replace all: the author is looking
 * at a cell and means that one.
 */
export function withReplacedIn(
    sheet: SheetDto,
    ref: string,
    query: string,
    replacement: string,
    options: FindOptions = {},
): SheetDto {
    const before = cellToInput(sheet.cells[ref]);
    const after = replaceIn(before, query, replacement, options);
    if (after === before) return sheet;

    const cells = { ...sheet.cells };
    const next = inputToCell(after, cells[ref]);
    if (null === next) {
        delete cells[ref];
    } else {
        cells[ref] = next;
    }

    return { ...sheet, cells };
}

/**
 * Replace across the whole sheet, reporting how many cells changed.
 *
 * CELLS changed, not occurrences: it is the number an author can check by
 * looking, and "3 cells" is a thing they can verify where "5 replacements"
 * is not.
 */
export function withReplacedAll(
    sheet: SheetDto,
    query: string,
    replacement: string,
    options: FindOptions = {},
): { sheet: SheetDto; count: number } {
    const matches = findMatches(sheet, query, options);
    if (0 === matches.length) return { sheet, count: 0 };

    const cells = { ...sheet.cells };
    let count = 0;

    for (const ref of matches) {
        const before = cellToInput(cells[ref]);
        const after = replaceIn(before, query, replacement, options);
        if (after === before) continue;

        // Through `inputToCell`, so a replaced formula is still a formula, a
        // replaced date still lands as a serial, and the cell's formatting
        // survives -- a replacement changes TEXT, not the look of the table.
        const next = inputToCell(after, cells[ref]);
        if (null === next) {
            delete cells[ref];
        } else {
            cells[ref] = next;
        }
        count += 1;
    }

    return { sheet: { ...sheet, cells }, count };
}
