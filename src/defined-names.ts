/**
 * Declaring a name for a range, and the rules a name has to obey (#2385).
 *
 * ## Why an author wants one
 *
 * A name is the only reference into a loop band that survives every editor.
 * `SUM(B2:B2)` is collapsed to `SUM(B2)` by LibreOffice on save, and a bare
 * reference is never widened afterwards because the evidence of intent is gone.
 * `SUM(items_amount)` needs no rewriting at all: when the band grows, the NAME's
 * range moves and the formula is left exactly as authored.
 *
 * ## Rules, not preferences
 *
 * A spreadsheet refuses some names, and for a reason worth keeping: a name that
 * looks like `B2` would shadow the cell `B2`, so a formula could not say which
 * it meant. These follow Excel's rules rather than inventing gentler ones,
 * because the workbook this produces is opened by Excel.
 */
import { parseRef, type SheetDocumentDto } from './sheet-document.model';

/** One declared name, as the panel lists it. */
export interface DefinedName {
    readonly name: string;
    /** The full `Sheet1!$B$2:$B$4`, as the document stores it. */
    readonly range: string;
}

/** Excel's cap. Long before this a name has stopped being useful. */
const MAX_LENGTH = 255;

/**
 * Letters, digits, underscore and full stop, not starting with a digit.
 * A backslash is legal in Excel too and deliberately not offered: it reads as
 * an escape everywhere else an author will paste this.
 */
const SHAPE = /^[A-Za-z_][A-Za-z0-9_.]*$/;

/**
 * ⚠️ `R` and `C` alone are reserved: Excel reads them as whole-row and
 * whole-column shorthand in R1C1 notation.
 */
const RESERVED = new Set(['R', 'C']);

/**
 * What is wrong with this name, or null when nothing is.
 *
 * Returns a SENTENCE rather than a code, because the only consumer is a person
 * reading it under an input box.
 */
export function nameProblem(name: string, taken: readonly string[] = []): string | null {
    const trimmed = name.trim();
    if ('' === trimmed) return 'A name is needed.';
    if (trimmed.length > MAX_LENGTH) return `A name may be at most ${MAX_LENGTH} characters.`;

    if (!SHAPE.test(trimmed)) {
        return 'Use letters, digits, underscores and full stops, starting with a letter or underscore.';
    }

    if (RESERVED.has(trimmed.toUpperCase())) {
        return `“${trimmed}” is reserved by the spreadsheet.`;
    }

    // ⚠️ The rule that actually bites: `Q4` is a perfectly natural name for a
    // quarter's figures, and it is also a cell.
    if (parseRef(trimmed.toUpperCase())) {
        return `“${trimmed}” is a cell reference, so a formula could not tell them apart.`;
    }

    const wanted = trimmed.toUpperCase();
    if (taken.some((t) => t.toUpperCase() === wanted)) {
        return `“${trimmed}” is already used.`;
    }

    return null;
}

/**
 * `Sheet1` + `B2:B4` → `Sheet1!$B$2:$B$4`.
 *
 * Absolute on purpose: a name denotes a fixed block, and every editor writes one
 * this way. A sheet whose name needs quoting gets them, with an embedded quote
 * doubled -- the same escaping the file format uses.
 */
export function scopedRange(sheet: string, range: string): string {
    const needsQuotes = !/^[A-Za-z0-9_.]+$/.test(sheet);
    const label = needsQuotes ? `'${sheet.replace(/'/g, "''")}'` : sheet;
    const absolute = range
        .split(':')
        .map((part) => part.replace(/^([A-Za-z]+)(\d+)$/, '$$$1$$$2'))
        .join(':');

    return `${label}!${absolute}`;
}

/** The declared names, sorted so the panel does not reorder itself on edit. */
export function definedNamesOf(doc: SheetDocumentDto | null | undefined): DefinedName[] {
    const names = doc?.definedNames;
    if (!names) return [];

    return Object.entries(names)
        .map(([name, range]) => ({ name, range }))
        .sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * The document with one more name.
 *
 * Returns a NEW document rather than mutating: the editor's undo stack keeps
 * the previous one, and a shared object would let undo restore a value that had
 * already been changed underneath it.
 */
export function withDefinedName(doc: SheetDocumentDto, name: string, range: string): SheetDocumentDto {
    return { ...doc, definedNames: { ...doc.definedNames, [name.trim()]: range } };
}

/**
 * The document without it.
 *
 * ⚠️ Drops the key entirely rather than leaving it empty, and drops the whole
 * map when it was the last one -- a `.dsheet` is a source file an operator reads
 * in a diff, and `"definedNames": {}` on every template is noise. The backend
 * omits it on the same rule.
 */
export function withoutDefinedName(doc: SheetDocumentDto, name: string): SheetDocumentDto {
    const rest: Record<string, string> = { ...doc.definedNames };
    delete rest[name];

    // ⚠️ Annotated, not inferred: the spread gives the literal a DEFINITE
    // `definedNames`, and `delete` is refused on a property that is not optional.
    // Spreading `doc` also carries every other field, so a document does not
    // quietly lose one here the way a rebuilt literal would.
    const next: SheetDocumentDto = { ...doc, definedNames: rest };
    if (0 === Object.keys(rest).length) delete next.definedNames;

    return next;
}
