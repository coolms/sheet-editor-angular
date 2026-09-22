/**
 * The NATIVE spreadsheet source format.
 *
 * A native spreadsheet template is a `.dsheet` -- a JSON grid document whose
 * cell values carry the same DTMPL tokens every other format uses -- and NOT an
 * `.xlsx`. The workbook is generation output, exactly as a native Word template
 * is a `.dtmpl` and never a `.docx`.
 *
 * Kept beside the other document features rather than inlined at the
 * registration so the backend constant
 * (`XlsxFormatProvider::DSHEET_MIME`) has one counterpart here to be
 * compared against, not several spellings to drift from.
 */
export const SHEET_DOCUMENT_MIME = 'application/x-coolms-sheet+json';

export const SHEET_DOCUMENT_EXT = '.dsheet';
