import { dateToSerial } from './number-format';
import {
    autoFilterOf,
    borderCss,
    frozenAt,
    withFreeze,
    borderCssAt,
    borderStateIn,
    cellToInput,
    withBorderPreset,
    columnToIndex,
    sheetLookup,
    filterBodyRows,
    filterColumns,
    isFilterHeader,
    shiftFormula,
    validationAt,
    validationOptions,
    validationRangeAt,
    withClearedRange,
    withDeletedColumn,
    withDeletedRow,
    withInsertedColumn,
    withInsertedRow,
    withAutoFilter,
    withoutAutoFilter,
    withValidation,
    withoutValidation,
    columnWidthFromPx,
    columnWidthOf,
    columnWidthToPx,
    gridExtent,
    MIN_COLUMN_WIDTH,
    indexToColumn,
    inputToCell,
    isKnownFormat,
    isMergeAnchor,
    mergeCovering,
    mergeSpan,
    parseRef,
    parseSheetDocument,
    normaliseColour,
    rangeBetween,
    rangeContains,
    withStyle,
    withWrap,
    refsInRange,
    safeSheetName,
    serialiseSheetDocument,
    SheetDocumentDto,
    SheetDto,
    withMerge,
    withNewSheet,
    withoutMerge,
    withoutSheet,
    withRenamedSheet,
    withBold,
    withColumnWidth,
    withNumberFormat,
} from './sheet-document.model';

/**
 * The `.dsheet` model the grid editor reads and writes.
 *
 * These are the parts that can be wrong without anything crashing: a lost
 * number format, a grid too small to type into, a garbage file silently
 * replaced with a blank one.
 */
describe('sheet document model', () => {
    describe('A1 arithmetic', () => {
        it('round-trips column letters past Z', () => {
            for (const [letters, index] of [['A', 1], ['Z', 26], ['AA', 27], ['AZ', 52], ['BA', 53]] as const) {
                expect(columnToIndex(letters)).toBe(index);
                expect(indexToColumn(index)).toBe(letters);
            }
        });

        it('parses an A1 reference and rejects what is not one', () => {
            expect(parseRef('B4')).toEqual({ column: 'B', row: 4 });
            expect(parseRef('AA10')).toEqual({ column: 'AA', row: 10 });
            // Row 0 does not exist in a spreadsheet, and a bare column is not a cell.
            expect(parseRef('A0')).toBeNull();
            expect(parseRef('A')).toBeNull();
            expect(parseRef('4')).toBeNull();
        });
    });

    describe('grid extent', () => {
        const sheet = (refs: string[]): SheetDto => ({
            cells: Object.fromEntries(refs.map(r => [r, { value: 'x' }])),
        });

        /**
         * The padding is what makes this an EDITOR. A document whose last cell
         * is B4 must still offer somewhere to add a fifth row.
         */
        it('reaches past the last used cell so there is room to type', () => {
            const { rows, cols } = gridExtent(sheet(['B4']), { minRows: 1, minCols: 1, padRows: 6, padCols: 3 });

            expect(rows).toBe(10);
            expect(cols).toBe(5);
        });

        it('covers a wide sheet by column letters, not by count', () => {
            const { cols } = gridExtent(sheet(['AA1']), { minRows: 1, minCols: 1, padRows: 0, padCols: 0 });

            expect(cols).toBe(27);
        });

        /** An empty document is what a NEW native template is. */
        it('still offers a usable grid for an empty document', () => {
            const { rows, cols } = gridExtent({ cells: {} });

            expect(rows).toBeGreaterThanOrEqual(20);
            expect(cols).toBeGreaterThanOrEqual(8);
        });
    });

    describe('cell ↔ input', () => {
        it('shows a formula with its leading = and a value plain', () => {
            expect(cellToInput({ formula: 'B4*C4' })).toBe('=B4*C4');
            expect(cellToInput({ value: '{var:total}' })).toBe('{var:total}');
            expect(cellToInput(undefined)).toBe('');
        });

        it('reads a leading = back as a formula, not a value', () => {
            expect(inputToCell('=SUM(A1:A3)', undefined)).toEqual({ formula: 'SUM(A1:A3)' });
            expect(inputToCell('plain', undefined)).toEqual({ value: 'plain' });
        });

        /**
         * THE one that matters. The number format is the author's TYPE
         * DECLARATION — `@` is what keeps an order number like `00412`
         * text instead of arithmetic. Editing the cell's TEXT must not discard
         * it, or the next generation silently promotes the value.
         */
        it('carries the number format and weight through an edit', () => {
            const previous = { value: '{var:order}', numberFormat: '@', bold: true };

            expect(inputToCell('{var:order.number}', previous)).toEqual({
                value: '{var:order.number}',
                numberFormat: '@',
                bold: true,
            });
        });

        it('drops an emptied cell entirely, keeping the document sparse', () => {
            expect(inputToCell('', { value: 'gone' })).toBeNull();
        });

        /** …unless it still carries formatting the author set deliberately. */
        it('keeps an emptied cell that carries formatting', () => {
            expect(inputToCell('', { value: 'gone', numberFormat: '@' })).toEqual({ numberFormat: '@' });
        });

        /**
         * Editing the TEXT of a cell must not undo its appearance. Every field
         * here is something an author set on purpose through the toolbar, and
         * retyping the words in the cell is not a request to undo any of it.
         */
        it('turns a typed date into the serial a workbook stores', () => {
            const cell = inputToCell('21/08/2026', { numberFormat: 'dd/mm/yyyy' });

            expect(cell).toEqual({ value: String(dateToSerial(2026, 8, 21)), numberFormat: 'dd/mm/yyyy' });
        });

        /**
         * Only where the cell SAYS it is a date. A version number typed into a
         * General cell is not a date, and turning `1.2.3` into a serial would
         * be a silent, unrecoverable edit.
         */
        it('leaves a date-shaped value alone in a cell that is not a date', () => {
            expect(inputToCell('21/08/2026', undefined)).toEqual({ value: '21/08/2026' });
        });

        it('keeps text that is not a date in a date cell', () => {
            expect(inputToCell('n/a', { numberFormat: 'dd/mm/yyyy' }))
                .toEqual({ value: 'n/a', numberFormat: 'dd/mm/yyyy' });
        });

        it('keeps every scrap of formatting through an edit of the text', () => {
            const styled = {
                value: 'before',
                numberFormat: '@',
                bold: true,
                italic: true,
                fontFamily: 'Georgia',
                fontSize: 14,
                color: '#FF0000',
                background: '#EEEEEE',
                align: 'center' as const,
                valign: 'middle' as const,
                wrap: true,
                link: 'https://example.test',
                borders: { top: 'thin' },
            };

            expect(inputToCell('after', styled)).toEqual({ ...styled, value: 'after' });
        });
    });

    describe('formatting', () => {
        it('applies and clears a number format, keeping the rest of the cell', () => {
            expect(withNumberFormat({ value: '{var:order}' }, '@'))
                .toEqual({ value: '{var:order}', numberFormat: '@' });

            expect(withNumberFormat({ value: '{var:order}', numberFormat: '@' }, undefined))
                .toEqual({ value: '{var:order}' });
        });

        /**
         * An author marks a column as Text BEFORE typing into it. That intent
         * has to survive, so a format creates the cell rather than being
         * dropped for having no value yet.
         */
        it('can format a cell that does not exist yet', () => {
            expect(withNumberFormat(undefined, '@')).toEqual({ numberFormat: '@' });
        });

        it('drops a cell once nothing is left worth storing', () => {
            expect(withNumberFormat({ numberFormat: '@' }, undefined)).toBeNull();
            expect(withBold({ bold: true }, false)).toBeNull();
        });

        it('toggles weight without disturbing the format', () => {
            expect(withBold({ value: 'Total', numberFormat: '@' }, true))
                .toEqual({ value: 'Total', numberFormat: '@', bold: true });
        });

        /**
         * A `.dsheet` may carry any OOXML code. The toolbar must be able to say
         * "this is not one of mine" so the UI can show it instead of
         * displaying General and overwriting it on the next unrelated change.
         */
        it('recognises which codes the menu can represent', () => {
            expect(isKnownFormat(undefined)).toBeTrue();
            expect(isKnownFormat('@')).toBeTrue();
            expect(isKnownFormat('#,##0.00\\ [$€-407]')).toBeFalse();
        });

        /**
         * Wrap and vertical alignment — the pair that makes a tall row
         * usable. Wrap is a boolean like bold, so it CLEARS to absent rather
         * than storing false: `wrapText` has a real default in OOXML and a
         * stored `false` would be a third state meaning the same as absent.
         */
        it('toggles wrap without disturbing the rest of the cell', () => {
            expect(withWrap({ value: 'A long label', align: 'center' }, true))
                .toEqual({ value: 'A long label', align: 'center', wrap: true });

            expect(withWrap({ value: 'A long label', wrap: true }, false))
                .toEqual({ value: 'A long label' });
        });

        it('drops a cell whose only content was wrap', () => {
            expect(withWrap({ wrap: true }, false)).toBeNull();
        });

        it('sets and clears a vertical alignment', () => {
            expect(withStyle({ value: 'Header' }, 'valign', 'middle'))
                .toEqual({ value: 'Header', valign: 'middle' });

            expect(withStyle({ value: 'Header', valign: 'middle' }, 'valign', undefined))
                .toEqual({ value: 'Header' });
        });

        /**
         * The two are independent: a cell may wrap without stating where its
         * text sits, and may state that without wrapping — a hand-set row
         * height creates the same spare space wrapping does.
         */
        it('keeps wrap and valign independent of each other', () => {
            const wrapped = withWrap({ value: 'x' }, true);
            expect(withStyle(wrapped ?? undefined, 'valign', 'top'))
                .toEqual({ value: 'x', wrap: true, valign: 'top' });

            expect(withWrap({ value: 'x', valign: 'top', wrap: true }, false))
                .toEqual({ value: 'x', valign: 'top' });
        });
    });

    describe('parsing', () => {
        it('reads a real document', () => {
            const { doc, ok } = parseSheetDocument(JSON.stringify({
                version: 1,
                sheets: { Invoice: { cells: { A1: { value: 'Hello' } } } },
            }));

            expect(ok).toBeTrue();
            expect(doc.sheets['Invoice'].cells['A1'].value).toBe('Hello');
        });

        it('treats an empty file as a new document', () => {
            const { doc, ok } = parseSheetDocument('   ');

            expect(ok).toBeTrue();
            expect(Object.keys(doc.sheets)).toEqual(['Sheet1']);
        });

        /**
         * A damaged file must REPORT itself rather than open as blank. Opening
         * blank and then saving would overwrite the operator's file with an
         * empty one — the editor is the only place they can repair it from.
         */
        it('reports a file it cannot understand instead of pretending it is empty', () => {
            for (const bad of ['not json at all', '[]', '{"sheets":null}', '{"sheets":{}}']) {
                expect(parseSheetDocument(bad).ok).withContext(bad).toBeFalse();
            }
        });

        it('survives its own serialisation', () => {
            const source = {
                version: 1,
                sheets: {
                    S: {
                        cells: { A1: { value: 'x' }, B2: { formula: 'A1*2', numberFormat: '#,##0.00' } },
                        merges: ['A1:D1'],
                    },
                },
            };

            const { doc, ok } = parseSheetDocument(serialiseSheetDocument(source));

            expect(ok).toBeTrue();
            expect(doc).toEqual(source);
        });
    });

    /**
     * THE data-loss bug, found by reading the stored bytes rather than trusting
     * a green "Saved" toast.
     *
     * A template minted by the backend carried `"cells": []` — PHP cannot tell
     * an empty map from an empty list, so `json_encode` emitted an array. `[]`
     * is neither null nor undefined, so the parser's `??=` left it alone, and
     * `cells['A1'] = …` then set a STRING KEY on a JS array. `JSON.stringify`
     * discards those: every cell typed into a brand-new native template was lost
     * on save, and the stored blob came back with the SAME CONTENT HASH.
     *
     * The assertion is on the SERIALISED form on purpose. Checking
     * `doc.sheets.S.cells['A1']` passes even on an array — the property is
     * really there in memory — so a test written that way would have gone green
     * against the bug. Only stringifying reproduces the loss.
     */
    describe('a minted document whose cells arrived as a JSON array', () => {
        const MINTED = '{"version":1,"sheets":{"S":{"cells":[]}}}';

        it('normalises cells to an object so a typed cell survives serialisation', () => {
            const { doc, ok } = parseSheetDocument(MINTED);
            expect(ok).toBeTrue();
            expect(Array.isArray(doc.sheets['S'].cells)).toBeFalse();

            doc.sheets['S'].cells['A1'] = { value: 'CELL-LOSS-PROBE' };

            expect(serialiseSheetDocument(doc)).toContain('CELL-LOSS-PROBE');
            expect(parseSheetDocument(serialiseSheetDocument(doc)).doc.sheets['S'].cells['A1'])
                .toEqual({ value: 'CELL-LOSS-PROBE' });
        });

        /** A readable file must not be reported as damaged just for this. */
        it('treats the minted shape as readable rather than corrupt', () => {
            expect(parseSheetDocument(MINTED).ok).toBeTrue();
        });
    });

    /**
     * The same loss one level up, reached by a name rather than by emptiness.
     *
     * Sheet names are operator-supplied and the rename validator accepts "0" and
     * "1". PHP coerces numeric string array keys to ints, so a document whose
     * sheets were named that way became a LIST and encoded as `"sheets": [ … ]`
     * — the names gone from the file entirely. `Object.keys` still yields
     * "0"/"1", so the tabs LOOK correct; adding a sheet then sets a string key
     * on a JS array and `JSON.stringify` drops it, exactly as with `cells`.
     *
     * The backend now casts on encode, so new files carry the object. This pins
     * the READ side, which is what repairs any file already written.
     */
    describe('a document whose sheets arrived as a JSON array', () => {
        const POSITIONAL = '{"version":1,"sheets":[{"cells":{"A1":{"value":"a"}}},{"cells":{}}]}';

        it('normalises sheets to an object, keeping the positional names', () => {
            const { doc, ok } = parseSheetDocument(POSITIONAL);

            expect(ok).toBeTrue();
            expect(Array.isArray(doc.sheets)).toBeFalse();
            expect(Object.keys(doc.sheets)).toEqual(['0', '1']);
            expect(doc.sheets['0'].cells['A1']).toEqual({ value: 'a' });
        });

        /**
         * Re-saving is what REPAIRS the file. Without the normalisation the
         * editor writes `"sheets": [ … ]` straight back out and the names never
         * become real, so a document stays in the broken shape however many
         * times it is opened.
         *
         * Note what is NOT asserted here: that adding a sheet survives. It does
         * either way, because `withNewSheet` SPREADS into a fresh object rather
         * than assigning into the array. That is a property of today's helper,
         * not a guarantee — which is the whole reason to normalise on read.
         */
        it('writes the sheets back out as an object, repairing the file', () => {
            const serialised = serialiseSheetDocument(parseSheetDocument(POSITIONAL).doc);

            expect(serialised).toContain('"sheets": {');
            expect(serialised).not.toContain('"sheets": [');
        });
    });

    describe('sheets', () => {
        const twoSheets = (): SheetDocumentDto => ({
            version: 1,
            sheets: { First: { cells: { A1: { value: 'a' } } }, Second: { cells: {} } },
        });

        /** The writer's rules, mirrored so the tab an author names is the tab they get. */
        it('applies the writer\'s name rules', () => {
            expect(safeSheetName('Invoice')).toBe('Invoice');
            expect(safeSheetName('A/B:C[D]*E?F\\G')).toBe('A-B-C-D--E-F-G');
            expect(safeSheetName('x'.repeat(40)).length).toBe(31);
            expect(safeSheetName('   ', 2)).toBe('Sheet3', 'an all-blank name falls back like the writer does');
        });

        it('adds a sheet and suffixes a name already taken', () => {
            const first = withNewSheet(twoSheets(), 'Third');
            expect(Object.keys(first.doc.sheets)).toEqual(['First', 'Second', 'Third']);
            expect(first.name).toBe('Third');

            const clash = withNewSheet(twoSheets(), 'First');
            expect(clash.name).toBe('First 2');
            expect(clash.doc.sheets['First'].cells['A1']).toEqual({ value: 'a' }, 'the original must not be overwritten');
        });

        /**
         * THE ordering trap. JS object keys iterate in insertion order, the
         * backend writes sheets in that order, and `SheetDocumentWriter` makes
         * index 0 the ACTIVE sheet — so a rename implemented as delete-then-add
         * would move the sheet to the end of the workbook, and renaming the
         * first one would hand the operator a different opening tab.
         */
        it('renames in place, keeping the sheet\'s position', () => {
            const renamed = withRenamedSheet(twoSheets(), 'First', 'Cover');

            expect(Object.keys(renamed.sheets)).toEqual(['Cover', 'Second']);
            expect(renamed.sheets['Cover'].cells['A1']).toEqual({ value: 'a' }, 'the content comes with the name');
        });

        it('refuses a rename that would collide, vanish, or do nothing', () => {
            const doc = twoSheets();

            expect(withRenamedSheet(doc, 'First', 'Second')).toBe(doc, 'a taken name');
            expect(withRenamedSheet(doc, 'First', 'First')).toBe(doc, 'a no-op');
            expect(withRenamedSheet(doc, 'Missing', 'X')).toBe(doc, 'an unknown sheet');
        });

        it('deletes a sheet but never the last one', () => {
            const doc = twoSheets();
            const one = withoutSheet(doc, 'Second');

            expect(Object.keys(one.sheets)).toEqual(['First']);
            // The writer throws on a document with no sheets, so an editor that
            // allowed this would produce a template that only fails at
            // generation time.
            expect(withoutSheet(one, 'First')).toBe(one);
            expect(withoutSheet(doc, 'Missing')).toBe(doc);
        });

        it('round-trips added sheets through serialisation', () => {
            const { doc } = withNewSheet(twoSheets(), 'Third');

            expect(Object.keys(parseSheetDocument(serialiseSheetDocument(doc)).doc.sheets))
                .toEqual(['First', 'Second', 'Third']);
        });
    });

    describe('merges', () => {
        /**
         * An author can shift-click in any direction, so the range has to be
         * normalised: `D4:A1` would never match the OOXML the renderer emits.
         */
        it('normalises a range whichever corner was picked first', () => {
            expect(rangeBetween('A1', 'D4')).toBe('A1:D4');
            expect(rangeBetween('D4', 'A1')).toBe('A1:D4');
            expect(rangeBetween('D1', 'A4')).toBe('A1:D4');
            expect(rangeBetween('A1', 'nonsense')).toBeNull();
        });

        it('knows which cells a merge covers and which one anchors it', () => {
            const sheet: SheetDto = { cells: {}, merges: ['B2:C3'] };

            expect(mergeCovering(sheet, 'B2')).toBe('B2:C3');
            expect(mergeCovering(sheet, 'C3')).toBe('B2:C3');
            expect(mergeCovering(sheet, 'D3')).toBeNull();
            expect(isMergeAnchor('B2:C3', 'B2')).toBeTrue();
            expect(isMergeAnchor('B2:C3', 'C3')).toBeFalse();
            expect(mergeSpan('B2:C3')).toEqual({ colspan: 2, rowspan: 2 });
        });

        /**
         * Clearing the covered cells is deliberate. A merged range keeps only
         * its top-left value — that is what the renderer does, because
         * PhpSpreadsheet's `mergeCells()` empties the rest — so a document
         * holding values under a merge would render differently from what the
 * grid shows. That divergence is what existed to close.
         */
        it('clears the cells a merge swallows, keeping the anchor', () => {
            const sheet: SheetDto = {
                cells: {
                    A1: { value: 'keep' },
                    B1: { value: 'swallowed' },
                    C1: { value: 'outside' },
                },
            };

            const merged = withMerge(sheet, 'A1:B1');

            expect(merged.cells['A1']).toEqual({ value: 'keep' });
            expect('B1' in merged.cells).toBeFalse();
            expect(merged.cells['C1']).toEqual({ value: 'outside' }, 'a cell outside the range is untouched');
            expect(merged.merges).toEqual(['A1:B1']);
            expect(sheet.cells['B1']).toEqual({ value: 'swallowed' }, 'the input sheet must not be mutated');
        });

        /** Two merges sharing a cell is a workbook Excel offers to repair. */
        it('drops a merge the new one overlaps', () => {
            const sheet: SheetDto = { cells: {}, merges: ['A1:B1', 'D1:E1'] };

            const merged = withMerge(sheet, 'B1:C1');

            expect(merged.merges).toEqual(['D1:E1', 'B1:C1']);
        });

        /** A click without a drag must not write `A1:A1` into the document. */
        it('refuses a single-cell range', () => {
            const sheet: SheetDto = { cells: {} };

            expect(withMerge(sheet, 'A1:A1')).toBe(sheet);
            expect(withMerge(sheet, 'not a range')).toBe(sheet);
        });

        it('removes a merge and drops the key with the last one', () => {
            const sheet: SheetDto = { cells: {}, merges: ['A1:B1', 'D1:E1'] };

            expect(withoutMerge(sheet, 'A1:B1').merges).toEqual(['D1:E1']);
            expect('merges' in withoutMerge(withoutMerge(sheet, 'A1:B1'), 'D1:E1')).toBeFalse();
        });

        it('survives serialisation with its merges intact', () => {
            const doc = parseSheetDocument('{"version":1,"sheets":{"S":{"cells":{}}}}').doc;
            doc.sheets['S'] = withMerge(doc.sheets['S'], 'A1:D1');

            expect(parseSheetDocument(serialiseSheetDocument(doc)).doc.sheets['S'].merges).toEqual(['A1:D1']);
        });
    });

    describe('column widths', () => {
        it('stores a width and reads it back, case-insensitively', () => {
            const sheet = withColumnWidth({ cells: {} }, 'a', 28);

            expect(sheet.columnWidths).toEqual({ A: 28 });
            expect(columnWidthOf(sheet, 'A')).toBe(28);
            expect(columnWidthOf(sheet, 'a')).toBe(28);
            expect(columnWidthOf(sheet, 'B')).toBeUndefined();
        });

        it('does not mutate the sheet it was given', () => {
            const original: SheetDto = { cells: {}, columnWidths: { A: 10 } };

            withColumnWidth(original, 'B', 20);

            expect(original.columnWidths).toEqual({ A: 10 });
        });

        /**
         * Clearing the last width must drop the KEY, not leave `{}` behind: the
         * backend writes the object only when it has entries, and a `.dsheet` is
         * a source file an author reads.
         */
        it('drops the columnWidths key once the last width is cleared', () => {
            const one = withColumnWidth({ cells: {} }, 'A', 28);
            const two = withColumnWidth(one, 'B', 12);

            expect(withColumnWidth(two, 'B', undefined).columnWidths).toEqual({ A: 28 });
            expect('columnWidths' in withColumnWidth(one, 'A', undefined)).toBeFalse();
        });

        /**
         * Excel reads width 0 as a HIDDEN column. Storing a stray 0 would make a
         * column vanish from the generated workbook with nothing in the editor to
         * explain it, so the model refuses rather than stores.
         */
        it('refuses a width that would hide or corrupt the column', () => {
            const sheet: SheetDto = { cells: {}, columnWidths: { A: 28 } };

            for (const bad of [0, -5, Number.NaN, Number.POSITIVE_INFINITY]) {
                expect(withColumnWidth(sheet, 'A', bad)).toBe(sheet);
            }
        });

        /**
         * Display only, and approximate on purpose — the editor renders in the
         * browser's font, not the workbook's. What must be exact is the STORED
         * number, which is why nothing converts on save.
         */
        it('converts a character width to approximate pixels', () => {
            expect(columnWidthToPx(10)).toBe(75);
            expect(columnWidthToPx(8.43)).toBe(64);
        });

        /**
         * The drag handle's direction. A round trip is the useful property: an
         * author who drags a column and lets go must not see the width shift
         * again on the next render.
         */
        it('converts pixels back to a character width', () => {
            expect(columnWidthFromPx(columnWidthToPx(10))).toBe(10);
            expect(columnWidthFromPx(75)).toBe(10);
        });

        /**
         * A drag produces a new value on every mousemove, and the raw division
         * yields things like 12.714285714285714 — which would be written into
         * the document and shown in the toolbar's width box. An author must end
         * up with a number they could have typed.
         */
        it('rounds to a width an author could have typed', () => {
            expect(columnWidthFromPx(94)).toBe(12.71);
        });

        /**
         * Dragging past the left edge must not produce a HIDDEN column. Width 0
         * is Excel's "hidden", and `withColumnWidth` refuses it — so without the
         * clamp the drag would silently stop having any effect at the exact
         * moment the author is trying hardest to make the column narrow.
         */
        it('clamps a drag past the left edge instead of hiding the column', () => {
            for (const px of [0, -200]) {
                expect(columnWidthFromPx(px)).toBe(MIN_COLUMN_WIDTH);
                expect(withColumnWidth({ cells: {} }, 'A', columnWidthFromPx(px)).columnWidths)
                    .toEqual({ A: MIN_COLUMN_WIDTH });
            }
        });
    });

    describe('colour normalisation', () => {
        /**
         * The editor and the backend must agree on ONE spelling. A browser
         * colour input emits lower case; `SheetCell::colour()` produces upper.
         * Without this the editor writes `#ffee00` into a file the backend
         * would rewrite as `#FFEE00` — a case-only diff on a line nobody
         * touched.
         */
        it('stores a colour in the same canonical form the backend parses to', () => {
            for (const spelling of ['#ffee00', 'ffee00', '#FFEE00', '#FfEe00']) {
                expect(withStyle({}, 'color', spelling)?.color).toBe('#FFEE00');
                expect(withStyle({}, 'background', spelling)?.background).toBe('#FFEE00');
            }
        });

        /** Only colours are folded — a font family is a NAME and case is part of it. */
        it('leaves non-colour styles exactly as given', () => {
            expect(withStyle({}, 'fontFamily', 'Times New Roman')?.fontFamily).toBe('Times New Roman');
            expect(withStyle({}, 'align', 'center')?.align).toBe('center');
        });

        /** An unparseable value passes through for the backend to refuse. */
        it('does not mangle something that is not a colour', () => {
            expect(normaliseColour('rgb(1,2,3)')).toBe('rgb(1,2,3)');
        });
    });

    describe('rangeContains', () => {
        /**
         * The fast path the grid uses for every rendered cell. It must agree
         * with the list version exactly — if the two ever disagree, the cheap
         * one is the one on screen.
         */
        it('agrees with refsInRange across a box', () => {
            const range = 'B2:D4';

            for (const ref of ['A1', 'B1', 'B2', 'C3', 'D4', 'E4', 'D5', 'B4', 'A3']) {
                expect(rangeContains(range, ref))
                    .withContext(ref)
                    .toBe(refsInRange(range).includes(ref));
            }
        });

        it('answers false for a reference or range it cannot parse', () => {
            expect(rangeContains('B2:D4', 'not-a-ref')).toBeFalse();
            expect(rangeContains('nonsense', 'B3')).toBeFalse();
        });
    });

    describe('auto filter', () => {
        it('declares one filter over a range, upper-cased', () => {
            const sheet: SheetDto = { cells: {} };

            expect(autoFilterOf(withAutoFilter(sheet, 'a1:c9'))).toBe('A1:C9');
        });

        /**
         * Singular, not a list: OOXML allows one `<autoFilter>` per worksheet,
         * so declaring a second must REPLACE the first rather than accumulate
         * into a workbook Excel offers to repair.
         */
        it('replaces the filter rather than keeping two', () => {
            const sheet = withAutoFilter({ cells: {} }, 'A1:C9');

            expect(autoFilterOf(withAutoFilter(sheet, 'E1:F4'))).toBe('E1:F4');
        });

        /** Same rule as a merge: a click without a drag is not a filter. */
        it('refuses a single cell and a range it cannot parse', () => {
            const sheet: SheetDto = { cells: {} };

            expect(withAutoFilter(sheet, 'A1:A1')).toBe(sheet);
            expect(withAutoFilter(sheet, 'not a range')).toBe(sheet);
        });

        it('removes the filter by dropping the key, not blanking it', () => {
            const sheet = withAutoFilter({ cells: {} }, 'A1:C9');

            expect('autoFilter' in withoutAutoFilter(sheet)).toBeFalse();
            expect(autoFilterOf(sheet)).toBe('A1:C9', 'the input sheet must not be mutated');
        });

        /** A filter stored by an older writer that no longer parses is ignored. */
        it('treats an unparseable stored range as no filter', () => {
            expect(autoFilterOf({ cells: {}, autoFilter: 'A1' })).toBeNull();
        });

        /**
         * The dropdowns sit on the range's TOP row — Excel's rule, and the
         * reason a filter must be declared with its header included.
         */
        it('puts the buttons on the top row of the range only', () => {
            const sheet = withAutoFilter({ cells: {} }, 'B2:D9');

            expect(isFilterHeader(sheet, 'B2')).toBeTrue();
            expect(isFilterHeader(sheet, 'D2')).toBeTrue();
            expect(isFilterHeader(sheet, 'E2')).withContext('outside the columns').toBeFalse();
            expect(isFilterHeader(sheet, 'B3')).withContext('a body row, not the header').toBeFalse();
            expect(isFilterHeader({ cells: {} }, 'B2')).withContext('no filter, no buttons').toBeFalse();
        });

        it('lists the columns it covers and the rows its dropdowns act on', () => {
            expect(filterColumns('B2:D9')).toEqual(['B', 'C', 'D']);
            expect(filterBodyRows('B2:D5')).toEqual([3, 4, 5]);
        });

        /**
         * A one-row filter has no body, which is the ORDINARY state of a
         * template whose table is a `{loop:}` band the backend has not expanded
         * yet. It must answer empty rather than treat the header as data.
         */
        it('has no body rows when the range is one row tall', () => {
            expect(filterBodyRows('B2:D2')).toEqual([]);
        });

        it('survives serialisation', () => {
            const doc = parseSheetDocument('{"version":1,"sheets":{"S":{"cells":{}}}}').doc;
            doc.sheets['S'] = withAutoFilter(doc.sheets['S'], 'A1:C9');

            const back = parseSheetDocument(serialiseSheetDocument(doc)).doc;

            expect(autoFilterOf(back.sheets['S'])).toBe('A1:C9');
        });
    });

    describe('inserting and deleting rows and columns', () => {
        const TABLE: SheetDto = {
            cells: {
                A1: { value: 'Item' }, B1: { value: 'Qty' },
                A2: { value: 'Widget' }, B2: { value: '2' },
                A3: { value: 'Gadget' }, B3: { value: '5' },
                B4: { formula: 'SUM(B2:B3)' },
            },
            columnWidths: { A: 28 },
            rowHeights: { 2: 30 },
            merges: ['A1:B1'],
            autoFilter: 'A1:B3',
            validations: { 'A2:A3': { type: 'list', values: ['Widget', 'Gadget'] } },
        };

        it('pushes rows down and moves everything keyed by row with them', () => {
            const next = withInsertedRow(TABLE, 2);

            expect(next.cells['A1']).toEqual({ value: 'Item' }, 'above the line, untouched');
            expect(next.cells['A2']).toBeUndefined('the new row is blank');
            expect(next.cells['A3']).toEqual({ value: 'Widget' }, 'pushed down');
            expect(next.rowHeights).toEqual({ 3: 30 });
            expect(next.merges).toEqual(['A1:B1'], 'above the line');
            expect(next.validations!['A3:A4']).toBeDefined('the rule follows its rows');
        });

        /**
         * The reason this is not just bookkeeping: a total that sums the wrong
         * range still looks like a number. `SUM(B2:B3)` over a table that grew
         * a row must become `SUM(B2:B4)`, not keep summing two of three lines.
         */
        it('grows a formula range when a row is inserted inside it', () => {
            const next = withInsertedRow(TABLE, 3);

            expect(next.cells['B5'].formula).toBe('SUM(B2:B4)');
        });

        it('grows the filter when a row is inserted inside it', () => {
            expect(withInsertedRow(TABLE, 3).autoFilter).toBe('A1:B4');
        });

        it('pulls rows up on delete, and takes the deleted row with it', () => {
            const next = withDeletedRow(TABLE, 2);

            expect(next.cells['A2']).toEqual({ value: 'Gadget' }, 'pulled up');
            expect(next.rowHeights).toBeUndefined('the sized row was the deleted one');
            expect(next.autoFilter).toBe('A1:B2', 'the filter shrinks');
            expect(next.cells['B3'].formula).toBe('SUM(B2:B2)');
        });

        /**
         * Excel answers `#REF!` for a formula pointing at a row that is gone,
         * and so does this. Shifting the reference to whatever moved INTO that
         * row would be silently wrong — the formula would keep computing, on
         * the wrong cell.
         */
        it('turns a reference to a deleted row into #REF!', () => {
            const sheet: SheetDto = { cells: { A1: { value: 'x' }, C1: { formula: 'A5*2' } } };

            expect(withDeletedRow(sheet, 5).cells['C1'].formula).toBe('#REF!*2');
        });

        it('shifts columns and everything keyed by column', () => {
            const next = withInsertedColumn(TABLE, 'A');

            expect(next.cells['B1']).toEqual({ value: 'Item' }, 'pushed right');
            expect(next.columnWidths).toEqual({ B: 28 });
            expect(next.merges).toEqual(['B1:C1']);
            expect(next.cells['C4'].formula).toBe('SUM(C2:C3)');
        });

        it('deletes a column, its cells and its width', () => {
            const next = withDeletedColumn(TABLE, 'A');

            expect(next.cells['A1']).toEqual({ value: 'Qty' }, 'B became A');
            expect(next.columnWidths).toBeUndefined();
            expect(next.validations).toBeUndefined('its only rule was in the deleted column');
            expect(next.autoFilter).toBe('A1:A3');
        });

        /** A merge reduced to one cell is no longer a merge. */
        it('drops a merge that shrinks to a single cell', () => {
            const sheet: SheetDto = { cells: {}, merges: ['A1:B1'] };

            expect(withDeletedColumn(sheet, 'B').merges).toBeUndefined();
        });

        /** Quoted text inside a formula is not a reference. */
        it('leaves a quoted reference alone', () => {
            expect(shiftFormula('IF(A1="B4","yes","no")', 'row', 1, 1))
                .toBe('IF(A2="B4","yes","no")');
        });

        it('refuses a line number that is not one', () => {
            expect(withInsertedRow(TABLE, 0)).toBe(TABLE);
            expect(withDeletedRow(TABLE, -1)).toBe(TABLE);
        });

        /** Contents only — an author clearing cells has not asked to dismantle the table. */
        it('clears contents without touching the layout', () => {
            const next = withClearedRange(TABLE, 'A2:B3');

            expect(next.cells['A2']).toBeUndefined();
            expect(next.cells['B3']).toBeUndefined();
            expect(next.cells['A1']).toEqual({ value: 'Item' });
            expect(next.merges).toEqual(['A1:B1']);
            expect(next.autoFilter).toBe('A1:B3');
            expect(next.validations!['A2:A3']).toBeDefined();
        });
    });

    describe('freezing panes', () => {
        const SHEET: SheetDto = { cells: { A1: { value: 'Item' } } };

        /**
         * One ref, not two counts, because that is what OOXML stores. `B2`
         * holds the first row and the first column.
         */
        it('reads a ref as the rows and columns it holds', () => {
            expect(frozenAt({ ...SHEET, freeze: 'B2' })).toEqual({ rows: 1, columns: 1 });
            expect(frozenAt({ ...SHEET, freeze: 'A3' })).toEqual({ rows: 2, columns: 0 });
            expect(frozenAt({ ...SHEET, freeze: 'C1' })).toEqual({ rows: 0, columns: 2 });
        });

        it('holds nothing when nothing says otherwise', () => {
            expect(frozenAt(SHEET)).toEqual({ rows: 0, columns: 0 });
            expect(frozenAt(undefined)).toEqual({ rows: 0, columns: 0 });
        });

        /**
         * Freezing above and left of the FIRST cell freezes nothing, and the
         * absence of the field is how the format says so -- the same answer
         * the backend and OOXML give, so a round trip through any of the three
         * says the same thing.
         */
        it('stores no freeze for A1, which freezes nothing', () => {
            expect(withFreeze(SHEET, 'A1').freeze).toBeUndefined();
            expect(withFreeze({ ...SHEET, freeze: 'B2' }, null).freeze).toBeUndefined();
        });

        it('freezes above and left of a cell', () => {
            expect(withFreeze(SHEET, 'C4').freeze).toBe('C4');
        });

        it('keeps the cells it is stored beside', () => {
            expect(withFreeze(SHEET, 'B2').cells['A1']).toEqual({ value: 'Item' });
        });

        it('survives serialisation', () => {
            const doc = parseSheetDocument('{"version":1,"sheets":{"S":{"cells":{}}}}').doc;
            doc.sheets['S'] = withFreeze(doc.sheets['S'], 'B2');

            const back = parseSheetDocument(serialiseSheetDocument(doc)).doc;

            expect(back.sheets['S'].freeze).toBe('B2');
        });
    });

    describe('borders', () => {
        const RANGE: SheetDto = { cells: { B2: { value: 'x' } } };

        /**
         * `outer` rules the OUTSIDE of the selection, not every cell in it.
         * That is what an author means by "box this table", and getting it
         * backwards is the kind of thing nobody reports — they just stop using
         * the control.
         */
        it('boxes a range without ruling the cells inside it', () => {
            const next = withBorderPreset(RANGE, 'B2:C3', 'outer');

            expect(next.cells['B2'].borders).toEqual({ top: 'thin', left: 'thin' });
            expect(next.cells['C2'].borders).toEqual({ top: 'thin', right: 'thin' });
            expect(next.cells['B3'].borders).toEqual({ bottom: 'thin', left: 'thin' });
            expect(next.cells['C3'].borders).toEqual({ bottom: 'thin', right: 'thin' });
        });

        it('rules every edge of every cell for all borders', () => {
            const next = withBorderPreset(RANGE, 'B2:C2', 'all');

            expect(next.cells['B2'].borders).toEqual({ top: 'thin', right: 'thin', bottom: 'thin', left: 'thin' });
            expect(next.cells['C2'].borders).toEqual({ top: 'thin', right: 'thin', bottom: 'thin', left: 'thin' });
        });

        /** One side means the side of the SELECTION — underline the table, not each row. */
        it('rules one side of the selection only', () => {
            const next = withBorderPreset(RANGE, 'B2:B4', 'bottom');

            expect(next.cells['B2'].borders).toBeUndefined();
            expect(next.cells['B3']).toBeUndefined('an untouched empty cell is not created');
            expect(next.cells['B4'].borders).toEqual({ bottom: 'thin' });
        });

        it('keeps the value under a ruled cell', () => {
            const next = withBorderPreset(RANGE, 'B2', 'all');

            expect(next.cells['B2'].value).toBe('x');
        });

        it('clears every edge in the range, including a perimeter', () => {
            const boxed = withBorderPreset(RANGE, 'B2:C3', 'outer');

            const cleared = withBorderPreset(boxed, 'B2:C3', 'none');

            expect(cleared.cells['B2'].borders).toBeUndefined();
            expect(cleared.cells['B2'].value).toBe('x', 'clearing borders is not clearing the cell');
            expect(cleared.cells['C3']).toBeUndefined('a cell left with nothing at all is dropped');
        });

        it('renders an edge as CSS, defaulting the colour', () => {
            expect(borderCss({ borders: { top: 'thin' } }, 'top')).toBe('1px solid #000000');
            expect(borderCss({ borders: { bottom: 'double #FF0000' } }, 'bottom')).toBe('3px double #FF0000');
            expect(borderCss({ borders: { top: 'thin' } }, 'left')).toBeNull();
            expect(borderCss({ borders: { top: 'nonsense' } }, 'top')).toBeNull();
            expect(borderCss(undefined, 'top')).toBeNull();
        });

        it('survives serialisation', () => {
            const doc = parseSheetDocument('{"version":1,"sheets":{"S":{"cells":{}}}}').doc;
            doc.sheets['S'] = withBorderPreset(doc.sheets['S'], 'A1:B2', 'outer');

            const back = parseSheetDocument(serialiseSheetDocument(doc)).doc;

            expect(back.sheets['S'].cells['A1'].borders).toEqual({ top: 'thin', left: 'thin' });
        });

        it('writes a colour only when it is not the black a bare style already means', () => {
            const black = withBorderPreset(RANGE, 'B2', 'top');
            const red = withBorderPreset(RANGE, 'B2', 'top', 'medium', '#ff0000');

            expect(black.cells['B2'].borders).toEqual({ top: 'thin' });
            expect(red.cells['B2'].borders).toEqual({ top: 'medium #FF0000' });
        });
    });

    /**
     * The grid is `border-collapse: collapse`, so the line between two cells is
     * ONE line and the browser picks a winner: at equal width and style, the
     * cell further up or further left. Every cell carries a grey gridline, so a
     * black `top` lost to the grey `bottom` above it and a black `left` lost to
     * the grey `right` beside it — "All borders" on a single cell drew its
     * right and bottom ONLY, which is exactly what was reported.
     */
    describe('rendered borders', () => {
        const SHEET: SheetDto = { cells: { C3: { value: 'x' } } };

        it('renders every edge an author ruled', () => {
            const ruled = withBorderPreset(SHEET, 'C3', 'all');

            expect(borderCssAt(ruled, 'C3')).toEqual({
                'border-top': '1px solid #000000',
                'border-right': '1px solid #000000',
                'border-bottom': '1px solid #000000',
                'border-left': '1px solid #000000',
            });
        });

        it('declares the same line on BOTH sides of a shared edge', () => {
            const ruled = withBorderPreset(SHEET, 'C3', 'all');

            // Without these two the black top and left lose the collapse
            // conflict to the neighbour's grey gridline and never appear.
            expect(borderCssAt(ruled, 'C2')['border-bottom'])
                .withContext('the cell above shares C3 top')
                .toBe('1px solid #000000');
            expect(borderCssAt(ruled, 'B3')['border-right'])
                .withContext('the cell beside shares C3 left')
                .toBe('1px solid #000000');
        });

        it('leaves a cell with no ruled neighbour alone', () => {
            const ruled = withBorderPreset(SHEET, 'C3', 'all');

            expect(borderCssAt(ruled, 'A1')).toEqual({});
        });

        it('does not look off the edge of the grid', () => {
            const ruled = withBorderPreset(SHEET, 'A1', 'all');

            expect(Object.keys(borderCssAt(ruled, 'A1')).length).toBe(4);
        });

        /** An explicit edge is the cell's own; only an ABSENT one is borrowed. */
        it('keeps the line a cell rules itself over the one across from it', () => {
            let sheet = withBorderPreset(SHEET, 'C3', 'top', 'thick');
            sheet = withBorderPreset(sheet, 'C2', 'bottom', 'dotted');

            expect(borderCssAt(sheet, 'C3')['border-top']).toBe('3px solid #000000');
            expect(borderCssAt(sheet, 'C2')['border-bottom']).toBe('1px dotted #000000');
        });
    });

    /**
     * The control has to show what the cells ARE. A cell with every edge ruled
     * whose control still reads "Borders…" tells the author their gesture did
     * not land — which is how this was reported.
     */
    describe('reading borders back', () => {
        const SHEET: SheetDto = { cells: { B2: { value: 'x' } } };

        it('names the gesture that produced what is there', () => {
            const all = withBorderPreset(SHEET, 'B2', 'all');
            const boxed = withBorderPreset(SHEET, 'B2:C3', 'outer');
            const underlined = withBorderPreset(SHEET, 'B2:B4', 'bottom');

            expect(borderStateIn(all, 'B2').preset).toBe('all');
            expect(borderStateIn(boxed, 'B2:C3').preset).toBe('outer');
            expect(borderStateIn(underlined, 'B2:B4').preset).toBe('bottom');
        });

        it('reports the line and the colour, so the palette opens on them', () => {
            const ruled = withBorderPreset(SHEET, 'B2', 'all', 'double', '#FF0000');

            expect(borderStateIn(ruled, 'B2')).toEqual({
                preset: 'all',
                style: 'double',
                colour: '#FF0000',
            });
        });

        it('defaults the colour of a bare style, as the renderer does', () => {
            const ruled = withBorderPreset(SHEET, 'B2', 'all');

            expect(borderStateIn(ruled, 'B2').colour).toBe('#000000');
        });

        it('says nothing about a selection with no borders at all', () => {
            expect(borderStateIn(SHEET, 'B2')).toEqual({ preset: null, style: null, colour: null });
        });

        /**
         * Guessing one would let the next gesture quietly rewrite the edges it
         * did not describe.
         */
        it('names no line when the edges disagree', () => {
            let sheet = withBorderPreset(SHEET, 'B2', 'top', 'thick');
            sheet = withBorderPreset(sheet, 'B2', 'bottom', 'dotted');

            const state = borderStateIn(sheet, 'B2');

            expect(state.style).toBeNull();
            expect(state.colour).toBeNull();
        });

        it('names no gesture for a shape no gesture makes', () => {
            const partial = withBorderPreset(SHEET, 'B2', 'all');

            // B2 is ruled, C2 is not: neither all nor any side describes that.
            expect(borderStateIn(partial, 'B2:C2').preset).toBeNull();
        });

        /**
         * On ONE cell the two gestures produce the same four edges. "all" is
         * the one the author reached for, so it is the one reported back.
         */
        it('prefers all over outer where they cannot be told apart', () => {
            const ruled = withBorderPreset(SHEET, 'B2', 'outer');

            expect(borderStateIn(ruled, 'B2').preset).toBe('all');
        });

        it('reports the edges the selection holds, not the ones it borrows', () => {
            const ruled = withBorderPreset(SHEET, 'B2', 'all');

            expect(borderStateIn(ruled, 'C2').preset).toBeNull();
        });
    });

    describe('form elements', () => {
        it('finds the rule covering a cell, and its range', () => {
            const sheet: SheetDto = {
                cells: {},
                validations: { 'B2:B9': { type: 'list', values: ['New', 'Open'] } },
            };

            expect(validationRangeAt(sheet, 'B5')).toBe('B2:B9');
            expect(validationAt(sheet, 'B5')?.values).toEqual(['New', 'Open']);
            expect(validationAt(sheet, 'C5')).toBeNull('outside the range');
        });

        /**
         * A single cell is a legal key here where it is not for a merge or a
         * filter — "this one cell is a dropdown" is the ordinary case — so the
         * lookup has to widen it rather than fail to parse it.
         */
        it('treats a single-cell key as a range of one', () => {
            const sheet: SheetDto = { cells: {}, validations: { B2: { type: 'checkbox' } } };

            expect(validationRangeAt(sheet, 'B2')).toBe('B2');
            expect(validationRangeAt(sheet, 'B3')).toBeNull();
        });

        /** A checkbox's options are its TYPE, never read from the document. */
        it('gives a checkbox its own two options whatever the file says', () => {
            expect(validationOptions({ type: 'checkbox', values: ['Ja', 'Nein'] })).toEqual(['TRUE', 'FALSE']);
            expect(validationOptions({ type: 'list', values: ['A', 'B'] })).toEqual(['A', 'B']);
            expect(validationOptions({ type: 'whole', min: '1' })).toEqual([], 'not a list-shaped rule');
        });

        /**
         * Two rules over one cell is a state whose behaviour depends on which
         * the reader applies, and Excel keeps only one. The author sees the rule
         * they just made — the one they were thinking about.
         */
        it('replaces a rule the new range overlaps', () => {
            const sheet: SheetDto = {
                cells: {},
                validations: { 'A1:A5': { type: 'checkbox' }, 'C1:C5': { type: 'checkbox' } },
            };

            const next = withValidation(sheet, 'A3:A9', { type: 'list', values: ['x'] });

            expect(Object.keys(next.validations!).sort()).toEqual(['A3:A9', 'C1:C5']);
        });

        /** A dropdown with nothing to offer is a cell nobody can type into. */
        it('refuses a list with no options', () => {
            const sheet: SheetDto = { cells: {} };

            expect(withValidation(sheet, 'A1:A5', { type: 'list', values: [] })).toBe(sheet);
            expect(withValidation(sheet, 'not a range', { type: 'checkbox' })).toBe(sheet);
        });

        it('removes a rule and drops the key with the last one', () => {
            const sheet = withValidation({ cells: {} }, 'A1:A5', { type: 'checkbox' });

            expect('validations' in withoutValidation(sheet, 'A1:A5')).toBeFalse();
        });

        it('survives serialisation', () => {
            const doc = parseSheetDocument('{"version":1,"sheets":{"S":{"cells":{}}}}').doc;
            doc.sheets['S'] = withValidation(doc.sheets['S'], 'B2:B9', { type: 'list', values: ['New', 'Open'] });

            const back = parseSheetDocument(serialiseSheetDocument(doc)).doc;

            expect(validationAt(back.sheets['S'], 'B5')?.values).toEqual(['New', 'Open']);
        });
    });

    describe('sheetLookup', () => {
        const sheet = (): SheetDto => ({
            cells: {},
            merges: ['A1:C1', 'E4:F6'],
            validations: {
                'B2:B9': { type: 'list', values: ['New', 'Open'] },
                'D3': { type: 'checkbox' },
            },
        });

        it('answers exactly what the scan it replaces answers', () => {
            const s = sheet();
            const index = sheetLookup(s);

            // Anchors, swallowed cells, and cells outside every range: the
            // three answers the grid actually asks for.
            for (const ref of ['A1', 'B1', 'C1', 'D1', 'E4', 'F6', 'E7', 'Z99', 'A2']) {
                expect(index.mergeCovering(ref)).withContext(ref).toBe(mergeCovering(s, ref));
            }
            for (const ref of ['B2', 'B9', 'B10', 'D3', 'D4', 'A1']) {
                expect(index.validationAt(ref)).withContext(ref).toEqual(validationAt(s, ref));
            }
        });

        it('gives the same answer when asked again -- the grid asks six times a cell', () => {
            const index = sheetLookup(sheet());

            expect(index.mergeCovering('B1')).toBe('A1:C1');
            expect(index.mergeCovering('B1')).toBe('A1:C1');
            expect(index.validationAt('B5')?.type).toBe('list');
            expect(index.validationAt('B5')?.type).toBe('list');

            // A miss is remembered too, and stays a miss.
            expect(index.mergeCovering('Z1')).toBeNull();
            expect(index.mergeCovering('Z1')).toBeNull();
        });

        /**
         *  The staleness guard. A lookup is a snapshot, and the dialog holds
         * it in a computed over the document so an edit throws it away. Were it
         * kept, the grid would go on drawing a merge the author had just undone.
         */
        it('is a SNAPSHOT -- a sheet edited after it was built needs a new one', () => {
            const before = sheet();
            const index = sheetLookup(before);
            expect(index.mergeCovering('B1')).toBe('A1:C1');

            const after = withoutMerge(before, 'A1:C1');

            expect(sheetLookup(after).mergeCovering('B1')).toBeNull();
            expect(mergeCovering(after, 'B1')).toBeNull();
        });

        it('keeps first-match-wins when ranges overlap', () => {
            const s: SheetDto = {
                cells: {},
                merges: ['A1:C1', 'A1:A3'],
                validations: {
                    'A1:C3': { type: 'checkbox' },
                    'A1:A9': { type: 'list', values: ['x'] },
                },
            };
            const index = sheetLookup(s);

            expect(index.mergeCovering('A1')).toBe(mergeCovering(s, 'A1'));
            expect(index.validationAt('A1')).toEqual(validationAt(s, 'A1'));
            expect(index.validationAt('A1')?.type).toBe('checkbox');
        });

        it('answers nothing for an absent sheet, and for a sheet with neither', () => {
            expect(sheetLookup(undefined).mergeCovering('A1')).toBeNull();
            expect(sheetLookup(undefined).validationAt('A1')).toBeNull();
            expect(sheetLookup({ cells: {} }).mergeCovering('A1')).toBeNull();
            expect(sheetLookup({ cells: {} }).validationAt('A1')).toBeNull();
        });

        it('ignores a range it cannot parse rather than counting it', () => {
            const index = sheetLookup({ cells: {}, merges: ['not a range', 'A1:B2'] });

            expect(index.mergeCovering('A1')).toBe('A1:B2');
            expect(index.mergeCovering('Z9')).toBeNull();
        });
    });
});
