import { HttpClientTestingModule, HttpTestingController } from '@angular/common/http/testing';
import { TestBed } from '@angular/core/testing';
import { DIALOG_DATA, DialogRef } from '@angular/cdk/dialog';
import { NgxsModule, Store } from '@ngxs/store';
import { AppConfigState } from '@coolms/core-angular';
import { SheetEditorDialogComponent } from './sheet-editor-dialog.component';
import { columnWidthFromPx, columnWidthToPx, gridExtent, indexToColumn } from './sheet-document.model';
import { dateToSerial } from './number-format';

/**
 * The grid surface for a native `.dsheet` (ADR-155).
 *
 * Rendered through TestBed rather than asserted on the model alone: the model
 * has its own spec, and what can go wrong HERE is the wiring — a grid that does
 * not reach the cells the document defines, an edit that does not reach the
 * document, or a save that writes something the backend cannot read.
 */
describe('SheetEditorDialogComponent', () => {
    let httpMock: HttpTestingController;

    const CONTENT_URL = 'https://api.test/vfs/content?path=%2Fdocs%2F.templates%2Finvoice.dsheet';

    function makeFixture() {
        const fixture = TestBed.createComponent(SheetEditorDialogComponent);
        fixture.detectChanges();

        return fixture;
    }

    /** Answer the load request the constructor fires. */
    function respondWith(doc: unknown): void {
        httpMock.expectOne(CONTENT_URL).flush({ content: JSON.stringify(doc) });
    }

    function cellInput(fixture: ReturnType<typeof makeFixture>, ref: string): HTMLInputElement | null {
        return fixture.nativeElement.querySelector(`input[aria-label="${ref}"]`);
    }

    /**
     * Found by LABEL, never by index. An earlier version of this spec used
     * `querySelectorAll('button')[2]`, and adding one toolbar button silently
     * turned every "click save" into "click cancel" — which a test asserting
     * that nothing was sent would have passed. Position is not identity.
     */
    function saveButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const buttons: HTMLButtonElement[] = Array.from(fixture.nativeElement.querySelectorAll('button'));
        const save = buttons.find(b => (b.textContent ?? '').trim().startsWith('Save'));
        if (!save) throw new Error('no Save button rendered');

        return save;
    }

    /** What the toolbar reports as the active cell. */
    function activeRef(fixture: ReturnType<typeof makeFixture>): string {
        return fixture.nativeElement.querySelector('.sheet-editor__active')?.textContent?.trim() ?? '';
    }

    /** Whether a cell is drawn as part of the selected range. */
    function inSelection(fixture: ReturnType<typeof makeFixture>, ref: string): boolean {
        const td = cellInput(fixture, ref)?.closest('td');

        return td?.classList.contains('sheet-editor__cell--in-range') ?? false;
    }

    /** As {@link saveButton} — by label, never by index. */
    function mergeButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const buttons: HTMLButtonElement[] = Array.from(fixture.nativeElement.querySelectorAll('button'));
        const merge = buttons.find(b => (b.textContent ?? '').trim().startsWith('Merge')
            || (b.textContent ?? '').trim().startsWith('Unmerge'));
        if (!merge) throw new Error('no Merge button rendered');

        return merge;
    }

    /**
     * A shift-click, as a real browser delivers one.
     *
     * NO `focus` event follows, and that is the point. Shift-click inside a
     * focused text input is the browser's native extend-the-text-selection
     * gesture: it highlights characters in the cell the author started from and
     * SUPPRESSES the focus change. The component cancels the event to take the
     * gesture over, so `mousedown` is the only thing that arrives.
     *
     * An earlier version of this helper also dispatched `focus`, which no real
     * shift-click produces. It passed while the feature did not work in the
     * admin at all — caught by clicking through it in a browser, not here. The
     * `defaultPrevented` assertion is what keeps this honest: if the component
     * stops cancelling, the real gesture silently reverts to selecting text.
     */
    function shiftClick(input: HTMLInputElement): void {
        const event = new MouseEvent('mousedown', { shiftKey: true, bubbles: true, cancelable: true });
        input.dispatchEvent(event);
        expect(event.defaultPrevented)
            .withContext('shift-click must be cancelled, or the browser selects text instead')
            .toBeTrue();
    }

    beforeEach(() => {
        TestBed.configureTestingModule({
            imports: [
                HttpClientTestingModule,
                SheetEditorDialogComponent,
                // AppConfigState must be registered: the component reads the
                // VFS content URL from the manifest via selectSnapshot, and an
                // unregistered state yields undefined there.
                NgxsModule.forRoot([AppConfigState]),
            ],
            providers: [
                { provide: DIALOG_DATA, useValue: { node: { path: '/docs/.templates/invoice.dsheet', name: 'invoice.dsheet', mimeType: 'application/x-coolms-sheet+json' } } },
                { provide: DialogRef, useValue: { close: (): void => {} } },
            ],
        });

        // The manifest is what turns a node path into a content URL; without it
        // the component never issues a request and every expectation below
        // would fail for the wrong reason.
        TestBed.inject(Store).reset({
            appConfig: {
                loaded: true,
                config: { manifest: { vfs: { fileContentUrl: 'https://api.test/vfs/content?path={path}' } } },
            },
        });

        httpMock = TestBed.inject(HttpTestingController);
    });

    afterEach(() => httpMock.verify());

    it('renders every cell the document defines, plus room to type past them', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                Invoice: {
                    cells: { A1: { value: 'Total' }, A2: { value: '5' }, B4: { formula: 'A2*2' } },
                },
            },
        });
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')?.value).toBe('Total');
        // A formula shows what it COMPUTES, not its own text -- the cell being
        // edited is the one that shows the formula, and none is focused here.
        expect(cellInput(fixture, 'B4')?.value).toBe('10');
        // The padding is what makes it an editor: there is somewhere below the
        // last used cell to add another line.
        expect(cellInput(fixture, 'A9')).not.toBeNull();
    });

    describe('a formula shows its result, and its text where that is the useful thing', () => {
        /** Focus a cell the way the grid does: the input's own focus event. */
        function focus(fixture: ReturnType<typeof makeFixture>, ref: string): void {
            cellInput(fixture, ref)!.dispatchEvent(new Event('focus'));
            fixture.detectChanges();
        }

        it('shows the formula in the cell being edited, and the result elsewhere', () => {
            const fixture = makeFixture();
            respondWith({
                version: 1,
                sheets: { S: { cells: { A1: { value: '4' }, A2: { formula: 'A1*3' } } } },
            });
            fixture.detectChanges();

            expect(cellInput(fixture, 'A2')?.value).toBe('12');

            // Focusing it swaps the result for the formula, which is the whole
            // reason this grid needs no separate formula bar.
            focus(fixture, 'A2');
            expect(cellInput(fixture, 'A2')?.value).toBe('=A1*3');
        });

        it('recomputes a dependent cell when its input changes', () => {
            const fixture = makeFixture();
            respondWith({
                version: 1,
                sheets: { S: { cells: { A1: { value: '4' }, A2: { formula: 'A1*3' } } } },
            });
            fixture.detectChanges();
            expect(cellInput(fixture, 'A2')?.value).toBe('12');

            const a1 = cellInput(fixture, 'A1')!;
            a1.value = '10';
            a1.dispatchEvent(new Event('change'));
            fixture.detectChanges();

            // The document reference is replaced on edit, which is what makes
            // the evaluation re-run. Without that this reads 12 for ever.
            expect(cellInput(fixture, 'A2')?.value).toBe('30');
        });

        it('shows an error code rather than a number it cannot produce', () => {
            const fixture = makeFixture();
            respondWith({ version: 1, sheets: { S: { cells: { A1: { formula: '1/0' } } } } });
            fixture.detectChanges();

            expect(cellInput(fixture, 'A1')?.value).toBe('#DIV/0!');
        });

        it('offers completions, and accepts one with Enter', () => {
            const fixture = makeFixture();
            respondWith({ version: 1, sheets: { S: { cells: {} } } });
            fixture.detectChanges();

            const input = cellInput(fixture, 'A1')!;
            // REAL focus, not a synthetic event. Caret tracking refuses an
            // input while another field holds the focus, and a synthetic focus
            // event leaves the document's active element on the body — which
            // is how these specs missed the point-mode defect entirely.
            input.focus();
            input.value = '=SU';
            input.setSelectionRange(3, 3);
            input.dispatchEvent(new Event('input'));
            fixture.detectChanges();

            const helper = fixture.nativeElement.querySelector('.sheet-editor__helper');
            expect(helper).withContext('the popup should be open').not.toBeNull();
            expect(helper.textContent).toContain('SUM');

            input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
            fixture.detectChanges();

            // The bracket comes with it and the caret sits inside, because the
            // next keystroke is an argument.
            expect(input.value).toBe('=SUM(');
            expect(input.selectionStart).toBe('=SUM('.length);
        });

        it('describes the function once the bracket is there', () => {
            const fixture = makeFixture();
            respondWith({ version: 1, sheets: { S: { cells: {} } } });
            fixture.detectChanges();

            const input = cellInput(fixture, 'A1')!;
            // REAL focus, not a synthetic event. Caret tracking refuses an
            // input while another field holds the focus, and a synthetic focus
            // event leaves the document's active element on the body — which
            // is how these specs missed the point-mode defect entirely.
            input.focus();
            input.value = '=ROUND(A1, ';
            input.setSelectionRange(input.value.length, input.value.length);
            input.dispatchEvent(new Event('input'));
            fixture.detectChanges();

            const helper = fixture.nativeElement.querySelector('.sheet-editor__helper');
            expect(helper).not.toBeNull();
            expect(helper.textContent).toContain('ROUND');
            expect(helper.textContent).toContain('digits');
        });

        it('Escape closes the popup and NOT the dialog', () => {
            const fixture = makeFixture();
            respondWith({ version: 1, sheets: { S: { cells: {} } } });
            fixture.detectChanges();

            const input = cellInput(fixture, 'A1')!;
            // REAL focus, not a synthetic event. Caret tracking refuses an
            // input while another field holds the focus, and a synthetic focus
            // event leaves the document's active element on the body — which
            // is how these specs missed the point-mode defect entirely.
            input.focus();
            input.value = '=SU';
            input.setSelectionRange(3, 3);
            input.dispatchEvent(new Event('input'));
            fixture.detectChanges();
            expect(fixture.nativeElement.querySelector('.sheet-editor__helper')).not.toBeNull();

            const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
            input.dispatchEvent(escape);
            fixture.detectChanges();

            expect(fixture.nativeElement.querySelector('.sheet-editor__helper')).toBeNull();
            // Stopped here, so the CDK dialog never sees it -- otherwise the
            // whole editor closes on a keystroke meant for a popup.
            expect(escape.defaultPrevented).toBe(true);
        });

        it('offers nothing for a cell that is not a formula', () => {
            const fixture = makeFixture();
            respondWith({ version: 1, sheets: { S: { cells: {} } } });
            fixture.detectChanges();

            const input = cellInput(fixture, 'A1')!;
            // REAL focus, not a synthetic event. Caret tracking refuses an
            // input while another field holds the focus, and a synthetic focus
            // event leaves the document's active element on the body — which
            // is how these specs missed the point-mode defect entirely.
            input.focus();
            input.value = 'SUM';
            input.setSelectionRange(3, 3);
            input.dispatchEvent(new Event('input'));
            fixture.detectChanges();

            expect(fixture.nativeElement.querySelector('.sheet-editor__helper')).toBeNull();
        });

        it('keeps showing the FORMULA when it waits on a template token', () => {
            const fixture = makeFixture();
            respondWith({
                version: 1,
                sheets: {
                    S: { cells: { A1: { value: '{var:order.total}' }, A2: { formula: 'A1*2' } } },
                },
            });
            fixture.detectChanges();

            // Not blank: there is no result yet, and an empty cell would read as
            // broken rather than as pending.
            expect(cellInput(fixture, 'A2')?.value).toBe('=A1*2');
        });
    });

    it('writes an edit back into the document it will save', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: {} } } });
        fixture.detectChanges();

        const input = cellInput(fixture, 'A1')!;
        input.value = '{var:company.name}';
        input.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        expect(save.request.method).toBe('PUT');
        const written: { sheets: Record<string, { cells: Record<string, { value?: string }> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['A1'].value).toBe('{var:company.name}');
        save.flush({ contentHash: 'x' });
    });

    /**
     * The #1977 rule, at the surface an operator touches. `@` declares the cell
     * text rather than arithmetic — editing the TEXT must not discard it, or
     * the next generation promotes `00412` to 412.
     */
    it('keeps a cell number format through an edit to its text', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { B2: { value: '{var:order}', numberFormat: '@' } } } } });
        fixture.detectChanges();

        const input = cellInput(fixture, 'B2')!;
        input.value = '{var:order.number}';
        input.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { cells: Record<string, { numberFormat?: string }> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['B2'].numberFormat).toBe('@');
        save.flush({ contentHash: 'x' });
    });

    /**
     * The format is the one property that changes MEANING rather than
     * appearance (#1977). Until now it could only be set by hand-editing JSON,
     * which is the operator this format exists for being sent to a text editor.
     */
    it('marks the focused cell as Text and saves that', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { B2: { value: '00412' } } } } });
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        const select: HTMLSelectElement = fixture.nativeElement.querySelector('.sheet-editor__format');
        select.value = '@';
        select.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { cells: Record<string, { numberFormat?: string; value?: string }> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['B2'].numberFormat).toBe('@');
        expect(written.sheets['S'].cells['B2'].value).toBe('00412', 'the value must be untouched');
        save.flush({ contentHash: 'x' });
    });

    /**
     * A `.dsheet` may carry any OOXML code. One the menu does not list must be
     * SHOWN, not displayed as General — otherwise the next unrelated toolbar
     * use silently overwrites a currency format the author hand-wrote.
     */
    it('surfaces a format code the menu does not list instead of hiding it', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { C3: { value: '9.99', numberFormat: '#,##0.00\\ [$€-407]' } } } } });
        fixture.detectChanges();

        cellInput(fixture, 'C3')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        const select: HTMLSelectElement = fixture.nativeElement.querySelector('.sheet-editor__format');
        expect(select.value).toBe('#,##0.00\\ [$€-407]');
        expect(select.textContent).toContain('Custom:');
    });

    /** Formatting a cell that has no content yet must still be recorded. */
    it('can format an empty cell, because intent precedes typing', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: {} } } });
        fixture.detectChanges();

        cellInput(fixture, 'A1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        const select: HTMLSelectElement = fixture.nativeElement.querySelector('.sheet-editor__format');
        select.value = '@';
        select.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { cells: Record<string, { numberFormat?: string }> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['A1'].numberFormat).toBe('@');
        save.flush({ contentHash: 'x' });
    });

    /**
     * The renderer has handled merges correctly since [#1998]; until now an
     * author could only create one by hand-editing JSON.
     *
     * Asserted on the RENDERED grid as well as the saved document, because the
     * two can disagree: a merge that stored correctly but drew as separate cells
     * would look to the author like it had not taken.
     */
    it('merges a shift-clicked range and renders it as one spanning cell', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Invoice' } } } } });
        fixture.detectChanges();

        cellInput(fixture, 'A1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();
        shiftClick(cellInput(fixture, 'D1')!);
        fixture.detectChanges();

        mergeButton(fixture).click();
        fixture.detectChanges();

        // The anchor spans the range, and the swallowed cells are GONE from the
        // DOM rather than hidden — a hidden input inside a merged region would
        // still be focusable and editable.
        const anchor = cellInput(fixture, 'A1')!.closest('td')!;
        expect(anchor.getAttribute('colspan')).toBe('4');
        expect(cellInput(fixture, 'B1')).toBeNull();
        expect(cellInput(fixture, 'C1')).toBeNull();
        expect(cellInput(fixture, 'D1')).toBeNull();
        // A different row must be untouched by a single-row merge.
        expect(cellInput(fixture, 'B2')).not.toBeNull();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { merges?: string[] }> } = JSON.parse(save.request.body.content);
        expect(written.sheets['S'].merges).toEqual(['A1:D1']);
        save.flush({ contentHash: 'x' });
    });

    /** Unmerging is the only way back, so the button has to do both jobs. */
    it('unmerges the merge under the cursor and restores the cells', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Invoice' } }, merges: ['A1:D1'] } } });
        fixture.detectChanges();

        expect(cellInput(fixture, 'B1')).toBeNull('the covered cell must start hidden');

        cellInput(fixture, 'A1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        const button = mergeButton(fixture);
        expect(button.textContent!.trim()).toBe('Unmerge', 'a merged cell offers the way out');
        button.click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'B1')).not.toBeNull();
        expect(cellInput(fixture, 'A1')!.closest('td')!.getAttribute('colspan')).toBe('1');

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { merges?: string[] }> } = JSON.parse(save.request.body.content);
        expect(written.sheets['S'].merges).toBeUndefined();
        save.flush({ contentHash: 'x' });
    });

    /**
     * A plain click must COLLAPSE the selection. Without this, clicking A1 then
     * clicking D1 would merge them — every navigation click in the grid would
     * arm the Merge button over a range the author never asked for.
     */
    it('collapses the selection on a plain click, so Merge needs a real range', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: {} } } });
        fixture.detectChanges();

        cellInput(fixture, 'A1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();
        cellInput(fixture, 'D1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        expect(mergeButton(fixture).disabled).toBeTrue();
    });

    /**
     * Column width was the last thing in the grid that could only be set by
     * hand-editing JSON. The backend has always written it — the author simply
     * had no way to say it.
     */
    it('stores a width for the focused cell\'s column and renders it', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { B2: { value: 'Description' } } } } });
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        const width: HTMLInputElement = fixture.nativeElement.querySelector('.sheet-editor__width');
        width.value = '28';
        width.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        // Rendered, not merely stored: a width the grid ignored would look to
        // the author like it had not been saved. B is the second data column,
        // and the first <col> is the row-header gutter.
        const cols: HTMLTableColElement[] = Array.from(fixture.nativeElement.querySelectorAll('colgroup col'));
        expect(cols[2].style.width).toBe('201px', '28 characters ≈ 28*7+5 px');

        // ...and LAID OUT at that width, which the style attribute alone does
        // not prove. A `<col>` width is only a SUGGESTION under `table-layout:
        // auto` — the browser may widen a column its content overflows — so this
        // asserts the real box, in real Chrome. Karma runs one; jsdom would not
        // have caught a table that ignored the column group.
        const cellWidth = cellInput(fixture, 'B2')!.closest('td')!.getBoundingClientRect().width;
        expect(Math.round(cellWidth)).toBe(201, 'the grid must lay the column out at the stored width');

        // A column the document says nothing about must be UNCHANGED. Making
        // widths binding meant switching the table to fixed layout, which
        // re-sizes every column — including the ones nobody asked to change — so
        // this pins the default against a regression no other assertion covers.
        const untouched = cellInput(fixture, 'C2')!.closest('td')!.getBoundingClientRect().width;
        expect(Math.round(untouched)).toBe(141, 'an unwidthed column keeps the grid default');

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { columnWidths?: Record<string, number> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].columnWidths).toEqual({ B: 28 });
        save.flush({ contentHash: 'x' });
    });

    /**
     * Excel reads width 0 as a HIDDEN column, so a stray zero must leave the
     * document alone rather than make a column disappear from the generated
     * workbook with nothing in the editor to explain it.
     *
     * A CELL IS EDITED FIRST, and that is load-bearing for the same reason it is
     * in the unreadable-file test below: Save is `[disabled]` until the document
     * is dirty, and refusing the zero leaves it clean — so the first draft of
     * this test sent no request at all and failed on `expectOne`. Dirtying the
     * document by other means is what lets the save through, and only then does
     * the assertion about the STORED width mean anything.
     */
    it('ignores a width of zero rather than hiding the column', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'x' } }, columnWidths: { A: 12 } } } });
        fixture.detectChanges();

        cellInput(fixture, 'A1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        const width: HTMLInputElement = fixture.nativeElement.querySelector('.sheet-editor__width');
        expect(width.value).toBe('12', 'the stored width must show');

        width.value = '0';
        width.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        // An unrelated edit, purely to enable Save.
        const cell = cellInput(fixture, 'A1')!;
        cell.value = 'edited';
        cell.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        const save = saveButton(fixture);
        expect(save.disabled).toBeFalse();
        save.click();

        const request = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { columnWidths?: Record<string, number> }> } =
            JSON.parse(request.request.body.content);
        expect(written.sheets['S'].columnWidths).toEqual({ A: 12 }, 'the zero must not have been stored');
        request.flush({ contentHash: 'x' });
    });

    /**
     * A file the editor cannot read must NOT be saved over. Opening it blank
     * and letting a save through would replace the operator's document with an
     * empty grid — and the editor is the only place they can repair it from.
     *
     * The cell is EDITED FIRST, and that is load-bearing: Save is `[disabled]`
     * until the document is dirty, so a version of this test that only clicked
     * the button passed with the guard deleted — it was asserting the disabled
     * attribute, not the refusal. Caught by mutating the guard and watching
     * nothing fail.
     */
    it('refuses to save over a file it could not parse, even once edited', () => {
        const fixture = makeFixture();
        httpMock.expectOne(CONTENT_URL).flush({ content: 'not a .dsheet at all' });
        fixture.detectChanges();

        const input = cellInput(fixture, 'A1')!;
        input.value = 'typed into a file we could not read';
        input.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        const save = saveButton(fixture);
        expect(save.disabled).toBeFalse();
        save.click();

        httpMock.expectNone(CONTENT_URL);
    });

    // ── header selection and resize ──────────────────────────────────────────

    function header(fixture: ReturnType<typeof makeFixture>, label: string): HTMLElement {
        const th: HTMLElement | null = fixture.nativeElement.querySelector(`th[aria-label="${label}"]`);
        if (!th) throw new Error(`no header labelled "${label}"`);

        return th;
    }

    /**
     * How many rows are in the DOM right now — the viewport window, not the
     * grid's height. Since #2067 those are different numbers, and conflating
     * them is how these specs started asserting the virtualiser instead of the
     * selection.
     */
    function renderedRows(fixture: ReturnType<typeof makeFixture>): number {
        return fixture.nativeElement.querySelectorAll('.sheet-editor__row-head').length;
    }

    /** As {@link renderedRows}, for the column axis virtualised in #2068. */
    function renderedCols(fixture: ReturnType<typeof makeFixture>): number {
        return fixture.nativeElement.querySelectorAll('.sheet-editor__col-head').length;
    }

    /** The refs the grid is currently highlighting as selected. */
    function highlighted(fixture: ReturnType<typeof makeFixture>): string[] {
        const cells: HTMLElement[] = Array.from(
            fixture.nativeElement.querySelectorAll('.sheet-editor__cell--in-range'));

        return cells
            .map(td => td.querySelector('input')?.getAttribute('aria-label') ?? '')
            .filter(Boolean);
    }

    function mousedown(el: HTMLElement, init: MouseEventInit = {}): MouseEvent {
        const event = new MouseEvent('mousedown', { bubbles: true, cancelable: true, ...init });
        el.dispatchEvent(event);

        return event;
    }

    const SMALL = { version: 1, sheets: { S: { cells: { A1: { value: 'a' }, B2: { value: 'b' } } } } };

    /**
     * `gridExtent` floors the grid whatever the document holds, so a two-cell
     * sheet still renders a spreadsheet's worth. The selection follows what is
     * ON SCREEN rather than what happens to be populated — an author dragging a
     * column selects the column they can see, including the empty room below
     * the last value that the grid deliberately provides for typing into.
     *
     * DERIVED rather than hard-coded. These specs are about the selection rule;
     * the floor is incidental to them, and writing it out meant three of them
     * failed the day it changed (#2066) while testing nothing that had broken.
     */
    const { rows: GRID_ROWS, cols: GRID_COLS } = gridExtent({ cells: {} });

    it('selects a whole column from its header', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const event = mousedown(header(fixture, 'Select column A'));
        fixture.detectChanges();

        // Cancelled, or the mousedown moves focus into a cell input whose focus
        // handler collapses the selection it just made.
        expect(event.defaultPrevented)
            .withContext('the header mousedown must be cancelled')
            .toBeTrue();

        // Counted against the RENDERED rows, not the grid's full height. Since
        // #2067 only the visible window exists in the DOM, so a count against
        // GRID_ROWS would be asserting how much got virtualised rather than
        // what got selected.
        const refs = highlighted(fixture);
        expect(refs.length).toBe(renderedRows(fixture));
        expect(refs.every(ref => ref.startsWith('A'))).toBeTrue();
        expect(refs).toContain('A1');
    });

    /**
     * The selection is a RANGE, not a set of rendered nodes.
     *
     * This is the regression virtualisation invites: highlight the visible
     * cells, then let the window move and find the selection has quietly become
     * whatever happened to be on screen when it was made. Selecting a column
     * and scrolling must reveal more of the SAME selection.
     */
    it('keeps a column selected in rows that were not rendered when it was made', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        mousedown(header(fixture, 'Select column A'));
        fixture.detectChanges();
        const before = highlighted(fixture);
        expect(before).not.toContain('A400', 'row 400 must be far outside the initial window');

        // Scroll deep into the sheet and let the window move.
        const body: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__body');
        body.scrollTop = 10_000;
        body.dispatchEvent(new Event('scroll'));
        fixture.detectChanges();

        const after = highlighted(fixture);
        expect(after.length).toBeGreaterThan(0);
        expect(after.every(ref => ref.startsWith('A'))).toBeTrue();
        expect(after).not.toContain('A1', 'the window should have moved off the top');
    });

    it('selects a whole row from its header', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        mousedown(header(fixture, 'Select row 2'));
        fixture.detectChanges();

        const refs = highlighted(fixture);
        expect(refs.length).toBe(renderedCols(fixture));
        expect(refs.every(ref => ref.endsWith('2'))).toBeTrue();
        expect(refs).toContain('A2');
    });

    it('marks the header of a fully selected column, and only that one', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        mousedown(header(fixture, 'Select column B'));
        fixture.detectChanges();

        expect(header(fixture, 'Select column B').classList).toContain('sheet-editor__head--selected');
        expect(header(fixture, 'Select column A').classList).not.toContain('sheet-editor__head--selected');
    });

    /**
     * Shift-click on a header widens the selection instead of replacing it —
     * the same rule the cells follow, so the two gestures do not disagree.
     */
    it('extends a column selection with shift-click', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        mousedown(header(fixture, 'Select column A'));
        fixture.detectChanges();
        mousedown(header(fixture, 'Select column B'), { shiftKey: true });
        fixture.detectChanges();

        const refs = highlighted(fixture);
        expect(refs.length).toBe(renderedRows(fixture) * 2);
        expect(refs.every(ref => /^[AB]\d+$/.test(ref))).toBeTrue();
    });

    it('selects the whole grid from the corner', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        mousedown(header(fixture, 'Select all cells'));
        fixture.detectChanges();

        expect(highlighted(fixture).length).toBe(renderedRows(fixture) * renderedCols(fixture));
    });

    /**
     * The drag writes a width into the DOCUMENT, which is the only thing that
     * survives the dialog. Asserting the rendered column would pass on a purely
     * visual resize that saved nothing.
     */
    it('stores a width when a column edge is dragged', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        // Measured, not assumed. This suite runs in a REAL headless Chrome, so
        // the header has real geometry and the component starts the drag from
        // whatever the column currently renders at — which for an unset width is
        // the stylesheet's default, not zero.
        const head = header(fixture, 'Select column A');
        const startPx = head.getBoundingClientRect().width;
        expect(startPx).toBeGreaterThan(0, 'the header must have real geometry to drag from');

        mousedown(head.querySelector('.sheet-editor__grip')!, { clientX: 0 });
        document.dispatchEvent(new MouseEvent('mousemove', { clientX: 70 }));
        document.dispatchEvent(new MouseEvent('mouseup'));
        fixture.detectChanges();

        saveButton(fixture).click();

        const request = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { columnWidths?: Record<string, number> }> } =
            JSON.parse(request.request.body.content);
        expect(written.sheets['S'].columnWidths?.['A']).toBe(columnWidthFromPx(startPx + 70));
        // And it genuinely GREW: the assertion above would also hold if the drag
        // delta were dropped and the width merely re-derived from the start.
        expect(written.sheets['S'].columnWidths?.['A']).toBeGreaterThan(columnWidthFromPx(startPx));
        request.flush({ contentHash: 'x' });
    });

    /**
     * The grip sits INSIDE the header whose mousedown selects the column. A
     * resize that also selected would fight its own highlight, and the author
     * would be unable to narrow a column without selecting it.
     */
    it('does not select the column when the resize grip is grabbed', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const grip: HTMLElement = header(fixture, 'Select column A')
            .querySelector('.sheet-editor__grip')!;
        mousedown(grip, { clientX: 0 });
        document.dispatchEvent(new MouseEvent('mouseup'));
        fixture.detectChanges();

        expect(highlighted(fixture)).toEqual([]);
    });

    /**
     * The listeners live on the DOCUMENT so a fast drag does not outrun a 6px
     * target. That makes tearing them down the component's problem: a dialog
     * closed mid-drag would otherwise keep resizing a sheet nobody is looking
     * at, forever.
     */
    it('stops tracking the pointer once destroyed mid-drag', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const grip: HTMLElement = header(fixture, 'Select column A')
            .querySelector('.sheet-editor__grip')!;
        mousedown(grip, { clientX: 0 });
        fixture.destroy();

        // Would throw on a destroyed component's signals if still subscribed.
        expect(() => document.dispatchEvent(new MouseEvent('mousemove', { clientX: 500 })))
            .not.toThrow();
    });

    /**
     * A plain click STARTS a new selection, even when a shift-click came before
     * it.
     *
     * `extendTo` cancels the browser's native extend-text gesture, and that
     * suppression means no `focus` fires — so the `extending` flag it sets to
     * guard against a stray focus survives until the author's NEXT plain click,
     * which then gets mistaken for that stray focus and leaves the anchor where
     * it was. The following shift-click then builds a range from a cell the
     * author left long ago.
     *
     * Found in the browser: selecting C2:C4 and inserting a checkbox put one in
     * column D as well, because the range really was C2:D4 — and inserting over
     * D replaced the dropdown that was there. Silent damage to cells nobody
     * selected.
     */
    it('starts a new selection on a plain click that follows a shift-click', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        // A range, the ordinary way.
        cellInput(fixture, 'A1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();
        shiftClick(cellInput(fixture, 'C1')!);
        fixture.detectChanges();
        expect(mergeButton(fixture).title).toContain('A1:C1');

        // A plain click elsewhere. In a real browser this is the FIRST focus
        // since the shift-click, because that one was cancelled.
        cellInput(fixture, 'B3')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        cellInput(fixture, 'B3')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        shiftClick(cellInput(fixture, 'D3')!);
        fixture.detectChanges();

        expect(mergeButton(fixture).title)
            .withContext('the range must start where the author last clicked')
            .toContain('B3:D3');
    });

    // ---- Form elements -----------------------------------------------------

    function insertSelect(fixture: ReturnType<typeof makeFixture>): HTMLSelectElement {
        const select = fixture.nativeElement.querySelector('select[aria-label="Insert form element"]');
        if (!(select instanceof HTMLSelectElement)) throw new Error('no Insert control rendered');

        return select;
    }

    function chooseInsert(fixture: ReturnType<typeof makeFixture>, value: string): void {
        const select = insertSelect(fixture);
        select.value = value;
        select.dispatchEvent(new Event('change'));
    }

    /**
     * A checkbox cell draws the CONTROL, not the word "TRUE" — that is what an
     * author drew a tick box for. The value underneath stays TRUE/FALSE, which
     * is exactly what the generated workbook carries.
     */
    it('turns the selected cells into checkboxes and writes their value', async () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Paid' } } } } });
        fixture.detectChanges();

        cellInput(fixture, 'A2')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();

        chooseInsert(fixture, 'checkbox');
        await fixture.whenStable();
        fixture.detectChanges();

        const box: HTMLInputElement = fixture.nativeElement.querySelector('input[type="checkbox"][aria-label="A2"]');
        expect(box).not.toBeNull('the cell must draw a tick box');
        expect(box.checked).toBeFalse();
        // The control REPLACES the text input rather than sitting beside it —
        // asserted on the type, because both carry the cell's aria-label and
        // `cellInput` finds either.
        expect(cellInput(fixture, 'A2')!.type).toBe('checkbox');

        box.click();
        fixture.detectChanges();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, {
            cells: Record<string, { value?: string }>;
            validations?: Record<string, { type: string }>;
        }> } = JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['A2'].value).toBe('TRUE');
        expect(written.sheets['S'].validations!['A2'].type).toBe('checkbox');
        save.flush({ contentHash: 'x' });
    });

    /**
     * A dropdown KEEPS its text input — an author may still need to type a
     * `{var:}` token into the cell — and gains an arrow offering the values.
     */
    it('offers a dropdown\'s options and writes the one picked', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                S: {
                    cells: { A1: { value: 'Status' } },
                    validations: { 'A2:A3': { type: 'list', values: ['New', 'Open', 'Closed'] } },
                },
            },
        });
        fixture.detectChanges();

        const handle: HTMLButtonElement = fixture.nativeElement.querySelector('button[aria-label="Options for A2"]');
        expect(handle).not.toBeNull();
        expect(cellInput(fixture, 'A2')).not.toBeNull('a dropdown cell can still be typed into');
        expect(fixture.nativeElement.querySelector('button[aria-label="Options for A4"]'))
            .toBeNull('outside the range, no control');

        handle.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        fixture.detectChanges();

        const options: HTMLElement[] = Array.from(fixture.nativeElement.querySelectorAll('.sheet-editor__option'));
        expect(options.map(o => o.textContent!.trim())).toEqual(['New', 'Open', 'Closed']);

        options[1].click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A2')!.value).toBe('Open');

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { cells: Record<string, { value?: string }> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['A2'].value).toBe('Open');
        save.flush({ contentHash: 'x' });
    });

    /** Removing is offered only when the focused cell actually has a rule. */
    it('offers Remove only where a rule applies, and clears it', async () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: 'x' } }, validations: { A2: { type: 'checkbox' } } } },
        });
        fixture.detectChanges();

        const values = () => Array.from(insertSelect(fixture).options).map(o => o.value);

        cellInput(fixture, 'B5')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();
        expect(values()).not.toContain('remove', 'no rule here, nothing to remove');

        const box: HTMLInputElement = fixture.nativeElement.querySelector('input[type="checkbox"][aria-label="A2"]');
        box.dispatchEvent(new Event('focus'));
        fixture.detectChanges();
        expect(values()).toContain('remove');

        chooseInsert(fixture, 'remove');
        await fixture.whenStable();
        fixture.detectChanges();

        expect(fixture.nativeElement.querySelector('input[type="checkbox"][aria-label="A2"]'))
            .toBeNull('the control is gone');
        expect(cellInput(fixture, 'A2')!.type).toBe('text', 'and the cell is an ordinary one again');
    });

    // ---- Point mode --------------------------------------------------------

    /** Type into a cell as a browser does: value, input event, caret at the end. */
    function typeFormula(fixture: ReturnType<typeof makeFixture>, ref: string, text: string): HTMLInputElement {
        const input = cellInput(fixture, ref)!;
        // REAL focus — see the note on the caret-driving helper specs. Point
        // mode turns on the difference between "this input has the focus" and
        // "somebody dispatched a focus event at it".
        input.focus();
        input.value = text;
        input.setSelectionRange(text.length, text.length);
        input.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        return input;
    }

    function outlined(fixture: ReturnType<typeof makeFixture>): string[] {
        const cells: HTMLElement[] = Array.from(
            fixture.nativeElement.querySelectorAll('td.sheet-editor__cell'),
        );

        return cells
            .filter(td => '' !== td.style.outline)
            .map(td => td.querySelector('input,textarea')?.getAttribute('aria-label') ?? '')
            .filter(label => '' !== label);
    }

    /**
     * A click, as a real browser delivers one to a cell input.
     *
     * The trailing `click` is load-bearing and was missing from the first
     * version of this helper. A real browser sends it EVEN WHEN the mousedown's
     * default was cancelled, and the cell's own `(click)` handler then fired —
     * handing the formula's editing state to the cell that was merely pointed
     * at. The specs all passed; the browser did not. A synthetic gesture has to
     * carry every event the real one does, or it tests a gesture nobody makes.
     */
    function pointClick(input: HTMLInputElement, shift = false): MouseEvent {
        const event = new MouseEvent('mousedown', { shiftKey: shift, bubbles: true, cancelable: true });
        input.dispatchEvent(event);
        input.dispatchEvent(new MouseEvent('click', { shiftKey: shift, bubbles: true }));

        return event;
    }

    it('outlines every cell a formula being typed points at', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        typeFormula(fixture, 'D1', '=SUM(A1:A3)');

        expect(outlined(fixture).sort()).toEqual(['A1', 'A2', 'A3']);
    });

    /** The outlines belong to the cell being edited, so they go with the focus. */
    it('drops the outlines when the formula stops being edited', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const input = typeFormula(fixture, 'D1', '=A1+A2');
        expect(outlined(fixture).length).toBe(2);

        input.dispatchEvent(new Event('blur'));
        fixture.detectChanges();

        expect(outlined(fixture)).toEqual([]);
    });

    it('writes a clicked cell into the formula instead of moving the cursor', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const editing = typeFormula(fixture, 'D1', '=');
        const event = pointClick(cellInput(fixture, 'B2')!);
        fixture.detectChanges();

        expect(event.defaultPrevented)
            .withContext('the click must not move focus out of the formula')
            .toBeTrue();
        expect(editing.value).toBe('=B2');
        // The outlines must SURVIVE the click. They did not: the pointed-at
        // cell's own click handler claimed the editing state, so the formula
        // stopped being the thing being edited the moment you pointed at
        // anything.
        expect(outlined(fixture)).toEqual(['B2']);
    });

    /** A second click SWAPS the reference; without that it would read `B2C3`. */
    it('replaces the reference on a second click', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const editing = typeFormula(fixture, 'D1', '=');
        pointClick(cellInput(fixture, 'B2')!);
        fixture.detectChanges();
        pointClick(cellInput(fixture, 'C3')!);
        fixture.detectChanges();

        expect(editing.value).toBe('=C3');
    });

    it('grows the reference into a range on a shift-click', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const editing = typeFormula(fixture, 'D1', '=SUM(');
        pointClick(cellInput(fixture, 'A1')!);
        fixture.detectChanges();
        pointClick(cellInput(fixture, 'A3')!, true);
        fixture.detectChanges();

        // No closing paren, because none was typed — the author is still
        // mid-formula, which is exactly when point mode is used.
        expect(editing.value).toBe('=SUM(A1:A3');
    });

    /**
     * THE case that keeps point mode safe. After a closing paren a reference
     * would be a syntax error, so the click has to keep its ordinary meaning —
     * otherwise clicking away from a finished formula would silently edit it.
     */
    it('leaves an ordinary click alone where a reference cannot go', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const editing = typeFormula(fixture, 'D1', '=SUM(A1:A3)');
        const event = pointClick(cellInput(fixture, 'B2')!);
        fixture.detectChanges();

        expect(event.defaultPrevented).withContext('navigation must still work').toBeFalse();
        expect(editing.value).toBe('=SUM(A1:A3)', 'the formula is untouched');
    });

    /** And a cell that is not a formula at all never enters point mode. */
    it('leaves an ordinary click alone when the cell holds plain text', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        typeFormula(fixture, 'D1', 'Total');
        const event = pointClick(cellInput(fixture, 'B2')!);

        expect(event.defaultPrevented).toBeFalse();
    });

    // ---- Context menu ------------------------------------------------------

    function rightClick(fixture: ReturnType<typeof makeFixture>, ref: string): void {
        const cell = cellInput(fixture, ref)!.closest('td')!;
        cell.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
        fixture.detectChanges();
    }

    function menuItem(fixture: ReturnType<typeof makeFixture>, label: string): HTMLButtonElement {
        const items: HTMLButtonElement[] = Array.from(
            fixture.nativeElement.querySelectorAll('.sheet-editor__context button'),
        );
        const match = items.find(i => (i.textContent ?? '').trim() === label);
        if (!match) {
            throw new Error(`no menu item "${label}" — have: ${items.map(i => i.textContent!.trim()).join(', ')}`);
        }

        return match;
    }

    /**
     * The menu exists because of what is under it: until now there was no way
     * to insert or delete a row or column AT ALL. A context menu that only
     * duplicated the toolbar would be decoration.
     */
    it('inserts a row above the cell, pushing the rest down', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Item' }, A2: { value: 'Widget' } } } } });
        fixture.detectChanges();

        rightClick(fixture, 'A2');
        menuItem(fixture, 'Insert row above').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')!.value).toBe('Item', 'above the line, untouched');
        expect(cellInput(fixture, 'A2')!.value).toBe('', 'the new row is blank');
        expect(cellInput(fixture, 'A3')!.value).toBe('Widget', 'pushed down');
        expect(fixture.nativeElement.querySelector('.sheet-editor__context')).toBeNull('the menu closes');
    });

    it('deletes the row under the cursor and pulls the rest up', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: 'Item' }, A2: { value: 'Widget' }, A3: { value: 'Gadget' } } } },
        });
        fixture.detectChanges();

        rightClick(fixture, 'A2');
        expect(menuItem(fixture, 'Delete row 2')).not.toBeNull('the label names what it will delete');
        menuItem(fixture, 'Delete row 2').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A2')!.value).toBe('Gadget');
        expect(cellInput(fixture, 'A3')!.value).toBe('');
    });

    it('inserts and deletes columns too', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Item' }, B1: { value: 'Qty' } } } } });
        fixture.detectChanges();

        rightClick(fixture, 'A1');
        menuItem(fixture, 'Insert column left').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')!.value).toBe('');
        expect(cellInput(fixture, 'B1')!.value).toBe('Item');

        rightClick(fixture, 'A1');
        menuItem(fixture, 'Delete column A').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')!.value).toBe('Item', 'back where it started');
    });

    /**
     * Saved through the real payload, because the point of a structural edit is
     * that it reaches the document — and the formula is the part that is
     * silently wrong when it does not: a total that sums the wrong range still
     * looks like a number.
     */
    it('rewrites a formula across an insert and saves the result', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: {
                B2: { value: '2' }, B3: { value: '5' }, B4: { formula: 'SUM(B2:B3)' },
            } } },
        });
        fixture.detectChanges();

        rightClick(fixture, 'B3');
        menuItem(fixture, 'Insert row above').click();
        fixture.detectChanges();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { cells: Record<string, { formula?: string }> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['B5'].formula).toBe('SUM(B2:B4)', 'the total still covers every line');
        save.flush({ contentHash: 'x' });
    });

    /**
     * The grid body is a scroll container, so a menu anchored BELOW a cell near
     * the bottom is clipped by it — and what gets cut is the end of the list,
     * which is how "Clear contents" through "Checkbox" became unreachable on a
     * low row. Found in the browser; no spec could see it, because nothing here
     * renders at a real size.
     *
     * The boxes are stubbed rather than laid out: karma gives the fixture a
     * real but arbitrary viewport, so asserting on genuine geometry would pin
     * the test to the runner's window instead of to the rule.
     */
    it('opens the context menu upward when it will not fit below', async () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const body: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__body');
        // A tall body whose BOTTOM sits just under the menu's top: no room
        // below, plenty above.
        body.getBoundingClientRect = () => ({ top: 0, bottom: 500, left: 0, right: 800 }) as DOMRect;

        rightClick(fixture, 'A2');
        const menu: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__context');
        menu.getBoundingClientRect = () => ({ top: 400, bottom: 740, left: 0, right: 200, height: 340 }) as DOMRect;

        await new Promise(resolve => requestAnimationFrame(() => resolve(null)));
        fixture.detectChanges();

        expect(menu.classList).toContain('sheet-editor__context--above');
    });

    /** Below is the default, and it must stay so wherever there IS room. */
    it('keeps the context menu below the cell when it fits', async () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const body: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__body');
        body.getBoundingClientRect = () => ({ top: 0, bottom: 900, left: 0, right: 800 }) as DOMRect;

        rightClick(fixture, 'A2');
        const menu: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__context');
        menu.getBoundingClientRect = () => ({ top: 60, bottom: 400, left: 0, right: 200, height: 340 }) as DOMRect;

        await new Promise(resolve => requestAnimationFrame(() => resolve(null)));
        fixture.detectChanges();

        expect(menu.classList).not.toContain('sheet-editor__context--above');
    });

    /**
     * A menu taller than the whole grid clips either way, and flipping it would
     * only move the unreachable half from the bottom to the top.
     */
    it('does not flip when there is no room above either', async () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const body: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__body');
        body.getBoundingClientRect = () => ({ top: 0, bottom: 200, left: 0, right: 800 }) as DOMRect;

        rightClick(fixture, 'A2');
        const menu: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__context');
        menu.getBoundingClientRect = () => ({ top: 40, bottom: 380, left: 0, right: 200, height: 340 }) as DOMRect;

        await new Promise(resolve => requestAnimationFrame(() => resolve(null)));
        fixture.detectChanges();

        expect(menu.classList).not.toContain('sheet-editor__context--above');
    });

    /** Contents only — the table around the cells was not what was selected. */
    it('clears contents without dismantling the layout', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: 'Item' }, A2: { value: 'Widget' } }, merges: ['A1:B1'] } },
        });
        fixture.detectChanges();

        rightClick(fixture, 'A2');
        menuItem(fixture, 'Clear contents').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A2')!.value).toBe('');
        expect(cellInput(fixture, 'A1')!.closest('td')!.getAttribute('colspan')).toBe('2', 'the merge survives');
    });

    // ---- Full screen -------------------------------------------------------

    /** As {@link saveButton} — by label, never by index. */
    function fullScreenButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const button = fixture.nativeElement.querySelector('button[aria-label="Full screen"]');
        if (!(button instanceof HTMLButtonElement)) throw new Error('no full-screen button rendered');

        return button;
    }

    function shell(fixture: ReturnType<typeof makeFixture>): HTMLElement {
        return fixture.nativeElement.querySelector('.sheet-editor');
    }

    /**
     * A grid is the surface that most wants the room — the default width shows
     * about eight columns of a sheet that has twenty-six.
     */
    it('fills the viewport when full screen is toggled, and goes back', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        expect(shell(fixture).classList).not.toContain('sheet-editor--full');
        expect(fullScreenButton(fixture).getAttribute('aria-pressed')).toBe('false');

        fullScreenButton(fixture).click();
        fixture.detectChanges();

        expect(shell(fixture).classList).toContain('sheet-editor--full');
        expect(fullScreenButton(fixture).getAttribute('aria-pressed')).toBe('true');

        fullScreenButton(fixture).click();
        fixture.detectChanges();

        expect(shell(fixture).classList).not.toContain('sheet-editor--full');
    });

    /**
     * F11 does it too, and the browser's own full screen is SUPPRESSED — its
     * version would leave the dialog exactly the size it was, with more black
     * around it.
     */
    it('toggles on F11 and cancels the browser default', () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const event = new KeyboardEvent('keydown', { key: 'F11', bubbles: true, cancelable: true });
        fixture.nativeElement.querySelector('.sheet-editor').dispatchEvent(event);
        fixture.detectChanges();

        expect(shell(fixture).classList).toContain('sheet-editor--full');
        expect(event.defaultPrevented)
            .withContext('the browser must not also go full screen')
            .toBeTrue();
    });

    /**
     * The one that can break silently. Both virtualisation windows are computed
     * from a MEASURED viewport, taken on load and on scroll and nowhere else.
     * Resizing without re-measuring leaves the grid rendering the old, smaller
     * window — blank rows below, missing columns right — until the author
     * happens to scroll, which on a sheet that now fits entirely may be never.
     */
    it('re-measures the viewport when the size changes, so the window grows with it', async () => {
        const fixture = makeFixture();
        respondWith(SMALL);
        fixture.detectChanges();

        const body: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__body');
        const rowsIn = () => fixture.nativeElement.querySelectorAll('.sheet-editor__row-head').length;

        // The jsdom-less karma DOM gives the body a real but small box; force a
        // tall one so the re-measure has something different to find.
        Object.defineProperty(body, 'clientHeight', { value: 4000, configurable: true });
        Object.defineProperty(body, 'clientWidth', { value: 3000, configurable: true });

        const before = rowsIn();

        fullScreenButton(fixture).click();
        await new Promise(resolve => requestAnimationFrame(() => resolve(null)));
        fixture.detectChanges();

        expect(rowsIn())
            .withContext('the row window must grow with the viewport, not wait for a scroll')
            .toBeGreaterThan(before);
    });

    // ---- Filtering ---------------------------------------------------------

    /** As {@link saveButton} — by label, never by index. */
    function filterButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const buttons: HTMLButtonElement[] = Array.from(fixture.nativeElement.querySelectorAll('button'));
        const filter = buttons.find(b => ['Filter', 'Remove filter'].includes((b.textContent ?? '').trim()));
        if (!filter) throw new Error('no Filter button rendered');

        return filter;
    }

    /** The dropdown handle on a filter's header cell, if that column has one. */
    function filterHandle(fixture: ReturnType<typeof makeFixture>, column: string): HTMLButtonElement | null {
        return fixture.nativeElement.querySelector(`button[aria-label="Filter column ${column}"]`);
    }

    /** One value's checkbox inside the open dropdown, found by its LABEL. */
    function filterOption(fixture: ReturnType<typeof makeFixture>, label: string): HTMLInputElement {
        const options: HTMLElement[] = Array.from(
            fixture.nativeElement.querySelectorAll('.sheet-editor__filter-option'),
        );
        const match = options.find(o => (o.textContent ?? '').trim() === label);
        if (!match) throw new Error(`no filter option labelled "${label}"`);

        return match.querySelector('input')!;
    }

    const TABLE = {
        version: 1,
        sheets: {
            S: {
                cells: {
                    A1: { value: 'Item' }, B1: { value: 'Qty' },
                    A2: { value: 'Widget' }, B2: { value: '2' },
                    A3: { value: 'Gadget' }, B3: { value: '5' },
                },
            },
        },
    };

    /**
     * The renderer has grown an `<autoFilter>` over a `{loop:}` band since this
     * arc; until now the native format could not declare one at all.
     *
     * Asserted on the RENDERED header as well as the saved document, for the
     * reason the merge case gives: a filter that stored correctly but drew no
     * dropdowns looks to the author like it never took.
     */
    it('declares a filter over the selection and puts a handle on its header row only', () => {
        const fixture = makeFixture();
        respondWith(TABLE);
        fixture.detectChanges();

        expect(filterHandle(fixture, 'A')).toBeNull('no filter, no handles');

        cellInput(fixture, 'A1')!.dispatchEvent(new Event('focus'));
        fixture.detectChanges();
        shiftClick(cellInput(fixture, 'B3')!);
        fixture.detectChanges();

        filterButton(fixture).click();
        fixture.detectChanges();

        expect(filterHandle(fixture, 'A')).not.toBeNull();
        expect(filterHandle(fixture, 'B')).not.toBeNull();
        // Two handles, not six: the body rows are data, not headers.
        expect(fixture.nativeElement.querySelectorAll('button[aria-label^="Filter column"]').length).toBe(2);

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { autoFilter?: string }> } = JSON.parse(save.request.body.content);
        expect(written.sheets['S'].autoFilter).toBe('A1:B3');
        save.flush({ contentHash: 'x' });
    });

    /**
     * Unchecking a value HIDES its rows, and the row leaves the DOM rather than
     * being hidden with CSS — a hidden input inside a filtered-out row would
     * still be focusable and editable, which is the same trap the merge case
     * names one feature over.
     *
     * The choice is NOT saved. A template holds tokens, not data, so which
     * values an author hid while editing says nothing about the document that
     * gets generated; writing it into the `.dsheet` would ship one person's
     * temporary view to every render.
     */
    it('hides the rows whose value was unchecked, and stores only the range', () => {
        const fixture = makeFixture();
        respondWith({ ...TABLE, sheets: { S: { ...TABLE.sheets.S, autoFilter: 'A1:B3' } } });
        fixture.detectChanges();

        filterHandle(fixture, 'A')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        fixture.detectChanges();

        filterOption(fixture, 'Widget').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A2')).toBeNull('the Widget row is filtered out');
        expect(cellInput(fixture, 'A3')).not.toBeNull('the Gadget row stays');
        expect(cellInput(fixture, 'A1')).not.toBeNull('the header is never filtered');

        // Nothing about the DOCUMENT changed, so there is nothing to save —
        // which is the sharpest statement that the exclusion is view state.
        expect(saveButton(fixture).disabled)
            .withContext('hiding a value must not mark the document dirty')
            .toBeTrue();

        // A real edit to a still-visible row, so a save DOES happen and the
        // payload can be inspected for what the filter did or did not leave in.
        const gadget = cellInput(fixture, 'A3')!;
        gadget.value = 'Gizmo';
        gadget.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { autoFilter?: string; cells: Record<string, unknown> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].autoFilter).toBe('A1:B3', 'the declaration is stored');
        expect(written.sheets['S'].cells['A2'])
            .toEqual({ value: 'Widget' }, 'a hidden row is still IN the document');
        expect(Object.keys(written.sheets['S']).sort())
            .toEqual(['autoFilter', 'cells'], 'no view state leaked into the .dsheet');
        save.flush({ contentHash: 'x' });
    });

    /**
     * Removing the filter must bring hidden rows back. A row left hidden by a
     * filter that no longer exists is unreachable — there would be no dropdown
     * left to restore it, and the author would see a document missing rows it
     * still contains.
     */
    it('restores every hidden row when the filter is removed', () => {
        const fixture = makeFixture();
        respondWith({ ...TABLE, sheets: { S: { ...TABLE.sheets.S, autoFilter: 'A1:B3' } } });
        fixture.detectChanges();

        filterHandle(fixture, 'A')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        fixture.detectChanges();
        filterOption(fixture, 'Widget').click();
        fixture.detectChanges();
        expect(cellInput(fixture, 'A2')).toBeNull();

        expect(filterButton(fixture).textContent!.trim()).toBe('Remove filter', 'the button offers the way out');
        filterButton(fixture).click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A2')).not.toBeNull();
        expect(filterHandle(fixture, 'A')).toBeNull('and the handles go with it');

        saveButton(fixture).click();

        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { autoFilter?: string }> } = JSON.parse(save.request.body.content);
        expect(written.sheets['S'].autoFilter).toBeUndefined();
        save.flush({ contentHash: 'x' });
    });

    /**
     * A filter over a `{loop:}` band has no body to list — the rows do not
     * exist until the backend expands it. The dropdown must say so rather than
     * render an empty checkbox list that reads as "this column has no values".
     */
    it('explains itself over a loop band instead of listing nothing', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                S: {
                    cells: { A1: { value: 'Item' }, A2: { value: '{loop:lines:l}{var:l.name}{endloop}' } },
                    autoFilter: 'A1:A2',
                },
            },
        });
        fixture.detectChanges();

        filterHandle(fixture, 'A')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        fixture.detectChanges();

        // One body row exists, so it IS listed -- what must not happen is the
        // list being empty while the band is there.
        expect(fixture.nativeElement.querySelectorAll('.sheet-editor__filter-option').length).toBe(1);
    });

    // ---- Borders ----------------------------------------------------------

    const SHEET = { version: 1, sheets: { S: { cells: { B2: { value: 'x' } } } } };

    /** The toolbar toggle, by label. */
    function bordersButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const button = fixture.nativeElement.querySelector('button[aria-label="Borders"]');
        if (!(button instanceof HTMLButtonElement)) throw new Error('no Borders button rendered');

        return button;
    }

    /** One gesture in the open palette, by label. */
    function borderGesture(fixture: ReturnType<typeof makeFixture>, label: string): HTMLButtonElement {
        const button = fixture.nativeElement.querySelector(`button[aria-label="${label}"]`);
        if (!(button instanceof HTMLButtonElement)) throw new Error(`no ${label} gesture rendered`);

        return button;
    }

    function cellBox(fixture: ReturnType<typeof makeFixture>, ref: string): HTMLTableCellElement {
        const td = cellInput(fixture, ref)?.closest('td');
        if (!(td instanceof HTMLTableCellElement)) throw new Error(`no cell rendered for ${ref}`);

        return td;
    }

    /**
     * Focus a cell and open the palette over it.
     *
     * The button is FOCUSED before it is clicked, because that is what a mouse
     * does and `click()` alone does not: a synthetic click leaves the cell
     * input holding the focus, so a test written without this would never see
     * the cell lose it — which is the whole thing being tested one case down.
     */
    function openBorders(fixture: ReturnType<typeof makeFixture>, ref: string): void {
        cellInput(fixture, ref)!.focus();
        fixture.detectChanges();
        bordersButton(fixture).focus();
        bordersButton(fixture).click();
        fixture.detectChanges();
    }

    /**
     * The grid is `border-collapse: collapse`, so the line between two cells is
     * ONE line and the browser resolves who draws it: at equal width and style,
     * the cell further up or further left. Every cell carries a grey gridline,
     * so declaring black on B2's top alone left the grey bottom of B1 winning —
     * "All borders" on a single cell showed its right and bottom ONLY.
     *
     * The assertion that matters is therefore on the NEIGHBOURS. Asserting B2's
     * own four edges would have passed before the fix: they were all in the
     * inline style, and the browser threw two of them away.
     */
    it('rules a cell on both sides of every line, or half of them never appear', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openBorders(fixture, 'B2');
        borderGesture(fixture, 'All borders').click();
        fixture.detectChanges();

        expect(cellBox(fixture, 'B1').style.borderBottomStyle)
            .withContext('the cell above shares B2 top')
            .toBe('solid');
        expect(cellBox(fixture, 'A2').style.borderRightStyle)
            .withContext('the cell beside shares B2 left')
            .toBe('solid');
        expect(cellBox(fixture, 'B2').style.borderTopStyle).toBe('solid');
        expect(cellBox(fixture, 'B2').style.borderLeftStyle).toBe('solid');
    });

    /**
     * Reaching for a toolbar control takes focus out of the cell. The input's
     * own `:focus` outline was the ONLY mark a single-cell selection had, so
     * the cell read as deselecting itself the moment the author went to rule
     * it — and the gesture then looked like it had nothing to act on.
     */
    it('keeps the selected cell marked once focus moves to the toolbar', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openBorders(fixture, 'B2');

        expect(document.activeElement).not.toBe(cellInput(fixture, 'B2'), 'the button took the focus');
        expect(cellBox(fixture, 'B2').classList).toContain('sheet-editor__cell--active');
    });

    /**
     * A palette that resets to its placeholder cannot say what the cells carry,
     * which reads as the last gesture having failed.
     */
    it('shows which gesture the selection already matches', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openBorders(fixture, 'B2');
        expect(borderGesture(fixture, 'All borders').getAttribute('aria-pressed')).toBe('false');

        borderGesture(fixture, 'All borders').click();
        fixture.detectChanges();
        expect(borderGesture(fixture, 'All borders').getAttribute('aria-pressed')).toBe('true');

        borderGesture(fixture, 'Clear borders').click();
        fixture.detectChanges();
        expect(borderGesture(fixture, 'All borders').getAttribute('aria-pressed')).toBe('false');
    });

    it('draws the line the palette is set to', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openBorders(fixture, 'B2');
        const line: HTMLSelectElement = fixture.nativeElement.querySelector('select[aria-label="Border line"]');
        line.value = 'dashed';
        line.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        borderGesture(fixture, 'All borders').click();
        fixture.detectChanges();

        expect(cellBox(fixture, 'B2').style.borderTopStyle).toBe('dashed');
    });

    // ---- The function catalogue -------------------------------------------

    /** The toolbar toggle, by label. */
    function functionsButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const button = fixture.nativeElement.querySelector('button[aria-label="Functions"]');
        if (!(button instanceof HTMLButtonElement)) throw new Error('no Functions button rendered');

        return button;
    }

    /** One entry in the open list, by the name it shows. */
    function functionEntry(fixture: ReturnType<typeof makeFixture>, name: string): HTMLButtonElement | null {
        const entries: HTMLButtonElement[] = Array.from(
            fixture.nativeElement.querySelectorAll('.sheet-editor__function'));

        return entries.find(e => e.querySelector('.sheet-editor__function-name')?.textContent?.trim() === name)
            ?? null;
    }

    /** Focus a cell and open the catalogue over it — see {@link openBorders}. */
    function openFunctions(fixture: ReturnType<typeof makeFixture>, ref: string): void {
        cellInput(fixture, ref)!.focus();
        fixture.detectChanges();
        functionsButton(fixture).focus();
        functionsButton(fixture).click();
        fixture.detectChanges();
    }

    /**
     * Type-ahead only serves an author who already knows the name. Nobody
     * discovers SUMIF by typing SUMIF, which is what a browsable list is for.
     */
    it('lists the catalogue on shelves, whole', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openFunctions(fixture, 'B2');

        const shelves: HTMLElement[] = Array.from(
            fixture.nativeElement.querySelectorAll('.sheet-editor__function-shelf'));
        expect(shelves.map(s => s.textContent!.trim())).toContain('Lookup');
        expect(functionEntry(fixture, 'SUMIF')).not.toBeNull();
        expect(functionEntry(fixture, 'VLOOKUP')).not.toBeNull();
    });

    /**
     * Searching the SUMMARY and not only the name is the half that makes the
     * list browsable: an author who wants a condition does not know that the
     * thing they want is spelled IF.
     */
    it('finds a function by what it does, not only by its name', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openFunctions(fixture, 'B2');
        const search: HTMLInputElement = fixture.nativeElement.querySelector('input[aria-label="Search functions"]');
        search.value = 'condition';
        search.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        expect(functionEntry(fixture, 'SUMIF')).withContext('summary says "meet a condition"').not.toBeNull();
        expect(functionEntry(fixture, 'VLOOKUP')).withContext('nothing to do with conditions').toBeNull();
    });

    it('says so when nothing matches, rather than showing an empty list', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openFunctions(fixture, 'B2');
        const search: HTMLInputElement = fixture.nativeElement.querySelector('input[aria-label="Search functions"]');
        search.value = 'zzzz';
        search.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        expect(fixture.nativeElement.querySelector('.sheet-editor__function-empty')).not.toBeNull();
    });

    it('starts a formula in an empty cell, with the caret between the brackets', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openFunctions(fixture, 'C3');
        functionEntry(fixture, 'SUM')!.click();
        fixture.detectChanges();

        const input = cellInput(fixture, 'C3')!;
        expect(input.value).toBe('=SUM(');
        expect(input.selectionStart).toBe('=SUM('.length);
        expect(document.activeElement).withContext('the caret is where the argument goes').toBe(input);
    });

    /**
     * A function dropped into a half-written formula goes AT THE CARET. Anything
     * else throws away what has been typed so far.
     */
    it('inserts into a formula already being written', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const input = cellInput(fixture, 'C3')!;
        input.focus();
        input.value = '=1+';
        input.setSelectionRange(3, 3);
        input.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        functionsButton(fixture).focus();
        functionsButton(fixture).click();
        fixture.detectChanges();
        functionEntry(fixture, 'MAX')!.click();
        fixture.detectChanges();

        expect(input.value).toBe('=1+MAX(');
    });

    /**
     * A cell holding a literal is REPLACED: "42" with SUM appended would be
     * `42SUM(`, which is not a formula and not what anyone meant.
     */
    it('replaces a literal rather than appending a function to it', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openFunctions(fixture, 'B2');
        expect(cellInput(fixture, 'B2')!.value).toBe('x', 'the cell holds a literal');
        functionEntry(fixture, 'SUM')!.click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'B2')!.value).toBe('=SUM(');
    });

    it('closes on Escape without closing the editor', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        openFunctions(fixture, 'B2');
        const menu: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__function-menu');
        const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
        menu.dispatchEvent(escape);
        fixture.detectChanges();

        expect(fixture.nativeElement.querySelector('.sheet-editor__function-menu')).toBeNull();
        expect(escape.defaultPrevented)
            .withContext('or the CDK dialog takes it and the whole editor closes')
            .toBeTrue();
    });

    // ---- Number formats ----------------------------------------------------

    /**
     * A `.xlsx` stores a date as the SERIAL its format describes, and the
     * importer keeps it that way on purpose — converting would discard the
     * format and make the value unarithmetic. So until the grid learned to
     * render a format, an imported invoice showed `46255` where the generated
     * document showed `21/08/2026`: the editor was not showing the document.
     */
    it('shows an imported date as a date, not as its serial number', () => {
        const serial = String(dateToSerial(2026, 8, 21));
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: serial, numberFormat: 'dd/mm/yyyy' } } } },
        });
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')!.value).toBe('21/08/2026');
    });

    it('shows a number the way its format will print it', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: '1234.5', numberFormat: '#,##0.00' } } } },
        });
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')!.value).toBe('1,234.50');
    });

    /**
     * The cell being EDITED shows what it stores — except a date, which shows
     * as a date because nobody wants to type a serial. Excel's formula bar
     * draws exactly this distinction.
     */
    it('edits a number as it is stored and a date as a date', () => {
        const serial = String(dateToSerial(2026, 8, 21));
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                S: {
                    cells: {
                        A1: { value: serial, numberFormat: 'dd/mm/yyyy' },
                        B1: { value: '1234.5', numberFormat: '#,##0.00' },
                    },
                },
            },
        });
        fixture.detectChanges();

        cellInput(fixture, 'A1')!.focus();
        fixture.detectChanges();
        expect(cellInput(fixture, 'A1')!.value).withContext('a date edits as a date').toBe('21/08/2026');

        cellInput(fixture, 'B1')!.focus();
        fixture.detectChanges();
        expect(cellInput(fixture, 'B1')!.value).withContext('a number edits as it is stored').toBe('1234.5');
    });

    /** The other half of the round trip: what is typed goes back as a serial. */
    it('stores a typed date as the serial the workbook wants', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: '', numberFormat: 'dd/mm/yyyy' } } } },
        });
        fixture.detectChanges();

        const input = cellInput(fixture, 'A1')!;
        input.focus();
        input.value = '21/08/2026';
        input.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        saveButton(fixture).click();
        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { cells: Record<string, { value?: string }> }> } =
            JSON.parse(save.request.body.content);
        expect(written.sheets['S'].cells['A1'].value).toBe(String(dateToSerial(2026, 8, 21)));
        save.flush({ contentHash: 'x' });
    });

    /** A computed value wears the cell's format too, as it will in the document. */
    it('formats what a formula computes', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                S: {
                    cells: {
                        A1: { value: '1200' },
                        A2: { value: '34.5' },
                        A3: { formula: 'SUM(A1:A2)', numberFormat: '#,##0.00' },
                    },
                },
            },
        });
        fixture.detectChanges();

        expect(cellInput(fixture, 'A3')!.value).toBe('1,234.50');
    });

    // ---- Committing, and moving on -----------------------------------------

    /** A key as the grid receives one, cancellable so the assertion can see it. */
    function press(el: Element, key: string, modifiers: Partial<KeyboardEventInit> = {}): KeyboardEvent {
        const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...modifiers });
        el.dispatchEvent(event);

        return event;
    }

    /** What the save request would carry for the sheet itself. */
    function savedSheet(fixture: ReturnType<typeof makeFixture>): { freeze?: string; conditionals?: unknown } {
        saveButton(fixture).click();
        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { freeze?: string; conditionals?: unknown }> } =
            JSON.parse(save.request.body.content);
        save.flush({ contentHash: 'x' });

        return written.sheets['S'];
    }

    /** What the save request would carry for one cell. */
    function savedCell(fixture: ReturnType<typeof makeFixture>, ref: string): { value?: string; formula?: string } | undefined {
        saveButton(fixture).click();
        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { cells: Record<string, { value?: string; formula?: string }> }> } =
            JSON.parse(save.request.body.content);
        save.flush({ contentHash: 'x' });

        return written.sheets['S'].cells[ref];
    }

    /**
     * ⚠️ The defect this was written for. A formula built ENTIRELY by clicking
     * cells is a value the browser does not consider the user to have edited,
     * so no `change` event ever fires for it — not on Enter, not on clicking
     * away. Measured in a real browser: `input`, then `blur`, and no `change`
     * at all. The formula was silently thrown away.
     */
    it('commits a formula built only by clicking cells', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b6 = cellInput(fixture, 'B6')!;
        b6.focus();
        b6.value = '=SUM(';
        b6.setSelectionRange(5, 5);
        b6.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        // Point mode: click, then shift-click, and NOTHING is typed.
        pointClick(cellInput(fixture, 'B2')!);
        shiftClick(cellInput(fixture, 'B4')!);
        fixture.detectChanges();
        expect(b6.value).toBe('=SUM(B2:B4');

        press(b6, 'Enter');
        fixture.detectChanges();

        expect(savedCell(fixture, 'B6')).toEqual({ formula: 'SUM(B2:B4' });
    });

    it('moves down on Enter and right on Tab, as every grid does', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        b2.value = '42';
        b2.dispatchEvent(new Event('input'));
        press(b2, 'Enter');
        fixture.detectChanges();

        expect(activeRef(fixture)).withContext('Enter goes down').toBe('B3');

        const b3 = cellInput(fixture, 'B3')!;
        b3.focus();
        press(b3, 'Tab');
        fixture.detectChanges();
        expect(activeRef(fixture)).withContext('Tab goes right').toBe('C3');

        press(cellInput(fixture, 'C3')!, 'Enter', { shiftKey: true });
        fixture.detectChanges();
        expect(activeRef(fixture)).withContext('Shift+Enter goes back up').toBe('C2');
    });

    it('stays put at the edges rather than walking off the grid', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const a1 = cellInput(fixture, 'A1')!;
        a1.focus();
        press(a1, 'Enter', { shiftKey: true });
        fixture.detectChanges();

        expect(activeRef(fixture)).toBe('A1');
    });

    /**
     * Escape reaches the CDK dialog otherwise and the whole editor closes,
     * taking every unsaved cell with it.
     */
    it('puts back what was there on Escape, without closing the editor', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        b2.value = 'typed over it';
        b2.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        const event = press(b2, 'Escape');
        fixture.detectChanges();

        expect(b2.value).toBe('x', 'reverted to what the cell holds');
        expect(event.defaultPrevented).withContext('and never reaches the dialog').toBeTrue();
    });

    it('lets Escape through when there is nothing to revert', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        fixture.detectChanges();

        expect(press(b2, 'Escape').defaultPrevented)
            .withContext('closing the editor is what Escape means when idle')
            .toBeFalse();
    });

    /**
     * Blur fires for every cell an author merely passes through. Rewriting each
     * one would mark a document dirty for having been looked at.
     */
    it('does not dirty a document just for visiting a cell', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        b2.blur();
        fixture.detectChanges();

        expect(fixture.nativeElement.querySelector('.sheet-editor__dirty')).toBeNull();
    });

    // ---- Dragging a selection ----------------------------------------------

    /** A drag: press in one cell, move over another, release. */
    function dragOver(fixture: ReturnType<typeof makeFixture>, from: string, to: string): void {
        cellInput(fixture, from)!.dispatchEvent(
            new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
        cellInput(fixture, from)!.focus();
        fixture.detectChanges();

        document.dispatchEvent(new MouseEvent('mousemove', {
            bubbles: true, buttons: 1,
        }));
        // The listener reads the cell under the pointer from the event target,
        // which a dispatched event only carries when it is dispatched ON it.
        cellInput(fixture, to)!.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, buttons: 1 }));
        fixture.detectChanges();
    }

    it('selects a range by dragging with the button held', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        dragOver(fixture, 'B2', 'C4');

        expect(activeRef(fixture)).toBe('C4');
        expect(inSelection(fixture, 'B3')).withContext('B2:C4 covers it').toBeTrue();
        expect(inSelection(fixture, 'D5')).withContext('outside the range').toBeFalse();
    });

    /**
     * ⚠️ Dragging mid-formula grows the REFERENCE. Found only by dragging in a
     * real browser: clearing the document's text selection to stop the drag
     * painting across cells also wiped the CARET of the focused formula input,
     * so the span was read as zero and the range landed in front of the
     * formula -- `B3:B4=SUM(B3`.
     */
    it('grows a formula reference when the drag is mid-formula', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const e8 = cellInput(fixture, 'E8')!;
        e8.focus();
        e8.value = '=SUM(';
        e8.setSelectionRange(5, 5);
        e8.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        pointClick(cellInput(fixture, 'B2')!);
        fixture.detectChanges();
        expect(e8.value).toBe('=SUM(B2');

        cellInput(fixture, 'B4')!.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, buttons: 1 }));
        fixture.detectChanges();

        expect(e8.value).toBe('=SUM(B2:B4');
    });

    /** A drag that never leaves its cell is the browser selecting TEXT. */
    it('leaves a drag inside one cell to the text it is over', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        dragOver(fixture, 'B2', 'B2');

        expect(fixture.nativeElement.querySelector('.sheet-editor__grid--dragging')).toBeNull();
    });

    it('stops extending once the button comes up', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        dragOver(fixture, 'B2', 'C4');
        document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        fixture.detectChanges();

        cellInput(fixture, 'D5')!.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, buttons: 1 }));
        fixture.detectChanges();

        expect(activeRef(fixture)).withContext('the released drag no longer follows').toBe('C4');
    });

    // ---- Moving about, and clearing --------------------------------------

    /** A clipboard event carrying text, as a browser delivers one. */
    function clipboardEvent(type: 'copy' | 'cut' | 'paste', text = ''): ClipboardEvent {
        const data = new DataTransfer();
        if ('' !== text) data.setData('text/plain', text);

        return new ClipboardEvent(type, { clipboardData: data, bubbles: true, cancelable: true });
    }

    function grid(fixture: ReturnType<typeof makeFixture>): HTMLElement {
        return fixture.nativeElement.querySelector('.sheet-editor__body');
    }

    /**
     * A spreadsheet has two states and this grid has one: every cell is an
     * input, always showing a caret. So the arrows belong to the GRID until
     * something is typed — otherwise they either never move between cells, or
     * always do, and neither is a spreadsheet.
     */
    it('moves the selection with the arrows while a cell is only selected', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        fixture.detectChanges();

        press(cellInput(fixture, 'B2')!, 'ArrowDown');
        fixture.detectChanges();
        expect(activeRef(fixture)).toBe('B3');

        press(cellInput(fixture, 'B3')!, 'ArrowRight');
        fixture.detectChanges();
        expect(activeRef(fixture)).toBe('C3');
    });

    it('extends the selection when Shift is held', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        fixture.detectChanges();
        press(cellInput(fixture, 'B2')!, 'ArrowDown', { shiftKey: true });
        fixture.detectChanges();

        expect(activeRef(fixture)).toBe('B3');
        expect(inSelection(fixture, 'B2')).withContext('the anchor stayed').toBeTrue();
    });

    /**
     * ⚠️ A plain arrow COLLAPSES a selection. The extend flag survives a
     * shift-click -- the prevented mousedown fires no focus to consume it -- so
     * the arrow inherited it, left the anchor behind, and went on extending a
     * range the author had just walked away from.
     */
    it('collapses a shift-clicked selection on a plain arrow', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        shiftClick(cellInput(fixture, 'B4')!);
        fixture.detectChanges();
        expect(inSelection(fixture, 'B3')).withContext('a range is selected').toBeTrue();

        press(cellInput(fixture, 'B4')!, 'ArrowDown');
        fixture.detectChanges();

        expect(activeRef(fixture)).toBe('B5');
        expect(inSelection(fixture, 'B3')).withContext('and the range is gone').toBeFalse();
    });

    /** Once there is something to edit, the arrows move the CARET. */
    it('leaves the arrows to the text once a cell is being edited', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        b2.value = '=SUM(A1';
        b2.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        const event = press(b2, 'ArrowLeft');
        fixture.detectChanges();

        expect(activeRef(fixture)).withContext('still on the cell being typed into').toBe('B2');
        expect(event.defaultPrevented).withContext('and the caret move was not stolen').toBeFalse();
    });

    /** F2 is how an author says so without typing anything. */
    it('hands the arrows to the text on F2', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        fixture.detectChanges();
        press(b2, 'F2');
        press(b2, 'ArrowDown');
        fixture.detectChanges();

        expect(activeRef(fixture)).toBe('B2');
    });

    it('clears the selected cells with Delete', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        fixture.detectChanges();
        press(cellInput(fixture, 'B2')!, 'Delete');
        fixture.detectChanges();

        expect(cellInput(fixture, 'B2')!.value).toBe('');
        expect(savedCell(fixture, 'B2')).toBeUndefined();
    });

    // ---- The clipboard ------------------------------------------------------

    it('copies the selection as the tab-separated text a spreadsheet expects', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: 'Item' }, B1: { value: 'Qty' }, A2: { value: 'Bolt' }, B2: { value: '4' } } } },
        });
        fixture.detectChanges();

        cellInput(fixture, 'A1')!.focus();
        shiftClick(cellInput(fixture, 'B2')!);
        fixture.detectChanges();

        const event = clipboardEvent('copy');
        grid(fixture).dispatchEvent(event);

        expect(event.clipboardData!.getData('text/plain')).toBe('Item\tQty\nBolt\t4');
        expect(event.defaultPrevented).toBeTrue();
    });

    /**
     * The author highlighted characters and means to copy those. A grid that
     * overrode that would make it impossible to copy half a formula.
     */
    it('leaves a copy alone when there is a text selection inside the cell', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        b2.setSelectionRange(0, 1);
        fixture.detectChanges();

        const event = clipboardEvent('copy');
        grid(fixture).dispatchEvent(event);

        expect(event.defaultPrevented).toBeFalse();
    });

    it('pastes a block and selects what arrived', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        cellInput(fixture, 'C3')!.focus();
        fixture.detectChanges();

        grid(fixture).dispatchEvent(clipboardEvent('paste', 'a\tb\nc\td'));
        fixture.detectChanges();

        expect(cellInput(fixture, 'C3')!.value).toBe('a');
        expect(cellInput(fixture, 'D4')!.value).toBe('d');
        expect(activeRef(fixture)).withContext('the far corner of what landed').toBe('D4');
    });

    /**
     * Copying a total down a column is the commonest thing anybody does in a
     * spreadsheet, and it only works if the references move with it.
     */
    it('moves the references in a formula pasted from this grid', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: '2' }, A2: { value: '3' }, A3: { formula: 'SUM(A1:A2)' } } } },
        });
        fixture.detectChanges();

        cellInput(fixture, 'A3')!.focus();
        fixture.detectChanges();
        const copy = clipboardEvent('copy');
        grid(fixture).dispatchEvent(copy);
        const text = copy.clipboardData!.getData('text/plain');

        cellInput(fixture, 'B3')!.focus();
        fixture.detectChanges();
        grid(fixture).dispatchEvent(clipboardEvent('paste', text));
        fixture.detectChanges();

        expect(savedCell(fixture, 'B3')).toEqual({ formula: 'SUM(B1:B2)' });
    });

    /**
     * A single value with no tabs or newlines belongs to the input: pasting a
     * word into a half-typed formula must put it at the caret.
     */
    it('leaves a plain value to the cell being edited', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        const b2 = cellInput(fixture, 'B2')!;
        b2.focus();
        b2.value = '=SUM(';
        b2.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        const event = clipboardEvent('paste', 'A1:A3');
        grid(fixture).dispatchEvent(event);
        fixture.detectChanges();

        expect(event.defaultPrevented).toBeFalse();
    });

    // ---- Freezing panes ----------------------------------------------------

    function freezeButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const buttons: HTMLButtonElement[] = Array.from(fixture.nativeElement.querySelectorAll('button'));
        const freeze = buttons.find(b => /^(Freeze|Unfreeze)$/.test((b.textContent ?? '').trim()));
        if (!freeze) throw new Error('no Freeze button rendered');

        return freeze;
    }

    it('freezes above and left of the selected cell, and lets go again', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        fixture.detectChanges();
        expect(freezeButton(fixture).textContent!.trim()).toBe('Freeze');

        freezeButton(fixture).click();
        fixture.detectChanges();

        expect(freezeButton(fixture).textContent!.trim()).toBe('Unfreeze');
        expect(savedSheet(fixture).freeze).toBe('B2');
    });

    /**
     * ⚠️ A frozen row has to be RENDERED to be sticky, and virtualisation would
     * otherwise drop it into the top spacer the moment the window moved past
     * it. That is the whole interaction between these two features.
     */
    it('keeps a frozen row on the page however far the grid is scrolled', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Header' } }, freeze: 'A2' } } });
        fixture.detectChanges();

        const body: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__body');
        body.scrollTop = 4000;
        body.dispatchEvent(new Event('scroll'));
        fixture.detectChanges();

        const frozen = cellInput(fixture, 'A1');
        expect(frozen).withContext('row 1 is still drawn').not.toBeNull();
        expect(frozen!.closest('td')!.classList).toContain('sheet-editor__cell--frozen-row');
    });

    it('draws a frozen row exactly once', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Header' } }, freeze: 'A2' } } });
        fixture.detectChanges();

        expect(fixture.nativeElement.querySelectorAll('input[aria-label="A1"]').length)
            .withContext('the window must not repeat what is frozen')
            .toBe(1);
    });

    /**
     * ⚠️ A merge reaches LEFTWARD when the window is widened to include its
     * anchor, and that widening would otherwise pull the window back over the
     * frozen block. A title merged across A1:D1 -- the commonest thing at the
     * top of a sheet -- rendered column A twice, two inputs claiming one cell.
     */
    it('does not redraw a frozen column that a merge reaches into', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                S: {
                    cells: { A1: { value: 'Quarterly orders' }, A2: { value: 'Item' }, B2: { value: 'Qty' } },
                    merges: ['A1:D1'],
                    freeze: 'B3',
                },
            },
        });
        fixture.detectChanges();

        expect(fixture.nativeElement.querySelectorAll('input[aria-label="A2"]').length).toBe(1);
        expect(fixture.nativeElement.querySelectorAll('input[aria-label="A1"]').length).toBe(1);
    });

    it('pins a frozen column and leaves the rest to scroll', () => {
        const fixture = makeFixture();
        respondWith({ version: 1, sheets: { S: { cells: { A1: { value: 'Item' }, B1: { value: 'Qty' } }, freeze: 'B1' } } });
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')!.closest('td')!.classList).toContain('sheet-editor__cell--frozen-col');
        expect(cellInput(fixture, 'B1')!.closest('td')!.classList).not.toContain('sheet-editor__cell--frozen-col');
    });

    // ---- Undo and redo -----------------------------------------------------

    function historyButton(fixture: ReturnType<typeof makeFixture>, label: 'Undo' | 'Redo'): HTMLButtonElement {
        const button = fixture.nativeElement.querySelector(`button[aria-label="${label}"]`);
        if (!(button instanceof HTMLButtonElement)) throw new Error(`no ${label} button rendered`);

        return button;
    }

    function type(fixture: ReturnType<typeof makeFixture>, ref: string, text: string): void {
        const input = cellInput(fixture, ref)!;
        input.focus();
        input.value = text;
        input.dispatchEvent(new Event('input'));
        input.dispatchEvent(new Event('change'));
        fixture.detectChanges();
    }

    it('takes back the last change, and puts it back again', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        expect(historyButton(fixture, 'Undo').disabled).withContext('nothing to undo yet').toBeTrue();

        type(fixture, 'B2', 'changed');
        expect(historyButton(fixture, 'Undo').disabled).toBeFalse();

        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();
        expect(cellInput(fixture, 'B2')!.value).toBe('x');

        historyButton(fixture, 'Redo').click();
        fixture.detectChanges();
        expect(cellInput(fixture, 'B2')!.value).toBe('changed');
    });

    it('walks back through several changes in order', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        type(fixture, 'B2', 'one');
        type(fixture, 'B3', 'two');

        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();
        expect([cellInput(fixture, 'B2')!.value, cellInput(fixture, 'B3')!.value]).toEqual(['one', '']);

        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();
        expect([cellInput(fixture, 'B2')!.value, cellInput(fixture, 'B3')!.value]).toEqual(['x', '']);
        expect(historyButton(fixture, 'Undo').disabled).withContext('back at the loaded document').toBeTrue();
    });

    /** Clearing a range is one change, not one per cell. */
    it('takes back a whole cleared selection at once', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: { S: { cells: { A1: { value: 'a' }, B1: { value: 'b' }, A2: { value: 'c' }, B2: { value: 'd' } } } },
        });
        fixture.detectChanges();

        cellInput(fixture, 'A1')!.focus();
        shiftClick(cellInput(fixture, 'B2')!);
        fixture.detectChanges();
        press(cellInput(fixture, 'B2')!, 'Delete');
        fixture.detectChanges();
        expect(cellInput(fixture, 'B2')!.value).toBe('');

        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();

        expect([
            cellInput(fixture, 'A1')!.value,
            cellInput(fixture, 'B1')!.value,
            cellInput(fixture, 'A2')!.value,
            cellInput(fixture, 'B2')!.value,
        ]).toEqual(['a', 'b', 'c', 'd']);
    });

    /** A new change abandons the branch that was undone away from. */
    it('drops the redo once a different change is made', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        type(fixture, 'B2', 'one');
        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();
        expect(historyButton(fixture, 'Redo').disabled).toBeFalse();

        type(fixture, 'B3', 'elsewhere');

        expect(historyButton(fixture, 'Redo').disabled).toBeTrue();
    });

    it('undoes on Ctrl+Z while a cell is only selected', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        type(fixture, 'B2', 'changed');
        cellInput(fixture, 'B3')!.focus();
        fixture.detectChanges();

        press(cellInput(fixture, 'B3')!, 'z', { ctrlKey: true });
        fixture.detectChanges();

        expect(cellInput(fixture, 'B2')!.value).toBe('x');
    });

    /**
     * While editing, Ctrl+Z is the INPUT's own undo: the author is taking back
     * characters, not the last thing that happened to the document.
     */
    it('leaves Ctrl+Z to the text while a cell is being edited', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        type(fixture, 'B3', 'committed');
        const input = cellInput(fixture, 'B2')!;
        input.focus();
        input.value = 'half typed';
        input.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        const event = press(input, 'z', { ctrlKey: true });
        fixture.detectChanges();

        expect(event.defaultPrevented).toBeFalse();
        expect(cellInput(fixture, 'B3')!.value).withContext('the document was not touched').toBe('committed');
    });

    /**
     * Snapshots are CLONES. Several write paths mutate the sheet in place and
     * only then publish a shallow copy, so a snapshot that shared structure
     * with the live document would be edited along with it and undo would
     * restore the state it was trying to leave.
     */
    it('remembers a state the next edit cannot reach into', () => {
        const fixture = makeFixture();
        respondWith(SHEET);
        fixture.detectChanges();

        type(fixture, 'B2', 'first');
        type(fixture, 'B2', 'second');
        type(fixture, 'B2', 'third');

        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();
        expect(cellInput(fixture, 'B2')!.value).toBe('second');

        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();
        expect(cellInput(fixture, 'B2')!.value).toBe('first');
    });

    // ---- Find and replace ---------------------------------------------------

    const TOKENS = {
        version: 1,
        sheets: {
            S: {
                cells: {
                    A1: { value: 'Order number' },
                    B1: { value: '{var:order.number}' },
                    A2: { value: 'order total' },
                    B2: { value: '{var:order.total}' },
                },
            },
        },
    };

    function findPanel(fixture: ReturnType<typeof makeFixture>): HTMLElement | null {
        return fixture.nativeElement.querySelector('.sheet-editor__find');
    }

    function findInput(fixture: ReturnType<typeof makeFixture>, label: string): HTMLInputElement {
        const input = fixture.nativeElement.querySelector(`.sheet-editor__find input[aria-label="${label}"]`);
        if (!(input instanceof HTMLInputElement)) throw new Error(`no ${label} field`);

        return input;
    }

    function findAction(fixture: ReturnType<typeof makeFixture>, label: string): HTMLButtonElement {
        const buttons: HTMLButtonElement[] = Array.from(
            fixture.nativeElement.querySelectorAll('.sheet-editor__find button'));
        const found = buttons.find(b => (b.textContent ?? '').trim() === label);
        if (!found) throw new Error(`no ${label} button`);

        return found;
    }

    function search(fixture: ReturnType<typeof makeFixture>, query: string): void {
        fixture.nativeElement.querySelector('button[aria-label="Find and replace"]').click();
        fixture.detectChanges();
        const input = findInput(fixture, 'Find');
        input.value = query;
        input.dispatchEvent(new Event('input'));
        fixture.detectChanges();
    }

    it('opens on the toolbar button and closes again', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        expect(findPanel(fixture)).toBeNull();

        fixture.nativeElement.querySelector('button[aria-label="Find and replace"]').click();
        fixture.detectChanges();
        expect(findPanel(fixture)).not.toBeNull();
    });

    /**
     * ⚠️ The browser's own find bar would open over a grid it cannot search:
     * the cells are inputs, and their text is not in the page.
     */
    it('takes Ctrl+F from the browser', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        const event = new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, cancelable: true });
        fixture.nativeElement.dispatchEvent(event);
        fixture.detectChanges();

        expect(findPanel(fixture)).not.toBeNull();
        expect(event.defaultPrevented).toBeTrue();
    });

    /**
     * A token has no computed value -- it is filled when the document is
     * generated -- so this is the search a template actually needs.
     */
    it('finds a token and says how many cells hold it', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        search(fixture, '{var:order');

        expect(fixture.nativeElement.querySelector('.sheet-editor__find-count').textContent.trim())
            .toBe('2 cells');
        expect(cellInput(fixture, 'B1')!.closest('td')!.classList).toContain('sheet-editor__cell--match');
        expect(cellInput(fixture, 'A1')!.closest('td')!.classList).not.toContain('sheet-editor__cell--match');
    });

    it('walks to the next match, and round again', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        search(fixture, '{var:order');

        findAction(fixture, 'Next').click();
        fixture.detectChanges();
        expect(activeRef(fixture)).toBe('B1');

        findAction(fixture, 'Next').click();
        fixture.detectChanges();
        expect(activeRef(fixture)).toBe('B2');

        findAction(fixture, 'Next').click();
        fixture.detectChanges();
        expect(activeRef(fixture)).withContext('wraps rather than stopping').toBe('B1');
    });

    it('replaces the one it is on and moves along', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        search(fixture, 'order.number');
        const replacement = findInput(fixture, 'Replace with');
        replacement.value = 'invoice.number';
        replacement.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        findAction(fixture, 'Replace').click();
        fixture.detectChanges();

        expect(savedCell(fixture, 'B1')).toEqual({ value: '{var:invoice.number}' });
    });

    it('replaces everywhere at once', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        search(fixture, 'order');
        const replacement = findInput(fixture, 'Replace with');
        replacement.value = 'invoice';
        replacement.dispatchEvent(new Event('input'));
        fixture.detectChanges();

        findAction(fixture, 'Replace all').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A1')!.value)
            .withContext('matched case-insensitively, rewritten with what was typed')
            .toBe('invoice number');
        expect(cellInput(fixture, 'A2')!.value).toBe('invoice total');
        expect(cellInput(fixture, 'B1')!.value).toBe('{var:invoice.number}');
    });

    /** One change, so one Ctrl+Z takes the whole sweep back. */
    it('makes replace all a single undo step', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        search(fixture, 'order');
        const replacement = findInput(fixture, 'Replace with');
        replacement.value = 'invoice';
        replacement.dispatchEvent(new Event('input'));
        fixture.detectChanges();
        findAction(fixture, 'Replace all').click();
        fixture.detectChanges();

        historyButton(fixture, 'Undo').click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'A2')!.value).toBe('order total');
        expect(cellInput(fixture, 'B1')!.value).toBe('{var:order.number}');
    });

    it('respects match case when asked', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        search(fixture, 'Order');
        expect(fixture.nativeElement.querySelector('.sheet-editor__find-count').textContent.trim())
            .withContext('A1, B1, A2 and B2 all hold it in some case')
            .toBe('4 cells');

        const box = findInput(fixture, 'Match case');
        box.checked = true;
        box.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        expect(fixture.nativeElement.querySelector('.sheet-editor__find-count').textContent.trim())
            .toBe('1 cell');
    });

    /** A sheet left painted with matches of a search nobody is running. */
    it('takes the highlights with it when it closes', () => {
        const fixture = makeFixture();
        respondWith(TOKENS);
        fixture.detectChanges();

        search(fixture, 'order');
        expect(cellInput(fixture, 'A2')!.closest('td')!.classList).toContain('sheet-editor__cell--match');

        const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
        findPanel(fixture)!.dispatchEvent(escape);
        fixture.detectChanges();

        expect(findPanel(fixture)).toBeNull();
        expect(cellInput(fixture, 'A2')!.closest('td')!.classList).not.toContain('sheet-editor__cell--match');
        expect(escape.defaultPrevented).withContext('or the dialog closes').toBeTrue();
    });

    // ---- Conditional formatting ---------------------------------------------

    function rulesPanel(fixture: ReturnType<typeof makeFixture>): HTMLElement | null {
        return fixture.nativeElement.querySelector('.sheet-editor__rules');
    }

    function openRules(fixture: ReturnType<typeof makeFixture>, ref: string): void {
        cellInput(fixture, ref)!.focus();
        fixture.detectChanges();
        fixture.nativeElement.querySelector('button[aria-label="Conditional formatting"]').click();
        fixture.detectChanges();
    }

    function ruleField(fixture: ReturnType<typeof makeFixture>, label: string): HTMLInputElement | HTMLSelectElement {
        const el = rulesPanel(fixture)!.querySelector(`[aria-label="${label}"]`);
        if (!(el instanceof HTMLInputElement || el instanceof HTMLSelectElement)) {
            throw new Error(`no ${label} field`);
        }

        return el;
    }

    function addRule(
        fixture: ReturnType<typeof makeFixture>,
        when: string,
        value: string,
        fill: string,
    ): void {
        const condition = ruleField(fixture, 'Condition');
        condition.value = when;
        condition.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        const box = ruleField(fixture, 'Condition value');
        box.value = value;
        box.dispatchEvent(new Event('input'));

        const colour = ruleField(fixture, 'Rule fill');
        colour.value = fill;
        colour.dispatchEvent(new Event('change'));
        fixture.detectChanges();

        const add = Array.from(rulesPanel(fixture)!.querySelectorAll('button'))
            .find(b => (b.textContent ?? '').trim() === 'Add rule');
        (add as HTMLButtonElement).click();
        fixture.detectChanges();
    }

    const NUMBERS = {
        version: 1,
        sheets: { S: { cells: { B2: { value: '95' }, B3: { value: '20' } } } },
    };

    it('paints the cells a rule claims, and leaves the rest', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        shiftClick(cellInput(fixture, 'B3')!);
        fixture.detectChanges();
        fixture.nativeElement.querySelector('button[aria-label="Conditional formatting"]').click();
        fixture.detectChanges();

        addRule(fixture, 'greaterThan', '90', '#ffcccc');

        expect(cellInput(fixture, 'B2')!.closest('td')!.style.background).toContain('rgb(255, 204, 204)');
        expect(cellInput(fixture, 'B3')!.closest('td')!.style.background).withContext('20 is not over 90').toBe('');
    });

    /**
     * ⚠️ A rule is judged on what the cell SHOWS -- a formula by its result.
     * Reading the formula's text would never match a number at all.
     */
    it('judges a formula by what it computes', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                S: {
                    cells: { A1: { value: '50' }, A2: { value: '50' }, B2: { formula: 'A1+A2' } },
                    conditionals: { B2: [{ when: 'greaterThan', value: '90', background: '#FFCCCC' }] },
                },
            },
        });
        fixture.detectChanges();

        expect(cellInput(fixture, 'B2')!.value).toBe('100');
        expect(cellInput(fixture, 'B2')!.closest('td')!.style.background).toContain('rgb(255, 204, 204)');
    });

    /**
     * ⚠️ The RULE wins over the cell's own fill. A conditional colour that lost
     * would show only on cells the author had left plain -- so a shaded table,
     * exactly where an overdue line needs highlighting, would show nothing.
     */
    it('lets a rule beat the fill already on the cell', () => {
        const fixture = makeFixture();
        respondWith({
            version: 1,
            sheets: {
                S: {
                    cells: { B2: { value: '95', background: '#EEEEEE' } },
                    conditionals: { B2: [{ when: 'greaterThan', value: '90', background: '#FFCCCC' }] },
                },
            },
        });
        fixture.detectChanges();

        expect(cellInput(fixture, 'B2')!.closest('td')!.style.background).toContain('rgb(255, 204, 204)');
    });

    it('lists a rule so it can be taken off again', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        openRules(fixture, 'B2');
        addRule(fixture, 'greaterThan', '90', '#ffcccc');

        const remove = rulesPanel(fixture)!.querySelector('button[aria-label^="Remove rule"]');
        expect(remove).not.toBeNull();

        (remove as HTMLButtonElement).click();
        fixture.detectChanges();

        expect(cellInput(fixture, 'B2')!.closest('td')!.style.background).toBe('');
    });

    it('saves the rules with the document', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        openRules(fixture, 'B2');
        addRule(fixture, 'greaterThan', '90', '#ffcccc');

        expect(savedSheet(fixture).conditionals).toEqual({
            B2: [{ when: 'greaterThan', value: '90', background: '#FFCCCC', color: '#842029', bold: false, italic: false }],
        });
    });

    /**
     * Naming a range, which is what lets a total follow rows that do not exist
     * when the template is written (#2385).
     *
     * The helpers have their own spec; what can go wrong HERE is the wiring --
     * a name that never reaches the document, or reaches it describing the
     * ACTIVE CELL rather than the selection an author had highlighted.
     */
    function openNames(fixture: ReturnType<typeof makeFixture>): void {
        const buttons: HTMLButtonElement[] = Array.from(fixture.nativeElement.querySelectorAll('button'));
        const names = buttons.find(b => (b.textContent ?? '').trim() === 'Names');
        if (!names) throw new Error('no Names button rendered');
        names.click();
        fixture.detectChanges();
    }

    function typeName(fixture: ReturnType<typeof makeFixture>, name: string): void {
        const input: HTMLInputElement = fixture.nativeElement.querySelector('input[aria-label="New name"]');
        input.value = name;
        input.dispatchEvent(new Event('input'));
        fixture.detectChanges();
    }

    function addNameButton(fixture: ReturnType<typeof makeFixture>): HTMLButtonElement {
        const buttons: HTMLButtonElement[] = Array.from(fixture.nativeElement.querySelectorAll('button'));
        const add = buttons.find(b => (b.textContent ?? '').trim().startsWith('Name '));
        if (!add) throw new Error('no add button rendered');

        return add;
    }

    /** What the save request would carry for the document's names. */
    function savedNames(fixture: ReturnType<typeof makeFixture>): Record<string, string> | undefined {
        saveButton(fixture).click();
        const save = httpMock.expectOne(CONTENT_URL);
        const written: { definedNames?: Record<string, string> } = JSON.parse(save.request.body.content);
        save.flush({ contentHash: 'x' });

        return written.definedNames;
    }

    it('names the SELECTION, not merely the active cell', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        // Select B2:B3, the shape an author highlights over a band's column.
        cellInput(fixture, 'B2')!.focus();
        fixture.detectChanges();
        cellInput(fixture, 'B3')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, shiftKey: true }));
        fixture.detectChanges();

        openNames(fixture);
        typeName(fixture, 'items_amount');
        addNameButton(fixture).click();
        fixture.detectChanges();

        expect(savedNames(fixture)).toEqual({ items_amount: 'S!$B$2:$B$3' });
    });

    it('refuses a name that would shadow a cell, and says why', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        fixture.detectChanges();
        openNames(fixture);
        typeName(fixture, 'Q4');

        expect(addNameButton(fixture).disabled).toBe(true);
        expect(fixture.nativeElement.querySelector('.sheet-editor__names-error')?.textContent)
            .toContain('cell reference');
    });

    it('lists a name so it can be taken off again', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        cellInput(fixture, 'B2')!.focus();
        fixture.detectChanges();
        openNames(fixture);
        typeName(fixture, 'one_cell');
        addNameButton(fixture).click();
        fixture.detectChanges();

        const remove: HTMLButtonElement = fixture.nativeElement
            .querySelector('button[aria-label="Remove the name one_cell"]');
        expect(remove).not.toBeNull();

        remove.click();
        fixture.detectChanges();

        // ⚠️ Absent, not empty: the backend omits the key too, so a template
        // that lost its last name is byte-identical to one that never had one.
        expect(savedNames(fixture)).toBeUndefined();
    });

    // -- Zoom ---------------------------------------------------------------
    //
    // Reported as "zoom over the spreadsheet zooms the whole dashboard": ctrl +
    // wheel is the browser's gesture until something takes it, and nothing did.
    // The half of these specs that matters is not that a number goes up -- it
    // is that the gesture is CANCELLED, and that the grid's own arithmetic
    // follows the new scale.

    function gridTable(fixture: ReturnType<typeof makeFixture>): HTMLElement {
        const table: HTMLElement | null = fixture.nativeElement.querySelector('.sheet-editor__grid');
        if (!table) throw new Error('no grid rendered');

        return table;
    }

    /** What the footer reports, which is the only zoom an author can see. */
    function zoomReadout(fixture: ReturnType<typeof makeFixture>): string {
        const button: HTMLElement | null = fixture.nativeElement.querySelector('button[aria-label="Zoom"]');

        return (button?.textContent ?? '').trim();
    }

    function zoomButton(fixture: ReturnType<typeof makeFixture>, label: string): HTMLButtonElement {
        const button: HTMLButtonElement | null = fixture.nativeElement.querySelector(`button[aria-label="${label}"]`);
        if (!button) throw new Error(`no ${label} button rendered`);

        return button;
    }

    /**
     * A ctrl + wheel, as a browser delivers one.
     *
     * `cancelable: true` is not decoration: an uncancellable event would report
     * `defaultPrevented === false` however well the component behaved, and the
     * assertion that matters here is exactly that flag.
     *
     * The anchor is given RELATIVE to the grid's top-left, because that is what
     * the component measures against.
     */
    function ctrlWheel(
        fixture: ReturnType<typeof makeFixture>,
        deltaY: number,
        at: { x: number; y: number } = { x: 0, y: 0 },
    ): WheelEvent {
        const body = grid(fixture);
        const box = body.getBoundingClientRect();
        const event = new WheelEvent('wheel', {
            deltaY,
            ctrlKey: true,
            bubbles: true,
            cancelable: true,
            clientX: box.left + at.x,
            clientY: box.top + at.y,
        });
        body.dispatchEvent(event);
        fixture.detectChanges();

        return event;
    }

    /** Step to a zoom through the footer, and FAIL rather than silently stop short. */
    function zoomTo(fixture: ReturnType<typeof makeFixture>, percent: number): void {
        const button = zoomButton(fixture, percent > 100 ? 'Zoom in' : 'Zoom out');
        for (let guard = 0; guard < 40 && zoomReadout(fixture) !== `${percent}%`; guard++) {
            button.click();
            fixture.detectChanges();
        }
        expect(zoomReadout(fixture))
            .withContext(`the zoom never reached ${percent}%`)
            .toBe(`${percent}%`);
    }

    function scrollGrid(fixture: ReturnType<typeof makeFixture>, px: number): void {
        const body = grid(fixture);
        body.scrollTop = px;
        body.dispatchEvent(new Event('scroll'));
        fixture.detectChanges();
    }

    it('zooms the GRID on ctrl+wheel, and takes the gesture from the browser', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        const event = ctrlWheel(fixture, -100);

        // ⚠️ THE bug. An uncancelled wheel stays the browser's, and the browser
        // zooms the whole admin -- shell, toolbar, dialog frame -- around a grid
        // that is still exactly the size it was.
        expect(event.defaultPrevented)
            .withContext('ctrl+wheel must be cancelled, or the browser zooms the admin instead')
            .toBeTrue();
        expect(zoomReadout(fixture)).toBe('110%');
        // Asked of the PAINT, not of the component: a signal that never reaches
        // a stylesheet is a number that changes while the grid does not.
        expect(Number(getComputedStyle(gridTable(fixture)).zoom))
            .withContext('the grid itself must carry the zoom')
            .toBeCloseTo(1.1, 2);
    });

    it('leaves a plain wheel to the browser, so the grid still scrolls', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        const event = new WheelEvent('wheel', { deltaY: 120, bubbles: true, cancelable: true });
        grid(fixture).dispatchEvent(event);
        fixture.detectChanges();

        expect(event.defaultPrevented)
            .withContext('cancelling a plain wheel would fix zoom by breaking scrolling')
            .toBeFalse();
        expect(zoomReadout(fixture)).toBe('100%');
    });

    it('leaves ctrl+wheel outside the grid alone', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        const footer: HTMLElement = fixture.nativeElement.querySelector('.sheet-editor__footer');
        const event = new WheelEvent('wheel', { deltaY: -100, ctrlKey: true, bubbles: true, cancelable: true });
        footer.dispatchEvent(event);
        fixture.detectChanges();

        // Over a surface with nothing to zoom, the gesture is still the
        // browser's -- taking it there would be taking it for nothing.
        expect(event.defaultPrevented).toBeFalse();
        expect(zoomReadout(fixture)).toBe('100%');
    });

    it('computes the row window in GRID pixels, so a zoomed sheet shows the rows it is scrolled to', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        scrollGrid(fixture, 2600);
        expect(cellInput(fixture, 'A101'))
            .withContext('2600px of 26px rows is row 101')
            .not.toBeNull();
        expect(cellInput(fixture, 'A51')).toBeNull();

        zoomTo(fixture, 200);
        scrollGrid(fixture, 2600);

        // ⚠️ The same scrollbar travel covers HALF the sheet at 200%. Reading
        // the raw pixels instead would build the window for row 101 while row 51
        // is what the grid is showing -- blanks above, missing cells below.
        expect(cellInput(fixture, 'A51'))
            .withContext('2600 screen px at 200% is 1300 grid px, which is row 51')
            .not.toBeNull();
        expect(cellInput(fixture, 'A101')).toBeNull();
    });

    it('keeps the grid point under the pointer where it is', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        scrollGrid(fixture, 1000);

        // Anchored at the very top of the viewport: whatever row is there has to
        // still be there afterwards, so the scrollbar must move with the scale.
        ctrlWheel(fixture, -100, { x: 0, y: 0 });

        expect(Math.abs(grid(fixture).scrollTop - 1100))
            .withContext(`scrollTop was ${grid(fixture).scrollTop}, expected about 1100`)
            .toBeLessThan(2);
    });

    it('stops at both ends of the range', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        for (let i = 0; i < 30; i++) ctrlWheel(fixture, -100);
        expect(zoomReadout(fixture)).toBe('200%');
        expect(zoomButton(fixture, 'Zoom in').disabled).toBeTrue();

        for (let i = 0; i < 40; i++) ctrlWheel(fixture, 100);
        expect(zoomReadout(fixture)).toBe('50%');
        expect(zoomButton(fixture, 'Zoom out').disabled).toBeTrue();
    });

    it('goes back to 100% from the readout', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        ctrlWheel(fixture, -100);
        zoomButton(fixture, 'Zoom').click();
        fixture.detectChanges();

        expect(zoomReadout(fixture)).toBe('100%');
        expect(Number(getComputedStyle(gridTable(fixture)).zoom)).toBeCloseTo(1, 2);
    });

    it('resizes a column by the GRID pixels the pointer crossed', () => {
        const fixture = makeFixture();
        respondWith(NUMBERS);
        fixture.detectChanges();

        zoomTo(fixture, 200);

        const grip: HTMLElement = fixture.nativeElement.querySelector('span[aria-label="Resize column B"]');
        grip.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX: 500 }));
        document.dispatchEvent(new MouseEvent('mousemove', { clientX: 600, buttons: 1 }));
        document.dispatchEvent(new MouseEvent('mouseup'));
        fixture.detectChanges();

        saveButton(fixture).click();
        const save = httpMock.expectOne(CONTENT_URL);
        const written: { sheets: Record<string, { columnWidths?: Record<string, number> }> } =
            JSON.parse(save.request.body.content);
        save.flush({ contentHash: 'x' });

        // ⚠️ The column starts at the default 141px and the pointer crossed 100
        // SCREEN pixels at 200%, which is 50 grid pixels. 191, not 241 -- and
        // certainly not the 382 that measuring a zoomed header and adding raw
        // pointer travel would store.
        expect(columnWidthToPx(written.sheets['S'].columnWidths?.['B'] ?? 0))
            .withContext('a drag at 200% must not resize twice as fast as the pointer')
            .toBeCloseTo(191, -1);
    });
});
