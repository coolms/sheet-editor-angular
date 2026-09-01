import { offeredFamilies, resetOfferedFamilies } from './font-families';

/**
 * The grid's font select, and what it does when the manifest does not arrive.
 *
 * The DERIVATION is tested where it lives, in `@coolms/document-engine`. What
 * is only true here is the caching: the sheet dialog is constructed on every
 * open, so a fetch per open would be a request nobody asked for -- and a
 * failure that stuck would cost the tab its font list for the session.
 */
describe('the sheet editor font families', () => {
    let fetchSpy: jasmine.Spy;

    const manifest = {
        version: 1,
        defaults: { family: 'Carlito', sizePt: 11 },
        families: {
            Carlito: { substitutes: ['Calibri', 'Carlito'], files: {} },
            Gelasio: { substitutes: ['Georgia', 'Gelasio'], files: {} },
        },
    };

    const ok = (): Promise<Response> => Promise.resolve({
        ok: true,
        status: 200,
        json: () => Promise.resolve(manifest),
    } as Response);

    beforeEach(() => {
        resetOfferedFamilies();
        fetchSpy = spyOn(window, 'fetch');
    });

    it('offers the names the manifest states, default family first', async () => {
        fetchSpy.and.callFake(ok);

        expect(await offeredFamilies()).toEqual(['Calibri', 'Georgia']);
    });

    it('fetches once however many times the dialog is opened', async () => {
        fetchSpy.and.callFake(ok);

        await offeredFamilies();
        await offeredFamilies();
        await offeredFamilies();

        expect(fetchSpy).toHaveBeenCalledTimes(1);
    });

    it('offers nothing rather than throwing when the manifest cannot be read', async () => {
        // An editor that will not open because a font list did not arrive is
        // the worse failure. The select stays at "Default", which is a
        // workbook that inherits.
        fetchSpy.and.returnValue(Promise.resolve({ ok: false, status: 503 } as Response));
        spyOn(console, 'error');

        expect(await offeredFamilies()).toEqual([]);
        expect(console.error).toHaveBeenCalled();
    });

    it('tries again on the next open after a failure', async () => {
        //  The memo is CLEARED on failure. Caching the empty list would cost
        // the tab its font select for the rest of the session over one 503.
        spyOn(console, 'error');
        fetchSpy.and.returnValue(Promise.resolve({ ok: false, status: 503 } as Response));
        expect(await offeredFamilies()).toEqual([]);

        fetchSpy.and.callFake(ok);

        expect(await offeredFamilies()).toEqual(['Calibri', 'Georgia']);
        expect(fetchSpy).toHaveBeenCalledTimes(2);
    });
});
