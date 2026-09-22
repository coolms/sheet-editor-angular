import { FileEditorRegistry } from '@coolms/ui-angular';
import { SHEET_DOCUMENT_MIME } from './sheet-document.constants';

/**
 * Why the `.dsheet` editor registration in `app.config.ts` must be EXACT.
 *
 * A native spreadsheet template is opened by the Documents library
 * through `FileEditorRegistry`; with no entry for its mime the operator gets
 * "No editor is registered for this template format" on a template the backend
 * can otherwise mint, fill and render.
 *
 * ## What this spec does NOT cover, and why
 *
 * It cannot assert that `app.config.ts` performs the registration, because
 * importing that module for its side effects pulls the whole application graph
 * into the karma build -- including `@coolms/designer`, whose `index.ts`
 * re-exports through ESM `./X.js` specifiers that the karma webpack build
 * cannot resolve (42 "Module not found" errors) even though `ng build` resolves
 * them fine. One such import takes the ENTIRE suite from 271 passing to zero
 * executed, so no spec may import `app.config`.
 *
 * What is testable is the property that makes the exact registration
 * load-bearing rather than tidy -- pinned here so a future reader does not
 * "simplify" it into an `application/json` fallback that cannot fire.
 */
describe('native spreadsheet template editor registration', () => {
 it('cannot inherit an editor from a wildcard, so the entry must be exact', () => {
 // The resolver's ONLY fallback is the mime's first segment plus `/*`.
 // For `application/x-coolms-sheet+json` that is `application/*`, which
 // nothing registers -- the `+json` suffix does not reach
 // `application/json`.
 expect(SHEET_DOCUMENT_MIME.split('/')[0] + '/*').toBe('application/*');
        expect(FileEditorRegistry.hasEditorForMime('application/*')).toBeFalse();
        expect(SHEET_DOCUMENT_MIME).not.toBe('application/json');
    });

 it('resolves once an exact entry exists, and misses without one', () => {
        const mime = 'application/x-coolms-sheet-spec+json';

        expect(FileEditorRegistry.hasEditorForMime(mime)).toBeFalse();

        class StubEditor {}
        FileEditorRegistry.register(mime, { component: StubEditor });

        expect(FileEditorRegistry.hasEditorForMime(mime)).toBeTrue();
    });

    /** The backend constant this must match: XlsxFormatProvider::DSHEET_MIME. */
 it('spells the mime the backend stamps on a native template', () => {
        expect(SHEET_DOCUMENT_MIME).toBe('application/x-coolms-sheet+json');
    });
});
