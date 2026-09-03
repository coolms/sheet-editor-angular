/**
 * Type-only, so nothing from the engine reaches the emitted bundle.
 *
 * `@coolms/document-engine` is declared an OPTIONAL peer, and a static import
 * made that a lie: ng-packagr emits one fesm bundle with no code splitting, so
 * a top-level import must resolve for every consumer while `optional` tells
 * npm not to install it. The module is loaded on demand below instead --
 * which costs nothing, because this path was already async and already had to
 * survive the manifest not arriving.
 */
import type { FontManifest } from '@coolms/document-engine';

/**
 * The families this editor offers, read from the platform's font manifest.
 *
 * ## Why a workbook cares what the document engine vendors
 *
 * The grid has no pagination engine, so the failure is quieter here than in the
 * document editor -- but it is the same failure. A name the platform ships
 * nothing for is painted with whatever the author's machine has, written into
 * the `.xlsx` as that name, and then resolved by Excel or LibreOffice to a
 * third thing. Offering only what we vendor means the cell an author sees is
 * the cell the workbook describes.
 *
 *  NAMES only. None of the 7.2MB of faces is fetched: a grid draws with the
 * browser's own copy of a family the platform also vendors, and the manifest is
 * consulted for what to put in the select, nothing more.
 */

let once: Promise<readonly string[]> | null = null;

/**
 * Drop the memo. For specs only.
 *
 * Module state outlives a test, so without this the second spec in a file
 * asserts against the first one's fetch and passes for the wrong reason.
 */
export function resetOfferedFamilies(): void {
    once = null;
}

export function offeredFamilies(): Promise<readonly string[]> {
    // Memoised on the module rather than per dialog: the sheet dialog is opened
    // and closed repeatedly, and re-fetching a static asset each time is a
    // request nobody asked for.
    return once ??= import('@coolms/document-engine')
        .then(async (engine) => engine.offeredFontFamilies(
            await fetchManifest(engine.FONT_MANIFEST_ASSET),
        ))
        .catch((error: unknown) => {
            // Not rethrown: an empty list leaves the select at "Default", which
            // is a workbook that inherits -- an editor that will not open
            // because a font list did not arrive would be the worse failure.
            // The optional peer being absent altogether lands here too, and
            // means the same thing to the person editing the sheet.
            console.error('[coolms-sheet-editor] the font manifest failed to load; only the default family is offered', error);

            // Cleared so the NEXT open tries again; a transient 503 should not
            // cost the tab its font select for the rest of the session.
            once = null;

            return [];
        });
}

async function fetchManifest(asset: string): Promise<FontManifest> {
    const response = await fetch(asset);
    if (!response.ok) {
        throw new Error(`Could not load ${asset}: HTTP ${response.status}`);
    }

    return await response.json() as FontManifest;
}
