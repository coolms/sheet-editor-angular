/**
 * The public surface of `@coolms/sheet-editor-angular`.
 *
 * Two modules, and both are public on purpose. The DIALOG is what an
 * application registers as a handler for spreadsheet MIME types; the MODEL is
 * the immutable document it edits -- refs, ranges, merges, formats and the
 * `with*` transforms -- which a caller needs to read a sheet, or to build one
 * without opening a dialog at all.
 */
export * from './sheet-document.model';
export * from './sheet-editor-dialog.component';

// -- Formulas ----------------------------------------------------------------
// The engine is public because it is useful without the grid: a caller that
// needs the VALUE of a template's cell -- a preview, a validation pass -- should
// not have to open a dialog to get one. `allFunctions()` and its shelved twin
// `functionsByCategory()` are what BOTH function surfaces read -- the type-ahead
// helper and the browsable list -- so there is one catalogue rather than three.
export * from './formula/values';
export * from './formula/ast';
export * from './formula/tokenise';
export * from './formula/parse';
export * from './formula/functions';
export * from './formula/evaluate';
export * from './formula/helper';
