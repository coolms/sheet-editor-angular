# @coolms/sheet-editor-angular

The CoolMS spreadsheet editor for Angular: a grid over an immutable
sheet-document model, with cell formats, merges, column and row sizing, and
multi-sheet editing.

## Install

```bash
npm install @coolms/sheet-editor-angular @coolms/ui-angular @coolms/core-angular
```

Angular 22, `@angular/cdk`, NGXS 22 and RxJS 7 are peers.

## Use

Register it as the handler for spreadsheet MIME types, and the file surfaces
open it without knowing what a spreadsheet is:

```ts
import { SheetEditorDialogComponent } from '@coolms/sheet-editor-angular';

FileEditorRegistry.register(SHEET_DOCUMENT_MIME, { component: SheetEditorDialogComponent });
```

The document model is public too, and separately useful — reading a sheet or
building one does not require opening a dialog:

```ts
import {
    parseSheetDocument, serialiseSheetDocument,
    withStyle, withMerge, withNumberFormat,
    parseRef, parseRange, rangeBetween,
} from '@coolms/sheet-editor-angular';
```

**The model is immutable.** Every mutator is a `with*` function returning a new
document rather than editing one in place, so undo, change detection and
"what changed" are properties of the data instead of bookkeeping in the
component.

## Why it is its own package

It lived in `@coolms/ui-angular` until it was ~3,850 lines, which is already
larger than several packages beside it. What settled it was what comes next: a
formula engine, an Excel/Sheets-style formula helper, filtering, and form
element types. All of that lands here, and inside the UI kit it would have
churned the kit's public surface — and the kit's version number — every time
a spreadsheet feature moved.

## Building it

```bash
npm --prefix ../core-angular run build
npm --prefix ../ui-angular run build
npm run build
```

Peers are consumed as BUILT output, never as sources: compiling a peer's
sources in would place them outside this package's `rootDir` and ship a second
copy of that peer to anyone installing both.

## Status

Not published, and no repository yet — `tools/publish-guard.sh` reports "no
tracked files" for it, which is the guard declining to certify what it cannot
read rather than a clean result.

## Licence

MIT.
