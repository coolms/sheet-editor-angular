# @coolms/sheet-editor-angular

The CoolMS spreadsheet editor for Angular: a grid over an immutable
sheet-document model, with cell formats, merges, column and row sizing, and
multi-sheet editing.

## Install

```bash
npm install @coolms/sheet-editor-angular @coolms/ui-angular @coolms/core-angular
```

Angular 22, `@angular/cdk`, NGXS 22 and RxJS 7 are peers.

`@coolms/document-engine` is an **optional** peer, used only for the font
catalogue the cell font picker reads.

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
larger than several packages beside it. What settled it was what came next —
and it has since landed here rather than in the kit: the formula engine and
its parser, an Excel/Sheets-style formula helper, filtering, conditional
formatting, number formats, defined names, find-and-replace and clipboard
handling. Inside the UI kit every one of those would have churned the kit's
public surface, and its version number, each time a spreadsheet feature moved.

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

A pre-release: the shape is still moving and it carries no
compatibility promise.

## Licence

MIT.
