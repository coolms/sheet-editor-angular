# Changelog

All notable changes to `@coolms/sheet-editor-angular` are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

This file starts at the version named below, which is what the registry
currently serves. Earlier alphas are deliberately not reconstructed: entries are written
in the same commit as the work they describe, and inventing the ones that
predate this file would be a worse record than not having them.

## Unreleased

### Changed

- The selected column or row header and the function helper's current item
  put their label in `--cms-selected-fg`, the foreground on the solid mark;
  the resize grip's hover reads `--cms-primary` -- a handle is an
  affordance, and affordances are the primary by ruling.
- The selection reads `--cms-selected`, the host theme's token for "this one
  is selected", instead of the blue `--cms-primary`: the active cell's ring,
  the range's inset, the selected column or row header (its label in
  `--cms-accent-fg`, the foreground for a solid mark) and the function
  helper's current item. The resize grip's hover reads `--cms-accent`, as the
  kit's own handle does. A lit filter button, the formula cell's text and the
  input's focus ring keep `--cms-primary`: a filter that is filtering, a
  formula and focus are not selections.

### Added

- Declares `bugs` so a page imported from this package, and the catalogue,
  know where a correction is filed. The registry filled the gap from GitHub when
  the manifest was silent; the declared field is the one that holds on any
  registry.

## 2.0.0-alpha.2 -- 2026-09-03

**A pre-release, carrying no compatibility promise.** Published under the
`alpha` dist-tag.

The spreadsheet editor: a grid over an immutable sheet-document model, with
cell formats, merges, column and row sizing, and multi-sheet editing. The model
is public and separately useful -- reading or building a sheet does not require
opening a dialog -- and every mutator is a `with*` function returning a new
document, so undo and change detection are properties of the data rather than
bookkeeping in the component.

Also here: the formula engine and its parser, an Excel/Sheets-style formula
helper, filtering, conditional formatting, number formats, defined names,
find-and-replace and clipboard handling.

### Fixed

- **`@coolms/document-engine` was declared an optional peer and imported
  unconditionally**, which made the package unbuildable from a clean install.
  Only the cell font picker reads it, and it is now fetched on demand -- the
  path was already asynchronous and already had to survive the font manifest
  not arriving, so a peer nobody installed lands where a failed fetch already
  landed: the select stays on "Default".
- The install instructions never named that optional peer.
