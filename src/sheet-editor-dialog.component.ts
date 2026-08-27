import { ChangeDetectionStrategy, Component, DestroyRef, ElementRef, NgZone, computed, inject, signal } from '@angular/core';
import { DIALOG_DATA, DialogRef } from '@angular/cdk/dialog';
import { HttpClient } from '@angular/common/http';
import { takeUntilDestroyed } from '@angular/core/rxjs-interop';
import { Store } from '@ngxs/store';
import { AppConfigState } from '@coolms/core-angular';
import { evaluateSheet } from './formula/evaluate';
import { displayValue, type CellValue } from './formula/values';
import { functionsByCategory, type FormulaFunction } from './formula/functions';
import {
    definedNamesOf, nameProblem, scopedRange, withDefinedName, withoutDefinedName,
} from './defined-names';
import { editForm, formatCellValue } from './number-format';
import { offeredFamilies } from './font-families';
import {
    findMatches, nextMatch, withReplacedAll, withReplacedIn, type FindOptions,
} from './find-replace';
import {
    CONDITIONAL_WHENS, conditionalsAt, lookFor, needsSecondValue, needsValue,
    withConditional, withoutConditional, type ConditionalWhen,
} from './conditional';
import {
    pasteOffset, pastedRange, parseClipboardText, rangeCorner, toClipboardText,
} from './clipboard';
import {
    applyCompletion, argumentLabel, helperAt, pointInsertAt, referencesIn, signatureParts, type HelperState,
} from './formula/helper';
import {
    autoFilterOf,
    BORDER_STYLES,
    borderCssAt,
    type BorderPreset,
    type BorderState,
    borderStateIn,
    DEFAULT_BORDER_COLOUR,
    cellToInput,
    columnToIndex,
    columnWidthFromPx,
    columnWidthOf,
    columnWidthToPx,
    filterBodyRows,
    filterColumns,
    frozenAt,
    withFreeze,
    gridExtent,
    indexToColumn,
    inputToCell,
    isFilterHeader,
    isKnownFormat,
    isMergeAnchor,
    refsInRange,
    type MergeBox,
    mergeCovering,
    mergeSpan,
    NUMBER_FORMATS,
    parseRange,
    parseRef,
    parseSheetDocument,
    rangeBetween,
    rangeContains,
    rowHeightFromPx,
    rowHeightToPx,
    safeSheetName,
    serialiseSheetDocument,
    type SheetCellDto,
    type SheetDocumentDto,
    type SheetDto,
    type SheetValidationDto,
    validationOptions,
    validationRangeAt,
    withBold,
    withColumnWidth,
    withItalic,
    sheetLookup,
    withMerge,
    withRowHeight,
    withBorderPreset,
    withAutoFilter,
    withClearedRange,
    withDeletedColumn,
    withDeletedRow,
    withInsertedColumn,
    withInsertedRow,
    withStyle,
    withValidation,
    withoutValidation,
    withPastedBlock,
    withWrap,
    withNewSheet,
    withNumberFormat,
    withoutAutoFilter,
    withoutMerge,
    withoutSheet,
    withRenamedSheet,
} from './sheet-document.model';
import { NativeDialogService, ToastService, type VfsNodeDto } from '@coolms/ui-angular';
import { NgStyle, NgTemplateOutlet } from '@angular/common';
import { CmsLoaderComponent } from '@coolms/core-angular';

/**
 * A grid surface for a native `.dsheet` template (ADR-155).
 *
 * ## Why a grid at all
 *
 * The backend can mint, fill and render a `.dsheet`, and until now the only way
 * to author one was to hand-edit JSON in CodeMirror. That is fine for a
 * developer and useless for the operator the format exists for: a spreadsheet
 * template is a GRID, and the whole argument for Tier 2 was that a document
 * workflow has to be first-class rather than upload-and-fill.
 *
 * ## What it deliberately is not
 *
 * Not a spreadsheet ENGINE. Nothing here evaluates a formula — `=B4*C4` is
 * stored, not computed, exactly as the model stores it and the renderer emits
 * it. Excel and LibreOffice compute it when the generated workbook opens.
 * Growing this into a calculator is the temptation ADR-155 names in its
 * consequences.
 *
 * ## What it must not lose
 *
 * A cell's `numberFormat` is the author's TYPE DECLARATION (#1977): `@` is what
 * keeps `00412` an order number rather than the integer 412. This edits a
 * cell's TEXT only and carries formatting through untouched — see
 * {@link inputToCell}. The format is not editable here yet; losing it silently
 * would be far worse than not offering it.
 */
/** The four arrows, as the directions they mean. */
const ARROWS: Readonly<Record<string, 'up' | 'down' | 'left' | 'right' | undefined>> = {
    ArrowUp: 'up',
    ArrowDown: 'down',
    ArrowLeft: 'left',
    ArrowRight: 'right',
};

@Component({
    selector: 'app-sheet-editor',
    standalone: true,
    imports: [CmsLoaderComponent, NgStyle, NgTemplateOutlet],
    changeDetection: ChangeDetectionStrategy.OnPush,
    // On the HOST, not the document: the shortcut belongs to this dialog while
    // it holds focus, and a document listener would keep firing after it
    // closed. `keydown.f11` is Angular's own key pseudo-event.
    host: {
        '(keydown.f11)': 'onFullScreenKey($event)',
        '(keydown.control.f)': 'onFindShortcut($event)',
        '(keydown.meta.f)': 'onFindShortcut($event)',
    },
    template: `
        <div class="sheet-editor" [class.sheet-editor--full]="fullScreen()">
            <div class="sheet-editor__header">
                <span class="sheet-editor__title">
                    <i class="bi bi-file-earmark-spreadsheet"></i>
                    {{ node.path }}
                    <!-- Always rendered, not only when a second sheet exists: it
                         is now the surface that CREATES the second one, so
                         hiding it at one sheet hid the way out of one sheet. -->
                    <select class="cms-input sheet-editor__sheet-pick"
                            aria-label="Active sheet"
                            [value]="activeSheet()"
                            (change)="selectSheet($any($event.target).value)">
                        @for (name of sheetNames(); track name) {
                            <option [value]="name" [selected]="name === activeSheet()">{{ name }}</option>
                        }
                    </select>
                    <button class="cms-btn cms-btn-sm" title="Add sheet" (click)="addSheet()">
                        <i class="bi bi-plus-lg"></i>
                    </button>
                    <button class="cms-btn cms-btn-sm" title="Rename sheet" (click)="renameSheet()">
                        <i class="bi bi-pencil"></i>
                    </button>
                    <button class="cms-btn cms-btn-sm"
                            title="Delete sheet"
                            [disabled]="sheetNames().length <= 1"
                            (click)="deleteSheet()">
                        <i class="bi bi-trash"></i>
                    </button>
                </span>
                <div class="sheet-editor__actions">
                    @if (dirty()) {
                        <span class="sheet-editor__dirty">unsaved changes</span>
                    }
                    <button class="cms-btn cms-btn-sm"
                            type="button"
                            title="Undo (Ctrl+Z)"
                            aria-label="Undo"
                            [disabled]="undoDepth() === 0"
                            (click)="undo()">
                        <i class="bi bi-arrow-counterclockwise"></i>
                    </button>
                    <button class="cms-btn cms-btn-sm"
                            type="button"
                            title="Redo (Ctrl+Y)"
                            aria-label="Redo"
                            [disabled]="redoDepth() === 0"
                            (click)="redo()">
                        <i class="bi bi-arrow-clockwise"></i>
                    </button>
                    <button class="cms-btn cms-btn-sm"
                            [title]="fullScreen() ? 'Exit full screen (F11)' : 'Full screen (F11)'"
                            [attr.aria-pressed]="fullScreen()"
                            aria-label="Full screen"
                            (click)="toggleFullScreen()">
                        <i class="bi"
                           [class.bi-arrows-fullscreen]="!fullScreen()"
                           [class.bi-fullscreen-exit]="fullScreen()"></i>
                    </button>
                    <button class="cms-btn cms-btn-sm" (click)="close()">
                        <i class="bi bi-x-lg"></i>
                    </button>
                </div>
            </div>

            <div class="sheet-editor__toolbar">
                <span class="sheet-editor__active">{{ activeRef() || '—' }}</span>
                <label class="sheet-editor__tool">
                    Format
                    <!-- Selection is expressed per-OPTION, not as [value] on the
                         select. A [value] binding is applied before the @if
                         below has attached the custom option, so the browser
                         finds no match and resets to '' — which showed a
                         hand-written currency code as "General" and would have
                         overwritten it on the next toolbar use. -->
                    <select class="cms-input sheet-editor__format"
                            [disabled]="!activeRef()"
                            (change)="applyFormat($any($event.target).value)">
                        @for (option of formats; track option.label) {
                            <option [value]="option.code ?? ''"
                                    [selected]="(activeFormat() ?? '') === (option.code ?? '')">
                                {{ option.label }}
                            </option>
                        }
                        @if (hasCustomFormat()) {
                            <!-- A code the menu does not list, shown verbatim so an
                                 author who hand-wrote it in the JSON can SEE it
                                 rather than have it read as General and lost on the
                                 next unrelated change. -->
                            <option [value]="activeFormat()" [selected]="true">Custom: {{ activeFormat() }}</option>
                        }
                    </select>
                </label>
                <button class="cms-btn cms-btn-sm"
                        title="Bold"
                        [class.cms-btn-primary]="activeBold()"
                        [disabled]="!activeRef()"
                        (click)="toggleBold()">
                    <i class="bi bi-type-bold"></i>
                </button>
                <button class="cms-btn cms-btn-sm"
                        title="Italic"
                        [class.cms-btn-primary]="activeItalic()"
                        [disabled]="!activeRef()"
                        (click)="toggleItalic()">
                    <i class="bi bi-type-italic"></i>
                </button>
                <label class="sheet-editor__tool">
                    Font
                    <!-- Per-OPTION selection, like Format above and for the same
                         reason: a [value] binding resolves before @for has
                         attached the options and silently falls back to ''. -->
                    <select class="cms-input sheet-editor__font"
                            [disabled]="!activeRef()"
                            (change)="applyStyle('fontFamily', $any($event.target).value)">
                        <option value="" [selected]="!activeStyle('fontFamily')">Default</option>
                        @for (family of fontFamilies(); track family) {
                            <option [value]="family" [selected]="activeStyle('fontFamily') === family">{{ family }}</option>
                        }
                    </select>
                </label>
                <input class="cms-input sheet-editor__size"
                       type="number" min="1" max="409" step="0.5"
                       placeholder="pt"
                       aria-label="Font size in points"
                       [disabled]="!activeRef()"
                       [value]="activeStyle('fontSize')"
                       (change)="applyStyle('fontSize', $any($event.target).value)" />
                <label class="sheet-editor__tool" title="Text colour">
                    <i class="bi bi-fonts"></i>
                    <input class="sheet-editor__swatch"
                           type="color"
                           aria-label="Text colour"
                           [disabled]="!activeRef()"
                           [value]="activeStyle('color') || '#000000'"
                           (change)="applyStyle('color', $any($event.target).value)" />
                </label>
                <label class="sheet-editor__tool" title="Fill colour">
                    <i class="bi bi-paint-bucket"></i>
                    <input class="sheet-editor__swatch"
                           type="color"
                           aria-label="Fill colour"
                           [disabled]="!activeRef()"
                           [value]="activeStyle('background') || '#ffffff'"
                           (change)="applyStyle('background', $any($event.target).value)" />
                </label>
                <!-- Clears BOTH colours. A colour input cannot express "none" —
                     it always reports a colour — so without this an author who
                     shaded a cell could never unshade it. -->
                <button class="cms-btn cms-btn-sm"
                        title="Clear colours"
                        [disabled]="!activeRef()"
                        (click)="clearColours()">
                    <i class="bi bi-eraser"></i>
                </button>
                @for (option of alignments; track option.value) {
                    <button class="cms-btn cms-btn-sm"
                            [title]="option.label"
                            [class.cms-btn-primary]="activeStyle('align') === option.value"
                            [disabled]="!activeRef()"
                            (click)="applyStyle('align', option.value)">
                        <i class="bi" [class]="option.icon"></i>
                    </button>
                }
                <!-- Wrap sits BEFORE the vertical buttons on purpose: it is what
                     makes a row taller than its text, and a vertical alignment
                     has nothing to do until something has. -->
                <button class="cms-btn cms-btn-sm"
                        title="Wrap text"
                        [class.cms-btn-primary]="activeWrap()"
                        [disabled]="!activeRef()"
                        (click)="toggleWrap()">
                    <i class="bi bi-text-wrap"></i>
                </button>
                <!-- A LINK is a text field, not a toggle (#2102): it carries a
                     URL, and often a DTMPL token instead of one, so there is
                     nothing to toggle and nothing to validate here — the
                     backend's writer is the gate before it enters a workbook.
                     Clearing the box removes the link. -->
                <input class="cms-input cms-input-sm sheet-editor__link"
                       type="text"
                       title="Hyperlink — a URL, or a token such as {var:order.url}"
                       placeholder="Link"
                       [disabled]="!activeRef()"
                       [value]="activeStyle('link')"
                       (change)="applyStyle('link', $any($event.target).value)" />
                @for (option of verticalAlignments; track option.value) {
                    <button class="cms-btn cms-btn-sm"
                            [title]="option.label"
                            [class.cms-btn-primary]="activeStyle('valign') === option.value"
                            [disabled]="!activeRef()"
                            (click)="applyStyle('valign', option.value)">
                        <i class="bi" [class]="option.icon"></i>
                    </button>
                }
                <label class="sheet-editor__tool">
                    Width
                    <!-- Bound with [value] rather than per-option like the format
                         select above: this is a plain number input with no
                         dynamically-added children, so nothing can attach after
                         the binding is applied. -->
                    <input class="cms-input sheet-editor__width"
                           type="number" min="1" step="0.5"
                           placeholder="auto"
                           [attr.aria-label]="'Width of column ' + (activeColumn() || '—')"
                           [disabled]="!activeRef()"
                           [value]="activeWidth()"
                           (change)="applyWidth($any($event.target).value)" />
                </label>
                <button class="cms-btn cms-btn-sm"
                        [class.cms-btn-primary]="mergeAtActive()"
                        [disabled]="!canMerge() && !mergeAtActive()"
                        [title]="mergeAtActive() ? 'Unmerge ' + mergeAtActive() : 'Merge ' + (selectionRange() ?? '')"
                        (click)="toggleMerge()">
                    <i class="bi" [class.bi-union]="!mergeAtActive()" [class.bi-subtract]="mergeAtActive()"></i>
                    {{ mergeAtActive() ? 'Unmerge' : 'Merge' }}
                </button>
                <button class="cms-btn cms-btn-sm"
                        [class.cms-btn-primary]="activeFilter()"
                        [disabled]="!canFilter()"
                        [title]="activeFilter()
                            ? 'Remove the filter on ' + activeFilter()
                            : 'Filter ' + (selectionRange() ?? '') + ' — its top row becomes the header'"
                        (click)="toggleFilter()">
                    <i class="bi" [class.bi-funnel]="!activeFilter()" [class.bi-funnel-fill]="activeFilter()"></i>
                    {{ activeFilter() ? 'Remove filter' : 'Filter' }}
                </button>
                <button class="cms-btn cms-btn-sm"
                        [class.cms-btn-primary]="namesOpen()"
                        [title]="'Name ' + (selectionRange() ?? activeRef())
                            + ' so a total can refer to it by name and follow it when rows are generated'"
                        (click)="toggleNames()">
                    <i class="bi bi-bookmark"></i>
                    Names
                </button>
                <button class="cms-btn cms-btn-sm"
                        type="button"
                        title="Colour cells by what is in them"
                        aria-label="Conditional formatting"
                        [class.cms-btn-primary]="rulesOpen()"
                        [disabled]="!activeRef()"
                        (click)="toggleRules()">
                    <i class="bi bi-palette"></i>
                </button>
                <button class="cms-btn cms-btn-sm"
                        type="button"
                        title="Find and replace (Ctrl+F)"
                        aria-label="Find and replace"
                        [class.cms-btn-primary]="findOpen()"
                        (click)="toggleFind()">
                    <i class="bi bi-search"></i>
                </button>
                <button class="cms-btn cms-btn-sm"
                        type="button"
                        [class.cms-btn-primary]="isFrozen()"
                        [disabled]="!activeRef() && !isFrozen()"
                        [title]="isFrozen()
                            ? 'Unfreeze the panes'
                            : 'Freeze the rows above and the columns left of ' + (activeRef() || 'the selected cell')"
                        (click)="toggleFreeze()">
                    <i class="bi" [class.bi-pin-angle]="!isFrozen()" [class.bi-pin-angle-fill]="isFrozen()"></i>
                    {{ isFrozen() ? 'Unfreeze' : 'Freeze' }}
                </button>
                <!-- A button and a palette rather than a <select>, for two
                     reasons reported within a minute of each other: a list of
                     words cannot DRAW a border, and a control that resets to
                     its placeholder cannot say which one the cells already
                     carry. Its own mousedown is swallowed so opening the
                     palette does not read as a click outside it. -->
                <div class="sheet-editor__tool sheet-editor__borders"
                     (mousedown)="$any($event).stopPropagation()">
                    <button class="cms-btn cms-btn-sm"
                            type="button"
                            title="Rule the edges of the selected cells"
                            aria-label="Borders"
                            [attr.aria-expanded]="borderMenuOpen()"
                            [class.cms-btn-primary]="borderMenuOpen()"
                            [disabled]="!activeRef()"
                            (click)="toggleBorderMenu()">
                        <i class="bi bi-border-all"></i>
                        Borders
                        <i class="bi bi-caret-down-fill"></i>
                    </button>
                    @if (borderMenuOpen()) {
                        <div class="sheet-editor__border-menu"
                             role="group"
                             aria-label="Borders"
                             (keydown)="onBorderMenuKey($event)">
                            <!-- Pressed marks what the SELECTION already has,
                                 which is the half a select could not show. -->
                            <div class="sheet-editor__border-row">
                                @for (shape of borderShapes; track shape.value) {
                                    <button class="cms-btn cms-btn-sm"
                                            type="button"
                                            [title]="shape.label"
                                            [attr.aria-label]="shape.label"
                                            [attr.aria-pressed]="borderState().preset === shape.value"
                                            [class.cms-btn-primary]="borderState().preset === shape.value"
                                            (click)="applyBorders(shape.value)">
                                        <i [class]="'bi ' + shape.icon"></i>
                                    </button>
                                }
                            </div>
                            <div class="sheet-editor__border-row">
                                @for (side of borderSides; track side.value) {
                                    <button class="cms-btn cms-btn-sm"
                                            type="button"
                                            [title]="side.label"
                                            [attr.aria-label]="side.label"
                                            [attr.aria-pressed]="borderState().preset === side.value"
                                            [class.cms-btn-primary]="borderState().preset === side.value"
                                            (click)="applyBorders(side.value)">
                                        <i [class]="'bi ' + side.icon"></i>
                                    </button>
                                }
                            </div>
                            <label class="sheet-editor__tool">
                                Line
                                <!-- Per-OPTION selection, not [value] on the
                                     select: the same lesson the Format control
                                     records one screen up. -->
                                <select class="cms-input sheet-editor__insert"
                                        aria-label="Border line"
                                        (change)="borderStyle.set($any($event.target).value)">
                                    @for (line of borderLines; track line.value) {
                                        <option [value]="line.value"
                                                [selected]="line.value === borderStyle()">{{ line.label }}</option>
                                    }
                                </select>
                            </label>
                            <label class="sheet-editor__tool">
                                Colour
                                <input class="sheet-editor__swatch"
                                       type="color"
                                       aria-label="Border colour"
                                       [value]="borderColour()"
                                       (change)="borderColour.set($any($event.target).value)" />
                            </label>
                        </div>
                    }
                </div>
                <label class="sheet-editor__tool" title="Insert a form element into the selected cells">
                    Form
                    <!-- Resets to '' after every choice, so it reads "Form
                         element" rather than showing a stale pick as if it
                         were the cell's current state. -->
                    <select class="cms-input sheet-editor__insert"
                            aria-label="Insert form element"
                            [disabled]="!activeRef()"
                            (change)="insertControl($any($event.target).value); $any($event.target).value = ''">
                        <option value="" selected>Form element…</option>
                        <option value="list">Dropdown…</option>
                        <option value="checkbox">Checkbox</option>
                        @if (validationAtActive()) {
                            <option value="remove">Remove</option>
                        }
                    </select>
                </label>
                <!-- The catalogue, browsable. Type-ahead only helps an
                     author who already knows the name; SUMIF cannot be
                     discovered by typing SUMIF. Same data as the helper popup
                     -- functionsByCategory() derives from the one list -- so
                     the two can never offer different functions. -->
                <div class="sheet-editor__tool sheet-editor__functions"
                     (mousedown)="$any($event).stopPropagation()">
                    <button class="cms-btn cms-btn-sm"
                            type="button"
                            title="Insert a function"
                            aria-label="Functions"
                            [attr.aria-expanded]="functionMenuOpen()"
                            [class.cms-btn-primary]="functionMenuOpen()"
                            [disabled]="!activeRef()"
                            (click)="toggleFunctionMenu()">
                        <span class="sheet-editor__sigma">&#931;</span>
                        <i class="bi bi-caret-down-fill"></i>
                    </button>
                    @if (functionMenuOpen()) {
                        <div class="sheet-editor__function-menu"
                             role="group"
                             aria-label="Functions"
                             (keydown)="onFunctionMenuKey($event)">
                            <input class="cms-input sheet-editor__function-search"
                                   type="search"
                                   autocomplete="off"
                                   placeholder="Search functions"
                                   aria-label="Search functions"
                                   [value]="functionQuery()"
                                   (input)="functionQuery.set($any($event.target).value)" />
                            <div class="sheet-editor__function-list">
                                @for (shelf of functionShelves(); track shelf.category) {
                                    <div class="sheet-editor__function-shelf">{{ shelf.category }}</div>
                                    @for (fn of shelf.functions; track fn.name) {
                                        <button class="sheet-editor__option sheet-editor__function"
                                                type="button"
                                                [title]="fn.signature"
                                                (click)="insertFunction(fn)">
                                            <span class="sheet-editor__function-name">{{ fn.name }}</span>
                                            <span class="sheet-editor__function-summary">{{ fn.summary }}</span>
                                        </button>
                                    }
                                }
                                @if (functionShelves().length === 0) {
                                    <div class="sheet-editor__function-empty">No function matches “{{ functionQuery() }}”.</div>
                                }
                            </div>
                        </div>
                    }
                </div>
                @if (hiddenRowCount() > 0) {
                    <span class="sheet-editor__tool-note">{{ hiddenRowCount() }} row(s) hidden while editing — not saved</span>
                }
                <span class="sheet-editor__hint">Click a header to select a column or row, or drag its edge to resize. Shift-click extends. Text (&#64;) keeps a value from becoming a number.</span>
            </div>

            @if (rulesOpen()) {
                <!-- Beside the find box and above the grid, for the same reason:
                     a panel that covers the cells it is about cannot be read
                     against them. -->
                <div class="sheet-editor__rules" role="group" aria-label="Conditional formatting"
                     (keydown)="onRulesKey($event)">
                    <div class="sheet-editor__rules-head">
                        Colour {{ selectionRange() ?? activeRef() }} when it is
                    </div>
                    <div class="sheet-editor__rules-row">
                        <select class="cms-input" aria-label="Condition"
                                (change)="ruleWhen.set($any($event.target).value)">
                            @for (option of conditionalWhens; track option.value) {
                                <option [value]="option.value"
                                        [selected]="option.value === ruleWhen()">{{ option.label }}</option>
                            }
                        </select>
                        @if (ruleNeedsValue()) {
                            <input class="cms-input sheet-editor__rules-value" type="text"
                                   aria-label="Condition value" placeholder="Value"
                                   [value]="ruleValue()"
                                   (input)="ruleValue.set($any($event.target).value)" />
                        }
                        @if (ruleNeedsSecondValue()) {
                            <input class="cms-input sheet-editor__rules-value" type="text"
                                   aria-label="Condition second value" placeholder="and"
                                   [value]="ruleValue2()"
                                   (input)="ruleValue2.set($any($event.target).value)" />
                        }
                        <label class="sheet-editor__find-option">
                            Fill
                            <input class="sheet-editor__swatch" type="color" aria-label="Rule fill"
                                   [value]="ruleBackground()"
                                   (change)="ruleBackground.set($any($event.target).value)" />
                        </label>
                        <label class="sheet-editor__find-option">
                            Text
                            <input class="sheet-editor__swatch" type="color" aria-label="Rule text colour"
                                   [value]="ruleColor()"
                                   (change)="ruleColor.set($any($event.target).value)" />
                        </label>
                        <button class="cms-btn cms-btn-sm"
                                type="button"
                                [class.cms-btn-primary]="ruleBold()"
                                aria-label="Rule bold"
                                (click)="ruleBold.set(!ruleBold())"><b>B</b></button>
                        <button class="cms-btn cms-btn-sm"
                                type="button"
                                [class.cms-btn-primary]="ruleItalic()"
                                aria-label="Rule italic"
                                (click)="ruleItalic.set(!ruleItalic())"><i>I</i></button>
                        <button class="cms-btn cms-btn-sm cms-btn-primary" type="button"
                                (click)="addRule()">Add rule</button>
                        <button class="cms-btn cms-btn-sm" type="button"
                                aria-label="Close conditional formatting"
                                (click)="closeRules()"><i class="bi bi-x-lg"></i></button>
                    </div>
                    <!-- What is already on this cell, because a rule nobody can
                         see is a rule nobody can remove. -->
                    @for (existing of rulesHere(); track existing.range + existing.index) {
                        <div class="sheet-editor__rules-existing">
                            <span class="sheet-editor__rules-swatch"
                                  [style.background]="existing.rule.background ?? 'transparent'"
                                  [style.color]="existing.rule.color ?? 'inherit'">Aa</span>
                            <span>{{ existing.range }} — {{ describeRule(existing.rule) }}</span>
                            <button class="cms-btn cms-btn-sm" type="button"
                                    [attr.aria-label]="'Remove rule on ' + existing.range"
                                    (click)="removeRule(existing.range, existing.index)">
                                <i class="bi bi-trash"></i>
                            </button>
                        </div>
                    }
                </div>
            }

            @if (namesOpen()) {
                <!-- A name is the only reference into a generated band that
                     survives every editor, so this is the thing to reach for
                     when a total has to cover rows that do not exist yet. -->
                <div class="sheet-editor__names">
                    <div class="sheet-editor__names-add">
                        <input class="cms-input sheet-editor__names-input"
                               type="text"
                               autocomplete="off"
                               placeholder="Name for {{ selectionRange() ?? activeRef() }}"
                               aria-label="New name"
                               [value]="newName()"
                               (input)="newName.set($any($event.target).value)"
                               (keydown.enter)="addName()" />
                        <button class="cms-btn cms-btn-sm cms-btn-primary"
                                type="button"
                                [disabled]="null !== nameError()"
                                (click)="addName()">Name {{ selectionRange() ?? activeRef() }}</button>
                        <button class="cms-btn cms-btn-sm" type="button"
                                aria-label="Close names"
                                (click)="closeNames()"><i class="bi bi-x-lg"></i></button>
                    </div>
                    @if ('' !== newName() && null !== nameError()) {
                        <p class="sheet-editor__names-error" role="alert">{{ nameError() }}</p>
                    }
                    <!-- What is already declared, because a name nobody can see
                         is a name nobody can remove. -->
                    @for (declared of definedNames(); track declared.name) {
                        <div class="sheet-editor__names-row">
                            <code>{{ declared.name }}</code>
                            <span>{{ declared.range }}</span>
                            <button class="cms-btn cms-btn-sm" type="button"
                                    [attr.aria-label]="'Remove the name ' + declared.name"
                                    (click)="removeName(declared.name)">
                                <i class="bi bi-trash"></i>
                            </button>
                        </div>
                    } @empty {
                        <p class="sheet-editor__names-empty">
                            No names yet. Select the band column and name it, then a total can say
                            SUM(that_name) and keep covering it however many rows are generated.
                        </p>
                    }
                </div>
            }

            @if (findOpen()) {
                <!-- Above the grid rather than floating over it: a find box that
                     hides the cell it just jumped to is a find box that has to
                     be moved before it can be read. -->
                <div class="sheet-editor__find" role="search" (keydown)="onFindKey($event)">
                    <input class="cms-input sheet-editor__find-input"
                           type="search"
                           autocomplete="off"
                           placeholder="Find in this sheet"
                           aria-label="Find"
                           [value]="findQuery()"
                           (input)="findQuery.set($any($event.target).value)" />
                    <input class="cms-input sheet-editor__find-input"
                           type="text"
                           autocomplete="off"
                           placeholder="Replace with"
                           aria-label="Replace with"
                           [value]="replaceWith()"
                           (input)="replaceWith.set($any($event.target).value)" />
                    <label class="sheet-editor__find-option">
                        <input type="checkbox" aria-label="Match case"
                               [checked]="matchCase()"
                               (change)="matchCase.set($any($event.target).checked)" />
                        Match case
                    </label>
                    <label class="sheet-editor__find-option">
                        <input type="checkbox" aria-label="Whole cell"
                               [checked]="wholeCell()"
                               (change)="wholeCell.set($any($event.target).checked)" />
                        Whole cell
                    </label>
                    <button class="cms-btn cms-btn-sm" type="button"
                            [disabled]="findCount() === 0"
                            (click)="findNext()">Next</button>
                    <button class="cms-btn cms-btn-sm" type="button"
                            [disabled]="findCount() === 0"
                            (click)="replaceOne()">Replace</button>
                    <button class="cms-btn cms-btn-sm" type="button"
                            [disabled]="findCount() === 0"
                            (click)="replaceAll()">Replace all</button>
                    <span class="sheet-editor__find-count">{{ findSummary() }}</span>
                    <button class="cms-btn cms-btn-sm" type="button"
                            aria-label="Close find"
                            (click)="closeFind()"><i class="bi bi-x-lg"></i></button>
                </div>
            }

            <!-- Any mousedown in the grid dismisses an open filter menu. The
                 menu itself stops propagation, so a click INSIDE it survives —
                 the same guard the formula helper uses one cell over. -->
            <div class="sheet-editor__body"
                 (scroll)="onGridScroll($event)"
                 (mousedown)="closeFilterMenu(); closeOptions(); closeContextMenu(); closeBorderMenu(); closeFunctionMenu()"
                 (copy)="onCopy($any($event), false)"
                 (cut)="onCopy($any($event), true)"
                 (paste)="onPaste($any($event))">
                @if (loading()) {
                    <cms-loader label="Opening the document" />
                } @else {
                    <table class="sheet-editor__grid"
                           [class.sheet-editor__grid--dragging]="dragging()">
                        <!-- Spacer COLS carry the width of the columns outside
                             the window, so the table stays as wide as the sheet
                             while only the visible slice is built. -->
                        <colgroup>
                            <col class="sheet-editor__row-head-col" />
                            @for (col of frozenColumns(); track col) { <col [style.width.px]="columnPx(col)" /> }
                            @if (leftSpacerPx() > 0) { <col [style.width.px]="leftSpacerPx()" /> }
                            @for (col of visibleColumns(); track col) {
                                <col [style.width.px]="columnPx(col)" />
                            }
                            @if (rightSpacerPx() > 0) { <col [style.width.px]="rightSpacerPx()" /> }
                        </colgroup>
                        <thead>
                            <tr>
                                <th class="sheet-editor__corner"
                                    role="button"
                                    aria-label="Select all cells"
                                    (mousedown)="selectAll($any($event))"></th>
                                @for (col of frozenColumns(); track col) {
                                    <th class="sheet-editor__col-head sheet-editor__col-head--frozen"
                                        role="button"
                                        [style.left.px]="frozenLeftPx(col)"
                                        [class.sheet-editor__head--selected]="isColumnSelected(col)"
                                        [attr.aria-label]="'Select column ' + col"
                                        (mousedown)="selectColumn(col, $any($event))">{{ col }}<!--
                                        --><span class="sheet-editor__grip"
                                              role="separator"
                                              [attr.aria-label]="'Resize column ' + col"
                                              (mousedown)="startResize(col, $any($event))"></span></th>
                                }
                                @if (leftSpacerPx() > 0) { <th class="sheet-editor__spacer-head"></th> }
                                @for (col of visibleColumns(); track col) {
                                    <th class="sheet-editor__col-head"
                                        role="button"
                                        [class.sheet-editor__head--selected]="isColumnSelected(col)"
                                        [attr.aria-label]="'Select column ' + col"
                                        (mousedown)="selectColumn(col, $any($event))">{{ col }}<!--
                                        --><span class="sheet-editor__grip"
                                              role="separator"
                                              [attr.aria-label]="'Resize column ' + col"
                                              (mousedown)="startResize(col, $any($event))"></span></th>
                                }
                                @if (rightSpacerPx() > 0) { <th class="sheet-editor__spacer-head"></th> }
                            </tr>
                        </thead>
                        <tbody>
                            <!-- The rows above the window, as one box of the
                                 height they would have occupied. This is what
                                 keeps the scrollbar honest while only the
                                 visible slice exists in the DOM. -->
                            <!-- ONE definition of a cell, instantiated by both axes. A frozen
                                 column and a windowed one differ only in where they STICK, and a
                                 second copy of a hundred lines of cell markup is a second copy to
                                 keep in step with the first. -->
                            <ng-template #cellTpl let-col let-row="row">
                                        <!-- A cell swallowed by a merge emits NO <td> at all. Rendering
                                             one and hiding it with CSS would leave a focusable input
                                             inside a region that does not exist in the workbook. -->
                                        @if (!isCovered(col + row)) {
                                        <td class="sheet-editor__cell"
                                            (contextmenu)="openContextMenu(col + row, $any($event))"
                                            [class.sheet-editor__cell--bold]="isBold(col + row)"
                                            [class.sheet-editor__cell--formula]="isFormula(col + row)"
                                            [class.sheet-editor__cell--error]="isFormulaError(col + row)"
                                            [class.sheet-editor__cell--unresolved]="isUnresolvedFormula(col + row)"
                                            [attr.title]="titleAt(col + row)"
                                            [class.sheet-editor__cell--merged]="isMerged(col + row)"
                                            [class.sheet-editor__cell--in-range]="isInSelection(col + row)"
                                            [class.sheet-editor__cell--active]="activeRef() === col + row"
                                            [class.sheet-editor__cell--match]="isMatch(col + row)"
                                            [class.sheet-editor__cell--frozen-row]="isFrozenRow(row)"
                                            [class.sheet-editor__cell--frozen-col]="isFrozenColumn(col)"
                                            [style.top.px]="isFrozenRow(row) ? frozenTopPx(row) : null"
                                            [style.left.px]="isFrozenColumn(col) ? frozenLeftPx(col) : null"
                                            [attr.colspan]="spanAt(col + row).colspan"
                                            [attr.rowspan]="spanAt(col + row).rowspan"
                                            [style.outline-offset.px]="-2"
                                            [ngStyle]="cellStyleAt(col + row)">
                                            <!-- ONE call per cell, not one per
                                                 property: six separate
                                                 [style.x] bindings each
                                                 re-read the cell and built a
                                                 fresh object every change
                                                 detection, which multiplies by
                                                 every rendered cell. -->
                                            <!-- A checkbox cell draws the CONTROL rather than
                                                 the text, because "TRUE" spelled out in a cell
                                                 is not what an author drew a tick box for. The
                                                 value underneath is still TRUE/FALSE, which is
                                                 exactly what the workbook carries. -->
                                            @if (isCheckbox(col + row)) {
                                            <input class="sheet-editor__checkbox"
                                                   type="checkbox"
                                                   [attr.aria-label]="col + row"
                                                   [checked]="isChecked(col + row)"
                                                   (focus)="focusCell(col + row)"
                                                   (change)="toggleCheckbox(col + row, $any($event.target).checked)" />
                                            } @else {
                                            <!-- A wrapped cell is a TEXTAREA,
                                                 because an <input> is
                                                 single-line by definition and
                                                 cannot show a wrap however it
                                                 is styled. Same bindings, so
                                                 the two differ only in whether
                                                 the text can break. -->
                                            @if (wrapAt(col + row)) {
                                            <textarea class="sheet-editor__input sheet-editor__input--wrap"
                                                      [attr.aria-label]="col + row"
                                                      [style]="styleAt(col + row)"
                                                      [value]="valueAt(col + row)"
                                                      (focus)="focusCell(col + row, $any($event.target))"
                                                      (input)="onCellInput($any($event.target))"
                                                      (keyup)="trackCaret($any($event.target))"
                                                      (click)="trackCaret($any($event.target))"
                                                      (keydown)="onCellKey(col + row, $any($event))"
                                                      (blur)="commitCell(col + row, $any($event.target))"
                                                      (change)="commitCell(col + row, $any($event.target))"
                                                      (mousedown)="extendTo(col + row, $any($event))"></textarea>
                                            } @else {
                                            <input class="sheet-editor__input"
                                                   [attr.aria-label]="col + row"
                                                   [style]="styleAt(col + row)"
                                                   [value]="valueAt(col + row)"
                                                   (focus)="focusCell(col + row, $any($event.target))"
                                                   (mousedown)="extendTo(col + row, $any($event))"
                                                   (input)="onCellInput($any($event.target))"
                                                   (keyup)="trackCaret($any($event.target))"
                                                   (click)="trackCaret($any($event.target))"
                                                   (keydown)="onCellKey(col + row, $any($event))"
                                                   (blur)="commitCell(col + row, $any($event.target))"
                                                   (change)="commitCell(col + row, $any($event.target))" />
                                            }
                                            }
                                            <!-- A dropdown cell keeps its text input — an author
                                                 may still need to type a token into it — and gains
                                                 an arrow that offers the allowed values. -->
                                            @if (hasOptions(col + row)) {
                                                <button type="button"
                                                        class="sheet-editor__options-btn"
                                                        [attr.aria-label]="'Options for ' + col + row"
                                                        [attr.aria-expanded]="openOptionsRef() === col + row"
                                                        (mousedown)="toggleOptions(col + row, $any($event))">
                                                    <i class="bi bi-caret-down-fill"></i>
                                                </button>
                                                @if (openOptionsRef() === col + row) {
                                                    <div class="sheet-editor__filter-menu"
                                                         role="listbox"
                                                         [attr.aria-label]="'Options for ' + col + row"
                                                         (mousedown)="$any($event).stopPropagation()">
                                                        <div class="sheet-editor__filter-list">
                                                            @for (option of optionsAt(col + row); track option) {
                                                                <button type="button"
                                                                        role="option"
                                                                        class="sheet-editor__option"
                                                                        [attr.aria-selected]="valueAt(col + row) === option"
                                                                        (click)="chooseOption(col + row, option)">{{ option }}</button>
                                                            }
                                                        </div>
                                                    </div>
                                                }
                                            }
                                            <!-- Anchored in the cell like the other two popups, so
                                                 it travels with the grid instead of needing a
                                                 bounding rect re-measured on every scroll. -->
                                            @if (contextRef() === col + row) {
                                                <div class="sheet-editor__context"
                                                     [class.sheet-editor__context--above]="contextAbove()"
                                                     role="menu"
                                                     [attr.aria-label]="'Actions for ' + col + row"
                                                     (mousedown)="$any($event).stopPropagation()">
                                                    <!-- Every item carries its glyph: the rest of this
                                                         editor's controls do, and a menu of bare labels
                                                         beside a toolbar of icons reads as unfinished.
                                                         The two deletes share the trash glyph on
                                                         purpose -- the LABEL says which, and inventing
                                                         a row-vs-column icon the set does not have
                                                         would say less, not more.
                                                         NO BACKTICKS in here: this comment sits inside
                                                         the template literal, so one ENDS it, and the
                                                         error surfaces hundreds of lines away. -->
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="insertRow('above')">
                                                        <i class="bi bi-arrow-bar-up"></i>Insert row above
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="insertRow('below')">
                                                        <i class="bi bi-arrow-bar-down"></i>Insert row below
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="insertColumn('left')">
                                                        <i class="bi bi-arrow-bar-left"></i>Insert column left
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="insertColumn('right')">
                                                        <i class="bi bi-arrow-bar-right"></i>Insert column right
                                                    </button>
                                                    <hr class="sheet-editor__context-rule" />
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="deleteRow()">
                                                        <i class="bi bi-trash"></i>Delete row {{ contextRowLabel() }}
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="deleteColumn()">
                                                        <i class="bi bi-trash"></i>Delete column {{ contextColumnLabel() }}
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="clearContents()">
                                                        <i class="bi bi-eraser"></i>Clear contents
                                                    </button>
                                                    <hr class="sheet-editor__context-rule" />
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            [disabled]="!canMerge() && !mergeAtActive()"
                                                            (click)="toggleMerge(); closeContextMenu()">
                                                        <i class="bi" [class.bi-union]="!mergeAtActive()"
                                                           [class.bi-subtract]="mergeAtActive()"></i>{{ mergeAtActive() ? 'Unmerge' : 'Merge cells' }}
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            [disabled]="!canFilter()"
                                                            (click)="toggleFilter(); closeContextMenu()">
                                                        <i class="bi" [class.bi-funnel]="!activeFilter()"
                                                           [class.bi-funnel-fill]="activeFilter()"></i>{{ activeFilter() ? 'Remove filter' : 'Create a filter' }}
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="closeContextMenu(); insertControl('list')">
                                                        <i class="bi bi-caret-down-square"></i>Dropdown…
                                                    </button>
                                                    <button type="button" role="menuitem" class="sheet-editor__option"
                                                            (click)="closeContextMenu(); insertControl('checkbox')">
                                                        <i class="bi bi-check-square"></i>Checkbox
                                                    </button>
                                                    @if (validationAtActive()) {
                                                        <button type="button" role="menuitem" class="sheet-editor__option"
                                                                (click)="closeContextMenu(); insertControl('remove')">
                                                            <i class="bi bi-x-square"></i>Remove form element
                                                        </button>
                                                    }
                                                </div>
                                            }
                                            <!-- The helper lives INSIDE the cell so it travels with it
                                                 when the grid scrolls. Positioning it from a bounding
                                                 rect would need re-measuring on every scroll and would
                                                 still drift during one. -->
                                            @if (activeRef() === col + row) {
                                                @if (helperState(); as state) {
                                                    <div class="sheet-editor__helper"
                                                         (mousedown)="$any($event).preventDefault()">
                                                        @if (state.kind === 'completions') {
                                                            @for (fn of state.matches; track fn.name; let i = $index) {
                                                                <button type="button"
                                                                        class="sheet-editor__helper-item"
                                                                        [class.sheet-editor__helper-item--on]="i === helperIndex()"
                                                                        (click)="acceptCompletion(fn)">
                                                                    <span class="sheet-editor__helper-name">{{ fn.name }}</span>
                                                                    <span class="sheet-editor__helper-hint">{{ fn.summary }}</span>
                                                                </button>
                                                            }
                                                        } @else {
                                                            <div class="sheet-editor__helper-sig">
                                                                <span class="sheet-editor__helper-name">{{ state.fn.name }}</span>(@for (part of signatureOf(state.fn); track part; let i = $index) {<span [class.sheet-editor__helper-arg--on]="isCurrentArgument(state, i)">{{ i > 0 ? ', ' : '' }}{{ part }}</span>})
                                                            </div>
                                                            <div class="sheet-editor__helper-hint">{{ state.fn.summary }}</div>
                                                        }
                                                    </div>
                                                }
                                            }
                                            <!-- The filter dropdown, anchored inside the header cell
                                                 for the same reason as the helper above: it travels
                                                 with the cell instead of being re-measured on scroll. -->
                                            @if (hasFilterButton(col + row)) {
                                                <button type="button"
                                                        class="sheet-editor__filter-btn"
                                                        [class.sheet-editor__filter-btn--on]="isColumnFiltered(col)"
                                                        [attr.aria-label]="'Filter column ' + col"
                                                        [attr.aria-expanded]="openFilterColumn() === col"
                                                        (mousedown)="toggleFilterMenu(col, $any($event))">
                                                    <i class="bi bi-funnel-fill"></i>
                                                </button>
                                                @if (openFilterColumn() === col) {
                                                    <div class="sheet-editor__filter-menu"
                                                         role="dialog"
                                                         [attr.aria-label]="'Filter column ' + col"
                                                         (mousedown)="$any($event).stopPropagation()">
                                                        <div class="sheet-editor__filter-head">
                                                            <span class="sheet-editor__filter-title">Show rows where {{ col }} is</span>
                                                            <button type="button" class="cms-btn cms-btn-sm"
                                                                    [disabled]="!isColumnFiltered(col)"
                                                                    (click)="showAllInColumn(col)">Select all</button>
                                                        </div>
                                                        <div class="sheet-editor__filter-list">
                                                            @for (option of filterValues(col); track option.value) {
                                                                <label class="sheet-editor__filter-option">
                                                                    <input type="checkbox"
                                                                           [checked]="option.shown"
                                                                           (change)="toggleFilterValue(col, option.value)" />
                                                                    <span class="sheet-editor__filter-value">{{ option.label }}</span>
                                                                </label>
                                                            } @empty {
                                                                <p class="sheet-editor__filter-empty">
                                                                    Nothing under the header yet. A filter over a
                                                                    <code>&#123;loop:&#125;</code> band still covers every row it generates.
                                                                </p>
                                                            }
                                                        </div>
                                                        <div class="sheet-editor__filter-foot">
                                                            <button type="button" class="cms-btn cms-btn-sm"
                                                                    (click)="closeFilterMenu()">Done</button>
                                                        </div>
                                                    </div>
                                                }
                                            }
                                        </td>
                                        }
                            </ng-template>

                            <ng-template #rowTpl let-row>
                                <tr [style.height.px]="rowPx(row)">
                                    <th class="sheet-editor__row-head"
                                        role="button"
                                        [class.sheet-editor__row-head--frozen]="isFrozenRow(row)"
                                        [style.top.px]="isFrozenRow(row) ? frozenTopPx(row) : null"
                                        [class.sheet-editor__head--selected]="isRowSelected(row)"
                                        [attr.aria-label]="'Select row ' + row"
                                        (mousedown)="selectRow(row, $any($event))">{{ row }}<!--
                                        --><span class="sheet-editor__grip sheet-editor__grip--row"
                                              role="separator"
                                              [attr.aria-label]="'Resize row ' + row"
                                              (mousedown)="startRowResize(row, $any($event))"></span></th>
                                    @for (col of frozenColumns(); track col) {
                                        <ng-container *ngTemplateOutlet="cellTpl; context: { $implicit: col, row: row }" />
                                    }
                                    @if (leftSpacerPx() > 0) { <td class="sheet-editor__spacer-cell"></td> }
                                    @for (col of visibleColumns(); track col) {
                                        <ng-container *ngTemplateOutlet="cellTpl; context: { $implicit: col, row: row }" />
                                    }
                                    @if (rightSpacerPx() > 0) { <td class="sheet-editor__spacer-cell"></td> }
                                </tr>
                            </ng-template>

                            <!-- Frozen rows come BEFORE the spacer: they are always rendered, and the
                                 spacer stands in only for what is neither frozen nor in the window. -->
                            @for (row of frozenRows(); track row) {
                                <ng-container *ngTemplateOutlet="rowTpl; context: { $implicit: row }" />
                            }
                            @if (topSpacerPx() > 0) {
                                <tr class="sheet-editor__spacer" [style.height.px]="topSpacerPx()">
                                    <td [attr.colspan]="gridColspan()"></td>
                                </tr>
                            }
                            @for (row of visibleRows(); track row) {
                                <ng-container *ngTemplateOutlet="rowTpl; context: { $implicit: row }" />
                            }
                            @if (bottomSpacerPx() > 0) {
                                <tr class="sheet-editor__spacer" [style.height.px]="bottomSpacerPx()">
                                    <td [attr.colspan]="gridColspan()"></td>
                                </tr>
                            }
                        </tbody>
                    </table>
                }
            </div>

            <div class="sheet-editor__footer">
                <div class="sheet-editor__foot-left">
                    <span class="sheet-editor__status">{{ statusText() }}</span>
                    <!-- Zoom belongs in the FOOTER, not the toolbar: the toolbar
                         is the surface that already wrapped once when six font
                         controls were added, and bottom-right is where every
                         office application puts a zoom readout. Ctrl + wheel over
                         the grid does the same thing -- this is the visible half
                         of it, and the way back to 100% without counting notches. -->
                    <div class="sheet-editor__zoom">
                        <button class="cms-btn cms-btn-sm" type="button"
                                title="Zoom out" aria-label="Zoom out"
                                [disabled]="zoomPercent() <= 50"
                                (click)="zoomOut()"><i class="bi bi-dash-lg"></i></button>
                        <button class="cms-btn cms-btn-sm sheet-editor__zoom-value" type="button"
                                title="Reset zoom to 100%" aria-label="Zoom"
                                [disabled]="100 === zoomPercent()"
                                (click)="resetZoom()">{{ zoomPercent() }}%</button>
                        <button class="cms-btn cms-btn-sm" type="button"
                                title="Zoom in" aria-label="Zoom in"
                                [disabled]="zoomPercent() >= 200"
                                (click)="zoomIn()"><i class="bi bi-plus-lg"></i></button>
                    </div>
                </div>
                <div class="d-flex gap-2">
                    <button class="cms-btn cms-btn-sm" (click)="close()">Cancel</button>
                    <button class="cms-btn cms-btn-primary cms-btn-sm"
                            [disabled]="saving() || !dirty()"
                            (click)="save()">
                        {{ saving() ? 'Saving…' : 'Save' }}
                    </button>
                </div>
            </div>
        </div>
    `,
    styles: [`
        .sheet-editor {
            display: flex; flex-direction: column;
            width: min(92vw, 1200px);
            height: min(85vh, 800px);
            background: var(--cms-surface);
            border-radius: var(--cms-radius-lg);
            overflow: hidden;
        }
        /* Full screen means the VIEWPORT, not the OS. A grid is the surface
           that most wants the room -- 1200px shows about eight columns of a
           sheet that has twenty-six -- and this is the same thing the office
           editors mean by it.
           Deliberately NOT the Fullscreen API: that needs a fresh user gesture
           every time, can be refused by policy, and would take the dialog out
           of the CDK overlay it lives in. Sizing the dialog is what an author
           actually wants and it cannot fail. The 100% here beats the min()
           above by specificity, so nothing needs !important. */
        .sheet-editor--full {
            width: 100vw; height: 100vh;
            max-width: 100vw; max-height: 100vh;
            border-radius: 0;
        }
        .sheet-editor__header, .sheet-editor__footer {
            display: flex; align-items: center; justify-content: space-between;
            padding: 10px 16px; flex-shrink: 0;
        }
        .sheet-editor__header { border-bottom: 1px solid var(--cms-border); }
        .sheet-editor__footer { border-top: 1px solid var(--cms-border); font-size: .8125rem; }
        .sheet-editor__title {
            font-size: .875rem; font-weight: 600;
            display: flex; align-items: center; gap: 8px;
        }
        .sheet-editor__sheet-pick { width: auto; padding: 2px 6px; font-size: .8125rem; }
        .sheet-editor__actions { display: flex; align-items: center; gap: 8px; }
        .sheet-editor__dirty { color: var(--cms-warning); font-size: .75rem; }
        .sheet-editor__status { color: var(--cms-text-muted); }
        .sheet-editor__foot-left { display: flex; align-items: center; gap: 12px; }
        .sheet-editor__zoom { display: flex; align-items: center; gap: 2px; }
        /* Fixed width and tabular figures so the footer does not shuffle
           sideways as the number goes 100% -> 95% -> 110% under the pointer. */
        .sheet-editor__zoom-value { min-width: 4rem; font-variant-numeric: tabular-nums; }
        .sheet-editor__note { padding: 16px; color: var(--cms-text-muted); }

        /* flex-wrap is load-bearing. Without it nothing wraps, so the LAST
           item absorbs every pixel of overflow — the hint was squeezed into a
           sliver and wrapped into eight lines, making the toolbar ~250px tall
           with the controls stranded in the middle of it. Adding six font
           controls is what tipped it over. */
        .sheet-editor__toolbar {
            display: flex; align-items: center; gap: 12px; flex-shrink: 0;
            flex-wrap: wrap;
            padding: 6px 16px;
            border-bottom: 1px solid var(--cms-border);
            font-size: .8125rem;
        }
        /* Controls keep their size and wrap whole rather than being squashed
           one at a time — the same rule the editor toolbar follows (#2061). */
        .sheet-editor__toolbar > *:not(.sheet-editor__hint) { flex-shrink: 0; }
        .sheet-editor__active {
            font-weight: 700; min-width: 3.5em;
            font-family: 'Courier New', monospace;
        }
        .sheet-editor__tool { display: flex; align-items: center; gap: 6px; }
        .sheet-editor__format { width: auto; padding: 2px 6px; font-size: .8125rem; }
        .sheet-editor__width { width: 5.5rem; padding: 2px 6px; font-size: .8125rem; }
        /* Allowed to shrink and to wrap to its own row, with a floor so it
           never becomes the vertical sliver described above. It is guidance,
           so it yields space to the controls rather than taking it. */
        .sheet-editor__hint {
            color: var(--cms-text-muted);
            margin-left: auto;
            flex: 1 1 16rem;
            min-width: 16rem;
        }

        /* min-height:0 lets the flex item shrink so the grid scrolls inside it
           rather than pushing the dialog taller than the viewport. */
        .sheet-editor__body { flex: 1; overflow: auto; min-height: 0; }

        /* table-layout:fixed is what makes a stored column width BINDING. Under
           the default auto layout a col width is only a suggestion and the
           browser sizes to content instead: measured, a column set to 28
           characters (201px) laid out at 141px, so the author saw no change and
           would reasonably conclude the width had not saved. */
        /* width:max-content is the other half of it. With table-layout:fixed and
           an auto width the table shrink-to-fits its container and divides that
           space among the columns, so a column asked for 201px laid out at 105px
           — the stored width still ignored, just differently. Sizing to the sum
           of the columns lets the body's overflow:auto scroll instead, which is
           what a grid should do. */
        /* While a drag is selecting CELLS it must not also be selecting
           characters: the native text selection started at mousedown and paints
           across every cell the pointer crosses. */
        .sheet-editor__grid--dragging, .sheet-editor__grid--dragging * { user-select: none; }
        /* The chosen zoom, applied to the GRID and to nothing else -- the
           header, toolbar and footer keep their size, which is what makes this
           a spreadsheet's zoom rather than the browser's. The zoom PROPERTY
           rather than a scale transform, for the reason the document editor
           gives where it does the same: zoom scales the LAYOUT, so the scroll
           box grows with the content and the sticky headers still stick, where
           a transform would paint a scaled picture over an unscaled extent.
           The variable is written onto the body by {@link applyZoom}. */
        .sheet-editor__grid {
            border-collapse: collapse; font-size: .8125rem;
            table-layout: fixed; width: max-content;
            zoom: var(--sheet-zoom, 1);
        }
        /* The default every column gets unless the document states one, chosen to
           match what the grid measured before fixed layout was introduced (the 140px
           input plus its border) so an untouched sheet looks unchanged. An inline
           width from the column group overrides it. */
        .sheet-editor__grid col { width: 141px; }
        .sheet-editor__grid col.sheet-editor__row-head-col { width: 44px; }
        .sheet-editor__corner, .sheet-editor__col-head, .sheet-editor__row-head {
            position: sticky; background: var(--cms-surface-2, var(--cms-surface));
            border: 1px solid var(--cms-border);
            font-weight: 600; color: var(--cms-text-muted);
            text-align: center; padding: 2px 6px;
        }
        .sheet-editor__col-head, .sheet-editor__corner { top: 0; z-index: 2; }
        .sheet-editor__row-head, .sheet-editor__corner { left: 0; z-index: 1; }
        .sheet-editor__corner { z-index: 5; }
        /* A frozen header outranks the ordinary ones on its own axis, and the
           corner outranks everything -- it is the only cell that is pinned on
           both. */
        .sheet-editor__col-head--frozen { z-index: 4; }
        .sheet-editor__row-head--frozen { z-index: 4; }
        /* ⚠️ TWO classes, not one, and that is the fix rather than the style:
           the plain .sheet-editor__cell rule sets position relative and is
           declared LATER in this sheet, so a single-class rule here lost the
           cascade and the cell stayed relative -- pinned by nothing, with a
           top the browser had no use for. Measured in the page, where the
           computed position read "relative" while every binding was correct.

           ⚠️ A BACKGROUND is not decoration either: a sticky cell floats over
           the rows scrolling beneath it, and a transparent one shows them
           through itself. An author's own fill is an inline style and wins. */
        .sheet-editor__cell.sheet-editor__cell--frozen-row,
        .sheet-editor__cell.sheet-editor__cell--frozen-col {
            position: sticky; z-index: 1; background: var(--cms-surface);
        }
        /* The block where the two overlap has to sit above both. */
        .sheet-editor__cell.sheet-editor__cell--frozen-row.sheet-editor__cell--frozen-col { z-index: 2; }
        .sheet-editor__row-head { min-width: 44px; }
        .sheet-editor__corner, .sheet-editor__col-head, .sheet-editor__row-head {
            cursor: pointer; user-select: none;
        }
        /* The header of a selected column or row, so the gesture is visible at
           the edge as well as in the highlighted cells — with a full column
           selected the cells may all be scrolled out of view. */
        .sheet-editor__head--selected {
            background: var(--cms-primary);
            color: var(--cms-text-inverse);
        }
        /* The resize target, sitting on the column's right edge and overhanging
           it by half its width so the grab area straddles the border the author
           is actually aiming at. position:sticky on the header is what makes
           this absolute box resolve against it. */
        .sheet-editor__grip {
            position: absolute; top: 0; bottom: 0; right: -3px;
            width: 6px; cursor: col-resize; z-index: 4;
        }
        .sheet-editor__grip:hover { background: var(--cms-primary); }
        /* The row equivalent, on the BOTTOM edge and resizing vertically. */
        .sheet-editor__grip--row {
            top: auto; left: 0; right: 0; bottom: -3px;
            width: auto; height: 6px; cursor: row-resize;
        }
        .sheet-editor__font { width: auto; padding: 2px 6px; font-size: .8125rem; }
        .sheet-editor__size { width: 4.5rem; padding: 2px 6px; font-size: .8125rem; }
        /* A colour input paints its own chrome in every browser; strip it back
           to the swatch so the toolbar reads as one row of controls. */
        .sheet-editor__swatch {
            width: 1.75rem; height: 1.5rem; padding: 0;
            border: 1px solid var(--cms-border); border-radius: var(--cms-radius-sm);
            background: none; cursor: pointer;
        }
        .sheet-editor__swatch::-webkit-color-swatch-wrapper { padding: 2px; }
        .sheet-editor__swatch::-webkit-color-swatch { border: 0; border-radius: 2px; }
        /* Stands in for the rows outside the window. No borders and no
           background: it is scroll height, not a cell, and painting it would
           draw one enormous box across the sheet. */
        .sheet-editor__spacer td { border: 0; padding: 0; }
        /* The column axis's spacers. Sticky like the real header so the grid
           does not tear when scrolled sideways, and unpainted for the same
           reason as the row spacer: they are width, not cells. */
        .sheet-editor__spacer-head {
            position: sticky; top: 0; z-index: 2;
            border: 0; padding: 0;
            background: var(--cms-surface-2, var(--cms-surface));
        }
        .sheet-editor__spacer-cell { border: 0; padding: 0; }
        /* Positioned so a wrapped cell's control can fill it exactly — see the
           wrap rule below for why a percentage height could not. */
        .sheet-editor__cell { border: 1px solid var(--cms-border); padding: 0; position: relative; }
        /* A merged cell reads as one box: the anchor already spans its columns
           and rows via colspan/rowspan, and the covered cells emit no td. */
        .sheet-editor__cell--merged { background: var(--cms-surface-2, transparent); }
        .sheet-editor__cell--in-range { box-shadow: inset 0 0 0 1px var(--cms-primary); }
        /* The selected cell keeps a mark of its own once focus goes to a
           TOOLBAR control. Until this existed the only mark was the input's
           :focus outline, so reaching for Borders read as the cell deselecting
           itself -- and the gesture then looked like it had nothing to act on.
           After --in-range deliberately: the active cell of a range is the one
           the toolbar reports. */
        .sheet-editor__cell--active { box-shadow: inset 0 0 0 2px var(--cms-primary); }
        /* 100%, not a fixed 140px: under table-layout:fixed the COLUMN sets the
           width, and an input that kept its own would leave dead space in a
           widened column and overflow a narrowed one. */
        .sheet-editor__input {
            border: 0; background: transparent; color: inherit;
            padding: 3px 6px; width: 100%; font: inherit; box-sizing: border-box;
        }
        /* The wrapped twin. No resize grip, because the row's height is the
           document's, not a drag handle's — an author sets it on the row
           header, and a corner grip here would silently disagree with the
           height the file stores. Full height so the control fills the row the
           auto-fit measured for it, rather than showing its own scrollbar
           inside a row that is already tall enough. */
        /* ⚠️ Absolutely positioned rather than height:100%, and that is a fix
           not a flourish. A percentage height needs a definite base and a table
           cell does not give one — measured, the control resolved to 45px
           inside a 111px row and silently clipped two of its own wrapped lines.
           Insetting to zero fills the cell's inner box exactly, whatever the
           row turned out to be. */
        /* Wide enough for a short URL or a token, narrow enough to leave
           the toolbar's buttons on one line. */
        .sheet-editor__link { width: 160px; }

        .sheet-editor__input--wrap {
            position: absolute; inset: 0; resize: none; overflow: hidden;
            white-space: pre-wrap; overflow-wrap: break-word; line-height: inherit;
        }
        .sheet-editor__input:focus { outline: 2px solid var(--cms-primary); outline-offset: -2px; }
        .sheet-editor__cell--bold .sheet-editor__input { font-weight: 700; }
        /* A formula shows what it COMPUTES; the cell being edited shows the
           formula itself. The colour is what tells the two apart from a literal
           that merely begins with an equals sign.
           (No backticks in here: this is inside a template literal, and one
           would end it -- the failure reads as a syntax error further down.) */
        .sheet-editor__cell--formula .sheet-editor__input { color: var(--cms-primary); }
        /* An error is not a value: it reads as a fault, not as content. */
        .sheet-editor__cell--error .sheet-editor__input { color: var(--cms-danger); }
        /* Waiting on template data. Dimmed rather than coloured, because there
           is nothing wrong with it — the answer simply is not knowable yet. */
        .sheet-editor__cell--unresolved .sheet-editor__input {
            color: var(--cms-text-muted); font-style: italic;
        }
        /* The formula helper. Anchored to the cell, above everything in the
           grid, and never inheriting the cell's own bold or colour -- it is
           chrome about the formula, not part of the sheet. */
        .sheet-editor__helper {
            position: absolute; top: 100%; left: 0; z-index: 30;
            min-width: 260px; max-width: 380px; max-height: 220px; overflow-y: auto;
            background: var(--cms-surface); color: var(--cms-text);
            border: 1px solid var(--cms-border); border-radius: var(--cms-radius, 4px);
            box-shadow: 0 4px 14px rgb(0 0 0 / 18%);
            font-weight: 400; font-style: normal; text-align: left;
        }
        .sheet-editor__helper-item {
            display: block; width: 100%; border: 0; background: transparent;
            padding: 4px 8px; text-align: left; font: inherit; color: inherit;
            cursor: pointer;
        }
        .sheet-editor__helper-item--on { background: var(--cms-primary); color: #fff; }
        .sheet-editor__helper-sig { padding: 4px 8px; }
        .sheet-editor__helper-name { font-weight: 600; }
        .sheet-editor__helper-hint { display: block; font-size: .85em; opacity: .75; }
        .sheet-editor__helper-arg--on { font-weight: 700; text-decoration: underline; }
        /* The filter button sits ON the header cell, over the right edge of its
           input, which is where every spreadsheet puts it. Absolutely
           positioned against the cell rather than in flow, so it cannot change
           the cell's height or push the text it labels.
           NO BACKTICKS in here -- a styles template literal ends at the first
           one, and the error surfaces hundreds of lines away. */
        .sheet-editor__filter-btn {
            position: absolute; top: 50%; right: 2px; transform: translateY(-50%);
            z-index: 5; display: flex; align-items: center; justify-content: center;
            width: 16px; height: 16px; padding: 0;
            border: 1px solid var(--cms-border); border-radius: 2px;
            background: var(--cms-surface); color: var(--cms-text-muted);
            font-size: 9px; line-height: 1; cursor: pointer;
        }
        .sheet-editor__filter-btn:hover { color: var(--cms-text); }
        /* Lit when this column is actually hiding something -- a filter that is
           merely DECLARED must look different from one that is doing something,
           or an author cannot tell why rows are missing. */
        .sheet-editor__filter-btn--on {
            background: var(--cms-primary); border-color: var(--cms-primary); color: #fff;
        }
        /* Opens RIGHTWARD from the cell's left edge, not leftward from its
           right one. Anchored to the right, a filter on column A -- which is
           where a table's first column nearly always is -- put the menu past
           the grid's left edge, where the scroll container CLIPPED its
           checkboxes and its Done button with nothing to scroll to. Opening
           right can also overflow, but the grid scrolls horizontally, so that
           menu is still reachable. Found in the browser; no spec sees paint. */
        .sheet-editor__filter-menu {
            position: absolute; top: 100%; left: 0; z-index: 30;
            min-width: 220px; max-width: 320px;
            background: var(--cms-surface); color: var(--cms-text);
            border: 1px solid var(--cms-border); border-radius: var(--cms-radius, 4px);
            box-shadow: 0 4px 14px rgb(0 0 0 / 18%);
            font-weight: 400; font-style: normal; text-align: left;
        }
        .sheet-editor__filter-head,
        .sheet-editor__filter-foot {
            display: flex; align-items: center; justify-content: space-between; gap: 8px;
            padding: 6px 8px;
        }
        .sheet-editor__filter-head { border-bottom: 1px solid var(--cms-border); }
        .sheet-editor__filter-foot { border-top: 1px solid var(--cms-border); }
        .sheet-editor__filter-title { font-size: .85em; opacity: .75; }
        .sheet-editor__filter-list { max-height: 220px; overflow-y: auto; padding: 4px 0; }
        .sheet-editor__filter-option {
            display: flex; align-items: center; gap: 6px;
            padding: 3px 8px; cursor: pointer;
        }
        .sheet-editor__filter-option:hover { background: var(--cms-surface-muted, rgb(0 0 0 / 5%)); }
        /* A long value must not widen the menu past its own max-width. */
        .sheet-editor__filter-value {
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }
        .sheet-editor__filter-empty {
            margin: 0; padding: 8px; font-size: .85em; color: var(--cms-text-muted);
        }
        .sheet-editor__tool-note { font-size: .85em; color: var(--cms-text-muted); }
        /* Same footprint and anchoring as the filter handle, a different glyph:
           one says "this column can be narrowed", the other "this cell has a
           list of allowed values". */
        .sheet-editor__options-btn {
            position: absolute; top: 50%; right: 2px; transform: translateY(-50%);
            z-index: 5; display: flex; align-items: center; justify-content: center;
            width: 16px; height: 16px; padding: 0;
            border: 1px solid var(--cms-border); border-radius: 2px;
            background: var(--cms-surface); color: var(--cms-text-muted);
            font-size: 9px; line-height: 1; cursor: pointer;
        }
        .sheet-editor__options-btn:hover { color: var(--cms-text); }
        .sheet-editor__option {
            display: block; width: 100%; border: 0; background: transparent;
            padding: 4px 8px; text-align: left; font: inherit; color: inherit;
            cursor: pointer;
        }
        .sheet-editor__option:hover { background: var(--cms-surface-muted, rgb(0 0 0 / 5%)); }
        .sheet-editor__option[aria-selected="true"] { font-weight: 600; }
        /* Centred in the cell rather than filling it: a tick box is a control,
           not text, so it does not take the input's full-bleed treatment. */
        .sheet-editor__checkbox {
            display: block; margin: 0 auto; width: 15px; height: 15px; cursor: pointer;
        }
        .sheet-editor__insert { width: auto; padding: 2px 6px; font-size: .8125rem; }
        /* Opens rightward and downward from the cell, for the reason the filter
           menu learned the hard way: anchored to the right edge it fell off the
           left of the grid on column A, where the scroll container clipped it
           with nothing to scroll to. */
        .sheet-editor__context {
            position: absolute; top: 100%; left: 0; z-index: 40;
            min-width: 190px;
            padding: 4px 0;
            background: var(--cms-surface); color: var(--cms-text);
            border: 1px solid var(--cms-border); border-radius: var(--cms-radius, 4px);
            box-shadow: 0 4px 14px rgb(0 0 0 / 18%);
            font-weight: 400; font-style: normal; text-align: left; white-space: nowrap;
        }
        /* Opens UPWARD when it would not fit below -- the grid body clips it
           otherwise, and what gets cut is the bottom of the list. Decided by
           measuring the rendered menu, not by counting its items. */
        .sheet-editor__context--above { top: auto; bottom: 100%; }
        /* The glyphs sit in a fixed column so the LABELS line up whatever each
           icon's own width is -- a ragged left edge is what makes an icon menu
           look worse than a plain one. */
        .sheet-editor__context .sheet-editor__option {
            display: flex; align-items: center; gap: 8px;
        }
        .sheet-editor__context .sheet-editor__option > .bi {
            flex: 0 0 auto; width: 14px; text-align: center;
            font-size: 12px; color: var(--cms-text-muted);
        }
        .sheet-editor__context .sheet-editor__option:disabled {
            opacity: .45; cursor: default; background: transparent;
        }
        .sheet-editor__context-rule {
            margin: 4px 0; border: 0; border-top: 1px solid var(--cms-border);
        }
        /* The palette hangs off its toolbar button. The toolbar neither scrolls
           nor clips, so this needs none of the measuring the cell popups do. */
        .sheet-editor__borders { position: relative; }
        .sheet-editor__border-menu {
            position: absolute; top: 100%; left: 0; z-index: 40;
            margin-top: 4px; padding: 8px;
            display: flex; flex-direction: column; gap: 6px;
            background: var(--cms-surface); color: var(--cms-text);
            border: 1px solid var(--cms-border); border-radius: var(--cms-radius, 4px);
            box-shadow: 0 4px 14px rgb(0 0 0 / 18%);
        }
        .sheet-editor__find {
            display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
            padding: 6px 12px;
            border-bottom: 1px solid var(--cms-border);
            background: var(--cms-surface-2, var(--cms-surface));
        }
        .sheet-editor__find-input { width: 180px; }
        .sheet-editor__rules {
            padding: 6px 12px;
            border-bottom: 1px solid var(--cms-border);
            background: var(--cms-surface-2, var(--cms-surface));
        }
        .sheet-editor__rules-head { font-size: .8125rem; color: var(--cms-text-muted); margin-bottom: 4px; }
        .sheet-editor__rules-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
        .sheet-editor__rules-value { width: 120px; }
        .sheet-editor__rules-existing {
            display: flex; align-items: center; gap: 8px;
            margin-top: 4px; font-size: .8125rem;
        }
        /* A sample of the rule rather than a description of it: an author
           recognises the colour they picked faster than they read its name. */
        .sheet-editor__rules-swatch {
            display: inline-block; min-width: 28px; padding: 0 6px;
            border: 1px solid var(--cms-border); border-radius: 2px;
            text-align: center;
        }
        .sheet-editor__find-option {
            display: flex; align-items: center; gap: 4px;
            font-size: .8125rem; color: var(--cms-text-muted); white-space: nowrap;
        }
        /* Same shell as the rules panel: both sit above the grid rather than
           over it, so neither hides the cells it is about. */
        .sheet-editor__names {
            padding: 6px 12px;
            border-bottom: 1px solid var(--cms-border);
            background: var(--cms-surface-2, var(--cms-surface));
        }
        .sheet-editor__names-add { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
        .sheet-editor__names-input { width: 200px; }
        .sheet-editor__names-error {
            margin: 4px 0 0; font-size: .8125rem; color: var(--cms-danger, #b42318);
        }
        .sheet-editor__names-empty {
            margin: 6px 0 0; font-size: .8125rem; color: var(--cms-text-muted); max-width: 60ch;
        }
        .sheet-editor__names-row {
            display: flex; align-items: center; gap: 8px;
            margin-top: 4px; font-size: .8125rem;
        }
        .sheet-editor__names-row code { color: var(--cms-text); }
        .sheet-editor__names-row span { color: var(--cms-text-muted); }
        .sheet-editor__find-count { font-size: .8125rem; color: var(--cms-text-muted); }
        /* Every match at once, not just the one being visited: the count says
           how many there are, and an author about to replace all of them should
           be able to SEE which cells that means. Two classes, so it beats the
           plain cell rule declared later -- the lesson the frozen panes left. */
        .sheet-editor__cell.sheet-editor__cell--match {
            background: var(--cms-warning-bg, rgb(255 214 0 / 22%));
        }
        .sheet-editor__border-row { display: flex; gap: 4px; }
        /* Square, because here the GLYPH is the label -- buttons sized to their
           text would be the list of words this replaced. */
        .sheet-editor__border-row .cms-btn { width: 32px; padding: 2px 0; text-align: center; }
        /* The one glyph every spreadsheet uses for this, and one no icon set
           has: a literal sigma beats a calculator that means something else. */
        .sheet-editor__sigma { font-weight: 700; font-size: 1.05em; line-height: 1; }
        .sheet-editor__functions { position: relative; }
        .sheet-editor__function-menu {
            position: absolute; top: 100%; right: 0; z-index: 40;
            margin-top: 4px; padding: 8px;
            display: flex; flex-direction: column; gap: 6px;
            width: 320px;
            background: var(--cms-surface); color: var(--cms-text);
            border: 1px solid var(--cms-border); border-radius: var(--cms-radius, 4px);
            box-shadow: 0 4px 14px rgb(0 0 0 / 18%);
        }
        /* Opens leftward, unlike the border palette: this one is wide and the
           control sits at the right-hand end of the toolbar. */
        .sheet-editor__function-search { width: 100%; }
        /* Bounded, because the catalogue does not fit on a screen and a menu
           that runs off the bottom is the bug the context menu already had. */
        .sheet-editor__function-list { max-height: 320px; overflow-y: auto; }
        .sheet-editor__function-shelf {
            padding: 6px 8px 2px; font-size: .75rem; font-weight: 600;
            text-transform: uppercase; letter-spacing: .04em;
            color: var(--cms-text-muted);
        }
        .sheet-editor__function {
            display: flex; flex-direction: column; align-items: flex-start; gap: 1px;
            text-align: left; white-space: normal;
        }
        .sheet-editor__function-name { font-weight: 600; font-family: var(--cms-font-mono, monospace); }
        .sheet-editor__function-summary { font-size: .75rem; color: var(--cms-text-muted); }
        .sheet-editor__function-empty { padding: 8px; font-size: .8125rem; color: var(--cms-text-muted); }
    `],
})
export class SheetEditorDialogComponent {
    private readonly dialogRef  = inject(DialogRef);
    private readonly data       = inject(DIALOG_DATA) as { node: VfsNodeDto };
    private readonly http       = inject(HttpClient);
    private readonly zone       = inject(NgZone);
    private readonly store      = inject(Store);
    private readonly toast      = inject(ToastService);
    private readonly dialogs    = inject(NativeDialogService);
    private readonly destroyRef = inject(DestroyRef);
    private readonly host       = inject(ElementRef<HTMLElement>);

    readonly node = this.data.node;

    protected readonly loading = signal(true);
    protected readonly saving  = signal(false);
    protected readonly dirty   = signal(false);

    private readonly doc = signal<SheetDocumentDto | null>(null);
    protected readonly activeSheet = signal('');

    /**
     * Refused rather than opened blank. A file the editor cannot parse is the
     * one state an operator most needs to repair, and opening it as an empty
     * grid invites a save that overwrites their document with nothing.
     */
    protected readonly unreadable = signal(false);

    protected readonly sheetNames = computed(() => Object.keys(this.doc()?.sheets ?? {}));

    /**
     * Rendered height of a row with no stated height, in CSS pixels (#2067).
     *
     * PINNED as a constant and applied to every row, because virtualisation
     * needs to know where row N is WITHOUT measuring it — a content-driven
     * height cannot be predicted, and a scroll position computed from a guess
     * puts the wrong rows under the pointer.
     *
     * ⚠️ This is 26px ≈ 19.5pt, and `SheetDocumentWriter::DEFAULT_ROW_HEIGHT`
     * is 12.8pt — so an untouched row is TALLER on the canvas than in the
     * generated workbook. That divergence predates this change and is left
     * alone deliberately: matching the workbook would cramp a row that has to
     * hold a text input, and matching the canvas would restyle every sheet
     * already generated. It is a decision, not an oversight, and it is the row
     * equivalent of the font mismatch #2052 fixed.
     */
    private static readonly DEFAULT_ROW_PX = 26;

    /** Rows kept rendered beyond the viewport, so a flick does not show blanks. */
    private static readonly ROW_BUFFER = 12;

    private static readonly GROW_COLS = 8;

    /** Distance from the edge, in px, at which more columns are added. */
    private static readonly GROW_THRESHOLD_PX = 240;

    /**
     * Columns the author has scrolled into being.
     *
     * Rows no longer need this — virtualisation means the row count costs
     * nothing to raise, so the floor is simply a spreadsheet's worth. COLUMNS
     * are still all rendered, so they still grow on demand.
     */
    private readonly grownCols = signal(0);
    private readonly grownRows = signal(0);

    /** Added per scroll to the edge. Rows are cheap to add; columns are wider. */
    private static readonly GROW_ROWS = 500;

    /**
     * Default rendered COLUMN width, mirroring `.sheet-editor__grid col`.
     *
     * Pinned for the same reason as the row height: the window is computed from
     * it, so paint and arithmetic have to be the same number.
     */
    private static readonly DEFAULT_COL_PX = 141;

    /** Columns kept rendered beyond the viewport. Fewer than rows — they are wider. */
    private static readonly COL_BUFFER = 3;

    /**
     * Viewport geometry as MEASURED, in the scroll container's own pixels.
     *
     * ⚠️ Raw on purpose, and the distinction is the whole reason zoom works
     * here. The grid carries a `zoom`, so one pixel of the scroll container is
     * `zoom` pixels of grid -- while every width and height in the document,
     * and therefore every window and spacer computed from one, is written in
     * GRID pixels. The conversion happens once, in the two computeds below,
     * rather than at each of the six places that ask where the window is.
     */
    private readonly scrollTopPx = signal(0);
    private readonly viewportHeightPx = signal(600);
    private readonly scrollLeftPx = signal(0);
    private readonly viewportWidthPx = signal(900);

    /**
     * The same geometry in GRID pixels -- the space the row heights and column
     * widths are in, and so the space the windows have to be computed in.
     *
     * At 200% the viewport covers half as many rows, and the row under a given
     * scroll position is half as far down. Reading the raw pixels instead would
     * render the window for somewhere else on the sheet: blank rows above the
     * ones actually on screen, and cells missing below them.
     */
    private readonly scrollTop = computed(() => this.scrollTopPx() / this.zoom());
    private readonly viewportHeight = computed(() => this.viewportHeightPx() / this.zoom());
    private readonly scrollLeft = computed(() => this.scrollLeftPx() / this.zoom());
    private readonly viewportWidth = computed(() => this.viewportWidthPx() / this.zoom());

    private readonly extent = computed(() => {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const base = sheet ? gridExtent(sheet) : { rows: 1000, cols: 26 };

        return { rows: base.rows + this.grownRows(), cols: base.cols + this.grownCols() };
    });

    /**
     * Rows that state a height, sorted — the only rows whose position differs
     * from the uniform stride.
     *
     * Kept as a sorted list rather than consulted per row: the map is sparse
     * (an author sets a handful of heights, not a thousand), so every offset
     * question is answered by walking a few entries instead of a thousand.
     */
    private readonly heightOverrides = computed(() => {
        const d = SheetEditorDialogComponent.DEFAULT_ROW_PX;

        return [...this.effectiveRowPx()]
            .filter(([, px]) => px !== d)
            .map(([row, px]) => ({ row, px }))
            .sort((a, b) => a.row - b.row);
    });

    /**
     * Every row whose height is not the default: auto-fitted wrapped rows and
     * rows an author sized by hand, in ONE map (#2085).
     *
     * ## Why they merge here rather than at each use
     *
     * Three things need a row's height — the offset arithmetic, the total
     * scroll height, and the row itself — and virtualisation only works while
     * all three agree. Merging once means a wrapped row cannot be tall in the
     * paint and default in the arithmetic, which is precisely the drift that
     * makes a scrollbar wander away from the rows under it.
     *
     * **A hand-set height WINS over the measurement**, which is Excel's rule,
     * not a preference: a row carrying `customHeight` keeps it and clips its
     * content, and only a row that never stated one auto-fits. Written in that
     * order — measured first, explicit second — so the last writer wins.
     */
    private readonly effectiveRowPx = computed<ReadonlyMap<number, number>>(() => {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const out = new Map<number, number>(this.wrappedRowPx());

        for (const [key, points] of Object.entries(sheet?.rowHeights ?? {})) {
            const row = Number(key);
            if (Number.isFinite(row) && row > 0) {
                out.set(row, rowHeightToPx(points));
            }
        }

        // A filtered-out row is height ZERO, and that is the whole mechanism.
        // The docblock above says the offset arithmetic, the scroll height and
        // the row itself only work while all three agree -- so hiding a row by
        // giving it no height makes all three agree for free, where a separate
        // "skip these rows" list in the renderer would have to be mirrored into
        // `offsetOf` and `totalHeight` by hand. Written LAST so a hidden row
        // stays hidden whatever height it states.
        for (const row of this.hiddenRows()) {
            out.set(row, 0);
        }

        return out;
    });

    /**
     * How tall each row with WRAPPED content needs to be, measured.
     *
     * ## Why the browser is asked instead of the text being modelled
     *
     * Line breaking is the browser's job and reimplementing it is the trap this
     * codebase keeps naming: word boundaries, long words that cannot break,
     * proportional glyph widths, the author's own font on the cell. A probe
     * given the column's width and the cell's font wraps exactly as the cell
     * will, because it IS the same engine doing it.
     *
     * Only cells that state `wrap` are measured, and a row takes its tallest —
     * so a sheet with no wrapping pays nothing, which is nearly every sheet.
     */
    private readonly wrappedRowPx = computed<ReadonlyMap<number, number>>(() => {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const out = new Map<number, number>();
        if (!sheet) {
            return out;
        }

        for (const [ref, cell] of Object.entries(sheet.cells)) {
            if (true !== cell.wrap) {
                continue;
            }
            const text = cell.value ?? cell.formula ?? '';
            const at = parseRef(ref);
            if ('' === text || !at) {
                continue;
            }

            const px = this.measureWrapped(text, this.columnPx(at.column), cell);
            out.set(at.row, Math.max(out.get(at.row) ?? 0, px));
        }

        return out;
    });

    /**
     * One reusable off-screen probe, inside the editor so it inherits the
     * grid's own font — the measurement has to be of THIS text in THIS face,
     * and a probe on `document.body` would be measuring the admin chrome's.
     */
    private measureProbe?: HTMLElement;

    /** The height a cell's text needs at a given column width, in px. */
    private measureWrapped(text: string, columnPx: number, cell: SheetCellDto): number {
        const host = this.host.nativeElement;
        let probe = this.measureProbe;
        if (!probe || !host.contains(probe)) {
            probe = document.createElement('div');
            probe.setAttribute('aria-hidden', 'true');
            // Off-screen rather than `display: none`, which measures as zero.
            // ⚠️ `box-sizing: border-box` matches .sheet-editor__input, and it
            // is load-bearing: content-box would wrap the probe at the full
            // column width while the real control wraps at the width MINUS its
            // 12px of padding, so the two would break at different words and
            // the measured height would be a different number of lines.
            probe.style.cssText = 'position:absolute;left:-9999px;top:0;visibility:hidden;'
                + 'box-sizing:border-box;white-space:pre-wrap;overflow-wrap:break-word;';
            host.appendChild(probe);
            this.measureProbe = probe;
        }

        // Every property re-set on every call: the probe is REUSED, so a value
        // left over from the last cell would be measured into this one.
        probe.style.width = columnPx + 'px';
        // Mirrors .sheet-editor__input, whose padding is part of the height a
        // row must find room for.
        probe.style.padding = '3px 6px';
        probe.style.font = '';
        probe.style.fontFamily = cell.fontFamily ?? '';
        probe.style.fontSize = undefined === cell.fontSize ? '' : cell.fontSize + 'pt';
        probe.style.fontWeight = true === cell.bold ? '700' : '';
        probe.style.fontStyle = true === cell.italic ? 'italic' : '';
        probe.textContent = text;

        // Plus the cell's own two 1px borders, which the probe does not carry
        // and the ROW still has to find room for. Without it the arithmetic is
        // two pixels short of the paint on every wrapped row, which is exactly
        // the drift that walks a scrollbar out of step with its rows.
        return Math.max(
            SheetEditorDialogComponent.DEFAULT_ROW_PX,
            probe.offsetHeight + SheetEditorDialogComponent.CELL_BORDER_PX,
        );
    }

    /** The `td`'s top plus bottom border — see {@link measureWrapped}. */
    private static readonly CELL_BORDER_PX = 2;

    /** Pixels from the top of the grid to the top of `row` (1-based). */
    private offsetOf(row: number): number {
        const d = SheetEditorDialogComponent.DEFAULT_ROW_PX;
        let extra = 0;
        for (const o of this.heightOverrides()) {
            if (o.row >= row) break;
            extra += o.px - d;
        }

        return (row - 1) * d + extra;
    }

    /** Total scrollable height of every row, rendered or not. */
    private readonly totalHeight = computed(() => {
        const d = SheetEditorDialogComponent.DEFAULT_ROW_PX;
        const extra = this.heightOverrides().reduce((sum, o) => sum + (o.px - d), 0);

        return this.extent().rows * d + extra;
    });

    /**
     * How many rows and columns this sheet holds in place.
     *
     * Read from the document rather than kept as view state: a freeze is part
     * of the workbook -- Excel stores it, our writer emits it, the importer
     * reads it back -- so it belongs to the file and not to this session.
     */
    protected readonly frozen = computed(() => frozenAt(this.doc()?.sheets[this.activeSheet()]));

    /**
     * The rows held in place at the top: always 1..N, whatever is scrolled to.
     *
     * ⚠️ They have to be RENDERED to be sticky, and virtualisation would
     * otherwise leave them inside the top spacer the moment the window moved
     * past them. That is the whole interaction between these two features: a
     * frozen row is a row the window may not drop.
     */
    protected readonly frozenRows = computed(() => {
        const { rows } = this.frozen();

        return Array.from({ length: Math.min(rows, this.extent().rows) }, (_, i) => i + 1);
    });

    protected readonly frozenColumns = computed(() => {
        const { columns } = this.frozen();

        return Array.from({ length: Math.min(columns, this.extent().cols) }, (_, i) => indexToColumn(i + 1));
    });

    /**
     * The rows actually rendered: the viewport, plus a buffer, plus any row a
     * MERGE reaches into.
     *
     * The merge clause is not an optimisation. A merged block is drawn by its
     * anchor with a `rowspan`, and the rows it covers emit no `<td>` at all —
     * so a window that started below an anchor would render rows whose cells
     * had been swallowed by a block that was no longer on the page, leaving a
     * hole. Merges are few, so widening to include them is cheap.
     */
    protected readonly visibleRows = computed(() => {
        const { rows } = this.extent();
        const buffer = SheetEditorDialogComponent.ROW_BUFFER;
        const d = SheetEditorDialogComponent.DEFAULT_ROW_PX;

        // Approximate the window from the uniform stride, then widen generously
        // — exactness here would need a search per scroll event, and the buffer
        // already absorbs the drift a handful of custom heights can introduce.
        let first = Math.max(1, Math.floor(this.scrollTop() / d) - buffer);
        let last = Math.min(rows, Math.ceil((this.scrollTop() + this.viewportHeight()) / d) + buffer);

        const sheet = this.doc()?.sheets[this.activeSheet()];
        for (const range of sheet?.merges ?? []) {
            const box = parseRange(range);
            if (!box || box.bottom < first || box.top > last) continue;
            first = Math.min(first, box.top);
            last = Math.max(last, box.bottom);
        }

        // Frozen rows are drawn separately and always: a window that also
        // contained them would render the same row twice, with two inputs
        // claiming one cell. AFTER the merge widening above, for the reason its
        // twin in `visibleColumns` spells out.
        first = Math.max(first, this.frozen().rows + 1);

        const window = Array.from({ length: Math.max(0, last - first + 1) }, (_, i) => first + i);

        const hidden = this.hiddenRows();
        if (0 === hidden.size) return window;

        // With rows collapsed to nothing, the uniform-stride approximation above
        // is measuring the wrong axis: it answers "which row NUMBER is at this
        // scroll position" when the layout now runs over the rows that are still
        // THERE. Filter first, then take the same window by INDEX -- the shown
        // rows are laid out contiguously, so the stride is honest again over
        // them. Only pays the full walk while a filter is actually hiding
        // something.
        const { rows: total } = this.extent();
        const shown: number[] = [];
        for (let row = 1; row <= total; row++) {
            if (!hidden.has(row)) shown.push(row);
        }

        const firstIndex = Math.max(0, Math.floor(this.scrollTop() / d) - buffer);
        const lastIndex = Math.min(shown.length - 1, Math.ceil((this.scrollTop() + this.viewportHeight()) / d) + buffer);

        return shown.slice(firstIndex, lastIndex + 1);
    });

    /**
     * Columns that state a width, sorted by INDEX — the column axis's twin of
     * {@link heightOverrides}, and sparse for the same reason.
     */
    private readonly widthOverrides = computed(() => {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const out: Array<{ index: number; px: number }> = [];
        for (const [letters, width] of Object.entries(sheet?.columnWidths ?? {})) {
            const index = columnToIndex(letters);
            if (index > 0) {
                out.push({ index, px: columnWidthToPx(width) });
            }
        }

        return out.sort((a, b) => a.index - b.index);
    });

    /** Pixels from the left of the grid to the left edge of column `index` (1-based). */
    private colOffsetOf(index: number): number {
        const d = SheetEditorDialogComponent.DEFAULT_COL_PX;
        let extra = 0;
        for (const o of this.widthOverrides()) {
            if (o.index >= index) break;
            extra += o.px - d;
        }

        return (index - 1) * d + extra;
    }

    private readonly totalWidth = computed(() => {
        const d = SheetEditorDialogComponent.DEFAULT_COL_PX;
        const extra = this.widthOverrides().reduce((sum, o) => sum + (o.px - d), 0);

        return this.extent().cols * d + extra;
    });

    /**
     * The columns actually rendered — the horizontal twin of
     * {@link visibleRows}, merge widening included for the same reason: a
     * merge's covered columns emit no `<td>`, so a window starting right of an
     * anchor would leave a hole.
     */
    protected readonly visibleColumns = computed(() => {
        const { cols } = this.extent();
        const buffer = SheetEditorDialogComponent.COL_BUFFER;
        const d = SheetEditorDialogComponent.DEFAULT_COL_PX;

        let first = Math.max(1, Math.floor(this.scrollLeft() / d) - buffer);
        let last = Math.min(cols, Math.ceil((this.scrollLeft() + this.viewportWidth()) / d) + buffer);

        const sheet = this.doc()?.sheets[this.activeSheet()];
        for (const range of sheet?.merges ?? []) {
            const box = parseRange(range);
            if (!box || box.right < first || box.left > last) continue;
            first = Math.min(first, box.left);
            last = Math.max(last, box.right);
        }

        // ⚠️ AFTER the merge widening, which reaches leftward and would
        // otherwise pull the window back over the frozen block: this sheet's
        // title is merged across A1:D1, so widening set `first` to 1 again and
        // column A was rendered twice -- once frozen, once in the window, two
        // inputs claiming one cell. A merge anchored in a frozen column is
        // still drawn, by the frozen block.
        first = Math.max(first, this.frozen().columns + 1);

        return Array.from({ length: Math.max(0, last - first + 1) }, (_, i) => indexToColumn(first + i));
    });

    /** Whether this row is one of the ones held at the top. */
    protected isFrozenRow(row: number): boolean {
        return row <= this.frozen().rows;
    }

    protected isFrozenColumn(column: string): boolean {
        return columnToIndex(column) <= this.frozen().columns;
    }

    protected isFrozen(): boolean {
        const { rows, columns } = this.frozen();

        return rows > 0 || columns > 0;
    }

    /**
     * Hold everything above and left of the selected cell, or let it all go.
     *
     * Above and LEFT of it, as every spreadsheet means by freezing: with B2
     * selected the header row and the first column stay while the rest scrolls.
     * Selecting A1 therefore freezes nothing, which is the same thing as
     * unfreezing and is stored the same way -- by leaving the field out.
     */
    protected toggleFreeze(): void {
        const at = this.isFrozen() ? null : this.activeRef();
        if (null !== at && '' === at) return;

        this.mutateSheet(sheet => withFreeze(sheet, at));
    }

    /**
     * How far from the top of the scroll box a frozen row sits.
     *
     * Under the column header, then stacked by the heights of the frozen rows
     * above it -- measured from the document's own row heights, so a frozen row
     * an author made taller still lands where it is drawn rather than where a
     * uniform stride would guess.
     */
    protected frozenTopPx(row: number): number {
        return SheetEditorDialogComponent.HEADER_PX + this.offsetOf(row) - this.offsetOf(1);
    }

    /** The horizontal twin: past the row header, then the frozen columns before it. */
    protected frozenLeftPx(column: string): number {
        return SheetEditorDialogComponent.ROW_HEAD_PX
            + this.colOffsetOf(columnToIndex(column)) - this.colOffsetOf(1);
    }

    /** The column-header row's own height, which every frozen row sits below. */
    private static readonly HEADER_PX = 25;
    /** Matches `.sheet-editor__row-head-col` in the styles. */
    private static readonly ROW_HEAD_PX = 44;

    /** Spacer widths standing in for the columns left and right of the window. */
    protected leftSpacerPx(): number {
        const first = this.visibleColumns()[0];
        if (undefined === first) return 0;

        // Minus what the frozen columns already occupy: they are rendered
        // before the spacer, so counting their width twice would push the
        // window one frozen block to the right of where it belongs.
        return Math.max(0, this.colOffsetOf(columnToIndex(first)) - this.colOffsetOf(this.frozen().columns + 1));
    }

    protected rightSpacerPx(): number {
        const visible = this.visibleColumns();
        const last = visible[visible.length - 1];

        return undefined === last
            ? 0
            : Math.max(0, this.totalWidth() - this.colOffsetOf(columnToIndex(last) + 1));
    }

    /**
     * Cells in one rendered row span three boxes: the left spacer, the visible
     * columns, and the right spacer — plus the row header. The row spacers have
     * to cover all of them or the table's column count disagrees between rows.
     */
    protected gridColspan(): number {
        return this.visibleColumns().length
            + this.frozenColumns().length
            + 1
            + (this.leftSpacerPx() > 0 ? 1 : 0)
            + (this.rightSpacerPx() > 0 ? 1 : 0);
    }

    /** Spacer heights that stand in for the rows above and below the window. */
    protected topSpacerPx(): number {
        // The frozen rows are rendered above this spacer, so their height is
        // already on the page -- see {@link leftSpacerPx} for the twin.
        return Math.max(0, this.offsetOf(this.visibleRows()[0] ?? 1) - this.offsetOf(this.frozen().rows + 1));
    }

    protected bottomSpacerPx(): number {
        const visible = this.visibleRows();
        const last = visible[visible.length - 1];

        return undefined === last ? 0 : Math.max(0, this.totalHeight() - this.offsetOf(last + 1));
    }

    /**
     * Track the viewport and extend the COLUMNS as their edge is approached.
     *
     * Rows need no such thing now: the full count is always addressable and
     * only the visible slice is built.
     */
    protected onGridScroll(event: Event): void {
        const el = event.target as HTMLElement;
        this.scrollTopPx.set(el.scrollTop);
        this.viewportHeightPx.set(el.clientHeight);
        this.scrollLeftPx.set(el.scrollLeft);
        this.viewportWidthPx.set(el.clientWidth);

        // Both axes grow on demand rather than starting at Excel's maximum.
        // Virtualisation makes the DOM cost of a bigger grid nothing, but the
        // SCROLL BOX is still real: 1,048,576 rows at 26px is 27 million pixels,
        // past what some browsers will lay out, and a scrollbar over a million
        // rows cannot be aimed anyway. Growing keeps the sheet effectively
        // endless while the scrollbar stays usable — which is what Sheets does,
        // and for the same reason.
        if (el.scrollHeight - el.scrollTop - el.clientHeight < SheetEditorDialogComponent.GROW_THRESHOLD_PX) {
            this.grownRows.update(n => n + SheetEditorDialogComponent.GROW_ROWS);
        }
        if (el.scrollWidth - el.scrollLeft - el.clientWidth < SheetEditorDialogComponent.GROW_THRESHOLD_PX) {
            this.grownCols.update(n => n + SheetEditorDialogComponent.GROW_COLS);
        }
    }

    /**
     * Zoom bounds and step.
     *
     * Sheets' own range. A step of 1.1 makes one wheel notch a nudge rather
     * than a jump, and eight of them cover the whole range -- which is about
     * how far a wrist turns in one gesture.
     */
    private static readonly ZOOM_MIN = 0.5;
    private static readonly ZOOM_MAX = 2;
    private static readonly ZOOM_STEP = 1.1;

    /** What the grid is scaled by. 1 is actual size. */
    private readonly zoom = signal(1);

    protected readonly zoomPercent = computed(() => Math.round(this.zoom() * 100));

    protected zoomIn(): void {
        this.stepZoom(SheetEditorDialogComponent.ZOOM_STEP);
    }

    protected zoomOut(): void {
        this.stepZoom(1 / SheetEditorDialogComponent.ZOOM_STEP);
    }

    protected resetZoom(): void {
        this.stepZoom(1 / this.zoom());
    }

    /** A button zooms about the middle of the viewport -- there is no pointer to zoom about. */
    private stepZoom(factor: number): void {
        const body = this.gridBody();
        this.applyZoom(
            this.zoom() * factor,
            (body?.clientWidth ?? 0) / 2,
            (body?.clientHeight ?? 0) / 2,
        );
    }

    /**
     * Ctrl + wheel over the grid zooms the GRID.
     *
     * ⚠️ `{ passive: false }` at the registration is the whole fix. Without a
     * cancellable listener the browser keeps the gesture and zooms the entire
     * admin -- shell, toolbar, dialog frame and all -- around a grid that is
     * still exactly the size it was. That is what an author reports as "zoom
     * does the wrong thing over the spreadsheet".
     *
     * A trackpad pinch arrives as precisely this event, which is why the
     * gesture is not handled separately.
     *
     * A PLAIN wheel is left entirely alone: it is the browser scrolling the
     * grid, which already works, and cancelling it would break scrolling to fix
     * zoom.
     */
    private readonly onWheel = (event: WheelEvent): void => {
        if (!event.ctrlKey && !event.metaKey) return;
        // A horizontal wheel held with ctrl says nothing about which way to
        // zoom, and `deltaY > 0 ? out : in` would read a flat 0 as "in".
        if (0 === event.deltaY) return;

        // Scoped to the grid, not the dialog: over the toolbar or the footer,
        // ctrl + wheel is still the browser's, and taking it there would be
        // taking a gesture from a surface that has nothing to zoom.
        const node = event.target;
        const body = node instanceof Element ? node.closest('.sheet-editor__body') : null;
        if (!(body instanceof HTMLElement)) return;

        event.preventDefault();

        const box = body.getBoundingClientRect();
        const factor = event.deltaY > 0
            ? 1 / SheetEditorDialogComponent.ZOOM_STEP
            : SheetEditorDialogComponent.ZOOM_STEP;

        // Back into the zone only now: a wheel over a grid fires dozens of
        // events a second and all but these are the browser's own scrolling,
        // which is none of Angular's business.
        this.zone.run(() => this.applyZoom(
            this.zoom() * factor,
            event.clientX - box.left,
            event.clientY - box.top,
        ));
    };

    private gridBody(): HTMLElement | null {
        const body = this.host.nativeElement.querySelector('.sheet-editor__body');

        return body instanceof HTMLElement ? body : null;
    }

    /**
     * Set the zoom, keeping the grid point under the anchor where it is.
     *
     * ⚠️ Written onto the DOM here rather than bound in the template, and the
     * ORDER is load-bearing. The new scroll position can only be set once the
     * grid has been laid out at the new scale: while zooming in, the scroll
     * extent grows, and a browser asked for a position past the old, smaller
     * extent silently clamps it -- so the sheet would drift towards the top on
     * every notch. Writing the variable and then reading `scrollLeft` flushes
     * that layout synchronously; a binding, applied on the next change
     * detection, could not, and would cost a frame of drift per notch inside a
     * gesture that fires dozens a second.
     *
     * Anchoring is what makes this feel like scaling the SHEET rather than
     * replacing it: the cell under the pointer stays under the pointer.
     */
    private applyZoom(next: number, anchorX: number, anchorY: number): void {
        const from = this.zoom();
        // Rounded to whole percents so the readout and the scale are the same
        // number -- a grid at 1.4641 displaying "146%" would be a lie an author
        // could catch by clicking the readout and watching nothing reset.
        const to = Math.min(
            SheetEditorDialogComponent.ZOOM_MAX,
            Math.max(SheetEditorDialogComponent.ZOOM_MIN, Math.round(next * 100) / 100),
        );
        if (to === from) return;

        this.zoom.set(to);

        const body = this.gridBody();
        if (null === body) return;

        const left = body.scrollLeft;
        const top = body.scrollTop;
        const ratio = to / from;

        body.style.setProperty('--sheet-zoom', String(to));
        body.scrollLeft = (left + anchorX) * ratio - anchorX;
        body.scrollTop = (top + anchorY) * ratio - anchorY;

        // Read BACK rather than assumed: at the edges the browser clamps what
        // was asked for, and a window computed from the position we wanted
        // instead of the one we got is a window over the wrong rows.
        this.scrollLeftPx.set(body.scrollLeft);
        this.scrollTopPx.set(body.scrollTop);
    }

    protected readonly rows = computed(() =>
        Array.from({ length: this.extent().rows }, (_, i) => i + 1));

    protected readonly columns = computed(() =>
        Array.from({ length: this.extent().cols }, (_, i) => indexToColumn(i + 1)));

    protected readonly statusText = computed(() => {
        if (this.unreadable()) return 'This file is not a readable .dsheet — saving would replace it.';
        const { rows, cols } = this.extent();

        return `${this.activeSheet()} · ${rows} × ${cols}`;
    });

    constructor() {
        this.load();
        // Names for the font select. Memoised across dialog opens, so this is
        // one fetch per tab and usually none at all.
        void offeredFamilies().then((families) => this.fontFamilies.set(families));
        // Registered once, not per drag: closing the dialog mid-resize would
        // otherwise leave mousemove/mouseup handlers on the document forever.
        this.destroyRef.onDestroy(() => this.cancelResize?.());

        // On the host rather than the grid because the grid does not exist yet
        // -- and it must be added by hand rather than as a template binding so
        // that `passive: false` can be stated. See {@link onWheel}: without it
        // the preventDefault is ignored and the browser zooms the admin.
        const host = this.host.nativeElement;
        this.zone.runOutsideAngular(() => host.addEventListener('wheel', this.onWheel, { passive: false }));
        this.destroyRef.onDestroy(() => host.removeEventListener('wheel', this.onWheel));
    }

    /**
     * Every formula on the active sheet, evaluated once per document change.
     *
     * A computed rather than a call per cell: `valueAt` runs for every rendered
     * cell on every change detection, and evaluating there would re-walk each
     * formula's whole dependency chain each time. `edit()` replaces the document
     * reference, which is what makes this re-run when a cell changes -- and why
     * editing one cell updates every formula that reads it.
     */
    private readonly evaluated = computed<ReadonlyMap<string, CellValue>>(() => {
        const doc = this.doc();

        return doc ? evaluateSheet(doc, this.activeSheet()) : new Map();
    });

    /**
     * What a cell shows.
     *
     * A formula shows its RESULT, except in the cell being edited, which shows
     * the formula so it can be changed -- the same swap a spreadsheet makes, and
     * the reason this grid needs no separate formula bar.
     *
     * An UNRESOLVED formula keeps showing its formula text rather than going
     * blank: it depends on a template token, so there is no result to show and
     * an empty cell would read as broken rather than as pending.
     */
    protected valueAt(ref: string): string {
        const cell = this.doc()?.sheets[this.activeSheet()]?.cells[ref];
        if (cell?.formula === undefined || this.activeRef() === ref) {
            return this.literalAt(ref, cell);
        }
        const value = this.evaluated().get(ref);
        if (value === undefined || value.kind === 'unresolved') {
            return this.literalAt(ref, cell);
        }

        // A computed value wears the cell's format too: a SUM under
        // `#,##0.00` shows as money, exactly as it will in the document.
        return formatCellValue(displayValue(value), cell.numberFormat);
    }

    /**
     * A cell that is not showing a computed value: what it holds, as it looks.
     *
     * ⚠️ The cell being EDITED shows what it stores; every other cell shows
     * what the document will show. The two differ for a date, and only for a
     * date: a `.xlsx` keeps one as the SERIAL its format describes, so an
     * imported invoice read `46255` in this grid while the generated document
     * read `21/08/2026`. The editor was not showing the document.
     *
     * The edit form of a date is still the date -- a spreadsheet's formula bar
     * shows `21/08/2026`, and nobody wants to type a serial -- while a number
     * under `#,##0.00` edits as `1234.5` and displays as `1,234.50`, which is
     * also what Excel does.
     */
    private literalAt(ref: string, cell: SheetCellDto | undefined): string {
        const raw = cellToInput(cell);
        if (cell?.formula !== undefined) return raw;

        return this.activeRef() === ref
            ? editForm(raw, cell?.numberFormat)
            : formatCellValue(raw, cell?.numberFormat);
    }

    /**
     * The hover text: the formula behind the result, and why a result is
     * missing when one is.
     */
    protected titleAt(ref: string): string | null {
        const cell = this.doc()?.sheets[this.activeSheet()]?.cells[ref];
        if (cell?.formula === undefined) return null;
        const value = this.evaluated().get(ref);
        if (value?.kind === 'unresolved') return `=${cell.formula}\n${value.because}`;

        return `=${cell.formula}`;
    }

    protected isFormula(ref: string): boolean {
        return this.doc()?.sheets[this.activeSheet()]?.cells[ref]?.formula !== undefined;
    }

    protected isFormulaError(ref: string): boolean {
        return this.evaluated().get(ref)?.kind === 'error';
    }

    protected isUnresolvedFormula(ref: string): boolean {
        return this.evaluated().get(ref)?.kind === 'unresolved';
    }

    protected isBold(ref: string): boolean {
        return this.doc()?.sheets[this.activeSheet()]?.cells[ref]?.bold === true;
    }

    protected readonly formats = NUMBER_FORMATS;

    /** The cell the toolbar acts on — focus IS selection in a grid of inputs. */
    protected readonly activeRef = signal('');

    private activeCell(): SheetCellDto | undefined {
        return this.doc()?.sheets[this.activeSheet()]?.cells[this.activeRef()];
    }

    protected activeFormat(): string | undefined {
        return this.activeCell()?.numberFormat;
    }

    protected activeBold(): boolean {
        return this.activeCell()?.bold === true;
    }

    /** Whether the active cell carries a code the menu cannot represent. */
    protected hasCustomFormat(): boolean {
        return !isKnownFormat(this.activeFormat());
    }

    protected applyFormat(code: string): void {
        // The select uses '' for General, because an <option> cannot carry
        // undefined — mapping it back here keeps the model's "absent means
        // General" rule rather than storing an empty format code.
        this.replaceActive(cell => withNumberFormat(cell, '' === code ? undefined : code));
    }

    protected toggleBold(): void {
        const next = !this.activeBold();
        this.replaceActive(cell => withBold(cell, next));
    }

    protected activeItalic(): boolean {
        return this.activeCell()?.italic === true;
    }

    protected toggleItalic(): void {
        const next = !this.activeItalic();
        this.replaceActive(cell => withItalic(cell, next));
    }

    /**
     * Families offered by name, not measured from the system, and READ FROM THE
     * PLATFORM MANIFEST rather than typed out (#2312).
     *
     * The name is what lands in the workbook, and the workbook is opened
     * somewhere else — so the useful list is the one Excel and LibreOffice both
     * resolve, not whatever happens to be installed on the author's machine.
     * An empty choice means "inherit", which is not the same as naming the
     * default: see the model's note on why the default is never stored.
     *
     * ⚠️ The literal this replaced offered Georgia, Verdana and Tahoma, which
     * the platform vendors nothing for. Installing a family is now a manifest
     * entry and four files; this list follows on its own.
     */
    protected readonly fontFamilies = signal<readonly string[]>([]);

    protected readonly alignments = [
        { value: 'left' as const, label: 'Align left', icon: 'bi-text-left' },
        { value: 'center' as const, label: 'Align centre', icon: 'bi-text-center' },
        { value: 'right' as const, label: 'Align right', icon: 'bi-text-right' },
    ];

    /**
     * Vertical alignment, which only becomes observable once a row is TALLER
     * than its content — wrapping and a hand-set row height being the two ways
     * to cause that. Offered next to wrap for exactly that reason.
     */
    protected readonly verticalAlignments = [
        { value: 'top' as const, label: 'Align top', icon: 'bi-align-top' },
        { value: 'middle' as const, label: 'Align middle', icon: 'bi-align-middle' },
        { value: 'bottom' as const, label: 'Align bottom', icon: 'bi-align-bottom' },
    ];

    protected activeWrap(): boolean {
        return this.activeCell()?.wrap === true;
    }

    protected toggleWrap(): void {
        const next = !this.activeWrap();
        this.replaceActive(cell => withWrap(cell, next));
    }

    /** The active cell's value for one style field, or '' so an input shows its placeholder. */
    protected activeStyle(key: 'fontFamily' | 'fontSize' | 'color' | 'background' | 'align' | 'valign' | 'link'): string {
        return String(this.activeCell()?.[key] ?? '');
    }

    /**
     * Apply one style to the active cell. An empty value CLEARS it.
     *
     * Clicking the alignment that is already set clears it too, which is how a
     * toggle should behave and the only way back to "inherit" — the buttons are
     * a three-way choice with no fourth button for "none".
     */
    protected applyStyle(
        key: 'fontFamily' | 'fontSize' | 'color' | 'background' | 'align' | 'valign' | 'link',
        raw: string,
    ): void {
        const trimmed = raw.trim();
        const cleared = '' === trimmed
            || (('align' === key || 'valign' === key) && this.activeStyle(key) === trimmed);

        this.replaceActive(cell => withStyle(
            cell,
            key,
            cleared
                ? undefined
                : ('fontSize' === key ? Number(trimmed) : trimmed) as never,
        ));
    }

    /**
     * A colour input can only report a colour, never "none", so removing one
     * needs its own control — without this an author who shaded a cell could
     * not unshade it.
     */
    protected clearColours(): void {
        this.replaceActive(cell => withStyle(withStyle(cell, 'color', undefined) ?? undefined, 'background', undefined));
    }

    /**
     * Rendered height for a row — the stated one, or the pinned default.
     *
     * Never null now. A content-sized row cannot be predicted, and
     * virtualisation has to know where row N begins without measuring it: if
     * the paint and the arithmetic disagree, the spacers are wrong and the
     * scrollbar drifts away from the rows under it.
     */
    protected rowPx(row: number): number {
        // The SAME map the offsets and the total height are built from — see
        // effectiveRowPx. Reading `rowHeights` directly here (as this did until
        // #2085) would paint an auto-fitted row at its wrapped height while the
        // arithmetic still believed it was 26px.
        return this.effectiveRowPx().get(row) ?? SheetEditorDialogComponent.DEFAULT_ROW_PX;
    }

    /** The row-header twin of {@link startResize}; see it for why the listeners go on the document. */
    protected startRowResize(row: number, event: MouseEvent): void {
        event.stopPropagation();
        event.preventDefault();

        const header = (event.target as HTMLElement | null)?.parentElement;
        // ⚠️ Both numbers are SCREEN pixels: a bounding rect reports the zoomed
        // size, and the pointer travels across the screen rather than across the
        // grid. The stored height is in the grid's own scale, so both are
        // brought back to it -- without this a drag at 200% resizes the row
        // twice as fast as the pointer moves, and the row ends up at a height
        // nobody dragged to.
        const zoom = this.zoom();
        const startPx = (header?.getBoundingClientRect().height ?? 0) / zoom;
        const startY = event.clientY;

        const move = (moved: MouseEvent): void =>
            this.writeRowHeight(row, rowHeightFromPx(startPx + (moved.clientY - startY) / zoom));

        const up = (): void => {
            document.removeEventListener('mousemove', move);
            document.removeEventListener('mouseup', up);
            this.cancelResize = null;
        };

        this.cancelResize?.();
        this.cancelResize = up;
        document.addEventListener('mousemove', move);
        document.addEventListener('mouseup', up);
    }

    private writeRowHeight(row: number, height: number | undefined): void {
        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        if (!doc || !sheet) return;

        const next = withRowHeight(sheet, row, height);
        if (next === sheet) return;

        doc.sheets[this.activeSheet()] = next;
        this.commit({ ...doc });
    }

    /**
     * The input's own styling, as ONE object for a single `[style]` binding.
     *
     * Returns null when the cell states nothing, which is the overwhelmingly
     * common case — an unstyled grid then binds null everywhere instead of a
     * fresh object per cell per change detection.
     */
    protected styleAt(ref: string): Record<string, string> | null {
        const cell = this.doc()?.sheets[this.activeSheet()]?.cells[ref];
        if (!cell) return null;

        const style: Record<string, string> = {};
        if (cell.fontFamily !== undefined) style['font-family'] = cell.fontFamily;
        if (cell.fontSize !== undefined) style['font-size'] = cell.fontSize + 'pt';
        if (cell.color !== undefined) style['color'] = cell.color;
        if (cell.align !== undefined) style['text-align'] = cell.align;
        if (cell.italic === true) style['font-style'] = 'italic';

        return Object.keys(style).length > 0 ? style : null;
    }

    /**
     * The cell's vertical alignment, which belongs on the TD rather than the
     * input (#2084).
     *
     * `vertical-align` on a table cell distributes whatever height the row has
     * spare — so it does nothing until a row is taller than its content, which
     * is precisely when an author asks for it: a hand-set row height, or a
     * merged cell spanning rows. The three values map to CSS one-for-one, so
     * there is no translation table to keep in step.
     */
    protected valignAt(ref: string): string | null {
        return this.doc()?.sheets[this.activeSheet()]?.cells[ref]?.valign ?? null;
    }

    /** Whether this cell wraps, and so is drawn as a multi-line control. */
    protected wrapAt(ref: string): boolean {
        return true === this.doc()?.sheets[this.activeSheet()]?.cells[ref]?.wrap;
    }

    /** The cell's fill, which sits on the td rather than the input. */
    /**
     * A cell's background: its rule's, if one claims it, else its own.
     *
     * ⚠️ The RULE WINS, which is what Excel does and what makes the feature
     * mean anything: a conditional fill that lost to the cell's own would show
     * only on cells the author had left plain, so a shaded table -- the kind
     * anybody would want an overdue line highlighted in -- would show nothing
     * at all.
     */
    protected fillAt(ref: string): string | null {
        return this.lookAt(ref)?.background
            ?? this.doc()?.sheets[this.activeSheet()]?.cells[ref]?.background
            ?? null;
    }

    private replaceActive(change: (cell: SheetCellDto | undefined) => SheetCellDto | null): void {
        const ref = this.activeRef();
        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        if ('' === ref || !doc || !sheet) return;

        const next = change(sheet.cells[ref]);
        if (next === null) {
            delete sheet.cells[ref];
        } else {
            sheet.cells[ref] = next;
        }

        this.commit({ ...doc });
    }

    /**
     * Commit the cell being left, and put the helper away.
     *
     * ⚠️ On BLUR and not on `change`, and that is the whole of the fix. The
     * browser fires `change` only for a value it considers the USER to have
     * edited -- and a value written by POINT MODE is not one. A formula built
     * entirely by clicking cells (`=SUM(` then click, then shift-click) fired
     * no `change` at any point: not on Enter, not on clicking away. It was
     * silently thrown away, which is what "I changed the range but can't
     * implement it" was.
     *
     * Blur always fires, and at blur the input still holds the RAW text: change
     * detection has not run yet, so the value has not been replaced by its
     * formatted form. Committing the formatted form would store `1,234.50` as
     * text.
     */
    protected commitCell(ref: string, input: HTMLInputElement | HTMLTextAreaElement): void {
        // Unchanged cells are left alone, which is what makes it safe to bind
        // this to BOTH `change` and `blur`: whichever arrives second finds the
        // text already stored and does nothing. It also means blur, which fires
        // for every cell an author merely passes through, cannot mark a
        // document dirty for having been looked at.
        if (input.value !== this.editTextFor(ref)) this.edit(ref, input.value);
        this.dismissHelper();
    }

    /** What a cell shows while it is being edited -- see {@link literalAt}. */
    private editTextFor(ref: string): string {
        const cell = this.doc()?.sheets[this.activeSheet()]?.cells[ref];
        const raw = cellToInput(cell);

        return cell?.formula !== undefined ? raw : editForm(raw, cell?.numberFormat);
    }

    protected edit(ref: string, raw: string): void {
        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        if (!doc || !sheet) return;

        const next = inputToCell(raw, sheet.cells[ref]);
        if (next === null) {
            delete sheet.cells[ref];
        } else {
            sheet.cells[ref] = next;
        }

        // Replace the reference so the computed extent re-runs: a cell typed
        // below the current last row must grow the grid, not vanish.
        this.commit({ ...doc });
    }

    protected selectSheet(name: string): void {
        this.activeSheet.set(name);
        this.activeRef.set('');
        this.anchorRef.set('');
        // Growth belongs to a scroll session, not to the document — carrying it
        // over would render extra blank columns onto a sheet just opened. The
        // scroll position resets with it, or the new sheet opens showing a
        // window computed for the old one's length.
        this.grownCols.set(0);
        this.grownRows.set(0);
        this.scrollTopPx.set(0);
        this.scrollLeftPx.set(0);
    }

    /**
     * The name rules are the WRITER's, mirrored.
     *
     * `safeSheetName` strips what Excel forbids and caps at 31; showing the
     * author the corrected name here means the tab they name is the tab they
     * get, instead of one the backend quietly rewrites at generation time.
     */
    private sheetNameError(value: string, allow?: string): string | null {
        const name = safeSheetName(value);
        if ('' === value.trim()) return 'A sheet needs a name.';
        if (name !== allow && name in (this.doc()?.sheets ?? {})) return `There is already a sheet called "${name}".`;

        return null;
    }

    protected async addSheet(): Promise<void> {
        const doc = this.doc();
        if (!doc) return;

        const typed = await this.dialogs.input({
            title: 'Add sheet',
            label: 'Sheet name',
            initialValue: 'Sheet' + (Object.keys(doc.sheets).length + 1),
            required: true,
            validator: (value: string) => this.sheetNameError(value),
        });
        if (null === typed) return;

        const { doc: next, name } = withNewSheet(doc, typed);
        this.commit(next);
        this.selectSheet(name);
    }

    protected async renameSheet(): Promise<void> {
        const doc = this.doc();
        const from = this.activeSheet();
        if (!doc || '' === from) return;

        const typed = await this.dialogs.input({
            title: 'Rename sheet',
            label: 'Sheet name',
            initialValue: from,
            required: true,
            validator: (value: string) => this.sheetNameError(value, from),
        });
        if (null === typed) return;

        const next = withRenamedSheet(doc, from, typed);
        if (next === doc) return;

        this.commit(next);
        this.selectSheet(safeSheetName(typed));
    }

    protected async deleteSheet(): Promise<void> {
        const doc = this.doc();
        const name = this.activeSheet();
        if (!doc || Object.keys(doc.sheets).length <= 1) return;

        const ok = await this.dialogs.confirm({
            title: 'Delete sheet',
            message: `Delete "${name}" and everything on it? This cannot be undone.`,
            confirmLabel: 'Delete',
            danger: true,
        });
        if (!ok) return;

        const next = withoutSheet(doc, name);
        if (next === doc) return;

        this.commit(next);
        this.selectSheet(Object.keys(next.sheets)[0] ?? '');
    }

    /**
     * The other end of a range selection.
     *
     * Focus alone cannot express a range in a grid of inputs — only one can be
     * focused — so a shift-click keeps this where it was and moves `activeRef`,
     * exactly as a spreadsheet does. A plain click collapses the two.
     */
    protected readonly anchorRef = signal('');

    /** Called on focus: a plain focus starts a fresh, collapsed selection. */
    protected focusCell(ref: string, input?: HTMLInputElement | HTMLTextAreaElement): void {
        // A cell arrived at is SELECTED, not yet edited -- so the arrows still
        // belong to the grid until something is typed.
        this.editing = false;
        const extend = ref === this.extendingFor;
        this.extendingFor = null;

        this.activeRef.set(ref);
        if (!extend) this.anchorRef.set(ref);
        this.editingInput = input;
        if (input) this.trackCaret(input);
    }

    // ── Formula helper ──────────────────────────────────────────────────────
    //
    // All the thinking is in `formula/helper.ts`, which is a pure function of
    // (text, caret). What is left here is plumbing: where the caret is, which
    // completion is highlighted, and what the keys do.

    private editingInput?: HTMLInputElement | HTMLTextAreaElement;
    private readonly helper = signal<HelperState>({ kind: 'none' });
    protected readonly helperIndex = signal(0);

    /**
     * The state to render, or null when there is nothing to say.
     *
     * The return type EXCLUDES `none` rather than being `HelperState | null`,
     * so the template's `@else` narrows to `signature` and can read `state.fn`.
     * With the wider type the else branch still admits `none`, and
     * `strictTemplates` rejects it -- correctly, and only at build time: `tsc`
     * does not check templates, so this compiles clean and fails the AOT build.
     */
    protected helperState(): Exclude<HelperState, { kind: 'none' }> | null {
        const state = this.helper();

        return state.kind === 'none' ? null : state;
    }

    protected signatureOf(fn: FormulaFunction): readonly string[] {
        return signatureParts(fn);
    }

    /**
     * Whether a signature part is the one being typed.
     *
     * Compared through `argumentLabel` rather than by index, so the repeating
     * last part of a variadic signature stays highlighted from the second
     * argument onwards instead of the highlight running off the end.
     */
    protected isCurrentArgument(state: HelperState, index: number): boolean {
        if (state.kind !== 'signature') return false;

        return signatureParts(state.fn)[index] === argumentLabel(state.fn, state.argumentIndex);
    }

    protected trackCaret(input: HTMLInputElement | HTMLTextAreaElement): void {
        // ⚠️ Only the FOCUSED input may claim the editing state.
        //
        // Point mode cancels its mousedown so focus never leaves the formula --
        // but the browser still delivers `click` to the cell that was pointed
        // at, and that cell's own `(click)="trackCaret(...)"` then fired here.
        // The formula's editing state moved to a cell the author had only
        // pointed at: the outlines vanished (its value is not a formula) and
        // the next point click would have written the reference INTO it.
        //
        // Found in the browser. The specs missed it because a synthetic
        // mousedown does not produce the click that follows a real one.
        if (input !== document.activeElement) return;

        this.editingInput = input;
        // Read from the LIVE input rather than the document: the outlines have
        // to follow what is being typed, and the model only learns about it on
        // `change`.
        this.editingText.set(input.value);
        const next = helperAt(input.value, input.selectionStart ?? input.value.length);
        const previous = this.helper();
        this.helper.set(next);
        // Keep the highlight where it was while the same prefix is being typed;
        // reset it whenever the offer changes, so Enter never accepts something
        // the author has not looked at.
        if (next.kind !== 'completions' || previous.kind !== 'completions'
            || previous.prefix !== next.prefix) {
            this.helperIndex.set(0);
        }
    }

    protected dismissHelper(): void {
        this.helper.set({ kind: 'none' });
        this.helperIndex.set(0);
        // The outlines belong to the cell being edited, so they go with the
        // focus. A point-mode click never reaches here: it cancels its own
        // mousedown, so no blur fires and the formula keeps its highlights.
        this.editingText.set('');
    }

    /**
     * The keys the helper owns while it is open.
     *
     * Escape is the one worth noting: without intercepting it, the CDK dialog
     * takes it and the whole editor closes -- losing the edit to a keystroke
     * the author meant for a popup.
     */
    /**
     * The keys a CELL owns: the helper's first, then the grid's own.
     *
     * Enter and Tab are handled here because nothing else was handling them.
     * An `<input>` outside a `<form>` does nothing at all on Enter -- no
     * implicit submission, so no `change` -- and the author pressed it,
     * watched nothing happen, and had no way to tell whether the edit had
     * landed. Measured in the browser before it was believed.
     */
    protected onCellKey(ref: string, event: KeyboardEvent): void {
        const input = event.target as HTMLInputElement | HTMLTextAreaElement;

        if (this.onFormulaKey(event)) return;

        // ⚠️ Escape reaches the CDK dialog otherwise, and the whole editor
        // closes -- taking every unsaved cell with it. Mid-edit it belongs to
        // the CELL, and it does what a spreadsheet does: puts back what was
        // there. With nothing to revert it still closes the editor.
        if ('Escape' === event.key) {
            const stored = this.editTextFor(ref);
            if (input.value === stored) return;
            event.preventDefault();
            event.stopPropagation();
            input.value = stored;
            this.dismissHelper();

            return;
        }

        // F2 is the spreadsheet's own "now I am editing this" -- it is how an
        // author says the arrow keys should move the caret and not the cell.
        if ('F2' === event.key) {
            event.preventDefault();
            this.editing = true;

            return;
        }

        if (this.onHistoryKey(event)) return;
        if (this.onGridKey(ref, event, input)) return;

        // Alt+Enter is a line break inside a wrapped cell, as it is in Excel.
        if ('Enter' === event.key && event.altKey) return;

        const step = 'Enter' === event.key
            ? (event.shiftKey ? 'up' : 'down')
            : 'Tab' === event.key ? (event.shiftKey ? 'left' : 'right') : null;
        if (null === step) return;

        event.preventDefault();
        this.commitAndMove(ref, step);
    }

    /**
     * Whether the focused cell is being EDITED rather than merely selected.
     *
     * A spreadsheet has two states and this grid has one: every cell is an
     * input, always showing a caret. So the distinction is inferred -- a cell
     * is being edited once something has been TYPED into it, or once F2 says
     * so -- and it decides who the arrow keys belong to. Without it, arrows
     * either never move between cells (unusable as a grid) or always do
     * (unusable as an editor).
     */
    private editing = false;

    /** Typing is what turns a selected cell into an edited one. */
    protected onCellInput(input: HTMLInputElement | HTMLTextAreaElement): void {
        this.editing = true;
        this.trackCaret(input);
    }

    /**
     * Ctrl+Z and Ctrl+Y, unless a cell is being edited.
     *
     * While editing, Ctrl+Z is the INPUT's own undo: the author is taking back
     * the characters they just typed, not the last thing that happened to the
     * document. Stealing it would make it impossible to undo a typo without
     * undoing the edit before it.
     */
    private onHistoryKey(event: KeyboardEvent): boolean {
        if (this.editing || event.altKey || !(event.ctrlKey || event.metaKey)) return false;

        const key = event.key.toLowerCase();
        const redo = 'y' === key || ('z' === key && event.shiftKey);
        if (!redo && 'z' !== key) return false;

        event.preventDefault();
        if (redo) this.redo(); else this.undo();

        return true;
    }

    /**
     * The keys that belong to the GRID rather than to the text in a cell.
     *
     * Only while the cell is merely selected: once it is being edited the
     * arrows move the caret and Delete removes a character, which is what an
     * author typing a formula expects and what F2 exists to declare.
     */
    private onGridKey(ref: string, event: KeyboardEvent, input: HTMLInputElement | HTMLTextAreaElement): boolean {
        if (this.editing || event.ctrlKey || event.metaKey || event.altKey) return false;

        const step = ARROWS[event.key];
        if (undefined !== step) {
            event.preventDefault();
            this.moveTo(ref, step, event.shiftKey);

            return true;
        }

        if ('Delete' === event.key || 'Backspace' === event.key) {
            event.preventDefault();
            const range = this.selectionRange() ?? ref;
            this.mutateSheet(sheet => withClearedRange(sheet, range));
            // The DOM still holds what was cleared: the input is not re-bound
            // while it has the focus, so without this the cell reads as full
            // until the author clicks away and back.
            input.value = '';

            return true;
        }

        return false;
    }

    /**
     * Move the selection, extending it when Shift is held.
     *
     * Extending moves the ACTIVE end and leaves the anchor, which is the same
     * shape shift-click and dragging already use -- three gestures, one idea.
     */
    private moveTo(ref: string, step: 'up' | 'down' | 'left' | 'right', extend: boolean): void {
        const cell = parseRef(ref);
        if (!cell) return;

        const column = columnToIndex(cell.column) + ('left' === step ? -1 : 'right' === step ? 1 : 0);
        const row = cell.row + ('up' === step ? -1 : 'down' === step ? 1 : 0);
        if (column < 1 || row < 1) return;

        const target = indexToColumn(column) + row;
        if (extend) {
            // Named so the focus that follows on the next frame preserves the
            // anchor rather than collapsing what was just extended.
            this.extendingFor = target;
            this.activeRef.set(target);
        } else {
            this.extendingFor = null;
            this.focusCell(target);
        }

        requestAnimationFrame(() => this.inputFor(target)?.focus());
    }

    // ---- The clipboard ------------------------------------------------------
    //
    // Through the browser's own copy/cut/paste EVENTS rather than
    // `navigator.clipboard`: the events carry the data with them, so there is
    // no permission to ask for and no reading of a clipboard the author did not
    // aim at this grid.

    /** What was last copied FROM here, so a formula pasted back can be offset. */
    private copied: { range: string; text: string } | null = null;

    /**
     * Ctrl+C and Ctrl+X over the grid.
     *
     * Declines when there is a text selection INSIDE the cell: the author
     * highlighted characters and means to copy those, and a grid that overrode
     * that would make it impossible to copy half a formula.
     */
    protected onCopy(event: ClipboardEvent, cut: boolean): void {
        const input = this.activeInput();
        if (input && input.selectionStart !== input.selectionEnd) return;

        const sheet = this.doc()?.sheets[this.activeSheet()];
        const range = this.selectionRange() ?? this.activeRef();
        if (!sheet || '' === range) return;

        const text = toClipboardText(sheet, range);
        event.clipboardData?.setData('text/plain', text);
        event.preventDefault();
        this.copied = { range, text };

        if (cut) this.mutateSheet(next => withClearedRange(next, range));
    }

    /**
     * Ctrl+V over the grid.
     *
     * A single value with no tabs or newlines is left to the browser: pasting a
     * word into a half-typed formula must put it at the caret, not overwrite
     * the cell. Anything shaped like a TABLE is the grid's.
     */
    protected onPaste(event: ClipboardEvent): void {
        const text = event.clipboardData?.getData('text/plain') ?? '';
        if ('' === text) return;
        if (this.editing && !/[\t\r\n]/.test(text)) return;

        const at = rangeCorner(this.selectionRange() ?? this.activeRef());
        if ('' === at) return;

        const block = parseClipboardText(text);
        if (0 === block.length) return;

        event.preventDefault();

        // Offset a formula only when this grid is where it came from. Text from
        // somewhere else has no origin to be relative to, so `=A1+1` pasted
        // from a mail is the characters somebody typed.
        const offset = null !== this.copied && this.copied.text === text
            ? pasteOffset(rangeCorner(this.copied.range), at)
            : null;

        this.mutateSheet(sheet => withPastedBlock(sheet, at, block, offset));

        // Select what arrived, which is both a confirmation and what an author
        // reaches for next -- to format it, or to move it again.
        const filled = pastedRange(at, block);
        this.anchorRef.set(at);
        this.activeRef.set(filled.split(':')[1] ?? at);
    }

    /**
     * Commit the cell and move to the next one, as Enter and Tab do everywhere.
     *
     * The commit goes through BLUR rather than calling `edit` directly, so
     * there is one commit path and not two that can disagree. Blur runs
     * synchronously here, before change detection, so the input still holds
     * what the author typed.
     */
    private commitAndMove(ref: string, step: 'up' | 'down' | 'left' | 'right'): void {
        const cell = parseRef(ref);
        if (!cell) return;

        this.inputFor(ref)?.blur();

        const column = columnToIndex(cell.column) + ('left' === step ? -1 : 'right' === step ? 1 : 0);
        const row = cell.row + ('up' === step ? -1 : 'down' === step ? 1 : 0);
        if (column < 1 || row < 1) return;

        const target = indexToColumn(column) + row;
        this.extendingFor = null;
        this.focusCell(target);

        // On the NEXT frame: the target may not be drawn yet -- a row below the
        // window, or one the commit just grew the sheet by. The selection has
        // moved either way, which is what the author sees.
        requestAnimationFrame(() => this.inputFor(target)?.focus());
    }

    /** Whether the helper consumed the key. */
    private onFormulaKey(event: KeyboardEvent): boolean {
        const state = this.helper();
        if (state.kind === 'none') return false;

        if (event.key === 'Escape') {
            event.preventDefault();
            event.stopPropagation();
            this.dismissHelper();

            return true;
        }

        if (state.kind !== 'completions') return false;

        if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
            event.preventDefault();
            const step = event.key === 'ArrowDown' ? 1 : -1;
            const count = state.matches.length;
            this.helperIndex.set((this.helperIndex() + step + count) % count);

            return true;
        }

        if (event.key === 'Enter' || event.key === 'Tab') {
            const chosen = state.matches[this.helperIndex()];
            if (!chosen) return false;
            event.preventDefault();
            this.acceptCompletion(chosen);

            return true;
        }

        return false;
    }

    /** Put the chosen function into the cell and leave the caret inside it. */
    protected acceptCompletion(fn: FormulaFunction): void {
        const state = this.helper();
        const input = this.editingInput;
        if (state.kind !== 'completions' || !input) return;

        const result = applyCompletion(input.value, state, fn);
        input.value = result.text;
        input.setSelectionRange(result.caret, result.caret);
        this.trackCaret(input);
    }

    /**
     * The cell a gesture just extended TO, or null.
     *
     * ⚠️ A ref and not a boolean, and that is the whole point. As a flag it was
     * set by shift-click and waited for a `focus` that a prevented mousedown
     * never fires -- so it sat there until the NEXT focus of any kind and ate
     * that one instead. A plain arrow after a shift-click inherited it, left
     * the anchor behind and went on extending a range the author had walked
     * away from; so did a click on a cell reached any way but by mouse.
     *
     * Naming the cell makes the guard exact: a stray focus for THAT cell keeps
     * the range, and a focus for any other collapses it, which is what a focus
     * on another cell means.
     *
     * mousedown is what carries the modifier — `focus` events do not.
     */
    private extendingFor: string | null = null;

    protected extendTo(ref: string, event: MouseEvent): void {
        // Left button only, and armed before anything else decides what this
        // click MEANS: a drag is the same gesture whether it ends up extending
        // the selection or growing a formula's reference.
        if (0 === event.button) this.startDrag(ref);

        // ⚠️ POINT MODE FIRST, and only where a reference could legally go.
        // While a formula is being typed, clicking a cell writes its reference
        // instead of moving the cursor -- what every spreadsheet does. The
        // condition is what keeps it safe: `pointInsertAt` answers null after a
        // number or a closing paren, so clicking away from a half-typed formula
        // still means "go there" rather than silently editing it.
        //
        // `preventDefault` keeps focus in the formula's own input, which is
        // also why no blur fires and the outlines survive the click.
        const span = this.pointTarget();
        if (null !== span && ref !== this.activeRef()) {
            event.preventDefault();
            // Shift-click grows the reference just written into a range, the
            // same gesture the grid already uses for selecting one.
            const anchor = event.shiftKey ? this.pointAnchor : null;
            this.pointAnchor = event.shiftKey ? anchor : ref;
            this.insertReference(
                null !== anchor ? (rangeBetween(anchor, ref) ?? ref) : ref,
                span,
            );

            return;
        }

        if (!event.shiftKey) {
            // ⚠️ A plain click ENDS any extension, and clearing the flag HERE
            // is what makes that true. The flag guards against a stray focus
            // arriving after a prevented shift-click; because that prevention
            // means no focus fires at all, the flag otherwise survived until
            // the author's next plain click -- which was then mistaken for the
            // stray one, so `focusCell` skipped moving the anchor and the
            // following shift-click built a range from a cell the author had
            // left long ago.
            //
            // Found by inserting a form element over C2:C4 and watching it
            // land on C2:D4, replacing the dropdown that was in D. Silent
            // damage to cells nobody selected.
            this.extendingFor = null;

            return;
        }

        // `preventDefault` is load-bearing, and only a real browser shows why.
        // Shift-click inside a focused text input is the NATIVE "extend the text
        // selection" gesture: the browser highlights characters in the cell the
        // author started from and suppresses the focus change, so without this
        // the second click did nothing at all to the selection. Synthetic
        // `mousedown`/`focus` events in a spec never reproduce it — the first
        // version of this passed its unit test and did not work in the admin.
        //
        // Suppressing the default also means no `focus` fires, so `activeRef` is
        // moved here and `anchorRef` is left exactly where it was — which is the
        // range. The `extending` flag stays as a guard for any browser that
        // focuses anyway.
        event.preventDefault();
        this.extendingFor = ref;
        this.activeRef.set(ref);
    }

    // ---- Dragging a selection ---------------------------------------------
    //
    // Shift-click extends, and so should holding the button down: it is the
    // gesture every grid has, and the one an author reaches for first.

    /** The cell a left-button drag began in, or null when none is running. */
    private dragFrom: string | null = null;
    /** The last cell the pointer was over, so one cell is not handled twice. */
    private dragLast: string | null = null;
    private stopDragging?: () => void;

    /** True once a drag has LEFT the cell it began in -- see {@link startDrag}. */
    protected readonly dragging = signal(false);

    /**
     * Watch the pointer, until the button comes up.
     *
     * ⚠️ The listener is registered OUTSIDE Angular and re-enters only when the
     * cell under the pointer CHANGES. A drag across one cell is hundreds of
     * mousemove events and at most one state change; running change detection
     * over a grid of inputs for each of them would make the grid stutter under
     * the very gesture being added.
     *
     * On `document`, not on the grid: a pointer dragged off the edge and
     * released outside must still end the drag, or the next plain move over the
     * grid would keep extending a selection nobody is holding.
     */
    private startDrag(ref: string): void {
        this.endDrag();
        this.destroyRef.onDestroy(() => this.stopDragging?.());
        this.dragFrom = ref;
        this.dragLast = ref;

        const move = (event: MouseEvent): void => {
            if (0 === (event.buttons & 1)) { this.endDrag(); return; }

            // `instanceof`, not a cast: a mousemove can arrive with a target
            // that is not an Element at all, and a cast asserting otherwise
            // throws inside the listener and kills the drag silently.
            const node = event.target;
            const target = node instanceof Element ? node.closest('[aria-label]') : null;
            const over = target?.getAttribute('aria-label') ?? null;
            // A header button carries a label too, and it is not a cell.
            if (null === over || null === parseRef(over) || over === this.dragLast) return;
            this.dragLast = over;
            this.zone.run(() => this.dragTo(over));
        };
        const up = (): void => this.endDrag();

        this.zone.runOutsideAngular(() => {
            document.addEventListener('mousemove', move);
            document.addEventListener('mouseup', up);
        });

        this.stopDragging = () => {
            document.removeEventListener('mousemove', move);
            document.removeEventListener('mouseup', up);
        };
    }

    private endDrag(): void {
        this.stopDragging?.();
        this.stopDragging = undefined;
        this.dragFrom = null;
        this.dragLast = null;
        if (this.dragging()) this.zone.run(() => this.dragging.set(false));
    }

    /**
     * The pointer has reached another cell with the button down.
     *
     * A drag that never leaves its own cell is the browser selecting TEXT,
     * which is exactly what dragging across a value should do -- so a grid
     * selection only begins once the pointer has left the cell it started in,
     * and the text selection made on the way is dropped at that moment.
     */
    private dragTo(ref: string): void {
        if (null === this.dragFrom) return;

        // ⚠️ POINT MODE FIRST, and before anything touches the document's
        // selection. `removeAllRanges()` wipes the caret of the focused formula
        // input, and the span below is read FROM that caret -- so clearing it
        // first made every dragged reference land at offset zero: `=SUM(`
        // dragged over B3:B4 produced `B3:B4=SUM(B3`. Point mode needs no
        // clearing anyway: its mousedown was cancelled, so the browser never
        // started selecting text.
        //
        // Mid-formula the drag grows the REFERENCE, which is the same gesture
        // meaning the same thing one layer down.
        const span = this.pointTarget();
        if (null !== span && null !== this.pointAnchor) {
            this.insertReference(rangeBetween(this.pointAnchor, ref) ?? ref, span);

            return;
        }

        if (!this.dragging()) {
            this.dragging.set(true);
            window.getSelection()?.removeAllRanges();
        }

        // `anchorRef` is where the mousedown left it, so moving the active end
        // is the whole of extending -- the same shape shift-click already uses.
        this.extendingFor = ref;
        this.activeRef.set(ref);
    }

    /** The selected range, or null when the selection is a single cell. */
    protected selectionRange(): string | null {
        const from = this.anchorRef();
        const to = this.activeRef();
        if ('' === from || '' === to || from === to) return null;

        return rangeBetween(from, to);
    }

    protected canMerge(): boolean {
        return null !== this.selectionRange();
    }

    // ---- Filtering -------------------------------------------------------
    //
    // Two halves that are deliberately NOT the same thing.
    //
    // The RANGE is a declaration and it is stored: it says which rows the
    // generated workbook's dropdowns cover, and the backend grows it over the
    // rows a `{loop:}` band adds, so a filter over a one-row band still filters
    // all fifty generated lines.
    //
    // The EXCLUDED VALUES are view state and are NOT stored. A template holds
    // tokens, not data -- hiding `{var:l.name}` in the editor says nothing
    // about the document that gets generated, and writing it into the `.dsheet`
    // would ship an author's temporary view to every render. Excel can store
    // criteria; a template has nothing true to put in them.

    // ---- Point mode ------------------------------------------------------
    //
    // While a formula is being typed, the grid stops being a place to navigate
    // and becomes a place to POINT: the ranges it already references are
    // outlined, and clicking a cell writes its reference into the formula.

    /**
     * What the cell being edited currently points at, coloured.
     *
     * A computed rather than a call per cell: this runs for every rendered cell
     * on every change detection, and re-tokenising the formula each time would
     * be quadratic in the visible grid. Colours cycle per DISTINCT reference,
     * which is what lets an author match `B3:B5` in the text to the outline on
     * the sheet.
     */
    private readonly pointedRefs = computed<readonly { box: MergeBox; colour: string }[]>(() => {
        const out: { box: MergeBox; colour: string }[] = [];
        const text = this.editingText();
        if ('' === text) return out;

        let colour = 0;
        for (const reference of referencesIn(text)) {
            // A qualified reference belongs to another sheet and outlining it
            // here would point at cells that have nothing to do with it.
            if (undefined !== reference.sheet && reference.sheet !== this.activeSheet()) continue;

            const range = reference.range.includes(':') ? reference.range : `${reference.range}:${reference.range}`;
            const box = parseRange(range);
            if (!box) continue;

            out.push({
                box,
                colour: SheetEditorDialogComponent.REF_COLOURS[colour % SheetEditorDialogComponent.REF_COLOURS.length],
            });
            colour++;
        }

        return out;
    });

    /** Distinct enough to tell apart, and readable on both themes. */
    private static readonly REF_COLOURS = ['#1a73e8', '#188038', '#c5221f', '#a142f4', '#e37400'];

    /**
     * The colour outlining this cell as a formula reference, if any.
     *
     * ⚠️ Tests BOXES rather than looking a cell up in a set of them. Building
     * that set meant walking every cell of every referenced range on every
     * keystroke -- `SUM(A1:A10000)` is ten thousand map entries per character
     * typed, and a mistyped `BD93:B5` made five thousand of them for a range
     * the author never meant. A reference is a rectangle; asking whether a cell
     * is inside one is four comparisons and no allocation at all.
     */
    protected pointedAt(ref: string): string | null {
        const boxes = this.pointedRefs();
        if (0 === boxes.length) return null;

        const cell = parseRef(ref);
        if (!cell) return null;
        const column = columnToIndex(cell.column);

        for (const { box, colour } of boxes) {
            if (cell.row >= box.top && cell.row <= box.bottom
                && column >= box.left && column <= box.right) {
                return colour;
            }
        }

        return null;
    }

    /** The live text of the cell being edited, or empty when nothing is. */
    private readonly editingText = signal('');

    /**
     * The cell a point-mode click started from, so a shift-click can grow it
     * into a range. Reset by every plain point click, exactly as the grid's own
     * selection anchor is.
     */
    private pointAnchor: string | null = null;

    /**
     * Where a click on another cell would write its reference, or null when a
     * click should keep its ordinary meaning.
     */
    private pointTarget(): { from: number; to: number } | null {
        const input = this.editingInput;
        if (undefined === input) return null;

        return pointInsertAt(input.value, input.selectionStart ?? input.value.length);
    }

    /**
     * Write a clicked cell's reference into the formula being edited.
     *
     * The DOM input is written directly and the model is left alone until the
     * ordinary `change` fires: this is mid-edit, and pushing every click
     * through the document would make each one an undo step and re-render the
     * grid under the author's cursor.
     */
    private insertReference(ref: string, span: { from: number; to: number }): void {
        const input = this.editingInput;
        if (undefined === input) return;

        const text = input.value;
        input.value = text.slice(0, span.from) + ref + text.slice(span.to);
        const caret = span.from + ref.length;
        input.setSelectionRange(caret, caret);
        input.focus();

        this.editingText.set(input.value);
        this.trackCaret(input);
    }

    // ---- Borders ---------------------------------------------------------
    //
    // A palette of glyphs, not a list of words. Both halves were reported at
    // once: a control that spells "All borders" cannot show what the line looks
    // like, and one that resets to its placeholder cannot show that the cell
    // already has it.

    /** Whether the palette is open. */
    protected readonly borderMenuOpen = signal(false);

    /** What the NEXT gesture draws -- synced from the cells when the palette opens. */
    protected readonly borderStyle = signal('thin');
    protected readonly borderColour = signal(DEFAULT_BORDER_COLOUR);

    /** The gestures that rule a whole selection. */
    protected readonly borderShapes: ReadonlyArray<{ value: BorderPreset; label: string; icon: string }> = [
        { value: 'all', label: 'All borders', icon: 'bi-border-all' },
        { value: 'outer', label: 'Outer border', icon: 'bi-border-outer' },
        { value: 'none', label: 'Clear borders', icon: 'bi-border' },
    ];

    /** The gestures that rule ONE side of it. */
    protected readonly borderSides: ReadonlyArray<{ value: BorderPreset; label: string; icon: string }> = [
        { value: 'top', label: 'Top', icon: 'bi-border-top' },
        { value: 'right', label: 'Right', icon: 'bi-border-right' },
        { value: 'bottom', label: 'Bottom', icon: 'bi-border-bottom' },
        { value: 'left', label: 'Left', icon: 'bi-border-left' },
    ];

    /** Only the lines the writer can express -- offering one it drops would lie. */
    protected readonly borderLines: ReadonlyArray<{ value: string; label: string }> =
        Object.keys(BORDER_STYLES).map(value => ({
            value,
            label: value.charAt(0).toUpperCase() + value.slice(1),
        }));

    /**
     * What the selection already carries.
     *
     * A computed rather than a method: it walks every cell of the selection and
     * a whole column is a legal selection. Read only from inside the open
     * palette, so a closed one costs nothing at all.
     */
    protected readonly borderState = computed<BorderState>(() => {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const range = this.selectionRange() ?? this.activeRef();

        return sheet && '' !== range
            ? borderStateIn(sheet, range)
            : { preset: null, style: null, colour: null };
    });

    protected toggleBorderMenu(): void {
        if (this.borderMenuOpen()) {
            this.closeBorderMenu();

            return;
        }

        // Open showing what the cells HAVE, so a red double rule is edited
        // rather than silently replaced by the thin black default.
        const state = this.borderState();
        if (null !== state.style) this.borderStyle.set(state.style);
        if (null !== state.colour) this.borderColour.set(state.colour);

        this.closeFilterMenu();
        this.closeOptions();
        this.closeContextMenu();
        this.closeFunctionMenu();
        this.borderMenuOpen.set(true);
    }

    protected closeBorderMenu(): void {
        this.borderMenuOpen.set(false);
    }

    /**
     * Escape belongs to the palette while it is open.
     *
     * Without intercepting it the CDK dialog takes it and the whole editor
     * closes -- the same trap the formula helper records.
     */
    protected onBorderMenuKey(event: KeyboardEvent): void {
        if ('Escape' !== event.key) return;
        event.preventDefault();
        event.stopPropagation();
        this.closeBorderMenu();
    }

    /**
     * Rule the selection.
     *
     * Acts on the SELECTION when there is one, so boxing a table is one
     * gesture; on the focused cell otherwise. The presets mean what they mean
     * in every other spreadsheet: `top` rules the top of the selection, not the
     * top of each cell in it.
     *
     * The palette stays OPEN afterwards: the pressed button moves to what the
     * gesture just produced, and that readout is the confirmation the author
     * gets that it landed.
     */
    protected applyBorders(preset: BorderPreset): void {
        const range = this.selectionRange() ?? this.activeRef();
        if ('' === range) return;

        this.mutateSheet(sheet => withBorderPreset(
            sheet,
            range,
            preset,
            this.borderStyle(),
            this.borderColour(),
        ));
    }

    /**
     * One cell's CSS border shorthands, as ONE object for a single binding.
     *
     * The same reason `styleAt` is one call and not six: this runs for every
     * rendered cell on every change detection, and a binding per edge would
     * re-read the cell four times and build four objects. Which edges those are
     * is `borderCssAt`'s problem -- and a shared one, since a line between two
     * cells belongs to neither.
     */
    protected borderAt(ref: string): Record<string, string> {
        const sheet = this.doc()?.sheets[this.activeSheet()];

        return sheet ? borderCssAt(sheet, ref) : {};
    }

    /**
     * EVERYTHING the cell's own `<td>` wears, as one object.
     *
     * ⚠️ One call and not seven, and this is the same lesson `styleAt` records
     * one screen up -- broken and re-learnt by measurement. The conditional
     * colour, weight, style and fill were four separate bindings, and each of
     * them called `lookAt`, which walks every conditional range on the sheet
     * AND re-derives the cell's displayed value. Four times, per cell, per
     * change-detection pass: measured at ~7ms of a 49ms pass over 588 cells,
     * for a question with one answer.
     *
     * A binding per property reads better in the template and costs a multiple
     * of the work in the only place this component cannot afford it.
     */
    protected cellStyleAt(ref: string): Record<string, string> {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const style: Record<string, string> = sheet ? { ...borderCssAt(sheet, ref) } : {};

        const look = this.lookAt(ref);
        const background = look?.background ?? sheet?.cells[ref]?.background;
        if (undefined !== background) style['background'] = background;
        if (undefined !== look?.color) style['color'] = look.color;
        if (true === look?.bold) style['font-weight'] = '700';
        if (true === look?.italic) style['font-style'] = 'italic';

        const valign = sheet?.cells[ref]?.valign;
        if (undefined !== valign) style['vertical-align'] = valign;

        const pointed = this.pointedAt(ref);
        if (null !== pointed) style['outline'] = `2px solid ${pointed}`;

        return style;
    }

    // ---- Function catalogue ----------------------------------------------
    //
    // Type-ahead only serves an author who already knows the name. Nobody
    // discovers SUMIF by typing SUMIF, which is why a browsable list is a
    // different feature and not a nicer version of the same one.

    protected readonly functionMenuOpen = signal(false);
    protected readonly functionQuery = signal('');

    /**
     * The catalogue, shelved, and narrowed by the search box.
     *
     * Matches on the NAME and on the summary both: an author looking for
     * "average" finds AVERAGE, and one looking for "condition" finds the whole
     * IF family without knowing a single one of their names -- which is the
     * entire point of a list you can browse.
     */
    protected readonly functionShelves = computed(() => {
        const query = this.functionQuery().trim().toLowerCase();
        const shelves = functionsByCategory();
        if ('' === query) return shelves;

        return shelves
            .map(shelf => ({
                category: shelf.category,
                functions: shelf.functions.filter(fn =>
                    fn.name.toLowerCase().includes(query)
                    || fn.summary.toLowerCase().includes(query)),
            }))
            .filter(shelf => shelf.functions.length > 0);
    });

    protected toggleFunctionMenu(): void {
        if (this.functionMenuOpen()) {
            this.closeFunctionMenu();

            return;
        }

        this.closeFilterMenu();
        this.closeOptions();
        this.closeContextMenu();
        this.closeBorderMenu();
        this.functionMenuOpen.set(true);
    }

    protected closeFunctionMenu(): void {
        this.functionMenuOpen.set(false);
        // The query is NOT kept: a list that reopens still filtered by what was
        // typed a quarter of an hour ago reads as a catalogue with five
        // functions in it.
        this.functionQuery.set('');
    }

    /** Escape belongs to the list while it is open -- see {@link onBorderMenuKey}. */
    protected onFunctionMenuKey(event: KeyboardEvent): void {
        if ('Escape' !== event.key) return;
        event.preventDefault();
        event.stopPropagation();
        this.closeFunctionMenu();
    }

    /**
     * Put a chosen function into the cell, with the caret between its brackets.
     *
     * Written straight into the DOM input and committed by the ordinary
     * `change`, exactly as a point-mode click is: this is mid-edit, and pushing
     * it through the document would make it an undo step of its own and
     * re-render the grid under the author's cursor.
     *
     * Inserts AT THE CARET when a cell is already being edited, so a function
     * can be dropped inside a formula that is half-written -- `=1+` then MAX
     * gives `=1+MAX(`, not a formula thrown away.
     */
    protected insertFunction(fn: FormulaFunction): void {
        // `activeRef` first and `editingInput` only as a fallback: the signal is
        // the authoritative selection, while the remembered element is a DOM
        // node the grid may since have recycled onto another cell as it
        // scrolled. They agree in the ordinary case and the signal is right in
        // the case where they do not.
        const input = this.activeInput() ?? this.editingInput;
        if (!input) return;

        this.closeFunctionMenu();

        // A cell holding a literal is REPLACED rather than appended to: "42"
        // with SUM dropped on the end would be `42SUM(`, which is not a formula
        // and not what anyone meant. A cell mid-formula keeps what it has.
        const editing = input.value.startsWith('=');
        const caret = editing ? (input.selectionStart ?? input.value.length) : 1;
        const before = editing ? input.value.slice(0, caret) : '=';
        const after = editing ? input.value.slice(input.selectionEnd ?? caret) : '';
        const insert = `${fn.name}(${0 === fn.maxArgs ? ')' : ''}`;

        input.value = before + insert + after;
        input.focus();
        const at = before.length + fn.name.length + 1;
        input.setSelectionRange(at, at);

        this.editingText.set(input.value);
        this.trackCaret(input);
    }

    /** The rendered control for the active cell, if the grid has one. */
    private activeInput(): HTMLInputElement | HTMLTextAreaElement | null {
        const ref = this.activeRef();

        return '' === ref ? null : this.inputFor(ref);
    }

    /** The rendered control for a cell, if the grid has drawn one. */
    private inputFor(ref: string): HTMLInputElement | HTMLTextAreaElement | null {
        const el = this.host.nativeElement.querySelector(`[aria-label="${ref}"]`);

        // A checkbox cell has a control and no text to put a formula in.
        return el instanceof HTMLTextAreaElement
            || (el instanceof HTMLInputElement && 'checkbox' !== el.type)
            ? el
            : null;
    }

    // ---- Context menu ----------------------------------------------------
    //
    // The menu exists because of what is UNDER it: until now there was no way
    // to insert or delete a row or column at all. Every other entry here also
    // has a toolbar control -- a context menu that only duplicated the toolbar
    // would be decoration.

    /** The cell whose menu is open, if any. */
    protected readonly contextRef = signal<string | null>(null);

    /**
     * Whether the menu opens UPWARD from the cell.
     *
     * The grid body is a scroll container with `overflow`, so a menu anchored
     * below a cell near the bottom is CLIPPED by it -- and the entries that get
     * cut are the ones furthest down the list, which is how "Clear contents"
     * through "Checkbox" became unreachable on row 15 of a short sheet.
     */
    protected readonly contextAbove = signal(false);

    protected openContextMenu(ref: string, event: MouseEvent): void {
        event.preventDefault();
        // Focus the cell first, so every action in the menu acts on the cell
        // the author actually right-clicked rather than on whatever the last
        // left-click happened to leave selected.
        this.focusCell(ref);
        this.closeFilterMenu();
        this.closeOptions();
        this.closeBorderMenu();
        this.closeFunctionMenu();
        this.contextAbove.set(false);
        this.contextRef.set(ref);
        this.placeContextMenu();
    }

    /**
     * Flip the menu above the cell when it would not fit below.
     *
     * MEASURED on the next frame rather than estimated from an item count:
     * the menu's height depends on which entries are present -- "Remove form
     * element" appears only where a rule applies -- and on whatever the
     * browser's own font metrics make of the labels. The same rule the rest of
     * this component follows for row heights and the viewport: ask the
     * element, do not model it.
     *
     * Flips only when flipping HELPS. A menu taller than the whole grid body
     * clips either way, and turning it upside down would just move the
     * unreachable half from the bottom to the top.
     */
    private placeContextMenu(): void {
        requestAnimationFrame(() => {
            const menu = this.host.nativeElement.querySelector('.sheet-editor__context');
            const body = this.host.nativeElement.querySelector('.sheet-editor__body');
            if (!(menu instanceof HTMLElement) || !(body instanceof HTMLElement)) return;

            const menuBox = menu.getBoundingClientRect();
            const bodyBox = body.getBoundingClientRect();
            const roomAbove = menuBox.top - bodyBox.top;

            this.contextAbove.set(menuBox.bottom > bodyBox.bottom && menuBox.height <= roomAbove);
        });
    }

    protected closeContextMenu(): void {
        this.contextRef.set(null);
        this.contextAbove.set(false);
    }

    /** The row and column the menu is acting on. */
    private contextCell(): { column: string; row: number } | null {
        const ref = this.contextRef();

        return null === ref ? null : parseRef(ref);
    }

    protected contextRowLabel(): string {
        return String(this.contextCell()?.row ?? '');
    }

    protected contextColumnLabel(): string {
        return this.contextCell()?.column ?? '';
    }

    protected insertRow(where: 'above' | 'below'): void {
        const cell = this.contextCell();
        if (null === cell) return;
        this.mutateSheet(sheet => withInsertedRow(sheet, 'above' === where ? cell.row : cell.row + 1));
    }

    protected insertColumn(where: 'left' | 'right'): void {
        const cell = this.contextCell();
        if (null === cell) return;
        const index = columnToIndex(cell.column);
        this.mutateSheet(sheet => withInsertedColumn(sheet, indexToColumn('left' === where ? index : index + 1)));
    }

    protected deleteRow(): void {
        const cell = this.contextCell();
        if (null === cell) return;
        this.mutateSheet(sheet => withDeletedRow(sheet, cell.row));
    }

    protected deleteColumn(): void {
        const cell = this.contextCell();
        if (null === cell) return;
        this.mutateSheet(sheet => withDeletedColumn(sheet, cell.column));
    }

    /** Contents only — the layout around the cells is not what was selected. */
    protected clearContents(): void {
        const range = this.selectionRange() ?? this.contextRef();
        if (null === range) return;
        this.mutateSheet(sheet => withClearedRange(sheet, range));
    }

    /**
     * Apply a pure transform to the active sheet.
     *
     * Every structural action goes through here so the close-the-menu, write,
     * mark-dirty sequence exists once. A transform that changes nothing returns
     * the same object and is treated as a no-op, which keeps the dirty flag
     * honest.
     */
    // ---- Conditional formatting ---------------------------------------------
    //
    // The last thing in this editor that changes the DOCUMENT rather than the
    // editing of it: "overdue in red" is decided by the author at template
    // time, because nobody is going to open forty generated invoices and
    // colour the late lines in by hand.

    protected readonly conditionalWhens = CONDITIONAL_WHENS;

    protected readonly rulesOpen = signal(false);
    protected readonly ruleWhen = signal<ConditionalWhen>('greaterThan');
    protected readonly ruleValue = signal('');
    protected readonly ruleValue2 = signal('');
    protected readonly ruleBackground = signal('#FFF3CD');
    protected readonly ruleColor = signal('#842029');
    protected readonly ruleBold = signal(false);
    protected readonly ruleItalic = signal(false);

    protected ruleNeedsValue(): boolean {
        return needsValue(this.ruleWhen());
    }

    protected ruleNeedsSecondValue(): boolean {
        return needsSecondValue(this.ruleWhen());
    }

    /** The rules already covering the selected cell. */
    protected readonly rulesHere = computed(() =>
        conditionalsAt(this.doc()?.sheets[this.activeSheet()], this.activeRef()));

    protected describeRule(rule: { when: ConditionalWhen; value?: string; value2?: string }): string {
        const label = CONDITIONAL_WHENS.find(w => w.value === rule.when)?.label ?? rule.when;
        if (!needsValue(rule.when)) return label.toLowerCase();

        return needsSecondValue(rule.when)
            ? `${label.toLowerCase()} ${rule.value} and ${rule.value2}`
            : `${label.toLowerCase()} ${rule.value}`;
    }

    protected toggleRules(): void {
        this.rulesOpen.set(!this.rulesOpen());
    }

    protected closeRules(): void {
        this.rulesOpen.set(false);
    }

    /** Escape belongs to the panel while it is open -- see {@link onBorderMenuKey}. */
    protected onRulesKey(event: KeyboardEvent): void {
        if ('Escape' !== event.key) return;
        event.preventDefault();
        event.stopPropagation();
        this.closeRules();
    }

    /**
     * Put the rule being described onto the selection.
     *
     * On the SELECTION when there is one, so a whole column is one gesture --
     * the same rule every other formatting control here follows.
     */
    protected addRule(): void {
        const range = this.selectionRange() ?? this.activeRef();
        if ('' === range) return;

        const when = this.ruleWhen();
        if (needsValue(when) && '' === this.ruleValue().trim()) return;
        if (needsSecondValue(when) && '' === this.ruleValue2().trim()) return;

        this.mutateSheet(sheet => withConditional(sheet, range, {
            when,
            value: needsValue(when) ? this.ruleValue().trim() : undefined,
            value2: needsSecondValue(when) ? this.ruleValue2().trim() : undefined,
            background: this.ruleBackground().toUpperCase(),
            color: this.ruleColor().toUpperCase(),
            bold: this.ruleBold(),
            italic: this.ruleItalic(),
        }));
    }

    protected removeRule(range: string, index: number): void {
        this.mutateSheet(sheet => withoutConditional(sheet, range, index));
    }

    /**
     * The look a rule gives one cell, against what that cell SHOWS.
     *
     * Its displayed value, so a formula is judged by its result -- "over 90"
     * means the computed 95, and reading `=B2*C2` as text would never match a
     * number at all. The opposite of what find-and-replace searches, and both
     * are right: that one is about content, this is about appearance once the
     * content has been worked out.
     */
    private lookAt(ref: string): { background?: string; color?: string; bold?: boolean; italic?: boolean } | null {
        const sheet = this.doc()?.sheets[this.activeSheet()];

        return undefined === sheet?.conditionals ? null : lookFor(sheet, ref, this.valueAt(ref));
    }

    protected conditionalColour(ref: string): string | null {
        return this.lookAt(ref)?.color ?? null;
    }

    protected conditionalBold(ref: string): boolean {
        return true === this.lookAt(ref)?.bold;
    }

    protected conditionalItalic(ref: string): boolean {
        return true === this.lookAt(ref)?.italic;
    }

    // ---- Find and replace --------------------------------------------------
    //
    // Over what a cell HOLDS, not what it computes -- see `find-replace.ts`.
    // The job this exists for is renaming a `{var:…}` token across a template,
    // and a token has no computed value to search.

    protected readonly findOpen = signal(false);
    protected readonly findQuery = signal('');
    protected readonly replaceWith = signal('');
    protected readonly matchCase = signal(false);
    protected readonly wholeCell = signal(false);

    private findOptions(): FindOptions {
        return { matchCase: this.matchCase(), wholeCell: this.wholeCell() };
    }

    /**
     * Every matching cell on the active sheet.
     *
     * A computed, so it re-runs when the document changes -- which is what
     * makes "replace" leave the remaining matches highlighted and the count
     * correct without anything having to remember to refresh it.
     */
    protected readonly findMatchesHere = computed(() => findMatches(
        this.doc()?.sheets[this.activeSheet()],
        this.findQuery(),
        { matchCase: this.matchCase(), wholeCell: this.wholeCell() },
    ));

    protected findCount(): number {
        return this.findMatchesHere().length;
    }

    protected findSummary(): string {
        if ('' === this.findQuery()) return '';
        const count = this.findCount();

        return 0 === count ? 'No matches' : `${count} cell${1 === count ? '' : 's'}`;
    }

    protected isMatch(ref: string): boolean {
        return this.findMatchesHere().includes(ref);
    }

    /**
     * Ctrl+F opens the box and puts the caret in it -- and OPENS it, never
     * closes it. Toggling on the shortcut would mean an author who pressed it
     * twice, as anyone does when a page has not responded yet, ended up with it
     * shut again.
     */
    protected onFindShortcut(event: Event): void {
        // The browser's own find bar would otherwise open over a grid it cannot
        // search: the cells are inputs, and their text is not in the page.
        event.preventDefault();
        if (!this.findOpen()) {
            this.toggleFind();

            return;
        }

        this.focusFindInput();
    }

    private focusFindInput(): void {
        requestAnimationFrame(() => {
            const input = this.host.nativeElement.querySelector('input[aria-label="Find"]');
            if (input instanceof HTMLInputElement) input.select();
        });
    }

    protected toggleFind(): void {
        if (this.findOpen()) {
            this.closeFind();

            return;
        }

        this.findOpen.set(true);
        // Focus on the next frame: the input does not exist until the panel has
        // been rendered, and a find box that has to be clicked before it can be
        // typed into is a find box nobody uses twice.
        this.focusFindInput();
    }

    protected closeFind(): void {
        this.findOpen.set(false);
        // The query goes with it, so the highlights do too -- leaving a sheet
        // painted with matches of a search nobody is running any more.
        this.findQuery.set('');
    }

    /** Enter finds the next; Escape closes, and must not reach the dialog. */
    protected onFindKey(event: KeyboardEvent): void {
        if ('Escape' === event.key) {
            event.preventDefault();
            event.stopPropagation();
            this.closeFind();

            return;
        }
        if ('Enter' !== event.key) return;

        event.preventDefault();
        this.findNext();
    }

    /** Go to the match after the selected cell, wrapping at the end. */
    protected findNext(): void {
        const target = nextMatch(this.findMatchesHere(), this.activeRef());
        if (null === target) return;

        this.extendingFor = null;
        this.focusCell(target);
        requestAnimationFrame(() => this.inputFor(target)?.focus());
    }

    /**
     * Replace in the selected cell, then move on.
     *
     * Replaces where the author is LOOKING -- the selected cell when it is a
     * match, otherwise the next one, which is what makes repeated pressing walk
     * the sheet rather than needing Next between every two.
     */
    protected replaceOne(): void {
        const matches = this.findMatchesHere();
        const at = matches.includes(this.activeRef()) ? this.activeRef() : nextMatch(matches, this.activeRef());
        if (null === at) return;

        const query = this.findQuery();
        const replacement = this.replaceWith();
        const options = this.findOptions();
        this.mutateSheet(sheet => withReplacedIn(sheet, at, query, replacement, options));

        this.findNext();
    }

    protected replaceAll(): void {
        const query = this.findQuery();
        const replacement = this.replaceWith();
        const options = this.findOptions();

        this.mutateSheet(sheet => withReplacedAll(sheet, query, replacement, options).sheet);
    }

    // ---- Undo and redo -----------------------------------------------------
    //
    // Snapshots of the whole document rather than a log of reversible actions.
    // A `.dsheet` is sparse and small -- a template, not a database -- so a
    // clone costs less than the machinery for inverting a dozen kinds of edit,
    // and it cannot get an inverse subtly wrong. The depth is capped because
    // the cost is memory rather than time.

    private static readonly UNDO_DEPTH = 50;

    private readonly undoStack: SheetDocumentDto[] = [];
    private readonly redoStack: SheetDocumentDto[] = [];

    /** Depths, as signals, because the toolbar buttons read them. */
    protected readonly undoDepth = signal(0);
    protected readonly redoDepth = signal(0);

    /**
     * The document as it stood after the LAST commit, cloned.
     *
     * ⚠️ Kept continuously rather than taken at the start of each edit, and
     * that is what makes one funnel possible. Several write paths mutate
     * `doc.sheets[...]` or `sheet.cells[...]` IN PLACE and only then publish a
     * shallow copy -- so by the time a commit runs, the "current" document has
     * already changed and a snapshot taken there would record the new state as
     * though it were the old one. Holding the previous state instead means the
     * eleven write sites need to know nothing about history.
     */
    private baseline: SheetDocumentDto | null = null;

    /** Publish a change, remembering what it replaced. */
    private commit(next: SheetDocumentDto): void {
        if (null !== this.baseline) {
            this.undoStack.push(this.baseline);
            if (this.undoStack.length > SheetEditorDialogComponent.UNDO_DEPTH) this.undoStack.shift();
            // A new change abandons the branch that was undone away from.
            this.redoStack.length = 0;
        }

        this.doc.set(next);
        this.baseline = structuredClone(next);
        this.dirty.set(true);
        this.tellHistory();
    }

    /** Start a fresh history: this document is where undo stops. */
    private rebase(doc: SheetDocumentDto): void {
        this.undoStack.length = 0;
        this.redoStack.length = 0;
        this.baseline = structuredClone(doc);
        this.tellHistory();
    }

    private tellHistory(): void {
        this.undoDepth.set(this.undoStack.length);
        this.redoDepth.set(this.redoStack.length);
    }

    /**
     * Step back one change.
     *
     * ⚠️ `dirty` is NOT cleared by undoing back to the loaded document. It
     * could be -- count the steps -- but a flag that says "saved" while the
     * file on disk might differ is the one failure worth avoiding, and undoing
     * to the start then saving costs nothing.
     */
    protected undo(): void {
        const target = this.undoStack.pop();
        const current = this.baseline;
        if (undefined === target || null === current) return;

        this.redoStack.push(current);
        this.doc.set(target);
        this.baseline = structuredClone(target);
        this.afterHistory();
    }

    protected redo(): void {
        const target = this.redoStack.pop();
        const current = this.baseline;
        if (undefined === target || null === current) return;

        this.undoStack.push(current);
        this.doc.set(target);
        this.baseline = structuredClone(target);
        this.afterHistory();
    }

    /** What both ends of the history have to put right afterwards. */
    private afterHistory(): void {
        this.dirty.set(true);
        this.tellHistory();
        // The helper describes text that may no longer be there, and a cell
        // being edited has just had its content replaced underneath it.
        this.dismissHelper();
        this.editing = false;
        this.closeContextMenu();
    }

    /** Whether the names panel is showing. */
    protected readonly namesOpen = signal(false);

    /** What the author is typing into it. */
    protected readonly newName = signal('');

    /** The names this document declares, sorted, for the panel's list. */
    protected readonly definedNames = computed(() => definedNamesOf(this.doc()));

    /**
     * Why the typed name cannot be used, or null.
     *
     * ⚠️ Recomputed against the names ALREADY declared, so "already used" is
     * answered before the button is pressed rather than after -- and the button
     * is disabled on the same signal, so the two cannot disagree.
     */
    protected readonly nameError = computed(
        () => nameProblem(this.newName(), this.definedNames().map(n => n.name)));

    protected toggleNames(): void {
        this.namesOpen.update(open => !open);
        if (!this.namesOpen()) this.newName.set('');
    }

    protected closeNames(): void {
        this.namesOpen.set(false);
        this.newName.set('');
    }

    /**
     * Declare the typed name for whatever is selected.
     *
     * The RANGE is the selection, not the active cell: naming a band's column
     * is the reason this exists, and that is a range every time.
     */
    protected addName(): void {
        if (null !== this.nameError()) return;

        const range = this.selectionRange() ?? this.activeRef();
        if ('' === range) return;

        const at = scopedRange(this.activeSheet(), range);
        this.mutateDoc(doc => withDefinedName(doc, this.newName(), at));
        this.newName.set('');
    }

    protected removeName(name: string): void {
        this.mutateDoc(doc => withoutDefinedName(doc, name));
    }

    /**
     * A DOCUMENT-level change, as {@see mutateSheet} is a sheet-level one.
     *
     * Defined names live on the workbook rather than a sheet, so they cannot go
     * through `mutateSheet` -- and going through `commit` directly would skip
     * the "nothing actually changed" check that keeps undo free of no-ops.
     */
    private mutateDoc(transform: (doc: SheetDocumentDto) => SheetDocumentDto): void {
        const doc = this.doc();
        if (!doc) return;

        const next = transform(doc);
        if (next === doc) return;

        this.commit(next);
    }

    private mutateSheet(transform: (sheet: SheetDto) => SheetDto): void {
        this.closeContextMenu();

        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        if (!doc || !sheet) return;

        const next = transform(sheet);
        if (next === sheet) return;

        doc.sheets[this.activeSheet()] = next;
        this.commit({ ...doc });
    }

    // ---- Form elements ---------------------------------------------------
    //
    // A generated workbook is often not the end of the journey: it goes to
    // someone who FILLS IT IN. These are the controls that say what may go in.
    //
    // The editor ships the two that Google Sheets' own Insert menu calls form
    // elements -- Dropdown and Tick box. The model expresses more (date, whole,
    // decimal, textLength with operators and bounds) and the writer emits all
    // of them, so a template can carry a numeric rule today; only the authoring
    // UI for those is not built yet.

    /** The rule covering the focused cell, with the range it came from. */
    protected validationAtActive(): { range: string; rule: SheetValidationDto } | null {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const ref = this.activeRef();
        if (!sheet || '' === ref) return null;
        const range = validationRangeAt(sheet, ref);

        return null === range ? null : { range, rule: sheet.validations![range] };
    }

    /** The rule covering any cell, for the grid to draw. */
    private ruleAt(ref: string): SheetValidationDto | null {
        return this.lookup().validationAt(ref);
    }

    protected isCheckbox(ref: string): boolean {
        return this.ruleAt(ref)?.type === 'checkbox';
    }

    /** A dropdown cell draws an arrow; a checkbox draws its own tick instead. */
    protected hasOptions(ref: string): boolean {
        const rule = this.ruleAt(ref);

        return null !== rule && rule.type === 'list';
    }

    protected isChecked(ref: string): boolean {
        return 'TRUE' === (this.doc()?.sheets[this.activeSheet()]?.cells[ref]?.value ?? '');
    }

    /** The options the cell's dropdown offers. */
    protected optionsAt(ref: string): string[] {
        const rule = this.ruleAt(ref);

        return null === rule ? [] : validationOptions(rule);
    }

    /** Which cell's option list is open, if any. */
    protected readonly openOptionsRef = signal<string | null>(null);

    protected toggleOptions(ref: string, event: MouseEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.openOptionsRef.update(open => (open === ref ? null : ref));
    }

    protected closeOptions(): void {
        this.openOptionsRef.set(null);
    }

    /** Pick an option, which is an ordinary edit -- undo and save see it. */
    protected chooseOption(ref: string, value: string): void {
        this.closeOptions();
        this.edit(ref, value);
    }

    protected toggleCheckbox(ref: string, checked: boolean): void {
        this.edit(ref, checked ? 'TRUE' : 'FALSE');
    }

    /**
     * The toolbar's one control for all of this: pick a type, or remove.
     *
     * A select rather than a row of buttons because the toolbar already wraps
     * to two lines, and because these are alternatives -- a range has one rule
     * or none. Resets itself afterwards so it always reads "Form element"
     * rather than showing a stale choice as if it were the cell's state.
     */
    protected async insertControl(choice: string): Promise<void> {
        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        // The SELECTION when there is one, so a whole column of a table gets
        // the dropdown in one action -- otherwise just the focused cell.
        const range = this.selectionRange() ?? this.activeRef();
        if (!doc || !sheet || '' === range) return;

        let next: SheetDto | null = null;

        if ('remove' === choice) {
            const existing = this.validationAtActive();
            next = null === existing ? null : withoutValidation(sheet, existing.range);
        } else if ('checkbox' === choice) {
            next = withValidation(sheet, range, { type: 'checkbox' });
            // A checkbox over cells holding nothing reads as unticked already;
            // writing FALSE makes the cell's value match what is drawn, which
            // is what the generated workbook will carry.
            for (const ref of refsInRange(range.includes(':') ? range : `${range}:${range}`)) {
                if (undefined === next.cells[ref]?.value) {
                    next = { ...next, cells: { ...next.cells, [ref]: { ...next.cells[ref], value: 'FALSE' } } };
                }
            }
        } else if ('list' === choice) {
            const typed = await this.dialogs.input({
                title: 'Dropdown options',
                label: 'One per line',
                multiline: true,
                placeholder: 'New\nOpen\nClosed',
                initialValue: (this.validationAtActive()?.rule.values ?? []).join('\n'),
                required: true,
                // ⚠️ A comma cannot be escaped in an OOXML inline list, so an
                // option containing one would silently become TWO options in
                // the generated workbook. Refused here, where the author can
                // still fix it, rather than dropped silently by the backend.
                validator: value => {
                    const options = this.splitOptions(value);
                    if (0 === options.length) return 'Give at least one option.';
                    if (options.some(o => o.includes(','))) return 'An option cannot contain a comma.';
                    if (options.join(',').length > 255) return 'Too long — Excel allows 255 characters of options.';

                    return null;
                },
            });
            if (null === typed) return;
            next = withValidation(sheet, range, { type: 'list', values: this.splitOptions(typed) });
        }

        if (null === next || next === sheet) return;

        doc.sheets[this.activeSheet()] = next;
        this.commit({ ...doc });
    }

    /** One option per line, blanks and duplicates removed — as the model does. */
    private splitOptions(raw: string): string[] {
        const out: string[] = [];
        for (const line of raw.split('\n')) {
            const value = line.trim();
            if ('' !== value && !out.includes(value)) out.push(value);
        }

        return out;
    }

    // ---- Full screen -----------------------------------------------------

    /**
     * Whether the dialog fills the viewport.
     *
     * A grid is the surface that most wants the room: the default 1200px shows
     * about eight columns of a sheet that has twenty-six, so an author working
     * a wide table spends the session scrolling sideways.
     *
     * Not the browser Fullscreen API -- see the `.sheet-editor--full` rule for
     * why. This is CSS, so it cannot be refused and needs no gesture.
     */
    protected readonly fullScreen = signal(false);

    protected toggleFullScreen(): void {
        this.fullScreen.update(on => !on);

        // ⚠️ Load-bearing. Both virtualisation windows are computed from a
        // MEASURED viewport that is otherwise only taken on load and on
        // scroll. Resizing without re-measuring leaves the grid rendering the
        // old, smaller window -- blank rows below and missing columns to the
        // right -- until the author happens to scroll, which on a sheet that
        // now fits entirely on screen may be never.
        this.measureViewport();
    }

    /**
     * F11 toggles it, and the browser's own full screen is suppressed while
     * the dialog has focus.
     *
     * Taking over a browser shortcut is worth justifying: inside a modal
     * spreadsheet, F11 meaning "make this editor bigger" is what every office
     * suite does, and the browser's version would leave the dialog exactly the
     * same 1200px it was, only with more black around it.
     */
    protected onFullScreenKey(event: Event): void {
        event.preventDefault();
        this.toggleFullScreen();
    }

    /** A range to filter, or a filter to remove — either enables the button. */
    protected canFilter(): boolean {
        return null !== this.selectionRange() || null !== this.activeFilter();
    }

    /** The filter declared on the active sheet, if any. */
    protected activeFilter(): string | null {
        const sheet = this.doc()?.sheets[this.activeSheet()];

        return sheet ? autoFilterOf(sheet) : null;
    }

    /** Column letter => the display values hidden in it. View state only. */
    private readonly filterExclusions = signal<ReadonlyMap<string, ReadonlySet<string>>>(new Map());

    /** Which column's dropdown is open, if one is. */
    protected readonly openFilterColumn = signal<string | null>(null);

    /**
     * Rows the current exclusions hide.
     *
     * Derived rather than stored so an EDIT re-answers it: change a cell to a
     * value that is filtered out and the row leaves, exactly as a spreadsheet
     * behaves. A stored set of row numbers would go stale the moment the sheet
     * changed under it.
     */
    private readonly hiddenRows = computed<ReadonlySet<number>>(() => {
        const range = this.activeFilter();
        const exclusions = this.filterExclusions();
        const hidden = new Set<number>();
        if (null === range || 0 === exclusions.size) return hidden;

        for (const row of filterBodyRows(range)) {
            for (const [column, excluded] of exclusions) {
                if (excluded.has(this.displayAt(column + row))) {
                    hidden.add(row);
                    break;
                }
            }
        }

        return hidden;
    });

    protected isRowHidden(row: number): boolean {
        return this.hiddenRows().has(row);
    }

    /**
     * What a cell shows for FILTERING purposes.
     *
     * {@link valueAt} swaps the focused cell to its formula text so it can be
     * edited; a filter must not see that swap, or focusing a cell would change
     * which value it counts as and rows would move under the cursor.
     */
    private displayAt(ref: string): string {
        const cell = this.doc()?.sheets[this.activeSheet()]?.cells[ref];
        if (cell?.formula === undefined) return cellToInput(cell);
        const value = this.evaluated().get(ref);

        return value === undefined || value.kind === 'unresolved'
            ? cellToInput(cell)
            : displayValue(value);
    }

    /** Create a filter over the selection, or remove the one already there. */
    protected toggleFilter(): void {
        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        if (!doc || !sheet) return;

        const existing = autoFilterOf(sheet);
        const range = this.selectionRange();

        let next: SheetDto;
        if (null !== existing) {
            next = withoutAutoFilter(sheet);
        } else if (null !== range) {
            next = withAutoFilter(sheet, range);
        } else {
            return;
        }

        if (next === sheet) return;

        // Removing the filter must clear what it was hiding. A row left hidden
        // by a filter that no longer exists is unreachable -- there would be no
        // dropdown left to bring it back.
        this.filterExclusions.set(new Map());
        this.openFilterColumn.set(null);

        doc.sheets[this.activeSheet()] = next;
        this.commit({ ...doc });
    }

    /** Whether this cell carries one of the filter's dropdown buttons. */
    protected hasFilterButton(ref: string): boolean {
        const sheet = this.doc()?.sheets[this.activeSheet()];

        return !!sheet && isFilterHeader(sheet, ref);
    }

    /** True when the column's dropdown is hiding something -- the button lights up. */
    protected isColumnFiltered(column: string): boolean {
        return (this.filterExclusions().get(column)?.size ?? 0) > 0;
    }

    protected toggleFilterMenu(column: string, event: MouseEvent): void {
        // The button sits inside a header CELL whose input would otherwise take
        // focus and collapse the selection, the same reason `extendTo` and the
        // header handlers call this.
        event.preventDefault();
        event.stopPropagation();
        this.openFilterColumn.update(open => (open === column ? null : column));
    }

    protected closeFilterMenu(): void {
        this.openFilterColumn.set(null);
    }

    /**
     * The distinct values in a column's body, in the order they first appear,
     * each with whether it is currently shown.
     *
     * First-appearance order rather than sorted: these are template cells, so
     * the list is short and an author recognises their own rows faster by
     * position than alphabetically. A blank shows as `(blank)` because an empty
     * label would be an unclickable checkbox.
     */
    protected filterValues(column: string): Array<{ value: string; label: string; shown: boolean }> {
        const range = this.activeFilter();
        if (null === range) return [];

        const excluded = this.filterExclusions().get(column) ?? new Set<string>();
        const seen = new Set<string>();
        const out: Array<{ value: string; label: string; shown: boolean }> = [];

        for (const row of filterBodyRows(range)) {
            const value = this.displayAt(column + row);
            if (seen.has(value)) continue;
            seen.add(value);
            out.push({ value, label: '' === value ? '(blank)' : value, shown: !excluded.has(value) });
        }

        return out;
    }

    protected toggleFilterValue(column: string, value: string): void {
        this.filterExclusions.update(current => {
            const next = new Map(current);
            const excluded = new Set(next.get(column) ?? []);
            if (excluded.has(value)) {
                excluded.delete(value);
            } else {
                excluded.add(value);
            }
            if (0 === excluded.size) {
                next.delete(column);
            } else {
                next.set(column, excluded);
            }

            return next;
        });
    }

    /** Clear one column's exclusions -- the "Select all" of a filter dropdown. */
    protected showAllInColumn(column: string): void {
        this.filterExclusions.update(current => {
            const next = new Map(current);
            next.delete(column);

            return next;
        });
    }

    /** How many rows the filter is hiding, for the toolbar's own hint. */
    protected hiddenRowCount(): number {
        return this.hiddenRows().size;
    }

    /** The columns whose headers carry a button, for the filter menu's lookup. */
    protected filterColumnsOf(): string[] {
        const range = this.activeFilter();

        return null === range ? [] : filterColumns(range);
    }

    /** The merge the focused cell belongs to, if any. */
    protected mergeAtActive(): string | null {
        const sheet = this.doc()?.sheets[this.activeSheet()];

        return sheet && '' !== this.activeRef() ? mergeCovering(sheet, this.activeRef()) : null;
    }

    /**
     * The merges and rules of the sheet on screen, indexed.
     *
     * ⚠️ Held in a `computed` so its LIFETIME is the document's: every write
     * commits a new document object, which drops this and builds a fresh one.
     * A lookup kept across an edit would draw merges the sheet no longer has.
     * {@link SheetLookup} has the measurements that made it necessary.
     */
    private readonly lookup = computed(() => sheetLookup(this.doc()?.sheets[this.activeSheet()]));

    protected isMerged(ref: string): boolean {
        return null !== this.lookup().mergeCovering(ref);
    }

    /** A cell swallowed by a merge it does not anchor renders no `<td>` at all. */
    protected isCovered(ref: string): boolean {
        const range = this.lookup().mergeCovering(ref);

        return null !== range && !isMergeAnchor(range, ref);
    }

    protected spanAt(ref: string): { colspan: number; rowspan: number } {
        const range = this.lookup().mergeCovering(ref);

        return null !== range && isMergeAnchor(range, ref) ? mergeSpan(range) : { colspan: 1, rowspan: 1 };
    }

    /**
     * Highlight for the cells a pending merge would cover.
     *
     * A bounds test rather than `refsInRange(...).includes(...)`: this runs for
     * every rendered cell on every change detection, and selecting a whole
     * column — one click on a header — would otherwise make it quadratic in the
     * sheet's size.
     */
    protected isInSelection(ref: string): boolean {
        const range = this.selectionRange();

        return null !== range && rangeContains(range, ref);
    }

    /**
     * Select a whole column from its header.
     *
     * `preventDefault` keeps the mousedown from moving focus into a cell input,
     * whose `focus` handler would immediately collapse the selection it just
     * made — the same interaction that made shift-click load-bearing above.
     *
     * `activeRef` takes the TOP cell and the anchor the bottom, so the toolbar
     * reads "A1" rather than "A20" while covering the identical range;
     * `rangeBetween` normalises either way.
     */
    protected selectColumn(column: string, event: MouseEvent): void {
        event.preventDefault();
        const from = event.shiftKey ? (parseRef(this.anchorRef())?.column ?? column) : column;
        this.anchorRef.set(from + this.extent().rows);
        this.activeRef.set(column + 1);
    }

    /** The row equivalent, spanning every column the grid shows. */
    protected selectRow(row: number, event: MouseEvent): void {
        event.preventDefault();
        const from = event.shiftKey ? (parseRef(this.anchorRef())?.row ?? row) : row;
        this.anchorRef.set(indexToColumn(this.extent().cols) + from);
        this.activeRef.set('A' + row);
    }

    /** The corner selects the whole grid, as it does in every spreadsheet. */
    protected selectAll(event: MouseEvent): void {
        event.preventDefault();
        const { rows, cols } = this.extent();
        this.anchorRef.set(indexToColumn(cols) + rows);
        this.activeRef.set('A1');
    }

    /**
     * Whether the selection covers a column ENTIRELY.
     *
     * Deliberately not "touches": the header takes a solid highlight, and using
     * it for a two-cell range that happens to sit in this column would claim a
     * selection the author does not have.
     */
    protected isColumnSelected(column: string): boolean {
        const box = this.selectionBox();
        if (!box) return false;
        const index = columnToIndex(column);

        return index >= box.left && index <= box.right
            && box.top <= 1 && box.bottom >= this.extent().rows;
    }

    protected isRowSelected(row: number): boolean {
        const box = this.selectionBox();

        return !!box && row >= box.top && row <= box.bottom
            && box.left <= 1 && box.right >= this.extent().cols;
    }

    private selectionBox(): ReturnType<typeof parseRange> {
        const range = this.selectionRange();

        return null === range ? null : parseRange(range);
    }

    /**
     * Drag a column's right edge to resize it.
     *
     * The starting width is MEASURED from the header rather than read from the
     * document: a column with no stored width still renders at the stylesheet's
     * default, and starting the drag from "undefined" would make the first
     * mousemove jump the column to whatever the pointer offset happened to be.
     *
     * Listeners go on the document, not the grip — a fast drag outranges a 6px
     * target long before the mouse comes up, and a resize that stops tracking
     * because the pointer left the handle is worse than no resize at all.
     */
    protected startResize(column: string, event: MouseEvent): void {
        // The grip lives INSIDE the header, whose own mousedown selects the
        // column. Without this a resize would also select it.
        event.stopPropagation();
        event.preventDefault();

        const header = (event.target as HTMLElement | null)?.parentElement;
        // SCREEN pixels on both counts -- see {@link startRowResize} for why
        // they are divided by the zoom before they reach a stored width.
        const zoom = this.zoom();
        const startPx = (header?.getBoundingClientRect().width ?? 0) / zoom;
        const startX = event.clientX;

        const move = (moved: MouseEvent): void =>
            this.writeColumnWidth(column, columnWidthFromPx(startPx + (moved.clientX - startX) / zoom));

        const up = (): void => {
            document.removeEventListener('mousemove', move);
            document.removeEventListener('mouseup', up);
            this.cancelResize = null;
        };

        this.cancelResize?.();
        this.cancelResize = up;
        document.addEventListener('mousemove', move);
        document.addEventListener('mouseup', up);
    }

    /** Torn down on destroy: a dialog closed mid-drag must not leave listeners on the document. */
    private cancelResize: (() => void) | null = null;

    /**
     * Merge the selection, or unmerge the merge under the cursor.
     *
     * Unmerge wins when the focused cell is already merged: with a collapsed
     * selection there is nothing to merge, and it is the only way back.
     */
    protected toggleMerge(): void {
        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        if (!doc || !sheet) return;

        const existing = this.mergeAtActive();
        const range = this.selectionRange();

        let next: SheetDto;
        if (null !== existing) {
            next = withoutMerge(sheet, existing);
        } else if (null !== range) {
            next = withMerge(sheet, range);
        } else {
            return;
        }

        if (next === sheet) return;

        doc.sheets[this.activeSheet()] = next;
        this.commit({ ...doc });
        this.anchorRef.set(this.activeRef());
    }

    /** The column the width control acts on — taken from the focused cell. */
    protected activeColumn(): string {
        return parseRef(this.activeRef())?.column ?? '';
    }

    /** Empty when the column has no stored width, so the input shows a placeholder. */
    protected activeWidth(): string {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const column = this.activeColumn();
        if (!sheet || '' === column) return '';

        return String(columnWidthOf(sheet, column) ?? '');
    }

    /**
     * Width for one column's `<col>`, in px, or null to let the table size it.
     *
     * Rendering this is what makes the control mean something: a stored width
     * the grid ignored would look like it had not been saved.
     */
    protected columnPx(column: string): number {
        const sheet = this.doc()?.sheets[this.activeSheet()];
        const width = sheet ? columnWidthOf(sheet, column) : undefined;

        // Never null since #2068: the window is computed from these widths, so
        // an implicit CSS default would make the arithmetic and the paint two
        // different numbers.
        return width === undefined
            ? SheetEditorDialogComponent.DEFAULT_COL_PX
            : columnWidthToPx(width);
    }

    /**
     * Store a width for the active column. Blank clears it.
     *
     * The model refuses a non-positive width — Excel reads 0 as a HIDDEN column
     * — so a stray `0` leaves the document untouched rather than making the
     * column disappear from the generated workbook.
     */
    protected applyWidth(raw: string): void {
        const column = this.activeColumn();
        if ('' === column) return;

        const trimmed = raw.trim();
        this.writeColumnWidth(column, '' === trimmed ? undefined : Number(trimmed));
    }

    /** Shared by the toolbar's number box and the headers' drag handle. */
    private writeColumnWidth(column: string, width: number | undefined): void {
        const doc = this.doc();
        const sheet = doc?.sheets[this.activeSheet()];
        if (!doc || !sheet) return;

        const next = withColumnWidth(sheet, column, width);
        if (next === sheet) return;

        doc.sheets[this.activeSheet()] = next;
        this.commit({ ...doc });
    }

    protected close(): void {
        this.dialogRef.close();
    }

    protected save(): void {
        const doc = this.doc();
        const url = this.contentUrl();
        if (!doc || !url || this.saving()) return;

        if (this.unreadable()) {
            this.toast.error('Refusing to save over a file this editor could not read.');

            return;
        }

        this.saving.set(true);
        this.http.put<{ contentHash: string }>(url, { content: serialiseSheetDocument(doc) })
            .pipe(takeUntilDestroyed(this.destroyRef))
            .subscribe({
                next: () => {
                    this.dirty.set(false);
                    this.saving.set(false);
                    this.toast.success('Saved', this.node.name);
                },
                error: () => {
                    this.saving.set(false);
                    this.toast.error('Save failed');
                },
            });
    }

    private load(): void {
        const url = this.contentUrl();
        if (!url) {
            this.apply('');

            return;
        }

        this.http.get<{ content: string }>(url)
            .pipe(takeUntilDestroyed(this.destroyRef))
            .subscribe({
                next: ({ content }) => this.apply(content),
                error: () => {
                    this.loading.set(false);
                    this.unreadable.set(true);
                    this.toast.error('Could not read this template.');
                },
            });
    }

    private apply(content: string): void {
        const { doc, ok } = parseSheetDocument(content);
        this.doc.set(doc);
        this.rebase(doc);
        this.activeSheet.set(Object.keys(doc.sheets)[0] ?? '');
        this.unreadable.set(!ok);
        this.loading.set(false);

        this.measureViewport();

        if (!ok) {
            this.toast.error('This file is not a readable .dsheet.');
        }
    }

    /**
     * MEASURED once the grid exists, not assumed.
     *
     * Both windows are computed from these, and a value SMALLER than the real
     * viewport leaves blank rows at the bottom and missing columns at the
     * right until the first scroll corrects it -- and the initial paint, or
     * the moment right after going full screen, is exactly when an author is
     * looking.
     *
     * On the next frame because the element must have been laid out at its new
     * size first; reading `clientHeight` in the same tick as the class change
     * returns the size it had a moment ago.
     */
    private measureViewport(): void {
        requestAnimationFrame(() => {
            const body = this.host.nativeElement.querySelector('.sheet-editor__body');
            if (body instanceof HTMLElement) {
                this.viewportHeightPx.set(body.clientHeight);
                this.viewportWidthPx.set(body.clientWidth);
            }
        });
    }

    private contentUrl(): string | undefined {
        return this.store.selectSnapshot(AppConfigState.manifest)
            ?.vfs?.fileContentUrl?.replace('{path}', encodeURIComponent(this.node.path));
    }
}
