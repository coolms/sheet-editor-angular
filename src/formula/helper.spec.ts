/**
 * The formula helper's state, as a function of a string and a caret.
 *
 * `|` marks the caret in every case below, which keeps the offsets readable —
 * an assertion written as `helperAt('=SUM(A1, ', 9)` is a puzzle, and a wrong
 * number in one looks exactly like a bug in the code.
 */
import { applyCompletion, argumentLabel, helperAt, pointInsertAt, referencesIn, signatureParts } from './helper';
import { lookupFunction } from './functions';

/** `'=SUM(|'` -> the text without the bar, and the caret where it stood. */
function at(marked: string): ReturnType<typeof helperAt> {
    const caret = marked.indexOf('|');
    if (caret < 0) throw new Error('the case must mark the caret with |');

    return helperAt(marked.replace('|', ''), caret);
}

describe('formula helper: what the caret is sitting inside', () => {
    it('offers nothing for a cell that is not a formula', () => {
        expect(at('SUM|').kind).toBe('none');
        expect(at('12|').kind).toBe('none');
    });

    it('offers nothing for a bare = with nothing typed yet', () => {
        // A spreadsheet waits for a letter before it starts guessing.
        expect(at('=|').kind).toBe('none');
    });

    it('completes a half-typed function name', () => {
        const state = at('=SU|');
        expect(state.kind).toBe('completions');
        if (state.kind === 'completions') {
            expect(state.prefix).toBe('SU');
            expect(state.matches.map((f) => f.name)).toContain('SUM');
            expect(state.matches.every((f) => f.name.startsWith('SU'))).toBe(true);
        }
    });

    it('completes case-insensitively, because nobody types in capitals', () => {
        const state = at('=su|');
        expect(state.kind).toBe('completions');
        if (state.kind === 'completions') {
            expect(state.matches.map((f) => f.name)).toContain('SUM');
        }
    });

    it('offers nothing when no function matches', () => {
        expect(at('=ZZZ|').kind).toBe('none');
    });

    it('describes a function once it has its bracket, rather than offering to replace it', () => {
        const state = at('=SUM(|');
        expect(state.kind).toBe('signature');
        if (state.kind === 'signature') {
            expect(state.fn.name).toBe('SUM');
            expect(state.argumentIndex).toBe(0);
        }
    });

    it('counts which argument the caret is in', () => {
        const first = at('=ROUND(A1|, 2)');
        const second = at('=ROUND(A1, 2|)');
        expect(first.kind === 'signature' && first.argumentIndex).toBe(0);
        expect(second.kind === 'signature' && second.argumentIndex).toBe(1);
    });

    it('describes the INNERMOST call when they are nested', () => {
        const state = at('=SUM(A1, MAX(B1, |))');
        expect(state.kind).toBe('signature');
        if (state.kind === 'signature') {
            expect(state.fn.name).toBe('MAX');
            expect(state.argumentIndex).toBe(1);
        }
    });

    it('returns to the outer call once the inner one is closed', () => {
        const state = at('=SUM(A1, MAX(B1, C1), |)');
        expect(state.kind).toBe('signature');
        if (state.kind === 'signature') {
            expect(state.fn.name).toBe('SUM');
            // Two commas at SUM's own depth; the ones inside MAX belong to MAX.
            expect(state.argumentIndex).toBe(2);
        }
    });

    it('says nothing once the whole call is closed', () => {
        expect(at('=SUM(A1)|').kind).toBe('none');
    });

    it('does not mistake ordinary brackets for a call', () => {
        expect(at('=(A1 + |').kind).toBe('none');
    });

    it('still describes a call whose brackets are unbalanced, which is the normal case mid-typing', () => {
        const state = at('=SUM(A1, MAX(B1|');
        expect(state.kind).toBe('signature');
        if (state.kind === 'signature') expect(state.fn.name).toBe('MAX');
    });

    it('handles a function that takes no arguments', () => {
        const state = at('=TRUE(|)');
        expect(state.kind).toBe('signature');
        if (state.kind === 'signature') expect(state.fn.name).toBe('TRUE');
    });
});

describe('accepting a completion', () => {
    it('brings the bracket with it and puts the caret inside', () => {
        const state = helperAt('=SU', 3);
        expect(state.kind).toBe('completions');
        if (state.kind !== 'completions') return;

        const result = applyCompletion('=SU', state, lookupFunction('SUM')!);
        expect(result.text).toBe('=SUM(');
        // Inside the bracket: the next keystroke is an argument.
        expect(result.caret).toBe('=SUM('.length);
    });

    it('closes a function that takes nothing, so it needs no second thought', () => {
        const state = helperAt('=TR', 3);
        if (state.kind !== 'completions') throw new Error('expected completions');

        const result = applyCompletion('=TR', state, lookupFunction('TRUE')!);
        expect(result.text).toBe('=TRUE()');
    });

    it('replaces only the name, keeping what follows', () => {
        const state = helperAt('=SU+1', 3);
        if (state.kind !== 'completions') throw new Error('expected completions');

        const result = applyCompletion('=SU+1', state, lookupFunction('SUM')!);
        expect(result.text).toBe('=SUM(+1');
    });
});

describe('signature parts, for highlighting the current argument', () => {
    it('splits a signature into its arguments', () => {
        expect(signatureParts(lookupFunction('ROUND')!)).toEqual(['number', '[digits]']);
    });

    it('has no parts for a function that takes nothing', () => {
        expect(signatureParts(lookupFunction('TRUE')!)).toEqual([]);
    });

    it('repeats the last argument of a variadic function', () => {
        // SUM(number1, [number2, …]) describes the fifth argument with its
        // second part rather than running off the end.
        const sum = lookupFunction('SUM')!;
        expect(argumentLabel(sum, 0)).toBe('number1');
        expect(argumentLabel(sum, 4)).toBe('[number2, …]');
    });
});

/** `'=SUM(|'` -> the point-mode span, and the caret where the bar stood. */
function pointAt(marked: string): ReturnType<typeof pointInsertAt> {
    const caret = marked.indexOf('|');
    if (caret < 0) throw new Error('the case must mark the caret with |');

    return pointInsertAt(marked.replace('|', ''), caret);
}

describe('formula references: what a formula points at', () => {
    it('finds nothing in a cell that is not a formula', () => {
        expect(referencesIn('B3')).toEqual([]);
        expect(referencesIn('')).toEqual([]);
    });

    /**
     * A `ref colon ref` triple is ONE reference. Returning the two ends
     * separately would outline the ends of a range and not its middle, which
     * is exactly the wrong picture of what SUM touches.
     */
    it('reads a range as one reference, not two', () => {
        expect(referencesIn('=SUM(B3:B5)').map(r => r.range)).toEqual(['B3:B5']);
    });

    it('reads several references in source order', () => {
        expect(referencesIn('=B2*C2+SUM(D1:D9)').map(r => r.range)).toEqual(['B2', 'C2', 'D1:D9']);
    });

    /**
     * The reason this reads TOKENS and not a regex: a quoted string is text and
     * `LOG10(` is a function name. Highlighting either would teach the author
     * to distrust the highlight.
     */
    it('ignores a reference-shaped thing that is not one', () => {
        expect(referencesIn('=IF(A1="B4","yes","no")').map(r => r.range)).toEqual(['A1']);
        expect(referencesIn('=LOG10(A1)').map(r => r.range)).toEqual(['A1']);
    });

    it('reports where each reference sits, so it can be replaced', () => {
        const [first] = referencesIn('=B2+1');

        expect(first.from).toBe(1);
        expect(first.to).toBe(3);
    });
});

describe('point mode: where a clicked cell would land', () => {
    it('does nothing outside a formula', () => {
        expect(pointAt('B3|')).toBeNull();
        expect(pointAt('|')).toBeNull();
    });

    it('inserts where a reference could legally go', () => {
        expect(pointAt('=|')).toEqual({ from: 1, to: 1 });
        expect(pointAt('=SUM(|')).toEqual({ from: 5, to: 5 });
        expect(pointAt('=1+|')).toEqual({ from: 3, to: 3 });
        expect(pointAt('=SUM(A1,|')).toEqual({ from: 8, to: 8 });
        expect(pointAt('=SUM(A1:|')).toEqual({ from: 8, to: 8 });
    });

    /**
     * Clicking again REPLACES the reference just placed. Without this rule the
     * second click produces `A1B2`, and the author reaches for backspace
     * between every two clicks.
     */
    it('replaces the reference the caret is parked on', () => {
        expect(pointAt('=A1|')).toEqual({ from: 1, to: 3 });
        expect(pointAt('=SUM(B12|')).toEqual({ from: 5, to: 8 });
    });

    /**
     * ⚠️ The caret INSIDE a token, which is not the same as the token before
     * it. Reported from the editor twice in one sitting: the reference landed
     * in the middle of what was already written, and the result was a formula
     * nobody could read back to what they had done.
     */
    it('replaces a reference the caret is anywhere inside, not just at its end', () => {
        // `=SUM(B|3` -- read as "just after the (", this inserted between the
        // B and the 3 and produced `=SUM(BD93`.
        expect(pointAt('=SUM(B|3')).toEqual({ from: 5, to: 7 });
        expect(pointAt('=SUM(|B3')).withContext('at its start').toEqual({ from: 5, to: 7 });
        // The WHOLE range, not the half the caret is in: see the case below.
        expect(pointAt('=SUM(A1:B|2)')).withContext('the far end of a range').toEqual({ from: 5, to: 10 });
    });

    /**
     * A RANGE is one reference. Read from the tokens it is three of them, and a
     * caret at its end replaced only the last: dragging the end of a range
     * wrote the new range inside the old one, `=SUM(B3:B3:B5`.
     */
    it('replaces a whole range, not the last reference in it', () => {
        expect(pointAt('=SUM(B3:B4|')).toEqual({ from: 5, to: 10 });
        expect(pointAt('=SUM(B3|:B4)')).withContext('from inside it').toEqual({ from: 5, to: 10 });
    });

    it('declines with the caret inside a function name', () => {
        // `=SU|M(B3:B5)` -- read as "just after the =", this inserted at 3 and
        // produced `=SUB4M(B3:B5)`.
        expect(pointAt('=SU|M(B3:B5)')).toBeNull();
        expect(pointAt('=|SUM(B3:B5)')).withContext('before the name').toBeNull();
        expect(pointAt('=SUM|(B3:B5)')).withContext('after the name').toBeNull();
    });

    it('declines with the caret inside a literal', () => {
        expect(pointAt('=1|23')).toBeNull();
        expect(pointAt('="ab|c"')).toBeNull();
    });

    /**
     * Everywhere else a click keeps its ordinary meaning. Getting this wrong is
     * worse than not having point mode at all: clicking away from a half-typed
     * formula would EDIT it instead of leaving it alone.
     */
    it('declines where a reference would be a syntax error', () => {
        expect(pointAt('=SUM(A1)|')).withContext('after a closing paren').toBeNull();
        expect(pointAt('=12|')).withContext('after a number').toBeNull();
        expect(pointAt('=SUM(A1)+1|')).toBeNull();
    });
});
