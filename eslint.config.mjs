// @ts-check
/**
 * ESLint config for `@coolms/document`.
 *
 * Lands with the move, not after it: these files were linted as
 * `src/app/coolms-*` until they left the admin's `src/` tree, and nothing
 * reports a directory leaving a tool's scope.
 *
 * **The relaxed `no-unsafe-*` tier is inherited deliberately.** The admin
 * exempted `src/app/coolms-*` because these sit close to vendored code, whose
 * typings hand back `any` at the seam. Those exemptions travel with the code
 * rather than being silently dropped or silently tightened.
 *
 * Rules otherwise come from the same `packages/eslint.config.base.mjs` factory
 * every other package uses.
 *
 * The base is VENDORED here as `eslint.config.base.mjs`, a byte-identical
 * copy of `packages/eslint.config.base.mjs`. That is what lets this package
 * lint inside its own repository, where the shared file does not exist.
 * `make check-fe` fails if a copy drifts; fix drift by editing the canonical
 * file and running `node tools/sync-eslint-base.mjs`, never by editing the
 * copy -- an edit in place is reverted by the next sync, silently.
 */

import createBaseConfig from './eslint.config.base.mjs';
import angular from 'angular-eslint';
import tseslint from 'typescript-eslint';
import globals from 'globals';

/** @type {import('typescript-eslint').ConfigArray} */
export default tseslint.config(
    ...createBaseConfig({ tseslint, globals }),
    { ignores: ['dist/**', 'src/**/*.spec.ts'] },
    {
        files: ['**/*.ts'],
        languageOptions: {
            parserOptions: {
                project: ['./tsconfig.lib.json'],
                tsconfigRootDir: import.meta.dirname,
            },
        },
        plugins: { '@angular-eslint': angular.tsPlugin },
        processor: angular.processInlineTemplates,
        rules: {
            ...angular.configs.tsRecommended.at(-1).rules,
            '@angular-eslint/component-selector': [
                'warn',
                { type: 'element', prefix: ['app', 'cms', 'coolms'], style: 'kebab-case' },
            ],
            '@angular-eslint/directive-selector': [
                'warn',
                { type: 'attribute', prefix: ['app', 'cms', 'coolms'], style: 'camelCase' },
            ],
            '@angular-eslint/prefer-on-push-component-change-detection': 'warn',
            '@typescript-eslint/unbound-method': ['error', { ignoreStatic: true }],

            '@typescript-eslint/no-explicit-any': 'off',
            '@typescript-eslint/no-unsafe-assignment': 'off',
            '@typescript-eslint/no-unsafe-member-access': 'off',
            '@typescript-eslint/no-unsafe-call': 'off',
            '@typescript-eslint/no-unsafe-argument': 'off',
            '@typescript-eslint/no-unsafe-return': 'off',

            '@typescript-eslint/no-base-to-string': 'warn',
            '@typescript-eslint/consistent-type-imports': 'warn',
            '@typescript-eslint/prefer-nullish-coalescing': 'warn',
            '@typescript-eslint/no-redundant-type-constituents': 'warn',
            '@typescript-eslint/no-unsafe-enum-comparison': 'warn',
            '@typescript-eslint/no-unnecessary-condition': 'warn',
            '@typescript-eslint/strict-boolean-expressions': 'warn',
            '@typescript-eslint/no-floating-promises': 'warn',
            '@typescript-eslint/no-misused-promises': 'warn',
            '@typescript-eslint/require-await': 'warn',
        },
    },
    { files: ['**/*.html'], ...tseslint.configs.disableTypeChecked },
    {
        files: ['**/*.html'],
        languageOptions: { parser: angular.templateParser },
        plugins: { '@angular-eslint/template': angular.templatePlugin },
        rules: {
            ...angular.configs.templateRecommended.at(-1).rules,
            '@angular-eslint/template/click-events-have-key-events': 'warn',
            '@angular-eslint/template/interactive-supports-focus': 'warn',
            '@angular-eslint/template/alt-text': 'warn',
        },
    },
);
