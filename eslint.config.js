import js from '@eslint/js';
import globals from 'globals';

const vitestGlobals = {
    describe: 'readonly',
    it: 'readonly',
    expect: 'readonly',
    beforeEach: 'readonly',
    afterEach: 'readonly',
    beforeAll: 'readonly',
    afterAll: 'readonly',
    vi: 'readonly'
};

export default [
    {
        ignores: ['dist/**', 'node_modules/**', 'coverage/**', 'vendor/**']
    },
    js.configs.recommended,
    {
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'module',
            globals: {
                ...globals.browser,
                ...globals.node,
                /** SheetJS aus vendor/ (xlsx.full.min.js per <script>), mehrere Tools */
                XLSX: 'readonly',
                /** JSZip aus vendor/ (jszip.min.js per <script>), z. B. Freistellung-Setup */
                JSZip: 'readonly'
            }
        },
        rules: {
            'no-unused-vars': [
                'warn',
                {
                    argsIgnorePattern: '^_',
                    varsIgnorePattern: '^_',
                    caughtErrors: 'none'
                }
            ],
            'no-empty': ['warn', { allowEmptyCatch: true }],
            'no-irregular-whitespace': 'off',
            'no-useless-escape': 'off',
            /**
             * Historisch gewachsene Module (>650 Zeilen) werden in ARCHITECTURE.md
             * und den Analyse-Docs nachgezogen – max-lines erzeugt sonst nur CI-Rauschen.
             * Neue Module bitte bewusst klein halten.
             */
            'max-lines': 'off'
        }
    },
    {
        files: ['tests/**/*.{js,mjs}', '**/*.test.{js,mjs}'],
        languageOptions: {
            globals: {
                ...globals.node,
                ...vitestGlobals
            }
        }
    }
];
