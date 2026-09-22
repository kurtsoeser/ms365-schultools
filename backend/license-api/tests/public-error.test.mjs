import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { publicErrorMessage } = require('../src/lib/http-utils.js');

describe('publicErrorMessage', () => {
    it('versteckt 500er und Graph-/AAD-Details', () => {
        expect(publicErrorMessage({ status: 500, message: 'AADSTS700016: secret' })).toBe(
            'Interner Fehler.'
        );
        expect(
            publicErrorMessage({ status: 401, message: 'graph.microsoft.com 401' }, 'Anmeldung fehlt.')
        ).toBe('Anmeldung fehlt.');
        expect(publicErrorMessage({ status: 403, message: 'Nur Betreiber-Konten.' })).toBe(
            'Nur Betreiber-Konten.'
        );
    });
});
