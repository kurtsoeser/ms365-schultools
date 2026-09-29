import { describe, it, expect } from 'vitest';
import { normalizeMailNickname } from '../src/shared/utils/mail-nickname.js';
import { buildMailNickFromLabel } from '../src/tools/schulstruktur-sync/schulstruktur-sync-naming.js';

describe('normalizeMailNickname', () => {
    it('mappt Umlaute auf ae/oe/ue/ss', () => {
        expect(normalizeMailNickname('Schüler-Gruppe')).toBe('schueler-gruppe');
        expect(normalizeMailNickname('Öko')).toBe('oeko');
        expect(normalizeMailNickname('Größe')).toBe('groesse');
    });

    it('kollabiert Bindestriche und begrenzt Länge', () => {
        expect(normalizeMailNickname('  Foo---Bar  ')).toBe('foo-bar');
        expect(normalizeMailNickname('a'.repeat(80)).length).toBe(64);
    });

    it('buildMailNickFromLabel nutzt dieselbe Normalisierung', () => {
        expect(buildMailNickFromLabel('Klasse Mädchen')).toBe('klasse-maedchen');
    });
});
