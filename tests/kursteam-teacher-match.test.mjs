import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('kursteam-teacher-match-logic', () => {
    function api() {
        return loadScript('src/tools/kursteams/kursteam-teacher-match-logic.js').ms365KursteamTeacherMatchLogic;
    }

    it('normalizeTeacherCode ersetzt Umlaute', () => {
        const { normalizeTeacherCode } = api();
        expect(normalizeTeacherCode('föde')).toBe('FOEDE');
        expect(normalizeTeacherCode('  Engel ')).toBe('ENGEL');
    });

    it('resolveTeacherMatch: exakter Code', () => {
        const { resolveTeacherMatch } = api();
        const hit = resolveTeacherMatch('ENGEL', [
            { code: 'engel', name: 'Anita Engelmann', email: 'anita.engelmann@hak-steyr.at' }
        ]);
        expect(hit.method).toBe('exact');
        expect(hit.email).toBe('anita.engelmann@hak-steyr.at');
    });

    it('resolveTeacherMatch: Umlaut-Normalisierung', () => {
        const { resolveTeacherMatch } = api();
        const hit = resolveTeacherMatch('FOEDE', [
            { code: 'FÖDE', name: 'Max Föde', email: 'max.foede@schule.at' }
        ]);
        expect(hit.method).toBe('exact');
        expect(hit.email).toBe('max.foede@schule.at');
    });

    it('resolveTeacherMatch: eindeutiger Nachname-Präfix', () => {
        const { resolveTeacherMatch } = api();
        const hit = resolveTeacherMatch('ENGEL', [
            { code: 'AE', name: 'Anita Engelmann', email: 'anita.engelmann@hak-steyr.at' },
            { code: 'XY', name: 'Max Muster', email: 'max@schule.at' }
        ]);
        expect(hit.method).toBe('namePrefix');
        expect(hit.email).toBe('anita.engelmann@hak-steyr.at');
    });

    it('resolveTeacherMatch: kein Treffer bei mehrdeutigen Namen', () => {
        const { resolveTeacherMatch } = api();
        const hit = resolveTeacherMatch('ENGEL', [
            { code: 'A1', name: 'Anita Engelmann', email: 'a@schule.at' },
            { code: 'A2', name: 'Bernd Engel', email: 'b@schule.at' }
        ]);
        expect(hit).toBeNull();
    });

    it('syncTeacherMappingFromTenant übernimmt fehlende und behält Unterrichts-Kürzel', () => {
        const { syncTeacherMappingFromTenant } = api();
        const r = syncTeacherMappingFromTenant(
            { KEEP: 'keep@schule.at' },
            ['ENGEL', 'FRECH', 'KEEP'],
            [
                { code: 'ENGEL', name: 'Anita Engelmann', email: 'anita.engelmann@hak-steyr.at' },
                { code: 'FRECH', name: 'Michaela Frech', email: 'michaela.frech@hak-steyr.at' }
            ]
        );
        expect(r.added).toBe(2);
        expect(r.mapping.ENGEL).toBe('anita.engelmann@hak-steyr.at');
        expect(r.mapping.FRECH).toBe('michaela.frech@hak-steyr.at');
        expect(r.mapping.KEEP).toBe('keep@schule.at');
    });
});
