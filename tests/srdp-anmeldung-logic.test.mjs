import { describe, it, expect } from 'vitest';
import {
    buildItemTitle,
    resolveVariant,
    buildPruefplanKurz,
    isSeminarWahlfach,
    validateAnmeldungDraft,
    filterAbschlussklassen,
    suggestLfsChoices,
    teacherChoiceLabels,
    parseChoiceLines
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-logic.js';
import { listTitleForYear, buildAnmeldungColumns, HAK_VARIANTS, PROFILE_IDS, getProfile } from '../src/tools/srdp-anmeldung/srdp-anmeldung-schema.js';

describe('srdp-anmeldung-logic', () => {
    it('buildItemTitle', () => {
        expect(buildItemTitle('Meier', 'Anna')).toBe('Meier, Anna');
        expect(buildItemTitle('Meier', '')).toBe('Meier');
        expect(buildItemTitle('', 'Anna')).toBe('Anna');
    });

    it('resolveVariant und Prüfplan', () => {
        expect(resolveVariant('Variante 2')?.key).toBe('2');
        expect(resolveVariant(1)?.label).toBe('Variante 1');
        expect(buildPruefplanKurz(3)).toContain('AM');
        expect(buildPruefplanKurz(3)).toContain('mündlich: BKO, Wahlfach');
        expect(HAK_VARIANTS).toHaveLength(3);
        expect(buildPruefplanKurz(1, 'htl')).toContain('SWP');
        expect(getProfile('hlw').variants).toHaveLength(2);
        expect(PROFILE_IDS).toEqual(expect.arrayContaining(['hak', 'htl', 'hlw', 'bafeb', 'ahs']));
    });

    it('isSeminarWahlfach', () => {
        expect(isSeminarWahlfach('Seminar')).toBe(true);
        expect(isSeminarWahlfach('Seminar DIGBIZ')).toBe(true);
        expect(isSeminarWahlfach('Recht')).toBe(false);
    });

    it('validateAnmeldungDraft', () => {
        const ok = validateAnmeldungDraft({
            variante: 'Variante 1',
            nachname: 'Huber',
            vorname: 'Max',
            lfs: 'ENWS',
            lehrerLfs: 'Müller',
            wahlfach: 'Recht',
            bestaetigung: true
        });
        expect(ok.ok).toBe(true);
        expect(ok.title).toBe('Huber, Max');
        expect(ok.pruefplanKurz).toMatch(/LFS/);

        const bad = validateAnmeldungDraft({
            variante: 'Variante 2',
            nachname: 'X',
            vorname: 'Y',
            lfs: 'ENWS',
            lehrerLfs: 'Z',
            wahlfach: 'Seminar',
            seminar: '',
            bestaetigung: true
        });
        expect(bad.ok).toBe(false);
        expect(bad.errors.some((e) => /Seminar/i.test(e))).toBe(true);
    });

    it('filterAbschlussklassen nach year', () => {
        const classes = [
            { code: '5AK', year: '2026' },
            { code: '4AK', year: '2027' },
            { code: '5BK', year: '2026' }
        ];
        expect(filterAbschlussklassen(classes, 2026)).toEqual(['5AK', '5BK']);
        expect(filterAbschlussklassen(classes, '2025')).toEqual([]);
    });

    it('suggestLfsChoices und teacherChoiceLabels', () => {
        const lfs = suggestLfsChoices([{ code: 'ENWS' }, { code: 'D' }], ['FRWS']);
        expect(lfs).toContain('ENWS');
        expect(lfs).toContain('FRWS');
        expect(teacherChoiceLabels([{ name: 'Berta A' }, { name: 'Anton B' }])).toEqual([
            'Anton B',
            'Berta A'
        ]);
    });

    it('parseChoiceLines', () => {
        expect(parseChoiceLines('Recht\n# Kommentar\nSeminar\nRecht')).toEqual(['Recht', 'Seminar']);
    });
});

describe('srdp-anmeldung-schema', () => {
    it('listTitleForYear inkl. Schulform', () => {
        expect(listTitleForYear(2026, 'hak')).toBe('sRDP-Anmeldungen HAK 2026');
        expect(listTitleForYear(2026, 'htl')).toBe('sRDP-Anmeldungen HTL 2026');
        expect(listTitleForYear(2026, 'ahs')).toBe('sRP-Anmeldungen AHS 2026');
        expect(() => listTitleForYear('26')).toThrow(/vierstellig/);
    });

    it('buildAnmeldungColumns mit Choices', () => {
        const cols = buildAnmeldungColumns({
            profileId: 'hak',
            klassen: ['5AK'],
            lehrer: ['Müller'],
            lfs: ['ENWS'],
            wahlfaecher: ['Recht'],
            seminare: ['DIGBIZ'],
            terminJahr: 2026
        });
        const klasse = cols.find((c) => c.name === 'Klasse');
        expect(klasse?.choice?.choices).toEqual(['5AK']);
        const seminar = cols.find((c) => c.name === 'Seminar');
        expect(seminar?.choice?.allowTextEntry).toBe(true);
        expect(cols.map((c) => c.name)).toContain('LehrerLFS');
    });

    it('HTL/HLW/BAFEB/AHS Spalten', () => {
        expect(buildAnmeldungColumns({ profileId: 'htl' }).map((c) => c.name)).toContain('SchwerpunktFach');
        expect(buildAnmeldungColumns({ profileId: 'hlw' }).map((c) => c.name)).toContain('Fachkolloquium');
        expect(buildAnmeldungColumns({ profileId: 'bafeb' }).map((c) => c.name)).toContain('Fachtheorie');
        expect(buildAnmeldungColumns({ profileId: 'ahs' }).map((c) => c.name)).toContain('HatABA');
        expect(buildAnmeldungColumns({ profileId: 'ahs' }).map((c) => c.name)).not.toContain('TitelDiplomarbeit');
    });
});
