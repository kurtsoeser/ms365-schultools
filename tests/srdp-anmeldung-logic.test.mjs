import { describe, it, expect } from 'vitest';
import {
    buildItemTitle,
    resolveVariant,
    buildPruefplanKurz,
    buildPruefplanDetail,
    isSeminarWahlfach,
    validateAnmeldungDraft,
    filterAbschlussklassen,
    suggestLfsChoices,
    teacherChoiceLabels,
    parseChoiceLines,
    teacherRolesOnItem,
    filterItemsForTeacher,
    filterItemsForKlassen,
    preferredSrdpTerminJahr,
    srdpYearsToProbe,
    computeAnmeldungStats,
    groupItemsBy,
    buildAnmeldungSummary,
    isSrdpAnmeldeFensterOffen,
    findAnmeldungBySchuelerEmail,
    splitPersonName,
    resolveSrdpOwnership,
    resolveStudentAssignedKlasse,
    klasseCodeFromGroupName
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-logic.js';
import { renderAntragWizard } from '../src/tools/srdp-anmeldung/srdp-anmeldung-wizard-ui.js';
import { listTitleForYear, buildAnmeldungColumns, HAK_VARIANTS, PROFILE_IDS, getProfile } from '../src/tools/srdp-anmeldung/srdp-anmeldung-schema.js';

describe('srdp-anmeldung-logic', () => {
    it('buildItemTitle', () => {
        expect(buildItemTitle('Meier', 'Anna')).toBe('Meier, Anna');
        expect(buildItemTitle('Meier', '')).toBe('Meier');
        expect(buildItemTitle('', 'Anna')).toBe('Anna');
    });

    it('splitPersonName und Ownership aus Konto', () => {
        expect(splitPersonName('Lena Bauer')).toEqual({ vorname: 'Lena', nachname: 'Bauer' });
        expect(splitPersonName('Bauer, Lena')).toEqual({ vorname: 'Lena', nachname: 'Bauer' });
        const own = resolveSrdpOwnership(
            { klasse: '5A' },
            { accountEmail: 'lena@schule.at', accountName: 'Lena Bauer', students: [] }
        );
        expect(own.schuelerEmail).toBe('lena@schule.at');
        expect(own.vorname).toBe('Lena');
        expect(own.nachname).toBe('Bauer');
        expect(own.klasse).toBe('5A');
    });

    it('resolveVariant und Prüfplan', () => {
        expect(resolveVariant('Variante 2')?.key).toBe('2');
        expect(resolveVariant(1)?.label).toBe('Variante 1');
        expect(buildPruefplanKurz(3)).toContain('AM');
        expect(buildPruefplanKurz(3)).toContain('mündlich: BKO, Wahlfach');
        const detail = buildPruefplanDetail(
            {
                variante: 'Variante 1',
                titelDiplomarbeit: 'Rock or Die',
                betreuungslehrerDA: 'Bon Scott',
                lfs: 'ENWS',
                lehrerLFS: 'Brian May',
                lehrerBKO: 'Brian Johnson',
                wahlfachMuendlich: 'Seminar …',
                seminar: 'Wirtschaftsrecht',
                lehrerWahlfach: 'JohnPaul Jones'
            },
            'hak'
        );
        expect(detail.schriftlich.find((r) => r.part.startsWith('LFS')).detail).toBe('ENWS · Brian May');
        expect(detail.muendlich.find((r) => r.part === 'BKO').detail).toBe('Brian Johnson');
        expect(detail.muendlich.find((r) => r.part === 'Wahlfach').detail).toContain('Wirtschaftsrecht');
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

    it('isSrdpAnmeldeFensterOffen und findAnmeldungBySchuelerEmail', () => {
        expect(isSrdpAnmeldeFensterOffen({})).toBe(true);
        expect(isSrdpAnmeldeFensterOffen({ anmeldeFrist: '2099-12-31' }, '2026-11-18')).toBe(true);
        expect(isSrdpAnmeldeFensterOffen({ anmeldeFrist: '2020-01-01' }, '2026-11-18')).toBe(false);
        expect(
            isSrdpAnmeldeFensterOffen(
                { anmeldeVon: '2026-11-20', anmeldeFrist: '2026-12-01' },
                '2026-11-18'
            )
        ).toBe(false);
        const hit = findAnmeldungBySchuelerEmail(
            [
                { itemId: '1', schuelerEmail: 'a@schule.at' },
                { itemId: '2', beantragtVon: 'b@schule.at' }
            ],
            'b@schule.at'
        );
        expect(hit && hit.itemId).toBe('2');
        expect(findAnmeldungBySchuelerEmail([], 'x@y.at')).toBe(null);
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
        // Kein Jahr-Match → Fallback auf alle Codes (Dropdown darf nicht leer sein)
        expect(filterAbschlussklassen(classes, '2025')).toEqual(['5AK', '4AK', '5BK']);
        expect(filterAbschlussklassen([{ code: '5AHAK' }, { code: '5BHAK' }], '2026')).toEqual([
            '5AHAK',
            '5BHAK'
        ]);
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

    it('teacherRolesOnItem und Stats', () => {
        const item = {
            klasse: '5AHAK',
            variante: 'Variante 1',
            nachname: 'Huber',
            vorname: 'Anna',
            betreuungslehrerDA: 'Mag. Müller',
            lehrerLFS: 'Mag. Schmidt',
            lehrerBKO: 'Mag. Hofer',
            wahlfachMuendlich: 'Recht',
            lehrerWahlfach: 'Mag. Müller'
        };
        expect(teacherRolesOnItem(item, 'Mag. Müller').map((r) => r.role)).toEqual(['da', 'wahlfach']);
        expect(filterItemsForTeacher([item], 'Mag. Schmidt')).toHaveLength(1);
        expect(filterItemsForKlassen([item, { klasse: '5a' }, { klasse: '4B' }], ['5A'])).toHaveLength(1);
        expect(preferredSrdpTerminJahr(new Date('2026-11-20'), '2026')).toBe('2027');
        expect(preferredSrdpTerminJahr(new Date('2026-11-20'), '2028')).toBe('2028');
        expect(preferredSrdpTerminJahr(new Date('2026-03-01'), '')).toBe('2026');
        expect(srdpYearsToProbe(new Date('2026-11-20'), '2026')[0]).toBe('2027');
        const stats = computeAnmeldungStats([item, { ...item, klasse: '5BHAK', variante: 'Variante 2' }]);
        expect(stats.total).toBe(2);
        expect(stats.byKlasse).toHaveLength(2);
        expect(buildAnmeldungSummary(item).rows.some((r) => r.label === 'Name')).toBe(true);
        const groups = groupItemsBy(
            [item, { ...item, klasse: '5BHAK', variante: 'Variante 2', wahlfachMuendlich: 'Seminar' }],
            'klasse'
        );
        expect(groups.map((g) => g.key)).toEqual(['5AHAK', '5BHAK']);
        expect(groupItemsBy([item], 'wahlfach')[0].key).toBeTruthy();
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

describe('Schülerklasse für die Anmeldung', () => {
    const email = 'schueler02.demo@ms365.schule';

    it('nimmt die Schülerzeile mit Klasse, auch wenn eine leere Demo-Zeile zuerst steht', () => {
        const klasse = resolveStudentAssignedKlasse({
            email,
            students: [
                { email, name: 'Schüler02 DEMO', klasse: '' },
                { email, name: 'Schüler02 DEMO', klasse: '5A' }
            ]
        });
        expect(klasse).toBe('5A');
    });

    it('liest 5A aus der Gruppenbezeichnung DEMO Klasse 5A', () => {
        expect(klasseCodeFromGroupName('Jahrgangsgruppe DEMO Klasse 5A / Abschluss 2027')).toBe('5A');
        const klasse = resolveStudentAssignedKlasse({
            email,
            students: [{ email, name: 'Schüler02 DEMO', klasse: '' }],
            memberGroups: [{ id: 'g1', displayName: 'DEMO Klasse 5A', mail: '' }],
            preferredCodes: ['5A', '5B']
        });
        expect(klasse).toBe('5A');
    });

    it('Formular bietet 5A an, statt die Auswahl zu sperren', () => {
        const html = renderAntragWizard({
            form: { wizardStep: 1, vorname: 'Schüler02', nachname: 'DEMO' },
            options: { jahr: '2027', profileLabel: 'HAK', klassen: ['5A', '5B'], variants: [] },
            accountEmail: email,
            stammdaten: {
                students: [{ email, name: 'Schüler02 DEMO', klasse: '5A' }],
                classes: [{ code: '5A', year: '2027' }]
            },
            memberGroups: []
        });
        expect(html).toContain('value="5A"');
        expect(html).toContain('selected');
        expect(html).not.toContain('noch keiner Klasse zugeordnet');
        expect(html).not.toContain('disabled');
    });
});
