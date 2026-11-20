import { describe, it, expect } from 'vitest';
import {
    navSectionsForRole,
    viewsForRole,
    filterItemsForView,
    applyFilterJump,
    isSrdpStaffRole,
    looksLikeDemoKlassen,
    looksLikeDemoSeminare,
    resolveFormOptions,
    roleLabel,
    terminJahrFromListTitle
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-state.js';
import {
    resolveSrdpOwnership,
    assertSrdpOwnership,
    parseCodeList,
    joinCodeList,
    countNacharbeitPending,
    klassenOhneAnmeldung
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-logic.js';
import { isSrdpDemoRoleUiEnabled, canUseSrdpRoleSwitcher, roleFromFlags, shellPreferredSrdpRole } from '../src/tools/srdp-anmeldung/srdp-anmeldung-entra-role.js';
import {
    DEMO_SEED_TAG,
    DEMO_STUDENT_PERSONA,
    DEMO_TERMIN_JAHR,
    buildDemoItems,
    buildDemoStammdaten,
    buildDemoWizard,
    isDemoSeedItem,
    stripDemoSeedItems
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-demo.js';
import {
    filterAbschlussklassen,
    findAnmeldungBySchuelerEmail
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-logic.js';
import { matchStudentByEmail } from '../src/tools/srdp-anmeldung/srdp-anmeldung-state.js';
import { mapAnmeldungFromItem } from '../src/tools/srdp-anmeldung/srdp-anmeldung-graph.js';

describe('srdp nav / staff', () => {
    it('Lehrer-Nav ist flach, Admin sektioniert', () => {
        const lehrer = navSectionsForRole('lehrer', null, { kvKlassen: ['5AHAK'] });
        expect(lehrer).toHaveLength(1);
        expect(lehrer[0].label).toBe('');
        expect(viewsForRole('lehrer', { kvKlassen: ['5AHAK'] }).map((v) => v.id)).toContain(
            'meine-klasse'
        );
        expect(viewsForRole('lehrer', { kvKlassen: [] }).map((v) => v.id)).not.toContain(
            'meine-klasse'
        );

        const admin = navSectionsForRole('admin');
        expect(admin.map((s) => s.id)).toEqual(['arbeit', 'auswertung', 'system']);
        expect(isSrdpStaffRole('admin')).toBe(true);
        expect(isSrdpStaffRole('lehrer')).toBe(false);
    });

    it('Role-Chips nur mit demoOverride', () => {
        expect(isSrdpDemoRoleUiEnabled(false, false)).toBe(false);
        expect(isSrdpDemoRoleUiEnabled(true, false)).toBe(false);
        expect(isSrdpDemoRoleUiEnabled(false, true)).toBe(true);
        expect(canUseSrdpRoleSwitcher(['lehrer', 'admin'], false)).toBe(false);
        expect(canUseSrdpRoleSwitcher(['lehrer', 'admin'], true)).toBe(true);
    });

    it('Persona: IT und Direktion sehen Verwaltung, Lehrkraft bleibt Lehrkraft', () => {
        expect(roleFromFlags({ isIt: true })).toBe('admin');
        expect(roleFromFlags({ isDirektion: true, isLehrer: true })).toBe('admin');
        expect(roleFromFlags({ isLehrer: true })).toBe('lehrer');
        expect(roleFromFlags({})).toBe('schueler');
        expect(shellPreferredSrdpRole(['lehrer', 'admin'], 'lehrer')).toBe('admin');
        expect(shellPreferredSrdpRole(['lehrer', 'admin'], 'lehrer', { allowSaved: true })).toBe('lehrer');
        expect(roleLabel('admin', 'designated-it')).toBe('Schul-IT');
        expect(roleLabel('admin', 'stammdaten-direktion')).toBe('Direktion');
        expect(roleLabel('lehrer')).toBe('Lehrkraft');
        expect(terminJahrFromListTitle('sRDP-Anmeldungen HAK 2027')).toBe('2027');
    });
});

describe('srdp filter / ownership', () => {
    it('Mehrfachfilter und Jump leeren Rest', () => {
        const items = [
            { klasse: '5AHAK', variante: 'Variante 1', nachname: 'A', betreuungslehrerDA: 'Mag. Müller' },
            { klasse: '5BHAK', variante: 'Variante 2', nachname: 'B', betreuungslehrerDA: 'Mag. Hofer' }
        ];
        const state = {
            filterKlasse: '5AHAK',
            filterVariante: '',
            filterLehrer: '',
            filterQuery: ''
        };
        expect(filterItemsForView(items, state)).toHaveLength(1);

        applyFilterJump(state, { variante: 'Variante 2' });
        expect(state.filterKlasse).toBe('');
        expect(parseCodeList(state.filterVariante)).toEqual(['Variante 2']);
        expect(filterItemsForView(items, state)).toHaveLength(1);
        expect(joinCodeList(['a', 'b'])).toBe('a; b');
    });

    it('Ownership Assert braucht E-Mail', () => {
        const own = resolveSrdpOwnership(
            { nachname: 'Huber', vorname: 'Anna' },
            { accountEmail: 'anna.huber@demo.schule.at' }
        );
        expect(own.schuelerEmail).toBe('anna.huber@demo.schule.at');
        expect(() => assertSrdpOwnership(own)).not.toThrow();
        expect(() => assertSrdpOwnership({ schuelerEmail: '' })).toThrow(/E-Mail/i);
    });

    it('Bestätigung ≠ Nacharbeit', () => {
        const items = [
            { bestaetigung: true, nacharbeitErledigt: false },
            { bestaetigung: true, nacharbeitErledigt: true },
            { bestaetigung: false, nacharbeitErledigt: false }
        ];
        expect(countNacharbeitPending(items)).toBe(1);
        expect(klassenOhneAnmeldung([{ klasse: '5AHAK' }], ['5AHAK', '5BHAK'])).toEqual(['5BHAK']);
    });
});

describe('srdp form options / demo scrub', () => {
    it('erkennt Demo-Klassen 5AHAK und bevorzugt Live-Stammdaten', () => {
        expect(looksLikeDemoKlassen(['5AHAK', '5BHAK'])).toBe(true);
        expect(looksLikeDemoKlassen(['5A', '5B'])).toBe(false);
        const opts = resolveFormOptions(
            { profileId: 'hak', jahr: '2027', klassen: ['5AHAK', '5BHAK'] },
            {
                classes: [
                    { code: '5A', year: 2027 },
                    { code: '5B', year: 2027 },
                    { code: '4A', year: 2028 }
                ]
            },
            'hak'
        );
        expect(opts.klassen).toEqual(['5A', '5B']);
    });

    it('nimmt Listen-Choices wenn Wizard nur Demo hat', () => {
        const opts = resolveFormOptions(
            {
                profileId: 'hak',
                jahr: '2027',
                klassen: ['5AHAK'],
                listKlassen: ['5A', '5B']
            },
            { classes: [] },
            'hak'
        );
        expect(opts.klassen).toEqual(['5A', '5B']);
    });

    it('blendet Demo-Klassen im Live-Betrieb aus', () => {
        const opts = resolveFormOptions(
            { profileId: 'hak', jahr: '2027', klassen: ['5AHAK', '5BHAK', '5CHAK'] },
            { classes: [] },
            'hak'
        );
        expect(opts.klassen).toEqual([]);
        const demo = resolveFormOptions(
            { profileId: 'hak', jahr: '2027', klassen: ['5AHAK', '5BHAK', '5CHAK'] },
            { classes: [] },
            'hak',
            { allowDemoKlassen: true }
        );
        expect(demo.klassen).toEqual(['5AHAK', '5BHAK', '5CHAK']);
    });

    it('ersetzt Demo-Seminare durch die Listen-Auswahl', () => {
        expect(looksLikeDemoSeminare('DIGBIZ\nWirtschaftsrecht')).toBe(true);
        const opts = resolveFormOptions(
            {
                profileId: 'hak',
                jahr: '2027',
                seminarText: 'DIGBIZ\nWirtschaftsrecht',
                listSeminare: ['Unternehmensrechnung', 'Volkswirtschaft']
            },
            { classes: [] },
            'hak'
        );
        expect(opts.seminare).toEqual(['Unternehmensrechnung', 'Volkswirtschaft']);
    });

    it('nimmt die Optionen aus Schritt 3, nicht die alte EN/ENWS-Liste', () => {
        const opts = resolveFormOptions(
            {
                profileId: 'hak',
                jahr: '2027',
                lfs: 'ENWS\nFRWS\nITWS\nSPWS',
                lfsText: 'EN\nENWS\nFR\nFRWS\nIT\nITWS\nSP\nSPWS',
                wahlfach: 'Religion / Ethik\nKultur',
                seminar: 'Rockology',
                listLfs: ['EN', 'Englisch'],
                listSeminare: ['DIGBIZ']
            },
            { classes: [] },
            'hak'
        );
        expect(opts.lfs).toEqual(['ENWS', 'FRWS', 'ITWS', 'SPWS']);
        expect(opts.wahlfaecher).toEqual(['Religion / Ethik', 'Kultur']);
        expect(opts.seminare).toEqual(['Rockology']);
    });

    it('ignoriert die Demo-LFS-Liste und nimmt die Listen-Auswahl', () => {
        const opts = resolveFormOptions(
            {
                profileId: 'hak',
                jahr: '2027',
                lfsText: 'EN\nENWS\nFR\nFRWS\nIT\nITWS\nSP\nSPWS',
                listLfs: ['ENWS', 'FRWS']
            },
            { classes: [] },
            'hak'
        );
        expect(opts.lfs).toEqual(['ENWS', 'FRWS']);
    });
});

describe('srdp demo seed', () => {
    it('Seed-Tag und Reset nur Seed-Zeilen', () => {
        const seed = buildDemoItems();
        expect(seed.every((it) => it.seedTag === DEMO_SEED_TAG)).toBe(true);
        const mixed = seed.concat([{ itemId: 'live-1', nachname: 'Live' }]);
        expect(stripDemoSeedItems(mixed)).toEqual([{ itemId: 'live-1', nachname: 'Live' }]);
        expect(isDemoSeedItem(seed[0])).toBe(true);
        expect(isDemoSeedItem({ itemId: 'live-1' })).toBe(false);
    });

    it('Schüler-Persona ist in den Stammdaten, aber noch nicht angemeldet (Formular zeigbar)', () => {
        const students = buildDemoStammdaten().students;
        const st = matchStudentByEmail(students, DEMO_STUDENT_PERSONA.email);
        expect(st).toBeTruthy();
        expect(st.klasse).toBe(DEMO_STUDENT_PERSONA.klasse);
        expect(DEMO_STUDENT_PERSONA.name).toBe(st.vorname + ' ' + st.nachname);
        expect(findAnmeldungBySchuelerEmail(buildDemoItems(), DEMO_STUDENT_PERSONA.email)).toBeFalsy();
    });

    it('Demo-Jahr ist in Wizard, Abschlussklassen und Seed-Zeilen dasselbe', () => {
        const wizard = buildDemoWizard();
        expect(wizard.jahr).toBe(DEMO_TERMIN_JAHR);
        expect(wizard.terminJahr).toBe(DEMO_TERMIN_JAHR);
        expect(DEMO_SEED_TAG).toContain(DEMO_TERMIN_JAHR);
        const klassen = filterAbschlussklassen(buildDemoStammdaten().classes, wizard.jahr);
        expect(klassen).toEqual(wizard.klassen);
        expect(buildDemoItems().every((it) => it.terminJahr === DEMO_TERMIN_JAHR)).toBe(true);
    });
});

describe('srdp graph map (SPO Live Sync)', () => {
    it('mapAnmeldungFromItem setzt itemId, Ownership und Nacharbeit', () => {
        const mapped = mapAnmeldungFromItem({
            id: 77,
            createdDateTime: '2026-03-01T10:00:00Z',
            webUrl: 'https://contoso.sharepoint.com/lists/1/77',
            createdBy: { user: { email: 'anna.huber@schule.at' } },
            fields: {
                Title: 'Huber, Anna',
                Klasse: '5AHAK',
                Nachname: 'Huber',
                Vorname: 'Anna',
                TitelDiplomarbeit: 'Tourismus 2030',
                BetreuungslehrerDA: 'Mag. Müller',
                Variante: 'Variante 1',
                PruefplanKurz: 'schriftlich: D, BFK, LFS',
                LFS: 'ENWS',
                LehrerLFS: 'Mag. Hofer',
                LehrerBKO: 'Mag. Berger',
                WahlfachMuendlich: 'Recht',
                Seminar: '',
                LehrerWahlfach: 'Mag. Weber',
                Bestaetigung: true,
                NacharbeitErledigt: false,
                SchuelerEmail: 'anna.huber@schule.at',
                Schulform: 'HAK',
                TerminJahr: '2026'
            }
        });
        expect(mapped.itemId).toBe('77');
        expect(mapped.schuelerEmail).toBe('anna.huber@schule.at');
        expect(mapped.beantragtVon).toBe('anna.huber@schule.at');
        expect(mapped.nacharbeitErledigt).toBe(false);
        expect(mapped.bestaetigung).toBe(true);
        expect(mapped.klasse).toBe('5AHAK');
        expect(mapped.variante).toBe('Variante 1');
        expect(mapped.webUrl).toContain('/77');
    });

    it('mapAnmeldungFromItem fällt auf createdBy zurück wenn SchuelerEmail fehlt', () => {
        const mapped = mapAnmeldungFromItem({
            id: '9',
            createdBy: { user: { email: 'max.mustermann@schule.at' } },
            fields: {
                Title: 'Mustermann, Max',
                Nachname: 'Mustermann',
                Vorname: 'Max',
                Bestaetigung: 'Ja',
                NacharbeitErledigt: '1'
            }
        });
        expect(mapped.itemId).toBe('9');
        expect(mapped.schuelerEmail).toBe('max.mustermann@schule.at');
        expect(mapped.nacharbeitErledigt).toBe(true);
        expect(mapped.bestaetigung).toBe(true);
    });
});
