import { describe, it, expect } from 'vitest';
import {
    navSectionsForRole,
    viewsForRole,
    filterItemsForView,
    applyFilterJump,
    isSrdpStaffRole
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-state.js';
import {
    resolveSrdpOwnership,
    assertSrdpOwnership,
    parseCodeList,
    joinCodeList,
    countNacharbeitPending,
    klassenOhneAnmeldung
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-logic.js';
import { isSrdpDemoRoleUiEnabled } from '../src/tools/srdp-anmeldung/srdp-anmeldung-entra-role.js';
import {
    DEMO_SEED_TAG,
    buildDemoItems,
    isDemoSeedItem,
    stripDemoSeedItems
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-demo.js';
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

describe('srdp demo seed', () => {
    it('Seed-Tag und Reset nur Seed-Zeilen', () => {
        const seed = buildDemoItems();
        expect(seed.every((it) => it.seedTag === DEMO_SEED_TAG)).toBe(true);
        const mixed = seed.concat([{ itemId: 'live-1', nachname: 'Live' }]);
        expect(stripDemoSeedItems(mixed)).toEqual([{ itemId: 'live-1', nachname: 'Live' }]);
        expect(isDemoSeedItem(seed[0])).toBe(true);
        expect(isDemoSeedItem({ itemId: 'live-1' })).toBe(false);
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
