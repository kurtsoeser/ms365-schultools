import { describe, it, expect } from 'vitest';
import {
    viewsForRole,
    navSectionsForRole,
    isAktStaffRole,
    scopeFromState,
    filterItems,
    statusLabel,
    resolveUserRelations,
    formatRelationSummary,
    hasAktKvCapability,
    canKvStellungnahme,
    itemTouchesKvKlassen
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-state.js';

describe('Schulaktivitäten Lehrer-UX Scope', () => {
    it('Lehrer-Nav ohne KV: Übersicht, Liste, Kalender, Neue Aktivität', () => {
        const views = viewsForRole('lehrer');
        expect(views.map((v) => v.id)).toEqual(['dashboard', 'liste', 'kalender', 'antrag']);
        expect(views.find((v) => v.id === 'liste')?.label).toBe('Meine Aktivitäten');
        expect(views.map((v) => v.id)).not.toContain('meine-klasse');
        expect(views.map((v) => v.id)).not.toContain('freigabe');
        expect(views.map((v) => v.id)).not.toContain('auslastung');
    });

    it('Lehrer mit KV-Zuordnung bekommt Meine Klasse – keinen Umschalter-Ersatz Freigabe', () => {
        const ids = viewsForRole('lehrer', ['lehrer', 'kv'], { kvKlassen: ['4A'] }).map((v) => v.id);
        expect(ids).toEqual(['dashboard', 'liste', 'kalender', 'meine-klasse', 'antrag']);
        expect(ids).not.toContain('freigabe');
    });

    it('KV-Rolle nutzt dieselbe Lehrerwelt + Meine Klasse', () => {
        const ids = viewsForRole('kv').map((v) => v.id);
        expect(ids).toContain('meine-klasse');
        expect(ids).toContain('liste');
        expect(ids).not.toContain('regeln');
        expect(ids).not.toContain('freigabe');
    });

    it('Admin: Arbeit/Auswertung/Regeln/System – ohne Meine Klasse / ohne Klassenübersicht', () => {
        const sections = navSectionsForRole('admin');
        expect(sections.map((s) => s.id)).toEqual(['arbeit', 'auswertung', 'regeln', 'system']);
        const ids = viewsForRole('admin').map((v) => v.id);
        expect(ids).toEqual([
            'dashboard',
            'freigabe',
            'liste',
            'kalender',
            'berichte',
            'regeln',
            'limits',
            'system'
        ]);
        expect(ids).not.toContain('meine-klasse');
        expect(ids).not.toContain('antrag');
        expect(ids).not.toContain('auslastung');
        expect(viewsForRole('admin').find((v) => v.id === 'liste')?.label).toBe('Alle Aktivitäten');
        expect(viewsForRole('admin').find((v) => v.id === 'freigabe')?.label).toBe('Entscheidungen');
        expect(viewsForRole('admin').find((v) => v.id === 'limits')?.label).toBe('Kostenschwellen');
    });

    it('Filter: Mehrfachauswahl Klasse/Art/Status/Lehrkraft', () => {
        const items = [
            {
                titel: '1A',
                klasseCodes: '1A',
                typ: 'Exkursion',
                status: 'genehmigt',
                lehrerCode: 'COL'
            },
            {
                titel: '2B',
                klasseCodes: '2B',
                typ: 'Wandertag',
                status: 'genehmigt',
                lehrerCode: 'MAY'
            },
            {
                titel: 'Beide',
                klasseCodes: '1A; 3C',
                typ: 'Exkursion',
                status: 'beantragt',
                lehrerCode: 'COL'
            }
        ];
        expect(filterItems(items, { klasse: '1A' }).map((i) => i.titel).sort()).toEqual([
            '1A',
            'Beide'
        ]);
        expect(filterItems(items, { klasse: '1A;2B' }).map((i) => i.titel).sort()).toEqual([
            '1A',
            '2B',
            'Beide'
        ]);
        expect(filterItems(items, { typ: 'Exkursion;Wandertag' }).map((i) => i.titel).sort()).toEqual([
            '1A',
            '2B',
            'Beide'
        ]);
        expect(filterItems(items, { status: 'genehmigt;beantragt' }).map((i) => i.titel).sort()).toEqual([
            '1A',
            '2B',
            'Beide'
        ]);
        expect(filterItems(items, { lehrer: 'COL;MAY' }).map((i) => i.titel).sort()).toEqual([
            '1A',
            '2B',
            'Beide'
        ]);
        expect(filterItems(items, { lehrer: 'COL', status: 'beantragt' }).map((i) => i.titel)).toEqual([
            'Beide'
        ]);
        expect(filterItems(items, { klasse: '9Z' })).toEqual([]);
    });

    it('isAktStaffRole nur admin', () => {
        expect(isAktStaffRole('lehrer')).toBe(false);
        expect(isAktStaffRole('kv')).toBe(false);
        expect(isAktStaffRole('admin')).toBe(true);
    });

    it('Meine Aktivitäten: Antragsteller, Verantwortliche, Begleitung – dedupliziert', () => {
        const state = {
            role: 'lehrer',
            accountEmail: 'dave@schule.at',
            stammdaten: { teachers: [{ code: 'COL', email: 'dave@schule.at', name: 'Dave Cole' }] }
        };
        const scope = scopeFromState(state, { onlyMine: true });
        const items = [
            {
                aktivitaetId: '1',
                titel: 'Organisiert',
                beantragtVon: 'dave@schule.at',
                lehrerEmail: 'dave@schule.at',
                lehrerCode: 'COL',
                status: 'beantragt'
            },
            {
                aktivitaetId: '2',
                titel: 'Nur Begleitung',
                beantragtVon: 'other@schule.at',
                lehrerEmail: 'other@schule.at',
                lehrerCode: 'MAY',
                begleitungCodes: 'COL;WEI',
                status: 'genehmigt'
            },
            {
                aktivitaetId: '3',
                titel: 'Fremd Klasse',
                beantragtVon: 'other@schule.at',
                lehrerEmail: 'other@schule.at',
                lehrerCode: 'MAY',
                klasseCodes: '4A',
                status: 'beantragt'
            }
        ];
        const filtered = filterItems(items, {}, scope);
        expect(filtered.map((i) => i.titel).sort()).toEqual(['Nur Begleitung', 'Organisiert']);
        const rels = resolveUserRelations(items[1], { email: 'dave@schule.at', teacherCode: 'COL' });
        expect(rels.map((r) => r.type)).toEqual(['companion']);
        expect(formatRelationSummary(rels)).toBe('Du begleitest');
        const multi = resolveUserRelations(items[0], { email: 'dave@schule.at', teacherCode: 'COL' });
        expect(multi.map((r) => r.type).sort()).toEqual(['applicant', 'responsible']);
    });

    it('KV ohne persönliche Beteiligung: nicht in Meine Aktivitäten, aber itemTouchesKvKlassen', () => {
        const item = {
            titel: 'Exkursion 4A',
            beantragtVon: 'other@schule.at',
            lehrerEmail: 'other@schule.at',
            klasseCodes: '4A',
            status: 'beantragt'
        };
        expect(itemTouchesKvKlassen(item, ['4A'])).toBe(true);
        const rels = resolveUserRelations(item, { email: 'kv@schule.at', teacherCode: 'COL' });
        expect(rels).toEqual([]);
    });

    it('KV-Capability und Stellungnahme ohne Rollenwechsel', () => {
        expect(
            hasAktKvCapability({
                role: 'lehrer',
                accountEmail: 'kv@schule.at',
                stammdaten: { classes: [{ code: '4A', headEmail: 'kv@schule.at' }] }
            })
        ).toBe(true);
        const item = { status: 'kv_stellungnahme', klasseCodes: '4A' };
        expect(
            canKvStellungnahme(item, {
                role: 'lehrer',
                planerRoles: ['lehrer', 'kv'],
                accountEmail: 'kv@schule.at',
                stammdaten: { classes: [{ code: '4A', headEmail: 'kv@schule.at' }] }
            })
        ).toBe(true);
    });

    it('Beteiligungsfilter companion', () => {
        const state = {
            role: 'lehrer',
            accountEmail: 'dave@schule.at',
            stammdaten: { teachers: [{ code: 'COL', email: 'dave@schule.at' }] }
        };
        const scope = scopeFromState(state, { onlyMine: true });
        const items = [
            {
                titel: 'Begleitung',
                beantragtVon: 'x@y.at',
                lehrerCode: 'MAY',
                begleitungCodes: 'COL',
                status: 'genehmigt'
            },
            {
                titel: 'Eigene',
                beantragtVon: 'dave@schule.at',
                lehrerCode: 'COL',
                status: 'beantragt'
            }
        ];
        const filtered = filterItems(items, { beteiligung: 'companion' }, scope);
        expect(filtered.map((i) => i.titel)).toEqual(['Begleitung']);
    });

    it('Statuslabels sind fachlich einheitlich', () => {
        expect(statusLabel('entwurf')).toBe('Entwurf');
        expect(statusLabel('beantragt')).toBe('Eingereicht');
        expect(statusLabel('kv_stellungnahme')).toBe('In Prüfung');
        expect(statusLabel('nachforderung')).toBe('Ergänzung erforderlich');
        expect(statusLabel('genehmigt')).toBe('Genehmigt');
        expect(statusLabel('abgelehnt')).toBe('Abgelehnt');
        expect(statusLabel('abgesagt')).toBe('Abgesagt');
    });

    it('Demo-Persona: IT-Account filtert als Dave (Meine Aktivitäten + KV-Klassen)', () => {
        const state = {
            role: 'lehrer',
            accountEmail: 'kurt@kurtrocks.onmicrosoft.com',
            demoPersonaEmail: 'dave.grohl@kurtrocks.onmicrosoft.com',
            stammdaten: {
                teachers: [
                    { code: 'GRO', email: 'dave.grohl@kurtrocks.onmicrosoft.com', name: 'Dave Grohl' }
                ],
                classes: [
                    { code: '2AK', headEmail: 'dave.grohl@kurtrocks.onmicrosoft.com' },
                    { code: '4AK', headEmail: 'dave.grohl@kurtrocks.onmicrosoft.com' }
                ]
            }
        };
        expect(hasAktKvCapability(state)).toBe(true);
        const scope = scopeFromState(state, { onlyMine: true });
        expect(scope.accountEmail).toBe('dave.grohl@kurtrocks.onmicrosoft.com');
        expect(scope.teacherCode).toBe('GRO');
        const items = [
            {
                titel: 'Von Dave',
                beantragtVon: 'dave.grohl@kurtrocks.onmicrosoft.com',
                lehrerCode: 'GRO',
                lehrerEmail: 'dave.grohl@kurtrocks.onmicrosoft.com',
                status: 'genehmigt'
            },
            {
                titel: 'Klasse ohne Dave',
                beantragtVon: 'other@kurtrocks.onmicrosoft.com',
                lehrerCode: 'WEI',
                klasseCodes: '2AK',
                status: 'kv_stellungnahme'
            },
            {
                titel: 'Fremd',
                beantragtVon: 'other@kurtrocks.onmicrosoft.com',
                lehrerCode: 'MAY',
                status: 'beantragt'
            }
        ];
        expect(filterItems(items, {}, scope).map((i) => i.titel)).toEqual(['Von Dave']);
        expect(
            canKvStellungnahme(items[1], state)
        ).toBe(true);
    });
});

describe('Validierung Microcopy', async () => {
    const { validateAktivitaet, vorlaufStatusForDraft, mapValidationIssueToStep } = await import(
        '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-logic.js'
    );

    it('Titel-Fehler menschlich formuliert', () => {
        const r = validateAktivitaet({
            draft: { titel: '', typ: 'Exkursion', klasseCode: '3BK', startdatum: '2099-12-01' },
            rules: { minVorlaufTage: 0 }
        });
        expect(r.errors.some((e) => /Titel/i.test(e))).toBe(true);
        expect(r.errors.join(' ')).not.toMatch(/Antrag heute/);
    });

    it('Vorlauf-Status ohne Debug-Text', () => {
        const st = vorlaufStatusForDraft(
            { startdatum: '2099-12-01' },
            { minVorlaufTage: 14, vorlaufModus: 'sperre' }
        );
        expect(st.message || '').not.toMatch(/DEBUG|console/i);
    });

    it('mapValidationIssueToStep liefert Schritt', () => {
        expect(mapValidationIssueToStep('Titel fehlt').step).toBe(1);
    });
});
