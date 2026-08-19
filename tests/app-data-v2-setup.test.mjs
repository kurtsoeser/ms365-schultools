import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function loadAppDataV2(store) {
    const full = join(projectRoot, 'src/shared/app-data-v2.js');
    const code = readFileSync(full, 'utf8');
    const sandbox = { console };
    sandbox.window = sandbox;
    sandbox.localStorage = {
        getItem(k) {
            return store.has(k) ? store.get(k) : null;
        },
        setItem(k, v) {
            store.set(k, String(v));
        },
        removeItem(k) {
            store.delete(k);
        }
    };
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox;
}

describe('app-data-v2 setup', () => {
    let store;

    beforeEach(() => {
        store = new Map();
    });

    it('VERSION is 4', () => {
        const ctx = loadAppDataV2(store);
        expect(ctx.ms365AppDataV2.VERSION).toBe(4);
    });

    it('normalizeSetup merges catalog links uniquely', () => {
        const ctx = loadAppDataV2(store);
        const n = ctx.ms365AppDataV2.normalizeSetup({
            catalogLinks: [
                { kind: 'subject', code: 'd', graphGroupId: 'g1' },
                { kind: 'subject', code: 'D', graphGroupId: 'g2' },
                { kind: 'arge', code: 'x', graphGroupId: '' }
            ]
        });
        expect(n.catalogLinks.length).toBe(2);
        expect(n.catalogLinks.find((x) => x.kind === 'subject').graphGroupId).toBe('g1');
    });

    it('normalizeSetup defaults group mail prefixes', () => {
        const ctx = loadAppDataV2(store);
        const n = ctx.ms365AppDataV2.normalizeSetup({});
        expect(n.subjectGroupMailPrefix).toBe('fach');
        expect(n.argeGroupMailPrefix).toBe('ag');
        const m = ctx.ms365AppDataV2.normalizeSetup({
            subjectGroupMailPrefix: 'fg',
            argeGroupMailPrefix: 'arbeits'
        });
        expect(m.subjectGroupMailPrefix).toBe('fg');
        expect(m.argeGroupMailPrefix).toBe('arbeits');
    });

    it('patchSetup merges directoryMatchByEmail by email key', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.patchSetup({
            directoryMatchByEmail: {
                'a@school.edu': {
                    graphUserId: 'id-a',
                    displayName: 'User A',
                    userPrincipalName: 'a@school.edu'
                }
            }
        });
        ctx.ms365AppDataV2.patchSetup({
            directoryMatchByEmail: {
                'b@school.edu': { notFound: true, checkedAt: '2026-01-01T00:00:00.000Z' }
            }
        });
        const s = ctx.ms365AppDataV2.getSetup();
        expect(s.directoryMatchByEmail['a@school.edu'].graphUserId).toBe('id-a');
        expect(s.directoryMatchByEmail['b@school.edu'].notFound).toBe(true);
    });

    it('mailNicknamePrefixSanitize keeps - _ . and strips MS-invalid chars', () => {
        const ctx = loadAppDataV2(store);
        const s = ctx.ms365AppDataV2.mailNicknamePrefixSanitize;
        expect(s('Fach-Sub_x.1', 24)).toBe('fach-sub_x.1');
        expect(s('a;b c', 24)).toBe('abc');
        const n = ctx.ms365AppDataV2.normalizeSetup({ subjectGroupMailPrefix: 'pre_fix-1' });
        expect(n.subjectGroupMailPrefix).toBe('pre_fix-1');
    });

    it('upsertClassTeam keeps Graph alias with underscores and identity without', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.upsertClassTeam({
            stableMailNickname: 'jg_2030_1a',
            mailNickname: 'jg_2030_1a',
            graphGroupId: 'gid-1a',
            classCode: '1A',
            displayName: 'Klasse 1A',
            abschlussJahr: '2030',
            mode: 'matched'
        });
        const first = ctx.ms365AppDataV2.getContainer().core.classTeams[0];
        expect(first.stableMailNickname).toBe('jg20301a');
        expect(first.mailNickname).toBe('jg_2030_1a');
        ctx.ms365AppDataV2.upsertClassTeam({
            stableMailNickname: 'jg_2030_1a',
            mailNickname: 'jg_2030_1a',
            graphGroupId: 'gid-1a',
            classCode: '1A',
            displayName: 'Klasse 1A',
            abschlussJahr: '2030',
            mode: 'matched'
        });
        const teams = ctx.ms365AppDataV2.getContainer().core.classTeams.filter((t) => t.classCode === '1A');
        expect(teams.length).toBe(1);
        expect(teams[0].mailNickname).toBe('jg_2030_1a');
    });

    it('normalizeSetup includes verwaltungGroupId and verwaltungDraft', () => {
        const ctx = loadAppDataV2(store);
        const n = ctx.ms365AppDataV2.normalizeSetup({});
        expect(n.matched.verwaltungGroupId).toBe(null);
        expect(n.verwaltungDraft.vwNewMailNick).toBe('verwaltung');
        expect(n.verwaltungDraft.vwNewDisplayName).toBe('Schulverwaltung');
    });

    it('normalizeSetup migrates layout7 wizardStep 3–7 to 4–8', () => {
        const ctx = loadAppDataV2(store);
        expect(
            ctx.ms365AppDataV2.normalizeSetup({
                wizardStep: 3,
                _einrichtungWizardLayout: 7
            }).wizardStep
        ).toBe(4);
        expect(
            ctx.ms365AppDataV2.normalizeSetup({
                wizardStep: 7,
                _einrichtungWizardLayout: 7
            }).wizardStep
        ).toBe(8);
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 2, _einrichtungWizardLayout: 7 }).wizardStep).toBe(2);
    });

    it('normalizeSetup migriert layout9/10 wizardStep 6–9 um +2 auf layout11', () => {
        const ctx = loadAppDataV2(store);
        // layout9 hatte 9 Schritte; Schritte 6–9 werden auf layout11 (11 Schritte) verschoben (+2)
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 9, _einrichtungWizardLayout: 9 }).wizardStep).toBe(11);
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 8, _einrichtungWizardLayout: 9 }).wizardStep).toBe(10);
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 6, _einrichtungWizardLayout: 9 }).wizardStep).toBe(8);
        // Schritte 1–5 bleiben unverändert
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 5, _einrichtungWizardLayout: 9 }).wizardStep).toBe(5);
        // layout10 verhält sich genauso
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 8, _einrichtungWizardLayout: 10 }).wizardStep).toBe(10);
        // layout11 (aktuell) – keine Migration
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 9, _einrichtungWizardLayout: 11 }).wizardStep).toBe(9);
        // Ergebnis trägt immer layout11
        const n = ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 3, _einrichtungWizardLayout: 9 });
        expect(n._einrichtungWizardLayout).toBe(11);
    });

    it('patchSetup preserves catalogLinks when omitted in partial', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.patchSetup({
            catalogLinks: [{ kind: 'subject', code: 'M', graphGroupId: 'id-m', mode: 'matched' }]
        });
        ctx.ms365AppDataV2.patchSetup({
            matched: { schuelerGroupId: 's1', lehrerGroupId: null }
        });
        const c = ctx.ms365AppDataV2.getContainer();
        expect(c.setup.matched.schuelerGroupId).toBe('s1');
        const subject = c.setup.catalogLinks.find((x) => x.kind === 'subject' && x.code === 'M');
        expect(subject && subject.graphGroupId).toBe('id-m');
        const sch = c.setup.catalogLinks.find((x) => x.kind === 'sammelgruppe' && x.code === 'schueler');
        expect(sch && sch.graphGroupId).toBe('s1');
    });

    it('getClassTeamGruppenmailForKlasse returns stable nick from registry', () => {
        store.set(
            'ms365-schooltool-data-v2',
            JSON.stringify({
                version: 3,
                core: {
                    domain: '',
                    subjects: [],
                    arges: [],
                    teachers: [],
                    admin: [],
                    classTeams: [
                        {
                            stableMailNickname: 'jg2031hma',
                            graphGroupId: 'gid-1',
                            classCode: 'HMA',
                            displayName: '1HMA',
                            abschlussJahr: '2031',
                            mode: 'created',
                            educationClassId: ''
                        }
                    ]
                },
                years: { current: '2025/26', byLabel: { '2025/26': { students: [], classes: [] } } },
                structure: { rows: [], memberships: {}, settings: {} },
                match: { links: {} },
                setup: {
                    wizardStep: 1,
                    completedSteps: [],
                    finishedAt: null,
                    lastVisitedAt: null,
                    matched: { schuelerGroupId: null, lehrerGroupId: null },
                    slgDraft: {
                        activeKind: 'schueler',
                        slgNewDisplayName: '',
                        slgNewMailNick: '',
                        slgNewDescription: '',
                        slgNewCreateTeam: false
                    },
                    catalogLinks: []
                },
                tenant: { cache: { rows: [], users: [], loadedAt: '' } }
            })
        );
        const ctx = loadAppDataV2(store);
        expect(ctx.ms365AppDataV2.getClassTeamGruppenmailForKlasse('1HMA')).toBe('jg2031hma');
        expect(ctx.ms365AppDataV2.getClassTeamGruppenmailForKlasse('HMA')).toBe('jg2031hma');
    });

    it('patchSetup merges verwaltungGroupId without wiping SLG matches', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.patchSetup({
            matched: { schuelerGroupId: 's1', lehrerGroupId: 'l1' }
        });
        ctx.ms365AppDataV2.patchSetup({
            matched: { verwaltungGroupId: 'v1' },
            verwaltungDraft: { vwNewDisplayName: 'Schulverwaltung', vwNewMailNick: 'verwaltung' }
        });
        const su = ctx.ms365AppDataV2.getSetup();
        expect(su.matched.schuelerGroupId).toBe('s1');
        expect(su.matched.lehrerGroupId).toBe('l1');
        expect(su.matched.verwaltungGroupId).toBe('v1');
        expect(su.verwaltungDraft.vwNewMailNick).toBe('verwaltung');
    });

    it('copies matched Sammelgruppen into catalogLinks', () => {
        const ctx = loadAppDataV2(store);
        const n = ctx.ms365AppDataV2.normalizeSetup({
            matched: { schuelerGroupId: 's-id', lehrerGroupId: 'l-id', verwaltungGroupId: 'v-id' }
        });
        const samm = n.catalogLinks.filter((x) => x.kind === 'sammelgruppe');
        expect(samm.map((x) => x.code).sort()).toEqual(['lehrer', 'schueler', 'verwaltung']);
        expect(samm.find((x) => x.code === 'schueler').graphGroupId).toBe('s-id');
        expect(samm.find((x) => x.code === 'verwaltung').graphGroupId).toBe('v-id');
    });

    it('unmatch via patchSetup clears catalog Sammelgruppe', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.patchSetup({
            matched: { schuelerGroupId: 's-id', lehrerGroupId: 'l-id' }
        });
        ctx.ms365AppDataV2.patchSetup({
            matched: { schuelerGroupId: null }
        });
        const su = ctx.ms365AppDataV2.getSetup();
        expect(su.matched.schuelerGroupId).toBe(null);
        expect(su.matched.lehrerGroupId).toBe('l-id');
        const sch = su.catalogLinks.find((x) => x.kind === 'sammelgruppe' && x.code === 'schueler');
        expect(sch && sch.graphGroupId).toBe('');
        const le = su.catalogLinks.find((x) => x.kind === 'sammelgruppe' && x.code === 'lehrer');
        expect(le && le.graphGroupId).toBe('l-id');
    });

    it('catalog sammelgruppe fills matched when matched is empty', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.patchSetup({
            catalogLinks: [{ kind: 'sammelgruppe', code: 'verwaltung', graphGroupId: 'verw-1', mode: 'matched' }]
        });
        const su = ctx.ms365AppDataV2.getSetup();
        expect(su.matched.verwaltungGroupId).toBe('verw-1');
        expect(ctx.ms365AppDataV2.getCatalogLink('sammelgruppe', 'verwaltung').graphGroupId).toBe('verw-1');
    });

    it('setCoreFromTenantSettings stores adminRoles', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.setCoreFromTenantSettings({
            domain: 'modeebensee.at',
            subjects: [],
            arges: [],
            teachers: [],
            admin: [{ role: 'Direktion', name: '', email: '' }],
            adminRoles: [{ code: 'DIREKTION', name: 'Direktion' }],
            students: [],
            classes: []
        });
        const c = ctx.ms365AppDataV2.getContainer();
        expect(c.core.adminRoles).toEqual([{ code: 'DIREKTION', name: 'Direktion' }]);
        expect(c.core.admin[0].role).toBe('Direktion');
    });

    it('normalizeSetup behält catalogLinks kind cohort und verwirft ungültige Jahre', () => {
        const ctx = loadAppDataV2(store);
        const n = ctx.ms365AppDataV2.normalizeSetup({
            catalogLinks: [
                {
                    kind: 'cohort',
                    code: '2026',
                    graphGroupId: 'g-c',
                    mode: 'matched',
                    mailNickname: 'maturajg2026',
                    displayName: 'Abschlussjahrgang 2026'
                },
                { kind: 'cohort', code: '26', graphGroupId: 'bad' },
                { kind: 'subject', code: 'M', graphGroupId: 'g-m' }
            ]
        });
        const coh = n.catalogLinks.filter((x) => x.kind === 'cohort');
        expect(coh.length).toBe(1);
        expect(coh[0].code).toBe('2026');
        expect(coh[0].graphGroupId).toBe('g-c');
        expect(coh[0].mailNickname).toBe('maturajg2026');
        expect(n.catalogLinks.find((x) => x.kind === 'subject').code).toBe('M');
    });

    it('getCatalogLink und upsertCatalogLink arbeiten mit kind cohort', () => {
        const ctx = loadAppDataV2(store);
        const saved = ctx.ms365AppDataV2.upsertCatalogLink({
            kind: 'cohort',
            code: '2027',
            graphGroupId: 'id-27',
            mode: 'created',
            mailNickname: 'maturajg2027'
        });
        expect(saved && saved.kind).toBe('cohort');
        const link = ctx.ms365AppDataV2.getCatalogLink('cohort', '2027');
        expect(link && link.graphGroupId).toBe('id-27');
        ctx.ms365AppDataV2.clearCatalogLinkGroup('cohort', '2027');
        expect(ctx.ms365AppDataV2.getCatalogLink('cohort', '2027').graphGroupId).toBe('');
    });

    it('normalizeSetup behält catalogLinks kind eltern getrennt von cohort', () => {
        const ctx = loadAppDataV2(store);
        const n = ctx.ms365AppDataV2.normalizeSetup({
            catalogLinks: [
                {
                    kind: 'eltern',
                    code: '2026',
                    graphGroupId: 'g-el',
                    mode: 'created',
                    mailNickname: 'elternjg2026'
                },
                {
                    kind: 'cohort',
                    code: '2026',
                    graphGroupId: 'g-co',
                    mode: 'matched'
                },
                { kind: 'eltern', code: '26', graphGroupId: 'bad' }
            ]
        });
        const el = n.catalogLinks.filter((x) => x.kind === 'eltern');
        const co = n.catalogLinks.filter((x) => x.kind === 'cohort');
        expect(el.length).toBe(1);
        expect(co.length).toBe(1);
        expect(el[0].graphGroupId).toBe('g-el');
        expect(co[0].graphGroupId).toBe('g-co');
    });

    it('getCatalogLink und upsertCatalogLink arbeiten mit kind eltern', () => {
        const ctx = loadAppDataV2(store);
        ctx.ms365AppDataV2.upsertCatalogLink({
            kind: 'cohort',
            code: '2026',
            graphGroupId: 'g-co',
            mode: 'matched'
        });
        const saved = ctx.ms365AppDataV2.upsertCatalogLink({
            kind: 'eltern',
            code: '2026',
            graphGroupId: 'g-el',
            mode: 'created',
            mailNickname: 'elternjg2026'
        });
        expect(saved && saved.kind).toBe('eltern');
        expect(ctx.ms365AppDataV2.getCatalogLink('eltern', '2026').graphGroupId).toBe('g-el');
        expect(ctx.ms365AppDataV2.getCatalogLink('cohort', '2026').graphGroupId).toBe('g-co');
        ctx.ms365AppDataV2.clearCatalogLinkGroup('eltern', '2026');
        expect(ctx.ms365AppDataV2.getCatalogLink('eltern', '2026').graphGroupId).toBe('');
        expect(ctx.ms365AppDataV2.getCatalogLink('cohort', '2026').graphGroupId).toBe('g-co');
    });
});
