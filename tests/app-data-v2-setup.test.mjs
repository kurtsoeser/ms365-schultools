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

    it('VERSION is 3', () => {
        const ctx = loadAppDataV2(store);
        expect(ctx.ms365AppDataV2.VERSION).toBe(3);
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

    it('normalizeSetup allows wizardStep 9 and layout 9', () => {
        const ctx = loadAppDataV2(store);
        const n = ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 9, _einrichtungWizardLayout: 9 });
        expect(n.wizardStep).toBe(9);
        expect(n._einrichtungWizardLayout).toBe(9);
        expect(ctx.ms365AppDataV2.normalizeSetup({ wizardStep: 8, _einrichtungWizardLayout: 8 }).wizardStep).toBe(8);
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
});
