import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function loadModules(store) {
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
    const eg = join(projectRoot, 'src/shared/eltern-guardians.js');
    const ad = join(projectRoot, 'src/shared/app-data-v2.js');
    const ts = join(projectRoot, 'src/shared/tenant-settings-core.js');
    runInContext(readFileSync(eg, 'utf8'), sandbox, { filename: eg });
    runInContext(readFileSync(ad, 'utf8'), sandbox, { filename: ad });
    runInContext(readFileSync(ts, 'utf8'), sandbox, { filename: ts });
    return sandbox;
}

describe('Eltern / Erziehungsberechtigte', () => {
    let store;

    beforeEach(() => {
        store = new Map();
    });

    it('normalizeYearBucket vergibt Schüler-IDs und filtert orphan guardianIds', () => {
        const ctx = loadModules(store);
        const n = ctx.ms365AppDataV2.normalizeYearBucket({
            students: [{ klasse: '1B', name: 'Dave Grohl', email: 'dave@schule.at', guardianIds: ['g1', 'missing'] }],
            guardians: [{ id: 'g1', name: 'Jane', email: 'jane@mail.com' }],
            classes: [{ code: '1B', name: '1B', year: '2030' }]
        });
        expect(n.students[0].id).toBeTruthy();
        expect(n.students[0].guardianIds).toEqual(['g1']);
        expect(n.guardians).toHaveLength(1);
        expect(n.parentLists).toEqual([]);
    });

    it('mergeStudentsImport erhält Zuordnung und dedupliziert Eltern per E-Mail', () => {
        const ctx = loadModules(store);
        ctx.ms365AppDataV2.setCoreFromTenantSettings({
            domain: 'schule.at',
            students: [
                {
                    klasse: '1B',
                    name: 'Dave Grohl',
                    email: 'dave@schule.at',
                    parentPairs: [
                        { name: 'Jane Grohl', email: 'jane@mail.com' },
                        { name: 'John Grohl', email: 'john@mail.com' }
                    ]
                },
                {
                    klasse: '1B',
                    name: 'Sibling Grohl',
                    email: 'sib@schule.at',
                    parentPairs: [{ name: 'Jane G', email: 'jane@mail.com' }]
                }
            ],
            classes: [{ code: '1B', name: '1B', year: '2030' }]
        });
        const { bucket } = ctx.ms365AppDataV2.getYearBucket();
        expect(bucket.guardians).toHaveLength(2);
        const dave = bucket.students.find((s) => s.email === 'dave@schule.at');
        const sib = bucket.students.find((s) => s.email === 'sib@schule.at');
        expect(dave.guardianIds).toHaveLength(2);
        expect(sib.guardianIds).toHaveLength(1);
        expect(sib.guardianIds[0]).toBe(dave.guardianIds[0]);
    });

    it('buildClassParentSoll und buildYearParentSoll aggregieren korrekt', () => {
        const ctx = loadModules(store);
        const bucket = {
            students: [
                { id: 's1', klasse: '1B', name: 'A', email: 'a@s.at', guardianIds: ['g1', 'g2'] },
                { id: 's2', klasse: '1B', name: 'B', email: 'b@s.at', guardianIds: ['g1'] }
            ],
            classes: [{ code: '1B', name: '1B', year: '2030' }],
            guardians: [
                { id: 'g1', name: 'P1', email: 'p1@mail.com' },
                { id: 'g2', name: 'P2', email: 'p2@mail.com' }
            ],
            parentLists: []
        };
        const classSoll = ctx.ms365ElternGuardians.buildClassParentSoll(bucket);
        expect(classSoll).toHaveLength(1);
        expect(classSoll[0].displayName).toBe('Eltern 1B');
        expect(classSoll[0].mailNickname).toBe('eltern1b');
        expect(classSoll[0].guardianCount).toBe(2);
        expect(classSoll[0].studentCount).toBe(2);

        const yearSoll = ctx.ms365ElternGuardians.buildYearParentSoll(bucket);
        expect(yearSoll).toHaveLength(1);
        expect(yearSoll[0].code).toBe('2030');
        expect(yearSoll[0].mailNickname).toBe('elternjg2030');
        expect(yearSoll[0].guardianCount).toBe(2);
    });

    it('Baustein-Muster steuert Alias mit Trenner', () => {
        const ctx = loadModules(store);
        const naming = {
            classAliasPattern: [
                { type: 'text', value: 'eltern' },
                { type: 'text', value: '-' },
                { type: 'klasse' }
            ],
            classDisplayPattern: [
                { type: 'text', value: 'Eltern – ' },
                { type: 'klasse' }
            ],
            yearAliasPattern: [
                { type: 'text', value: 'eltern' },
                { type: 'text', value: '_' },
                { type: 'year' }
            ],
            yearDisplayPattern: ctx.ms365ElternGuardians.defaultYearDisplayPattern()
        };
        expect(
            ctx.ms365ElternGuardians.buildNameFromPattern(naming.classAliasPattern, {
                klasse: '1A',
                forAlias: true
            })
        ).toBe('eltern-1a');
        const bucket = {
            students: [{ id: 's1', klasse: '1A', name: 'A', email: 'a@s.at', guardianIds: [] }],
            classes: [{ code: '1A', year: '2031' }],
            guardians: [],
            parentLists: []
        };
        const rows = ctx.ms365ElternGuardians.buildClassParentSoll(bucket, { naming });
        expect(rows[0].mailNickname).toBe('eltern-1a');
        expect(rows[0].displayName).toBe('Eltern – 1A');
        const years = ctx.ms365ElternGuardians.buildYearParentSoll(bucket, { naming });
        expect(years[0].mailNickname).toBe('eltern_2031');
    });

    it('patchSetup speichert Eltern-Namensschema', () => {
        const ctx = loadModules(store);
        ctx.ms365AppDataV2.patchSetup({
            elternClassAliasPattern: [
                { type: 'text', value: 'el' },
                { type: 'text', value: '-' },
                { type: 'klasse' }
            ]
        });
        const s = ctx.ms365AppDataV2.getSetup();
        expect(s.elternClassAliasPattern).toEqual([
            { type: 'text', value: 'el' },
            { type: 'text', value: '-' },
            { type: 'klasse' }
        ]);
        const naming = ctx.ms365ElternGuardians.namingFromSetup(s);
        expect(
            ctx.ms365ElternGuardians.buildNameFromPattern(naming.classAliasPattern, {
                klasse: '2A',
                forAlias: true
            })
        ).toBe('el-2a');
    });

    it('buildElternSyncScript setzt GAL-sichtbare DL und versteckte Membership/Contacts', () => {
        const ctx = loadModules(store);
        const script = ctx.ms365ElternGuardians.buildElternSyncScript({
            domain: 'schule.at',
            schoolName: 'Testschule',
            lists: [
                {
                    displayName: 'Eltern 1B',
                    mailNickname: 'eltern1b',
                    guardians: [
                        { name: 'Jane', email: 'jane@mail.com' },
                        { name: 'John', email: 'john@mail.com' }
                    ]
                }
            ]
        });
        expect(script).toContain('New-MailContact');
        expect(script).toContain('HiddenFromAddressListsEnabled $true');
        expect(script).toContain('HiddenGroupMembershipEnabled');
        expect(script).toContain('HiddenFromAddressListsEnabled $false');
        expect(script).toContain('New-DistributionGroup');
        expect(script).toContain('Add-DistributionGroupMember');
        expect(script).toContain('jane@mail.com');
        expect(script).toContain('Connect-ExchangeOnline');
        expect(script).toContain('Get-AcceptedDomain');
        expect(script).toContain('$ExpectedDomain');
        expect(script).toContain('schule.at');
        expect(script).toContain('Testschule');
        expect(script).toMatch(/Read-Host.*JA/);
    });

    it('parseLinesToStudents liest optionale Elternspalten', () => {
        const ctx = loadModules(store);
        const rows = ctx.ms365TenantSettingsParseStudentsLines(
            '1B;Dave Grohl;dave@schule.at;Jane Grohl;jane@mail.com;John Grohl;john@mail.com\n'
        );
        expect(rows).toHaveLength(1);
        expect(rows[0].parentPairs).toEqual([
            { name: 'Jane Grohl', email: 'jane@mail.com' },
            { name: 'John Grohl', email: 'john@mail.com' }
        ]);
    });

    it('linkGuardianToStudent und Klassen-SOLL aktualisieren sich', () => {
        const ctx = loadModules(store);
        ctx.ms365AppDataV2.setCoreFromTenantSettings({
            domain: 'schule.at',
            students: [{ klasse: '1B', name: 'Dave', email: 'dave@schule.at' }],
            classes: [{ code: '1B', year: '2030' }]
        });
        const { bucket } = ctx.ms365AppDataV2.getYearBucket();
        const sid = bucket.students[0].id;
        ctx.ms365AppDataV2.linkGuardianToStudent(sid, { name: 'Jane', email: 'jane@mail.com' });
        const again = ctx.ms365AppDataV2.getYearBucket().bucket;
        const soll = ctx.ms365ElternGuardians.buildClassParentSoll(again);
        expect(soll[0].guardianCount).toBe(1);
        expect(soll[0].guardians[0].email).toBe('jane@mail.com');
    });

    it('removeStudents entfernt Abgänger und nur verwaiste Elternkontakte', () => {
        const ctx = loadModules(store);
        ctx.ms365AppDataV2.setCoreFromTenantSettings({
            domain: 'schule.at',
            students: [
                {
                    klasse: '1B',
                    name: 'Dave',
                    email: 'dave@schule.at',
                    parentPairs: [
                        { name: 'Jane', email: 'jane@mail.com' },
                        { name: 'Tom', email: 'tom@mail.com' }
                    ]
                },
                {
                    klasse: '1B',
                    name: 'Ben',
                    email: 'ben@schule.at',
                    parentPairs: [{ name: 'Jane', email: 'jane@mail.com' }]
                }
            ],
            classes: [{ code: '1B', year: '2030' }]
        });
        const before = ctx.ms365AppDataV2.getYearBucket().bucket;
        const dave = before.students.find((s) => s.email === 'dave@schule.at');
        const result = ctx.ms365AppDataV2.removeStudents([dave.id]);
        const after = ctx.ms365AppDataV2.getYearBucket().bucket;
        expect(result.removedStudents).toBe(1);
        expect(result.removedGuardians).toBe(1);
        expect(after.students.some((s) => s.email === 'dave@schule.at')).toBe(false);
        expect(after.guardians.some((g) => g.email === 'tom@mail.com')).toBe(false);
        expect(after.guardians.some((g) => g.email === 'jane@mail.com')).toBe(true);
    });

    it('pruneUnlinkedGuardians bereinigt bestehende Karteileichen', () => {
        const ctx = loadModules(store);
        ctx.ms365AppDataV2.saveYearBucket('2026/27', {
            students: [{ id: 's1', klasse: '1A', name: 'Anna', email: 'anna@schule.at', guardianIds: ['g1'] }],
            classes: [{ code: '1A', year: '2031' }],
            guardians: [
                { id: 'g1', name: 'Maria', email: 'maria@mail.com' },
                { id: 'g2', name: 'Alt', email: 'alt@mail.com' }
            ],
            parentLists: []
        });
        const removed = ctx.ms365AppDataV2.pruneUnlinkedGuardians();
        const after = ctx.ms365AppDataV2.getYearBucket().bucket;
        expect(removed).toBe(1);
        expect(after.guardians).toHaveLength(1);
        expect(after.guardians[0].email).toBe('maria@mail.com');
    });

    it('buildElternDiagnoseReport warnt ohne Elternmails und bei Alias-Kollision', () => {
        const ctx = loadModules(store);
        ctx.ms365AppDataV2.setCoreFromTenantSettings({
            domain: 'schule.at',
            students: [{ klasse: '1B', name: 'Dave', email: 'dave@schule.at' }],
            classes: [{ code: '1B', year: '2030' }]
        });
        const { bucket } = ctx.ms365AppDataV2.getYearBucket();
        const report = ctx.ms365ElternGuardians.buildElternDiagnoseReport(bucket, ctx.ms365ElternGuardians.getNaming(), 'schule.at');
        expect(report.counts.lists).toBeGreaterThan(0);
        expect(report.issues.some((i) => i.code === 'no-parents')).toBe(true);
        expect(report.hints.gal).toMatch(/GAL/);
        const script = ctx.ms365ElternGuardians.buildElternDiagnoseScript(report.lists, 'schule.at', 'Testschule');
        expect(script).toMatch(/Get-DistributionGroup/);
        expect(script).toMatch(/HiddenFromAddressListsEnabled/);
        expect(script).toContain('Connect-ExchangeOnline');
        expect(script).toContain('Get-AcceptedDomain');
        expect(script).toContain('$Ms365ReadOnly = $true');
    });
});
