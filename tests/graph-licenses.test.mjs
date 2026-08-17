import { describe, expect, it } from 'vitest';
import {
    applyStudentImportSelection,
    applyTeacherImportSelection,
    buildAssignableSkuOptions,
    buildLicenseFilterOptions,
    buildStudentImportPreview,
    buildTeacherImportPreview,
    facultyUserPlanSkuIds,
    remainingPrepaidUnits,
    studentUserPlanSkuIds,
    resolveSku,
    splitPersonName,
    suggestKlasseFromUser,
    suggestTeacherCode,
    summarizeUserLicenses,
    userMatchesLicenseFilter
} from '../src/shared/graph-licenses.js';

const A3_FAC = '4b590615-0888-425a-a965-b3bf7789848d';
const A1_FAC = '94763226-9b3c-4e75-a931-5c89701abe66';
const A1_STU = '314c4481-f395-4525-be8b-2ec4bb1e9d91';
const A3_STU = '7cfd9a2b-e110-4c39-bf20-c6a3f36a3121';

describe('graph-licenses SKU-Katalog', () => {
    it('erkennt Microsoft 365 A3 für Lehrpersonal', () => {
        const s = resolveSku(A3_FAC);
        expect(s.audience).toBe('faculty');
        expect(s.family).toBe('a3');
        expect(s.userPlan).toBe(true);
        expect(s.shortLabel).toMatch(/A3/);
    });

    it('erkennt Office 365 A1 für Lehrpersonal', () => {
        const s = resolveSku(A1_FAC);
        expect(s.family).toBe('a1');
        expect(s.audience).toBe('faculty');
        expect(s.userPlan).toBe(true);
    });

    it('klassifiziert unbekannte skuPartNumber heuristisch', () => {
        const s = resolveSku('aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', 'M365EDU_A3_FACULTY_CUSTOM');
        expect(s.audience).toBe('faculty');
        expect(s.family).toBe('a3');
        expect(s.userPlan).toBe(true);
    });

    it('liefert faculty-Userplan-SKU-IDs', () => {
        const ids = facultyUserPlanSkuIds();
        expect(ids).toContain(A3_FAC);
        expect(ids).toContain(A1_FAC);
        expect(ids).not.toContain(A1_STU);
    });

    it('liefert student-Userplan-SKU-IDs', () => {
        const ids = studentUserPlanSkuIds();
        expect(ids).toContain(A1_STU);
        expect(ids).toContain(A3_STU);
        expect(ids).not.toContain(A3_FAC);
    });
});

describe('graph-licenses Benutzer-Zusammenfassung und Filter', () => {
    const teacher = {
        assignedLicenses: [{ skuId: A3_FAC }]
    };
    const student = {
        assignedLicenses: [{ skuId: A1_STU }]
    };
    const bare = { assignedLicenses: [] };

    it('summarizeUserLicenses für A3-Lehrkraft', () => {
        const sum = summarizeUserLicenses(teacher);
        expect(sum.hasFacultyUserPlan).toBe(true);
        expect(sum.hasStudent).toBe(false);
        expect(sum.facultyFamilies).toEqual(['a3']);
        expect(sum.primaryLabel).toMatch(/A3/);
    });

    it('userMatchesLicenseFilter', () => {
        expect(userMatchesLicenseFilter(teacher, 'faculty')).toBe(true);
        expect(userMatchesLicenseFilter(teacher, 'faculty-a3')).toBe(true);
        expect(userMatchesLicenseFilter(teacher, 'faculty-a1')).toBe(false);
        expect(userMatchesLicenseFilter(teacher, 'student')).toBe(false);
        expect(userMatchesLicenseFilter(student, 'student')).toBe(true);
        expect(userMatchesLicenseFilter(student, 'student-a1')).toBe(true);
        expect(userMatchesLicenseFilter(student, 'student-a3')).toBe(false);
        expect(userMatchesLicenseFilter(teacher, 'student-a1')).toBe(false);
        expect(userMatchesLicenseFilter(bare, 'none')).toBe(true);
        expect(userMatchesLicenseFilter(teacher, 'none')).toBe(false);
        expect(userMatchesLicenseFilter(teacher, 'sku:' + A3_FAC)).toBe(true);
        expect(userMatchesLicenseFilter(teacher, '')).toBe(true);
    });

    it('buildLicenseFilterOptions zählt Gruppen', () => {
        const opts = buildLicenseFilterOptions([teacher, student, bare]);
        const byVal = Object.fromEntries(opts.map((o) => [o.value, o.label]));
        expect(byVal.faculty).toMatch(/1/);
        expect(byVal.student).toMatch(/1/);
        expect(byVal.none).toMatch(/1/);
        expect(byVal['faculty-a3']).toMatch(/1/);
        expect(byVal['student-a1']).toMatch(/1/);
        expect(opts.some((o) => o.value === 'sku:' + A3_FAC)).toBe(true);
    });
});

describe('graph-licenses Kürzel', () => {
    it('splitPersonName nimmt Nachnamen aus displayName', () => {
        expect(splitPersonName('Max Mustermann')).toEqual({ given: 'Max', surname: 'Mustermann' });
        expect(splitPersonName('Mag. Dr. Anna Beispiel')).toEqual({ given: 'Anna', surname: 'Beispiel' });
    });

    it('suggestTeacherCode: 3 Buchstaben, Umlaute, Kollision', () => {
        expect(suggestTeacherCode('Max Mustermann')).toBe('MUS');
        expect(suggestTeacherCode('Eva Müller')).toBe('MUE');
        expect(suggestTeacherCode('Anna Beispiel', 'Anna', 'Beispiel', ['BEI'])).toBe('BEA');
        const used = new Set(['mus', 'mum', 'musm']);
        const c = suggestTeacherCode('Max Mustermann', 'Max', 'Mustermann', used);
        expect(c).toBeTruthy();
        expect(used.has(c.toLowerCase())).toBe(false);
    });
});

describe('graph-licenses Import-Vorschau und Übernahme', () => {
    const users = [
        {
            id: 'u1',
            displayName: 'Max Mustermann',
            givenName: 'Max',
            surname: 'Mustermann',
            mail: 'max.mustermann@schule.at',
            userPrincipalName: 'max.mustermann@schule.at',
            accountEnabled: true,
            userType: 'Member',
            assignedLicenses: [{ skuId: A3_FAC }]
        },
        {
            id: 'u2',
            displayName: 'Lisa Schüler',
            mail: 'lisa@schule.at',
            userPrincipalName: 'lisa@schule.at',
            accountEnabled: true,
            assignedLicenses: [{ skuId: A1_STU }]
        },
        {
            id: 'u3',
            displayName: 'Alt Lehrer',
            mail: 'alt@schule.at',
            userPrincipalName: 'alt@schule.at',
            accountEnabled: true,
            assignedLicenses: [{ skuId: A1_FAC }]
        },
        {
            id: 'u4',
            displayName: 'Inaktiv',
            mail: 'inaktiv@schule.at',
            accountEnabled: false,
            assignedLicenses: [{ skuId: A3_FAC }]
        }
    ];

    it('Vorschau: nur Lehrpersonal-Userpläne, Schüler raus, Inaktive optional', () => {
        const rows = buildTeacherImportPreview(users, [{ code: 'ALT', name: 'Alt Lehrer', email: 'alt@schule.at' }], null, {
            activeOnly: true,
            guests: false,
            families: ['a1', 'a3', 'a5']
        });
        expect(rows.map((r) => r.email).sort()).toEqual(['alt@schule.at', 'max.mustermann@schule.at']);
        const neu = rows.find((r) => r.email === 'max.mustermann@schule.at');
        expect(neu.selected).toBe(true);
        expect(neu.alreadyInList).toBe(false);
        expect(neu.code).toBe('MUS');
        const alt = rows.find((r) => r.email === 'alt@schule.at');
        expect(alt.alreadyInList).toBe(true);
        expect(alt.selected).toBe(false);
        expect(alt.code).toBe('ALT');
    });

    it('Übernahme ergänzt neue Zeilen und aktualisiert Namen, ohne Duplikate', () => {
        const preview = buildTeacherImportPreview(users, [{ code: 'ALT', name: 'Alter Name', email: 'alt@schule.at' }], null, {
            activeOnly: true,
            families: ['a1', 'a3', 'a5']
        });
        preview.forEach((r) => {
            r.selected = true;
        });
        const result = applyTeacherImportSelection(
            [{ code: 'ALT', name: 'Alter Name', email: 'alt@schule.at' }],
            preview
        );
        expect(result.added.length).toBe(1);
        expect(result.added[0].code).toBe('MUS');
        expect(result.added[0].email).toBe('max.mustermann@schule.at');
        expect(result.updated.length).toBe(1);
        expect(result.updated[0].name).toBe('Alt Lehrer');
        expect(result.teachers.length).toBe(2);
        expect(result.directoryMatches['max.mustermann@schule.at'].graphUserId).toBe('u1');
    });
});

describe('graph-licenses Schüler-Import', () => {
    const A3_STU_ID = '7cfd9a2b-e110-4c39-bf20-c6a3f36a3121';
    const users = [
        {
            id: 's1',
            displayName: 'Lisa Beispiel',
            givenName: 'Lisa',
            surname: 'Beispiel',
            mail: 'lisa@schule.at',
            userPrincipalName: 'lisa@schule.at',
            department: '1A',
            accountEnabled: true,
            assignedLicenses: [{ skuId: A1_STU }]
        },
        {
            id: 's2',
            displayName: 'Tom Alt',
            mail: 'tom@schule.at',
            userPrincipalName: 'tom@schule.at',
            department: '2B',
            accountEnabled: true,
            assignedLicenses: [{ skuId: A3_STU_ID }]
        },
        {
            id: 't1',
            displayName: 'Max Lehrer',
            mail: 'max@schule.at',
            assignedLicenses: [{ skuId: A3_FAC }]
        }
    ];

    it('suggestKlasseFromUser liest 1A aus department', () => {
        expect(suggestKlasseFromUser({ department: '1A' })).toBe('1A');
        expect(suggestKlasseFromUser({ department: '5HMA' })).toBe('5HMA');
        expect(suggestKlasseFromUser({ department: 'Schülerinnen der Unterstufe' })).toBe('');
    });

    it('Vorschau: nur Schüler-Userpläne, Lehrkräfte raus, Klasse aus Abteilung', () => {
        const rows = buildStudentImportPreview(
            users,
            [{ klasse: '2B', name: 'Tom Alt', email: 'tom@schule.at' }],
            null,
            { activeOnly: true, guests: false, families: ['a1', 'a3', 'a5'] }
        );
        expect(rows.map((r) => r.email).sort()).toEqual(['lisa@schule.at', 'tom@schule.at']);
        const neu = rows.find((r) => r.email === 'lisa@schule.at');
        expect(neu.selected).toBe(true);
        expect(neu.klasse).toBe('1A');
        const alt = rows.find((r) => r.email === 'tom@schule.at');
        expect(alt.alreadyInList).toBe(true);
        expect(alt.selected).toBe(false);
        expect(alt.klasse).toBe('2B');
    });

    it('Übernahme ergänzt neue Zeilen und füllt leere Klasse nach', () => {
        const preview = buildStudentImportPreview(
            users,
            [{ klasse: '', name: 'Tom Alt', email: 'tom@schule.at' }],
            null,
            { activeOnly: true, families: ['a1', 'a3', 'a5'] }
        );
        preview.forEach((r) => {
            r.selected = true;
        });
        const result = applyStudentImportSelection(
            [{ klasse: '', name: 'Alter Name', email: 'tom@schule.at' }],
            preview
        );
        expect(result.added.length).toBe(1);
        expect(result.added[0].email).toBe('lisa@schule.at');
        expect(result.added[0].klasse).toBe('1A');
        expect(result.updated.length).toBe(1);
        expect(result.updated[0].name).toBe('Tom Alt');
        expect(result.updated[0].klasse).toBe('2B');
        expect(result.students.length).toBe(2);
        expect(result.directoryMatches['lisa@schule.at'].graphUserId).toBe('s1');
    });
});

describe('graph-licenses Zuweisung (freie Sitze)', () => {
    it('remainingPrepaidUnits rechnet enabled minus consumed', () => {
        expect(remainingPrepaidUnits(null)).toBe(null);
        expect(
            remainingPrepaidUnits({ prepaidUnits: { enabled: 100 }, consumedUnits: 37 })
        ).toBe(63);
        expect(
            remainingPrepaidUnits({ prepaidUnits: { enabled: 10 }, consumedUnits: 10 })
        ).toBe(0);
        expect(remainingPrepaidUnits({ prepaidUnits: {}, consumedUnits: 2 })).toBe(null);
    });

    it('buildAssignableSkuOptions lässt ausgeschöpfte und bereits zugewiesene SKUs weg', () => {
        const opts = buildAssignableSkuOptions(
            [
                {
                    skuId: A3_FAC,
                    skuPartNumber: 'M365EDU_A3_FACULTY',
                    capabilityStatus: 'Enabled',
                    prepaidUnits: { enabled: 50 },
                    consumedUnits: 12
                },
                {
                    skuId: A1_FAC,
                    skuPartNumber: 'STANDARDWOFFPACK_FACULTY',
                    capabilityStatus: 'Enabled',
                    prepaidUnits: { enabled: 20 },
                    consumedUnits: 20
                },
                {
                    skuId: A1_STU,
                    skuPartNumber: 'STANDARDWOFFPACK_STUDENT',
                    capabilityStatus: 'Enabled',
                    prepaidUnits: { enabled: 200 },
                    consumedUnits: 10
                }
            ],
            [A3_FAC]
        );
        expect(opts.map((o) => o.skuId)).toEqual([A1_STU]);
        expect(opts[0].remaining).toBe(190);
    });

    it('ohne Tenant-SKUs fällt auf den Education-Katalog zurück', () => {
        const opts = buildAssignableSkuOptions([], [A3_FAC], { fallbackCatalog: true });
        expect(opts.length).toBeGreaterThan(3);
        expect(opts.some((o) => o.skuId === A3_FAC)).toBe(false);
        expect(opts.some((o) => o.skuId === A1_STU)).toBe(true);
        expect(opts.some((o) => o.skuId === A3_STU)).toBe(true);
    });

    it('ohne Tenant-SKUs und ohne Fallback bleibt die Liste leer', () => {
        expect(buildAssignableSkuOptions([], [], { fallbackCatalog: false })).toEqual([]);
    });
});
