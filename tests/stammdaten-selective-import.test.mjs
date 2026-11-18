import { describe, it, expect } from 'vitest';
import {
    buildSelectiveImportPlan,
    mergeSchooltoolSelective,
    buildSelectiveLocalStoragePatch,
    describeSelectiveImport
} from '../src/shared/stammdaten-sharepoint-selective-import.js';

function backupWithV2(v2, exportedAt) {
    return {
        kind: 'ms365-browser-backup-v1',
        exportedAt: exportedAt || '2026-01-01T12:00:00.000Z',
        localStorage: {
            'ms365-schooltool-data-v2': v2,
            'ms365-dashboard-tool-access-v1': { tools: { a: true } }
        }
    };
}

function v2(overrides) {
    const base = {
        version: 4,
        core: {
            schoolName: 'Lokal',
            domain: 'lokal.at',
            teachers: [{ name: 'Lehrer L' }],
            subjects: [{ code: 'M' }]
        },
        years: {
            current: '2025/26',
            byLabel: {
                '2025/26': {
                    students: [],
                    classes: [{ name: '1A' }],
                    guardians: [],
                    parentLists: [],
                    studentCouncil: []
                }
            }
        },
        setup: { matched: { schuelerGroupId: 'local-g' } }
    };
    return Object.assign(base, overrides || {});
}

describe('stammdaten-selective-import', () => {
    it('erkennt fehlende Schüler lokal vs. SharePoint und empfiehlt Import', () => {
        const local = backupWithV2(v2(), '2026-03-10T10:00:00.000Z');
        const remote = backupWithV2(
            v2({
                core: {
                    schoolName: 'Lokal',
                    domain: 'lokal.at',
                    teachers: [{ name: 'Lehrer L' }],
                    subjects: [{ code: 'M' }]
                },
                years: {
                    current: '2025/26',
                    byLabel: {
                        '2025/26': {
                            students: [
                                { name: 'Anna Test', email: 'anna@demo.at' },
                                { name: 'Ben Demo', email: 'ben@demo.at' }
                            ],
                            classes: [{ name: '1A' }],
                            guardians: [],
                            parentLists: [],
                            studentCouncil: []
                        }
                    }
                }
            }),
            '2026-03-01T10:00:00.000Z'
        );

        const plan = buildSelectiveImportPlan(local, remote, { newerSide: 'local' });
        const students = plan.sections.find((s) => s.id === 'students');
        expect(students).toBeTruthy();
        expect(students.status).toBe('differs');
        expect(students.remote.total).toBe(2);
        expect(students.local.total).toBe(0);
        expect(students.recommended).toBe(true);
        expect(students.remote.samples.length).toBeGreaterThan(0);
    });

    it('übernimmt Schulprofil inkl. Logo aus setup.schoolProfile', () => {
        const localV2 = v2({
            setup: {
                matched: { schuelerGroupId: 'local-g' },
                schoolProfile: { schoolCode: '111', logoDataUrl: '', city: 'Graz' }
            }
        });
        const remoteV2 = v2({
            core: {
                schoolName: 'Remote Schule',
                domain: 'remote.at',
                verifiedEmailDomains: ['alt.remote.at'],
                teachers: [{ name: 'Lehrer R' }],
                subjects: [{ code: 'D' }]
            },
            setup: {
                matched: { schuelerGroupId: 'remote-g' },
                schoolProfile: {
                    schoolCode: '999',
                    logoDataUrl: 'data:image/png;base64,remoteLogo',
                    logoFileName: 'remote.png',
                    city: 'Linz',
                    street: 'Hauptplatz 1'
                }
            }
        });
        const merged = mergeSchooltoolSelective(localV2, remoteV2, ['schoolMeta']);
        expect(merged.core.schoolName).toBe('Remote Schule');
        expect(merged.core.domain).toBe('remote.at');
        expect(merged.core.verifiedEmailDomains).toEqual(['alt.remote.at']);
        expect(merged.setup.schoolProfile.logoDataUrl).toContain('remoteLogo');
        expect(merged.setup.schoolProfile.city).toBe('Linz');
        expect(merged.setup.schoolProfile.schoolCode).toBe('999');
        // matched Groups bleiben lokal, weil nur schoolMeta gewählt wurde
        expect(merged.setup.matched.schuelerGroupId).toBe('local-g');
        expect(merged.core.teachers[0].name).toBe('Lehrer L');
    });

    it('merged nur Schüler und lässt Lehrer lokal', () => {
        const localV2 = v2({
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        students: [],
                        classes: [{ name: '1A' }],
                        guardians: [],
                        parentLists: [],
                        studentCouncil: []
                    }
                }
            }
        });
        const remoteV2 = v2({
            core: {
                schoolName: 'Remote Schule',
                domain: 'remote.at',
                teachers: [{ name: 'Lehrer R' }],
                subjects: [{ code: 'D' }]
            },
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        students: [{ name: 'Clara' }],
                        classes: [{ name: '2B' }],
                        guardians: [],
                        parentLists: [],
                        studentCouncil: []
                    }
                }
            }
        });

        const merged = mergeSchooltoolSelective(localV2, remoteV2, ['students']);
        expect(merged.years.byLabel['2025/26'].students).toHaveLength(1);
        expect(merged.years.byLabel['2025/26'].students[0].name).toBe('Clara');
        expect(merged.years.byLabel['2025/26'].classes[0].name).toBe('1A');
        expect(merged.core.teachers[0].name).toBe('Lehrer L');
        expect(merged.core.schoolName).toBe('Lokal');
    });

    it('baut localStorage-Patch nur für gewählte Bereiche', () => {
        const local = backupWithV2(v2());
        const remote = backupWithV2(
            v2({
                years: {
                    current: '2025/26',
                    byLabel: {
                        '2025/26': {
                            students: [{ name: 'Dana' }],
                            classes: [{ name: '1A' }],
                            guardians: [],
                            parentLists: [],
                            studentCouncil: []
                        }
                    }
                }
            })
        );
        remote.localStorage['ms365-dashboard-tool-access-v1'] = { tools: { b: true } };

        const patch = buildSelectiveLocalStoragePatch(local, remote, ['students']);
        expect(Object.keys(patch)).toEqual(['ms365-schooltool-data-v2']);
        expect(patch['ms365-schooltool-data-v2'].years.byLabel['2025/26'].students[0].name).toBe(
            'Dana'
        );

        const patchDash = buildSelectiveLocalStoragePatch(local, remote, ['bundle-dashboard']);
        expect(patchDash['ms365-dashboard-tool-access-v1']).toEqual({ tools: { b: true } });
        expect(patchDash['ms365-schooltool-data-v2']).toBeUndefined();
    });

    it('describeSelectiveImport liefert Vorschauzeilen', () => {
        const local = backupWithV2(v2());
        const remote = backupWithV2(
            v2({
                years: {
                    current: '2025/26',
                    byLabel: {
                        '2025/26': {
                            students: [{ name: 'Eva' }],
                            classes: [{ name: '1A' }],
                            guardians: [],
                            parentLists: [],
                            studentCouncil: []
                        }
                    }
                }
            })
        );
        const desc = describeSelectiveImport(local, remote, ['students']);
        expect(desc.sectionIds).toContain('students');
        expect(desc.previewLines.some((l) => /Schüler/i.test(l))).toBe(true);
        expect(desc.writtenKeys).toContain('ms365-schooltool-data-v2');
    });
});
