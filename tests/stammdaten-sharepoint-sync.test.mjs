import { describe, it, expect } from 'vitest';
import {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    IT_LIBRARY_TITLE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    describeRemoteBackup,
    isBroadSiteAudience,
    entraGroupLogonName,
    buildItLibraryPlan,
    isItLibraryConfigured,
    collectItLibraryLinkHints,
    normalizeItLibraryMeta,
    compareBackupPayloads,
    isLikelyFreshLocalBackup,
    formatBackupCompareDe,
    isAutoSyncIgnoredChangeSource,
    shouldApplyRemoteBackup,
    formatSyncStatusDe,
    SPO_ROLE
} from '../src/shared/stammdaten-sharepoint-sync-logic.js';

describe('stammdaten-sharepoint-sync-logic', () => {
    it('baut Ordner/Datei-Pfad', () => {
        expect(buildDriveRelativePath('', CURRENT_FILE)).toBe(CURRENT_FILE);
        expect(buildDriveRelativePath(DEFAULT_FOLDER, CURRENT_FILE)).toBe('Backups/ms365-stammdaten-aktuell.json');
        expect(buildDriveRelativePath('/a/b/', 'x.json')).toBe('a/b/x.json');
    });

    it('encodiert Graph root-Pfad', () => {
        expect(encodeDriveRootPath('Backups/ms365-stammdaten-aktuell.json')).toBe(
            'root:/Backups/ms365-stammdaten-aktuell.json:'
        );
        expect(encodeDriveRootPath('Ordner mit Leerzeichen/a.json')).toContain('%20');
    });

    it('beschreibt Meta', () => {
        const s = describeRemoteBackup({
            schoolName: 'Testschule',
            exportedAt: '2026-09-11T10:00:00.000Z',
            keyCount: 12
        });
        expect(s).toContain('Testschule');
        expect(s).toContain('12 Schlüssel');
    });

    it('erkennt breite Site-Rollen', () => {
        expect(isBroadSiteAudience({ Title: 'Intranet Visitors' })).toBe(true);
        expect(isBroadSiteAudience({ Title: 'Intranet Members' })).toBe(true);
        expect(isBroadSiteAudience({ Title: 'Intranet Owners' })).toBe(false);
        expect(isBroadSiteAudience({ LoginName: 'c:0o.c|…', Title: 'Schulverwaltung' })).toBe(false);
    });

    it('baut IT-Bibliothek-Plan', () => {
        expect(IT_LIBRARY_TITLE).toContain('IT');
        expect(entraGroupLogonName('aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee')).toContain('federateddirectoryclaimprovider');
        const bad = buildItLibraryPlan({});
        expect(bad.ok).toBe(false);
        const ok = buildItLibraryPlan({ itGroupId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' });
        expect(ok.ok).toBe(true);
        expect(ok.roleDefId).toBe(SPO_ROLE.contribute);
    });

    it('erkennt eingerichtete IT-Bibliothek', () => {
        expect(isItLibraryConfigured(null)).toBe(false);
        expect(isItLibraryConfigured({})).toBe(false);
        expect(isItLibraryConfigured({ driveId: 'abc' })).toBe(true);
    });

    it('vergleicht Browser-Backups', () => {
        const a = {
            exportedAt: '2026-09-20T10:00:00.000Z',
            contentFingerprint: 'abc:3',
            schoolName: 'A',
            localStorage: { 'ms365-schooltool-data-v2': { x: 1 }, 'ms365-tenant-settings-v1': {} }
        };
        const b = {
            exportedAt: '2026-09-21T12:00:00.000Z',
            contentFingerprint: 'def:3',
            schoolName: 'A',
            localStorage: { 'ms365-schooltool-data-v2': { x: 2 }, 'ms365-tenant-settings-v1': {} }
        };
        const same = compareBackupPayloads(a, a);
        expect(same.identical).toBe(true);
        const diff = compareBackupPayloads(a, b);
        expect(diff.identical).toBe(false);
        expect(diff.newerSide).toBe('remote');
        expect(diff.changedCount).toBe(1);
        expect(formatBackupCompareDe(diff, { remoteLastModified: '2026-09-21T12:05:00.000Z' })).toMatch(
            /SharePoint/
        );
        expect(isLikelyFreshLocalBackup({ localStorage: {} })).toBe(true);
        expect(
            isLikelyFreshLocalBackup({
                localStorage: {
                    'ms365-schooltool-data-v2': JSON.stringify({ years: { '2025/26': {} } }),
                    'ms365-tenant-settings-v1': '{}'
                }
            })
        ).toBe(false);
    });

    it('sammelt Auto-Verknüpfungs-Hinweise', () => {
        const hints = collectItLibraryLinkHints({
            setup: { intranetSiteUrl: 'https://schule.sharepoint.com/sites/intranet' },
            formDraft: { libraryTitle: 'MS365-IT-Stammdaten', itGroup: 'verwaltung@schule.at' }
        });
        expect(hints.hasMinimum).toBe(true);
        expect(hints.siteUrl).toContain('intranet');
        expect(hints.itGroupMail).toBe('verwaltung@schule.at');
        expect(normalizeItLibraryMeta({ driveId: 'x', listTitle: ' Lib ' }).listTitle).toBe('Lib');
    });

    it('entscheidet Session-Pull anhand von Versionen', () => {
        expect(shouldApplyRemoteBackup({ remoteExists: false }).apply).toBe(false);
        expect(shouldApplyRemoteBackup({ remoteExists: true, localDirty: true }).reason).toBe('local-dirty');
        expect(
            shouldApplyRemoteBackup({
                remoteExists: true,
                remoteLastModified: '2026-09-21T10:00:00Z',
                localRemoteLastModified: '2026-09-21T10:00:00Z'
            }).apply
        ).toBe(false);
        expect(
            shouldApplyRemoteBackup({
                remoteExists: true,
                remoteLastModified: '2026-09-21T12:00:00Z',
                localRemoteLastModified: '2026-09-21T10:00:00Z'
            }).apply
        ).toBe(true);
        expect(
            shouldApplyRemoteBackup({
                remoteExists: true,
                remoteLastModified: '2026-09-21T12:00:00Z'
            }).reason
        ).toBe('never-synced');
    });

    it('ignoriert Auto-Push-Quellen vom Sync selbst', () => {
        expect(isAutoSyncIgnoredChangeSource('browser-backup-import')).toBe(true);
        expect(isAutoSyncIgnoredChangeSource('spo-auto-pull')).toBe(true);
        expect(isAutoSyncIgnoredChangeSource('render')).toBe(true);
        expect(isAutoSyncIgnoredChangeSource('autosave')).toBe(false);
    });

    it('formatiert Sync-Status', () => {
        expect(formatSyncStatusDe({ ready: false })).toMatch(/nicht eingerichtet/);
        expect(
            formatSyncStatusDe({
                ready: true,
                dirty: true,
                libraryTitle: 'MS365-IT-Stammdaten'
            })
        ).toMatch(/ausstehend/);
        expect(
            formatSyncStatusDe({
                ready: true,
                lastAt: '2026-09-21T10:00:00.000Z',
                lastDirection: 'push'
            })
        ).toMatch(/gesichert/);
    });
});
