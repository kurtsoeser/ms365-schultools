import { describe, it, expect } from 'vitest';
import {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    IT_LIBRARY_TITLE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    normalizeDriveListFolder,
    sortDriveBrowserItems,
    sortDriveBrowserItemsByColumn,
    isLikelyImportableBackupFileName,
    CONFIG_FOLDER,
    describeRemoteBackup,
    isBroadSiteAudience,
    entraGroupLogonName,
    buildItLibraryPlan,
    isItLibraryConfigured,
    resolveItLibraryBrowserHref,
    collectItLibraryLinkHints,
    uniqueItLibrarySiteUrls,
    formatItLibraryLinkSkipDe,
    scoreItLibraryDiscoverySite,
    normalizeItLibraryMeta,
    compareBackupPayloads,
    isLikelyFreshLocalBackup,
    formatBackupCompareDe,
    isAutoSyncIgnoredChangeSource,
    shouldApplyRemoteBackup,
    formatSyncStatusDe,
    SPO_ROLE,
    designHintSummaryDe,
    designHintBulletsDe,
    designHintDe
} from '../src/shared/stammdaten-sharepoint-sync-logic.js';

describe('stammdaten-sharepoint-sync-logic', () => {
    it('baut Ordner/Datei-Pfad', () => {
        expect(buildDriveRelativePath('', CURRENT_FILE)).toBe(CURRENT_FILE);
        expect(buildDriveRelativePath(DEFAULT_FOLDER, CURRENT_FILE)).toBe('Backups/ms365-stammdaten-aktuell.json');
        expect(buildDriveRelativePath('/a/b/', 'x.json')).toBe('a/b/x.json');
    });

    it('normalisiert Drive-Listing-Pfade', () => {
        expect(normalizeDriveListFolder('')).toBe('');
        expect(normalizeDriveListFolder('/Backups/')).toBe('Backups');
        expect(normalizeDriveListFolder('config\\manifest')).toBe('config/manifest');
    });

    it('sortiert Ordner vor Dateien', () => {
        const sorted = sortDriveBrowserItems([
            { name: 'z.json', file: {} },
            { name: 'Backups', folder: {} },
            { name: 'config', folder: {} },
            { name: 'a.json', file: {} }
        ]);
        expect(sorted.map((x) => x.name)).toEqual(['Backups', 'config', 'a.json', 'z.json']);
    });

    it('erkennt importierbare Backup-Dateinamen', () => {
        expect(
            isLikelyImportableBackupFileName('ms365-browser-backup-2026-10-08-Schule.json', { folder: 'Backups' })
        ).toBe(true);
        expect(isLikelyImportableBackupFileName(CURRENT_FILE, { folder: 'Backups' })).toBe(true);
        expect(isLikelyImportableBackupFileName('manifest.json', { folder: CONFIG_FOLDER })).toBe(false);
        expect(isLikelyImportableBackupFileName('permissions-freistellung.json', { folder: 'config' })).toBe(false);
    });

    it('sortiert Browse-Spalten mit Richtung', () => {
        const rows = [
            { name: 'b.json', file: {}, size: 200, lastModifiedDateTime: '2026-01-02T10:00:00Z' },
            { name: 'a.json', file: {}, size: 100, lastModifiedDateTime: '2026-01-03T10:00:00Z' }
        ];
        expect(sortDriveBrowserItemsByColumn(rows, 'size', 1).map((x) => x.name)).toEqual(['a.json', 'b.json']);
        expect(sortDriveBrowserItemsByColumn(rows, 'modified', -1).map((x) => x.name)).toEqual(['a.json', 'b.json']);
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

    it('liefert kompakte Hero-Hinweise', () => {
        expect(designHintSummaryDe()).toContain(IT_LIBRARY_TITLE);
        expect(designHintBulletsDe().length).toBeGreaterThanOrEqual(4);
        expect(designHintDe()).toContain(designHintBulletsDe()[0]);
    });

    it('priorisiert Bibliotheks-URL vor Backup-Datei für Browser-Link', () => {
        const lib = 'https://contoso.sharepoint.com/sites/s/MS365-IT-Stammdaten';
        const file = 'https://contoso.sharepoint.com/sites/s/MS365-IT-Stammdaten/Backups/ms365-stammdaten-aktuell.json';
        expect(resolveItLibraryBrowserHref({ webUrl: lib }, { webUrl: file })).toBe(lib);
        expect(resolveItLibraryBrowserHref(null, { webUrl: file })).toBe(file);
        expect(resolveItLibraryBrowserHref({ webUrl: lib }, null)).toBe(lib);
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

        const fromSchoolHub = collectItLibraryLinkHints({
            setup: { schoolIntranetSiteUrl: 'https://schule.sharepoint.com/sites/schulhub' }
        });
        expect(fromSchoolHub.hasMinimum).toBe(true);
        expect(fromSchoolHub.siteUrl).toContain('schulhub');

        const both = collectItLibraryLinkHints({
            setup: {
                intranetSiteUrl: 'https://schule.sharepoint.com/sites/intranet',
                schoolIntranetSiteUrl: 'https://schule.sharepoint.com/sites/schulhub'
            }
        });
        expect(both.siteUrls.length).toBe(2);
        expect(both.siteUrls[0]).toContain('intranet');
        expect(both.siteUrls[1]).toContain('schulhub');
        expect(uniqueItLibrarySiteUrls('https://a/sites/x/', 'https://a/sites/x', 'https://b/sites/y')).toEqual([
            'https://a/sites/x',
            'https://b/sites/y'
        ]);
        expect(formatItLibraryLinkSkipDe('no-hints')).toMatch(/Site-URL/i);
        expect(formatItLibraryLinkSkipDe('library-not-found')).toMatch(/MS365-IT-Stammdaten/);
    });

    it('bewertet Discovery-Site-Kandidaten', () => {
        expect(scoreItLibraryDiscoverySite('MS365-IT-Stammdaten', 'https://x/sites/a')).toBe(0);
        expect(scoreItLibraryDiscoverySite('Schultools', 'https://x/sites/schultools')).toBeLessThan(
            scoreItLibraryDiscoverySite('Allgemein', 'https://x/sites/allgemein')
        );
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
