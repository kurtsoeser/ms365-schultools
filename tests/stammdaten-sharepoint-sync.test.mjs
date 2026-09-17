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
});
