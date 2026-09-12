import { describe, it, expect } from 'vitest';
import {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    describeRemoteBackup
} from '../src/shared/stammdaten-sharepoint-sync-logic.js';

describe('stammdaten-sharepoint-sync-logic', () => {
    it('baut Ordner/Datei-Pfad', () => {
        expect(buildDriveRelativePath('', CURRENT_FILE)).toBe(CURRENT_FILE);
        expect(buildDriveRelativePath(DEFAULT_FOLDER, CURRENT_FILE)).toBe(
            'MS365-Schulverwaltung/ms365-stammdaten-aktuell.json'
        );
        expect(buildDriveRelativePath('/a/b/', 'x.json')).toBe('a/b/x.json');
    });

    it('encodiert Graph root-Pfad', () => {
        expect(encodeDriveRootPath('MS365-Schulverwaltung/ms365-stammdaten-aktuell.json')).toBe(
            'root:/MS365-Schulverwaltung/ms365-stammdaten-aktuell.json:'
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
});
