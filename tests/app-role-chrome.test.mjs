import { describe, it, expect } from 'vitest';
import {
    appRoleChromeModel,
    appRoleChromeText,
    appRoleChipLabel,
    appRoleSourceLabel,
    matchSchoolAppFromPath
} from '../src/shared/app-role-chrome.js';
import {
    readAppRoleInventory,
    summarizePermissionRecord
} from '../src/shared/app-role-inventory.js';

describe('app-role-chrome', () => {
    it('erkennt die Schul-App an der Seite', () => {
        const app = matchSchoolAppFromPath('/tools/lehrer-freistellung-planer.html');
        expect(app && app.id).toBe('lehrer-freistellung');
        expect(matchSchoolAppFromPath('/index.html')).toBe(null);
    });

    it('ohne Konto bleibt die Zeile offen', () => {
        const model = appRoleChromeModel(
            matchSchoolAppFromPath('lehrer-freistellung-planer.html'),
            { email: '' }
        );
        expect(model.pending).toBe(true);
        expect(appRoleChromeText(model)).toBe('Freistellungen Lehrkräfte · Konto wird gelesen');
    });

    it('Direktor zeigt Rolle und Quelle dieser App', () => {
        const model = appRoleChromeModel(
            matchSchoolAppFromPath('lehrer-freistellung-planer.html'),
            { email: 'direktor@ms365.schule', role: 'direktion', source: 'stammdaten-direktion' }
        );
        expect(appRoleSourceLabel('stammdaten-direktion')).toBe('Schulleitung (Stammdaten)');
        expect(appRoleChipLabel(model)).toBe('Direktion');
        expect(appRoleChromeText(model)).toBe(
            'Freistellungen Lehrkräfte · Direktion · Schulleitung (Stammdaten)'
        );
    });

    it('ohne echte Quelle bleibt nur das Rollenwort', () => {
        const model = appRoleChromeModel(matchSchoolAppFromPath('srdp-anmeldung.html'), {
            email: 'schueler01.demo@ms365.schule',
            role: 'schueler',
            source: 'default'
        });
        expect(model.sourceLabel).toBe('');
        expect(appRoleChipLabel(model)).toBe('Schüler/in');
    });
});

describe('app-role-inventory', () => {
    it('fasst Gruppen und Personen zusammen und lässt nackte IDs weg', () => {
        const lines = summarizePermissionRecord({
            groupLehrerId: '11111111-1111-1111-1111-111111111111',
            groupLehrerName: 'DEMO Lehrer:innen',
            emailDirektion: 'direktor@ms365.schule',
            direktionUsers: [{ displayName: 'Direktor', mail: 'direktor@ms365.schule' }]
        });
        expect(lines).toContain('Lehrkräfte-Gruppe: DEMO Lehrer:innen');
        expect(lines).toContain('Genehmiger: direktor@ms365.schule');
        expect(lines.some((line) => line.includes('Direktor'))).toBe(true);
        expect(lines.some((line) => line.includes('11111111'))).toBe(false);
    });

    it('liest die Lehrer-Freistellung aus dem Browser-Speicher', () => {
        const map = {
            'ms365-lfr-perms-v1': JSON.stringify({
                groupLehrerName: 'DEMO Lehrer:innen',
                emailDirektion: 'direktor@ms365.schule'
            }),
            'ms365-dashboard-audience-groups-v1': JSON.stringify({
                groupLehrerName: 'DEMO Lehrer:innen'
            })
        };
        const data = readAppRoleInventory({
            getItem(key) {
                return Object.prototype.hasOwnProperty.call(map, key) ? map[key] : null;
            }
        });
        const row = data.apps.find((app) => app.id === 'lehrer-freistellung');
        expect(row && row.configured).toBe(true);
        expect(row.lines.join(' | ')).toContain('direktor@ms365.schule');
        expect(data.principals.join(' | ')).toContain('DEMO Lehrer:innen');
    });
});
