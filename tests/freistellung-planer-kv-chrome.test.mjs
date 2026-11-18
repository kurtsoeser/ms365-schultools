import { describe, it, expect } from 'vitest';
import {
    viewsForRole,
    navSectionsForRole,
    useKvPlanerChrome,
    useDirektionPlanerChrome,
    useStudentPlanerChrome,
    canDeleteFreistellung,
    hasFrKvCapability
} from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('freistellung KV chrome', () => {
    it('useKvPlanerChrome nur für Rolle kv ohne Demo-Umschalter', () => {
        expect(useKvPlanerChrome({ role: 'kv', demoRoleOverride: false })).toBe(true);
        expect(useKvPlanerChrome({ role: 'kv', demoRoleOverride: true })).toBe(false);
        expect(useKvPlanerChrome({ role: 'direktion', demoRoleOverride: false })).toBe(false);
        expect(useStudentPlanerChrome({ role: 'kv', demoRoleOverride: false })).toBe(false);
    });

    it('viewsForRole für KV bleibt flach und vollständig (Capability, kein Dauer-Modus)', () => {
        const state = { role: 'kv', demoRoleOverride: false, entraGroupsConfigured: true };
        const views = viewsForRole('kv', state);
        expect(views.map((v) => v.id)).toEqual([
            'dashboard',
            'liste',
            'kalender',
            'freigabe',
            'antrag'
        ]);
    });

    it('viewsForRole für KV im Demo-Modus behält Arbeits-Navigation', () => {
        const state = { role: 'kv', demoRoleOverride: true };
        const views = viewsForRole('kv', state);
        expect(views.some((v) => v.id === 'liste')).toBe(true);
        expect(views.some((v) => v.id === 'freigabe')).toBe(true);
    });

    it('viewsForRole für Schüler: Meine Anträge, Kalender und Antrag stellen', () => {
        const views = viewsForRole('schueler', { role: 'schueler' });
        expect(views.map((v) => v.id)).toEqual(['meine', 'kalender', 'antrag']);
    });

    it('hasFrKvCapability erkennt Stammdaten-KV ohne App-Modus', () => {
        expect(hasFrKvCapability({ role: 'schueler', kvMatch: { classCode: '1AK' } })).toBe(true);
        expect(hasFrKvCapability({ role: 'schueler', planerRoles: ['kv'] })).toBe(true);
        expect(hasFrKvCapability({ role: 'schueler', planerRoles: [] })).toBe(false);
    });
});

describe('freistellung Direktion chrome', () => {
    it('useDirektionPlanerChrome nur für Rolle direktion ohne Demo-Umschalter', () => {
        expect(useDirektionPlanerChrome({ role: 'direktion', demoRoleOverride: false })).toBe(true);
        expect(useDirektionPlanerChrome({ role: 'direktion', demoRoleOverride: true })).toBe(false);
        expect(useDirektionPlanerChrome({ role: 'kv', demoRoleOverride: false })).toBe(false);
    });

    it('navSectionsForRole für Direktion sektioniert Arbeit/Auswertung/System', () => {
        const sections = navSectionsForRole('direktion', ['direktion'], { planerAdminAccess: true });
        expect(sections.map((s) => s.id)).toEqual(['arbeit', 'auswertung', 'system']);
        const flat = sections.flatMap((s) => s.items || []).map((v) => v.id);
        expect(flat).toContain('dashboard');
        expect(flat).toContain('freigabe');
        expect(flat).toContain('bericht');
        expect(flat).toContain('administration');
    });

    it('viewsForRole für Direktion enthält Administration', () => {
        const state = { role: 'direktion', demoRoleOverride: false, planerAdminAccess: true };
        const views = viewsForRole('direktion', state);
        expect(views.some((v) => v.id === 'administration')).toBe(true);
        expect(views.some((v) => v.id === 'dashboard')).toBe(true);
    });

    it('canDeleteFreistellung nur für Direktion/Admin', () => {
        expect(canDeleteFreistellung({ role: 'direktion' })).toBe(true);
        expect(canDeleteFreistellung({ role: 'kv' })).toBe(false);
        expect(canDeleteFreistellung({ role: 'schueler' })).toBe(false);
        expect(canDeleteFreistellung(null)).toBe(false);
    });
});
