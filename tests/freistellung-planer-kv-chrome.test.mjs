import { describe, it, expect } from 'vitest';
import {
    viewsForRole,
    useKvPlanerChrome,
    useDirektionPlanerChrome,
    useStudentPlanerChrome
} from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('freistellung KV chrome', () => {
    it('useKvPlanerChrome nur für Rolle kv ohne Demo-Umschalter', () => {
        expect(useKvPlanerChrome({ role: 'kv', demoRoleOverride: false })).toBe(true);
        expect(useKvPlanerChrome({ role: 'kv', demoRoleOverride: true })).toBe(false);
        expect(useKvPlanerChrome({ role: 'direktion', demoRoleOverride: false })).toBe(false);
        expect(useStudentPlanerChrome({ role: 'kv', demoRoleOverride: false })).toBe(false);
    });

    it('viewsForRole für KV-Chrome nur Übersicht', () => {
        const state = { role: 'kv', demoRoleOverride: false, entraGroupsConfigured: true };
        const views = viewsForRole('kv', state);
        expect(views.map((v) => v.id)).toEqual(['dashboard']);
    });

    it('viewsForRole für KV im Demo-Modus behält Staff-Navigation', () => {
        const state = { role: 'kv', demoRoleOverride: true };
        const views = viewsForRole('kv', state);
        expect(views.some((v) => v.id === 'liste')).toBe(true);
        expect(views.some((v) => v.id === 'bericht')).toBe(true);
    });

    it('viewsForRole für Schüler: Meine Anträge und Antrag stellen', () => {
        const views = viewsForRole('schueler', { role: 'schueler' });
        expect(views.map((v) => v.id)).toEqual(['meine', 'antrag']);
    });
});

describe('freistellung Direktion chrome', () => {
    it('useDirektionPlanerChrome nur für Rolle direktion ohne Demo-Umschalter', () => {
        expect(useDirektionPlanerChrome({ role: 'direktion', demoRoleOverride: false })).toBe(true);
        expect(useDirektionPlanerChrome({ role: 'direktion', demoRoleOverride: true })).toBe(false);
        expect(useDirektionPlanerChrome({ role: 'kv', demoRoleOverride: false })).toBe(false);
    });

    it('viewsForRole für Direktion enthält Administration', () => {
        const state = { role: 'direktion', demoRoleOverride: false };
        const views = viewsForRole('direktion', state);
        expect(views.some((v) => v.id === 'administration')).toBe(true);
        expect(views.some((v) => v.id === 'dashboard')).toBe(true);
    });
});
