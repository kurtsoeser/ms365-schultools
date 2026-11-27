import { describe, it, expect } from 'vitest';
import {
    isToolVisibleForView,
    resolveActiveView,
    listAvailableViews,
    DASHBOARD_LEHRER_SCHUELER_APPS,
    DASHBOARD_PORTAL_SECTIONS_BY_VIEW,
    DASHBOARD_TOOL_LABELS
} from '../src/shared/dashboard-audience-catalog.js';
import {
    isTeacherDashboardRole,
    resolveNonItDashboardPersona
} from '../src/shared/dashboard-audience-resolve.js';

describe('dashboard-audience-catalog', () => {
    it('IT sieht alle Werkzeuge', () => {
        expect(isToolVisibleForView('jahrgang', {}, 'it')).toBe(true);
        expect(isToolVisibleForView('schularbeiten-planer', {}, 'it')).toBe(true);
    });

    it('Lehrkraft und Schüler sehen Portal-fähige Apps inkl. Matura und Teams-Antrag', () => {
        expect(isToolVisibleForView('schularbeiten-planer', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('freistellung-planer', {}, 'schueler')).toBe(true);
        expect(isToolVisibleForView('teams-antrag', {}, 'schueler')).toBe(true);
        expect(isToolVisibleForView('teams-antrag', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('diplomarbeiten', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('diplomarbeiten', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('diplomarbeiten', {}, 'it')).toBe(true);
        expect(isToolVisibleForView('schulaktivitaeten-planer', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('elternbrief', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('einsammler', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('krankmeldung', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('kollegium-umfragen', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('lehrer-freistellung-planer', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('projektwochen', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('projektwochen', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('ausgaben', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('srdp-anmeldung', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('srdp-anmeldung', {}, 'schueler')).toBe(true);
        expect(isToolVisibleForView('kursteams', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('kursteam-einzeln', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('kursteam-einzeln', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('kursteam-einzeln', {}, 'it')).toBe(true);
        expect(isToolVisibleForView('unterrichtsteams-katalog', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('unterrichtsteams-katalog', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('unterrichtsteams-katalog', {}, 'it')).toBe(true);
        expect(isToolVisibleForView('wahlfach-zuteilung', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('schulsprecherwahl', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('schulsprecherwahl', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('schulsprecherwahl', {}, 'it')).toBe(true);
        expect(isToolVisibleForView('sharepoint-intranet-hub', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('sharepoint-intranet-hub', {}, 'schueler')).toBe(true);
        expect(isToolVisibleForView('elternbrief', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('klassenchats', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('klassenchats', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('klassenchats', {}, 'it')).toBe(true);
        expect(isToolVisibleForView('jahrgang', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('sharepoint-liste-srdp', {}, 'lehrer')).toBe(false);
    });

    it('IT-Matrix-Override steuert Portal-Sichtbarkeit innerhalb der Whitelist', () => {
        expect(
            isToolVisibleForView('kursteams', {}, 'lehrer', { audience: 'it' })
        ).toBe(false);
        expect(
            isToolVisibleForView('ausgaben', {}, 'lehrer', { audience: ['lehrer'] })
        ).toBe(true);
        expect(
            isToolVisibleForView('jahrgang', {}, 'lehrer', { audience: 'lehrer' })
        ).toBe(false);
    });

    it('Planer-Apps immer in Lehrkraft/Schüler-Ansicht (Rechte in der App)', () => {
        expect(isToolVisibleForView('schularbeiten-planer', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('freistellung-planer', {}, 'schueler')).toBe(true);
    });

    it('Portal-Sektionen enthalten Maturaanmeldung und erweiterte Alltags-Apps', () => {
        const lehrerIds = DASHBOARD_PORTAL_SECTIONS_BY_VIEW.lehrer.flatMap((s) => s.toolIds);
        const schuelerIds = DASHBOARD_PORTAL_SECTIONS_BY_VIEW.schueler.flatMap((s) => s.toolIds);
        expect(lehrerIds).toContain('srdp-anmeldung');
        expect(lehrerIds).toContain('kursteams');
        expect(lehrerIds).toContain('wahlfach-zuteilung');
        const lehrerBetrieb = DASHBOARD_PORTAL_SECTIONS_BY_VIEW.lehrer.find((s) => s.id === 'betrieb');
        const lehrerKollegium = DASHBOARD_PORTAL_SECTIONS_BY_VIEW.lehrer.find((s) => s.id === 'kollegium');
        const lehrerUnterricht = DASHBOARD_PORTAL_SECTIONS_BY_VIEW.lehrer.find((s) => s.id === 'unterricht');
        const lehrerAntraege = DASHBOARD_PORTAL_SECTIONS_BY_VIEW.lehrer.find((s) => s.id === 'antraege');
        expect(lehrerBetrieb.toolIds).toContain('teams-antrag');
        expect(lehrerKollegium.toolIds).not.toContain('teams-antrag');
        expect(lehrerAntraege.label).toBe('Anträge und Formulare');
        expect(lehrerAntraege.toolIds).toEqual([
            'lehrer-freistellung-planer',
            'freistellung-planer',
            'srdp-anmeldung',
            'wahlfach-zuteilung'
        ]);
        expect(lehrerUnterricht.toolIds).not.toContain('wahlfach-zuteilung');
        expect(lehrerUnterricht.toolIds).not.toContain('lehrer-freistellung-planer');
        expect(lehrerUnterricht.toolIds).not.toContain('freistellung-planer');
        expect(lehrerUnterricht.toolIds).not.toContain('srdp-anmeldung');
        expect(lehrerUnterricht.toolIds).not.toContain('kursteam-einzeln');
        expect(lehrerUnterricht.toolIds).not.toContain('unterrichtsteams-katalog');
        expect(lehrerUnterricht.toolIds).not.toContain('diplomarbeiten');
        expect(lehrerIds).not.toContain('diplomarbeiten');
        expect(DASHBOARD_LEHRER_SCHUELER_APPS).not.toContain('kursteam-einzeln');
        expect(DASHBOARD_LEHRER_SCHUELER_APPS).not.toContain('unterrichtsteams-katalog');
        expect(DASHBOARD_LEHRER_SCHUELER_APPS).not.toContain('diplomarbeiten');
        expect(schuelerIds).toContain('srdp-anmeldung');
        expect(schuelerIds).toContain('wahlfach-anmeldung');
        expect(schuelerIds).not.toContain('wahlfach-zuteilung');
        expect(DASHBOARD_LEHRER_SCHUELER_APPS).toContain('srdp-anmeldung');
        expect(DASHBOARD_LEHRER_SCHUELER_APPS).not.toContain('klassenchats');
        expect(lehrerIds).not.toContain('klassenchats');
        expect(schuelerIds).not.toContain('klassenchats');
        expect(DASHBOARD_LEHRER_SCHUELER_APPS).not.toContain('schulsprecherwahl');
        expect(lehrerIds).not.toContain('schulsprecherwahl');
        expect(schuelerIds).not.toContain('schulsprecherwahl');
        expect(DASHBOARD_LEHRER_SCHUELER_APPS).not.toContain('elternsprechtag-bookings');
        expect(lehrerIds).not.toContain('elternsprechtag-bookings');
        expect(schuelerIds).not.toContain('elternsprechtag-bookings');
        expect(lehrerIds).toContain('elternsprechtag');
        expect(schuelerIds).not.toContain('elternsprechtag');
        const lehrerEltern = DASHBOARD_PORTAL_SECTIONS_BY_VIEW.lehrer.find((s) => s.id === 'eltern');
        expect(lehrerEltern.toolIds).toContain('elternsprechtag');
        expect(isToolVisibleForView('elternsprechtag-bookings', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('elternsprechtag-bookings', {}, 'it')).toBe(true);
        expect(DASHBOARD_TOOL_LABELS['srdp-anmeldung']).toMatch(/Maturaanmeldung/i);
    });

    it('resolveActiveView: IT-Default vor Lehrkraft', () => {
        expect(resolveActiveView(['it', 'lehrer', 'schueler'], null, { isIt: true })).toBe('it');
        expect(resolveActiveView(['it', 'lehrer'], 'lehrer', { isIt: true })).toBe('lehrer');
        expect(resolveActiveView(['lehrer', 'schueler'], null, { isIt: false })).toBe('lehrer');
        expect(resolveActiveView(['schueler'], null, { isIt: false })).toBe('schueler');
    });

    it('listAvailableViews: eine Hauptansicht pro Stufe', () => {
        expect(listAvailableViews({ isIt: true, isLehrer: true, isSchueler: true })).toEqual(['it']);
        expect(listAvailableViews({ isIt: false, isLehrer: true, isSchueler: false })).toEqual(['lehrer']);
        expect(listAvailableViews({ isIt: false, isLehrer: false, isSchueler: true })).toEqual(['schueler']);
    });
});

describe('dashboard-audience-resolve persona', () => {
    it('erkennt Lehrkraft über Stammdaten/Planer', () => {
        expect(isTeacherDashboardRole(['lehrer'], [], null)).toBe(true);
        expect(isTeacherDashboardRole([], ['kv'], null)).toBe(true);
        expect(isTeacherDashboardRole([], [], { teacherMatch: { email: 'a@b.c' } })).toBe(true);
        expect(isTeacherDashboardRole([], [], { kvMatch: true, kvKlassen: ['4A'] })).toBe(true);
        expect(isTeacherDashboardRole([], [], { schulleitungMatch: true })).toBe(true);
        expect(isTeacherDashboardRole(['schueler'], [], { studentMatch: {} })).toBe(false);
    });

    it('Stammdaten-Lehrkraft gewinnt gegen fehlende Entra-Gruppenmitgliedschaft', () => {
        const d = resolveNonItDashboardPersona({
            entraDashPersona: null,
            groupsConfigured: true,
            isTeacher: true
        });
        expect(d.isLehrer).toBe(true);
        expect(d.isSchueler).toBe(false);
        expect(d.personaSource).toBe('stammdaten-lehrer-entra-miss');
    });

    it('Stammdaten-Lehrkraft gewinnt gegen reine Schüler-Entra-Gruppe', () => {
        const d = resolveNonItDashboardPersona({
            entraDashPersona: 'schueler',
            groupsConfigured: true,
            isTeacher: true
        });
        expect(d.isLehrer).toBe(true);
        expect(d.personaSource).toBe('stammdaten-lehrer-over-entra-schueler');
    });

    it('ohne Lehrkraft-Signal und ohne Schüler-Match: unbekannt (nicht Schüler/in)', () => {
        const d = resolveNonItDashboardPersona({
            entraDashPersona: null,
            groupsConfigured: true,
            isTeacher: false
        });
        expect(d.isSchueler).toBe(false);
        expect(d.isLehrer).toBe(false);
        expect(d.unknownUser).toBe(true);
        expect(d.personaSource).toBe('dashboard-entra-none');
    });

    it('Stammdaten-Schüler bleibt Schüler-Portal', () => {
        const d = resolveNonItDashboardPersona({
            entraDashPersona: null,
            groupsConfigured: false,
            isTeacher: false,
            studentMatch: true
        });
        expect(d.isSchueler).toBe(true);
        expect(d.unknownUser).toBeFalsy();
        expect(d.personaSource).toBe('default-schueler');
    });
});
