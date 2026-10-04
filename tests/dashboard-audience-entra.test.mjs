import { describe, it, expect } from 'vitest';
import { personaFromDashboardEntraGroups } from '../src/shared/dashboard-audience-entra.js';
import {
    normalizeDashboardAudienceGroups,
    dashboardAudienceGroupsConfigured
} from '../src/shared/dashboard-audience-groups-store.js';

const LEHRER = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const SCHUELER = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';

describe('dashboard-audience-entra', () => {
    it('personaFromDashboardEntraGroups priorisiert Lehrer', () => {
        const cfg = normalizeDashboardAudienceGroups({
            groupLehrerId: LEHRER,
            groupSchuelerId: SCHUELER
        });
        const both = new Set([LEHRER, SCHUELER].map((x) => x.toLowerCase()));
        expect(personaFromDashboardEntraGroups(both, cfg)).toBe('lehrer');
    });

    it('personaFromDashboardEntraGroups erkennt Schüler', () => {
        const cfg = normalizeDashboardAudienceGroups({
            groupLehrerId: LEHRER,
            groupSchuelerId: SCHUELER
        });
        expect(personaFromDashboardEntraGroups(new Set([SCHUELER.toLowerCase()]), cfg)).toBe('schueler');
    });

    it('dashboardAudienceGroupsConfigured', () => {
        expect(dashboardAudienceGroupsConfigured({})).toBe(false);
        expect(dashboardAudienceGroupsConfigured({ groupLehrerId: LEHRER })).toBe(true);
    });
});
