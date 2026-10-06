import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import {
    loadSchoolAudienceGroups,
    normalizeSchoolAudienceGroups,
    overlaySchoolAudienceOnPermissions
} from '../src/shared/school-audience-groups.js';

const LEHRER = '11111111-1111-1111-1111-111111111111';
const SCHUELER = '22222222-2222-2222-2222-222222222222';

describe('school-audience-groups', () => {
    /** @type {Record<string, string>} */
    let store;

    beforeEach(() => {
        store = {};
        vi.stubGlobal('localStorage', {
            getItem(k) {
                return Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null;
            },
            setItem(k, v) {
                store[k] = String(v);
            },
            removeItem(k) {
                delete store[k];
            },
            clear() {
                store = {};
            }
        });
        vi.stubGlobal('ms365AppDataV2', undefined);
    });

    afterEach(() => {
        vi.unstubAllGlobals();
    });

    it('liest Lehrer/Schüler aus Stammdaten setup.matched', () => {
        vi.stubGlobal('ms365AppDataV2', {
            getSetup: () => ({
                matched: { lehrerGroupId: LEHRER, schuelerGroupId: SCHUELER },
                catalogLinks: [
                    {
                        kind: 'sammelgruppe',
                        code: 'lehrer',
                        displayName: 'Alle Lehrkräfte',
                        graphGroupId: LEHRER
                    }
                ]
            }),
            getCatalogLink: () => null
        });
        const cfg = loadSchoolAudienceGroups();
        expect(cfg.groupLehrerId).toBe(LEHRER);
        expect(cfg.groupSchuelerId).toBe(SCHUELER);
        expect(cfg.groupLehrerName).toBe('Alle Lehrkräfte');
    });

    it('fallback auf legacy dashboard-audience-groups', () => {
        localStorage.setItem(
            'ms365-dashboard-audience-groups-v1',
            JSON.stringify({ groupLehrerId: LEHRER, groupLehrerName: 'Legacy Lehrer' })
        );
        expect(loadSchoolAudienceGroups().groupLehrerId).toBe(LEHRER);
    });

    it('overlaySchoolAudienceOnPermissions überschreibt Planer-Gruppen', () => {
        vi.stubGlobal('ms365AppDataV2', {
            getSetup: () => ({ matched: { lehrerGroupId: LEHRER, schuelerGroupId: SCHUELER }, catalogLinks: [] }),
            getCatalogLink: () => null
        });
        const merged = overlaySchoolAudienceOnPermissions({
            groupLehrerId: '',
            groupSchuelerId: 'old-id'
        });
        expect(merged.groupLehrerId).toBe(LEHRER);
        expect(merged.groupSchuelerId).toBe(SCHUELER);
    });

    it('normalizeSchoolAudienceGroups filtert ungültige GUIDs', () => {
        expect(normalizeSchoolAudienceGroups({ groupLehrerId: 'keine-guid' }).groupLehrerId).toBe('');
    });
});
