import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
    saveLiveSchoolStammdatenCache,
    loadLiveSchoolStammdatenCache,
    mergeSchoolStammdatenLocalAndLive,
    loadEffectiveTeachersForPersona,
    mergeLiveTeachersIntoLocalTenantIfEmpty,
    isSparseOrDemoOnlyTeacherList,
    preferRicherTeacherList
} from '../src/shared/live-school-stammdaten-cache.js';

describe('live-school-stammdaten-cache', () => {
    const store = new Map();

    beforeEach(() => {
        store.clear();
        vi.stubGlobal('sessionStorage', {
            getItem: (k) => (store.has(k) ? store.get(k) : null),
            setItem: (k, v) => store.set(k, String(v)),
            removeItem: (k) => store.delete(k)
        });
        vi.stubGlobal('window', {
            ms365TenantSettingsLoad: () => ({ teachers: [], classes: [], domain: 'kurtrocks.com' }),
            ms365TenantSettingsSave: vi.fn()
        });
    });

    afterEach(() => {
        vi.unstubAllGlobals();
    });

    it('speichert und lädt Live-Lehrer', () => {
        saveLiveSchoolStammdatenCache({
            teachers: [{ code: 'GRO', email: 'dave.grohl@kurtrocks.com', name: 'Dave Grohl' }],
            classes: [{ code: '2AK' }],
            siteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools'
        });
        const live = loadLiveSchoolStammdatenCache();
        expect(live.teachers).toHaveLength(1);
        expect(live.teachers[0].email).toBe('dave.grohl@kurtrocks.com');
        expect(live.classes[0].code).toBe('2AK');
    });

    it('merge bevorzugt Live', () => {
        const m = mergeSchoolStammdatenLocalAndLive(
            { teachers: [{ code: 'OLD' }], classes: [] },
            { teachers: [{ code: 'NEW' }], classes: [{ code: '1A' }] }
        );
        expect(m.teachers[0].code).toBe('NEW');
        expect(m.classes[0].code).toBe('1A');
    });

    it('loadEffectiveTeachersForPersona nutzt Live-Cache wenn Tenant leer', () => {
        saveLiveSchoolStammdatenCache({
            teachers: [{ code: 'GRO', email: 'dave.grohl@kurtrocks.com' }]
        });
        const teachers = loadEffectiveTeachersForPersona();
        expect(teachers).toHaveLength(1);
        expect(teachers[0].code).toBe('GRO');
    });

    it('mergeLiveTeachersIntoLocalTenantIfEmpty schreibt bei leerer Liste', () => {
        const save = vi.fn();
        window.ms365TenantSettingsSave = save;
        const r = mergeLiveTeachersIntoLocalTenantIfEmpty(
            [{ code: 'GRO', email: 'dave.grohl@kurtrocks.com' }],
            [{ code: '2AK' }]
        );
        expect(r.saved).toBe(true);
        expect(save).toHaveBeenCalled();
        const arg = save.mock.calls[0][0];
        expect(arg.teachers[0].code).toBe('GRO');
        expect(arg.classes[0].code).toBe('2AK');
    });

    it('Demo-Einzelkonto blockiert SharePoint-Volliste nicht', () => {
        expect(
            isSparseOrDemoOnlyTeacherList([
                { code: 'L01', name: 'Lehrer01 DEMO', email: 'lehrer01.demo@ms365.schule' }
            ])
        ).toBe(true);
        const richer = preferRicherTeacherList(
            [{ code: 'L01', name: 'Lehrer01 DEMO', email: 'lehrer01.demo@ms365.schule' }],
            [
                { code: 'GRO', email: 'dave.grohl@kurtrocks.com' },
                { code: 'MAY', email: 'anna@kurtrocks.com' },
                { code: 'BAU', email: 'julia@kurtrocks.com' }
            ]
        );
        expect(richer).toHaveLength(3);
        expect(richer[0].code).toBe('GRO');
    });

    it('merge ersetzt Demo-Mini-Stand durch Live-Kollegium', () => {
        const save = vi.fn();
        window.ms365TenantSettingsLoad = () => ({
            teachers: [{ code: 'L01', name: 'Lehrer01 DEMO', email: 'lehrer01.demo@ms365.schule' }],
            classes: [],
            domain: 'kurtrocks.com'
        });
        window.ms365TenantSettingsSave = save;
        const live = Array.from({ length: 10 }, (_, i) => ({
            code: 'T' + i,
            email: 't' + i + '@kurtrocks.com',
            name: 'Lehrer ' + i
        }));
        const r = mergeLiveTeachersIntoLocalTenantIfEmpty(live, [{ code: '1AK' }]);
        expect(r.saved).toBe(true);
        expect(save.mock.calls[0][0].teachers).toHaveLength(10);
    });
});
