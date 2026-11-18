import { describe, it, expect } from 'vitest';
import {
    mapSchularbeitFromItem,
    mapSchularbeitToFields,
    isSaAccessDeniedError
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-graph.js';
import {
    canEditSchularbeit,
    canAdminDecide,
    canDeleteSchularbeit
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-state.js';

describe('schularbeiten-planer graph mapping', () => {
    it('mapSchularbeitFromItem liest itemId und Kernfelder', () => {
        const sa = mapSchularbeitFromItem({
            id: 42,
            fields: {
                SchularbeitId: 'sa-abc',
                Titel: 'SA Mathematik',
                Thema: 'Gleichungen',
                Title: 'Schularbeit - M - 3AK',
                FachCode: 'M',
                KlasseCode: '3AK',
                LehrerCode: 'BAU',
                LehrerEmail: 'Bau@Schule.at',
                Datum: '2026-11-20',
                BeginnUhrzeit: '08:00',
                DauerMinuten: 100,
                Semester: 'WS',
                Status: 'fixiert',
                BeantragtVon: 'bau@schule.at',
                Schuljahr: '2026/27'
            }
        });
        expect(sa.itemId).toBe('42');
        expect(sa.schularbeitId).toBe('sa-abc');
        expect(sa.titel).toBe('SA Mathematik');
        expect(sa.thema).toBe('Gleichungen');
        expect(sa.fachCode).toBe('M');
        expect(sa.klasseCode).toBe('3AK');
        expect(sa.lehrerEmail).toBe('bau@schule.at');
        expect(sa.status).toBe('fixiert');
        expect(sa.schuljahr).toBe('2026/27');
    });

    it('mapSchularbeitToFields schreibt Graph-Felder inkl. Ownership', () => {
        const fields = mapSchularbeitToFields(
            {
                schularbeitId: 'sa-xyz',
                titel: 'Probe',
                thema: 'Thema A',
                fachCode: 'D',
                klasseCode: '1AK',
                lehrerCode: 'MUE',
                lehrerEmail: 'mue@schule.at',
                datum: '2026-12-01',
                beginnUhrzeit: '09:50',
                dauerMinuten: 50,
                semester: 'WS',
                status: 'beantragt',
                beantragtVon: 'mue@schule.at',
                schuljahr: '2026/27'
            },
            { labels: { fach: { D: 'Deutsch' }, klasse: { '1AK': '1AK' } } }
        );
        expect(fields.SchularbeitId).toBe('sa-xyz');
        expect(fields.Titel).toBe('Probe');
        expect(fields.FachCode).toBe('D');
        expect(fields.KlasseCode).toBe('1AK');
        expect(fields.LehrerEmail).toBe('mue@schule.at');
        expect(fields.BeantragtVon).toBe('mue@schule.at');
        expect(fields.Status).toBe('beantragt');
        expect(fields.Schuljahr).toBe('2026/27');
        expect(fields.Title).toMatch(/^Schularbeit -/);
        expect(fields.FixiertAm).toBeUndefined();
    });

    it('isSaAccessDeniedError erkennt 403/Forbidden', () => {
        expect(isSaAccessDeniedError(new Error('403 Forbidden'))).toBe(true);
        expect(isSaAccessDeniedError({ message: 'Access denied' })).toBe(true);
        expect(isSaAccessDeniedError(new Error('Liste fehlt'))).toBe(false);
    });
});

describe('schularbeiten-planer Schüler-Read', () => {
    const sa = {
        itemId: '1',
        status: 'beantragt',
        lehrerCode: 'BAU',
        lehrerEmail: 'bau@schule.at'
    };
    const schuelerScope = {
        role: 'schueler',
        teacherMatch: null,
        accountEmail: 'max@schule.at',
        studentMatch: { klasse: '3AK' }
    };

    it('Schüler darf nicht editieren/löschen/entscheiden', () => {
        expect(canEditSchularbeit(sa, schuelerScope)).toBe(false);
        expect(canDeleteSchularbeit(sa, schuelerScope)).toBe(false);
        expect(canAdminDecide(sa, schuelerScope)).toBe(false);
    });

    it('Admin darf beantragte SA entscheiden', () => {
        expect(
            canAdminDecide(sa, {
                role: 'admin',
                accountEmail: 'admin@schule.at'
            })
        ).toBe(true);
    });
});
