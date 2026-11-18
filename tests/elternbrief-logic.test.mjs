import { describe, expect, it } from 'vitest';
import {
    briefLinkAktiv,
    parseEmpfaengerText,
    parseKlasseCodes,
    progressStats,
    resolveLeseStatus,
    validateBrief
} from '../src/tools/elternbrief/elternbrief-logic.js';
import { LESE_STATUS, STATUS } from '../src/tools/elternbrief/elternbrief-model.js';
import {
    BRIEFE_COLUMNS,
    LESE_COLUMNS,
    LIST_TITLES
} from '../src/tools/elternbrief/elternbrief-schema.js';
import { VORLAGEN, draftFromVorlage, getVorlage } from '../src/tools/elternbrief/elternbrief-templates.js';
import {
    buildPublicPayload,
    decodePublicToken,
    encodePublicToken,
    validatePublicPayload
} from '../src/tools/elternbrief/elternbrief-token.js';

describe('elternbrief-logic', () => {
    it('validiert Betreff und Inhalt', () => {
        expect(validateBrief({ betreff: '', inhalt: 'kurz' }).ok).toBe(false);
        expect(validateBrief({ betreff: 'Info', inhalt: 'Liebe Eltern, …' }).ok).toBe(true);
    });

    it('parst Klassen und Empfänger', () => {
        expect(parseKlasseCodes('3A; 2B,1C')).toEqual(['3A', '2B', '1C']);
        const rows = parseEmpfaengerText('Anna;3A;a@x.at\nBen Hofer;3A');
        expect(rows).toHaveLength(2);
        expect(rows[0].elternEmail).toBe('a@x.at');
        expect(rows[1].schuelerName).toBe('Ben Hofer');
    });

    it('berechnet Fortschritt und Ablauf', () => {
        const brief = { id: 'eb-1', faelligAm: '2020-01-01' };
        const emp = [
            { briefId: 'eb-1', status: LESE_STATUS.OFFEN },
            { briefId: 'eb-1', status: LESE_STATUS.GELESEN },
            { briefId: 'eb-2', status: LESE_STATUS.OFFEN }
        ];
        const st = progressStats(brief, emp);
        expect(st.total).toBe(2);
        expect(st.gelesen).toBe(1);
        expect(st.abgelaufen).toBe(1);
        expect(resolveLeseStatus(brief, { status: LESE_STATUS.OFFEN }, '2026-01-01')).toBe(
            LESE_STATUS.ABGELAUFEN
        );
    });

    it('erkennt aktive Links', () => {
        expect(briefLinkAktiv({ status: STATUS.GESENDET })).toBe(true);
        expect(briefLinkAktiv({ status: STATUS.ENTWURF })).toBe(false);
    });
});

describe('elternbrief-templates', () => {
    it('liefert Vorlagen und Drafts', () => {
        expect(VORLAGEN.length).toBeGreaterThanOrEqual(5);
        expect(getVorlage('elternabend')).toBeTruthy();
        const d = draftFromVorlage('schularbeitenplan');
        expect(d.betreff).toMatch(/Schularbeiten/i);
        expect(d.vorlageId).toBe('schularbeitenplan');
    });
});

describe('elternbrief-token', () => {
    it('kodiert und dekodiert Public-Payload', () => {
        const payload = buildPublicPayload({
            brief: {
                id: 'eb-1',
                betreff: 'Test',
                inhalt: 'Hallo',
                status: 'gesendet',
                lesebestaetigung: true
            },
            empfaenger: { id: 'el-1', schuelerName: 'Anna', klasseCode: '3A' },
            schoolName: 'Demo'
        });
        expect(validatePublicPayload(payload).ok).toBe(true);
        const token = encodePublicToken(payload);
        const decoded = decodePublicToken(token);
        expect(decoded.ok).toBe(true);
        expect(decoded.payload.bid).toBe('eb-1');
        expect(decoded.payload.eid).toBe('el-1');
    });
});

describe('elternbrief-schema', () => {
    it('hat Listen und Spalten', () => {
        expect(LIST_TITLES.briefe).toBe('Elternbriefe');
        expect(LIST_TITLES.lesebestaetigungen).toBe('ElternbriefLesebestaetigungen');
        expect(BRIEFE_COLUMNS.some((c) => c.name === 'BriefId')).toBe(true);
        expect(BRIEFE_COLUMNS.some((c) => c.name === 'AnhangName')).toBe(true);
        expect(LESE_COLUMNS.some((c) => c.name === 'LeseId')).toBe(true);
    });
});
