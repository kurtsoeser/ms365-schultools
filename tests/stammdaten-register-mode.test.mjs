import { describe, it, expect } from 'vitest';
import { parseRegisterHash, buildRegisterHash, isRegisterModeHash } from '../src/shared/stammdaten-register-mode.js';

describe('stammdaten-register-mode', () => {
    it('mappt Legacy-Hashes auf Tabs', () => {
        expect(parseRegisterHash('#sync').tabBtnId).toBe('tabMainStammdaten');
        expect(parseRegisterHash('#import').tabBtnId).toBe('tabMainSchueler');
        expect(parseRegisterHash('#intranet').tabBtnId).toBe('tabMainStammdaten');
        expect(parseRegisterHash('#pflegen').tabBtnId).toBe('');
    });

    it('Tab-Hashes', () => {
        const p = parseRegisterHash('#klassen');
        expect(p.tabBtnId).toBe('tabMainKlassen');
        expect(parseRegisterHash('#werkzeuge').tabBtnId).toBe('tabMainStammdaten');
        expect(parseRegisterHash('#datenlandkarte').tabBtnId).toBe('tabMainDatenlandkarte');
        expect(buildRegisterHash('tabMainDatenlandkarte')).toBe('datenlandkarte');
    });

    it('buildRegisterHash nur Tab', () => {
        expect(buildRegisterHash('tabMainSchueler')).toBe('schueler');
        expect(buildRegisterHash('')).toBe('');
    });

    it('isRegisterModeHash', () => {
        expect(isRegisterModeHash('sync')).toBe(true);
        expect(isRegisterModeHash('klassen')).toBe(false);
    });
});
