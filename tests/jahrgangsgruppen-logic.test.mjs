import { describe, expect, it } from 'vitest';
import {
    sanitizeNick,
    rowKey,
    domainFromMail,
    isDirektionRole,
    classCodeExists,
    remapStudentKlassen,
    deriveNickFallback
} from '../src/tools/jahrgangsgruppen/jahrgangsgruppen-logic.js';
import { psEscapeSingle, buildClassSmtpPs1 } from '../src/tools/jahrgangsgruppen/jahrgangsgruppen-smtp.js';

describe('jahrgangsgruppen-logic', () => {
    it('sanitizeNick / rowKey / domainFromMail', () => {
        expect(sanitizeNick('Ab-C!!')).toBe('ab-c');
        expect(rowKey({ code: '1a' })).toBe('1A');
        expect(domainFromMail('a@schule.at')).toBe('schule.at');
    });

    it('isDirektionRole / classCodeExists / remapStudentKlassen', () => {
        expect(isDirektionRole('Direktorin')).toBe(true);
        expect(classCodeExists([{ code: '1A' }, { code: '1B' }], '1a', null)).toBe(true);
        expect(classCodeExists([{ code: '1A' }], '1A', '1A')).toBe(false);
        const next = remapStudentKlassen([{ klasse: '1A', email: 'x@y.z' }], '1A', '2A');
        expect(next[0].klasse).toBe('2A');
    });

    it('deriveNickFallback folgt dem jg+Jahr+Code-Schema', () => {
        expect(deriveNickFallback({ year: '2026', code: '1A' })).toBe('jg20261a');
    });
});

describe('jahrgangsgruppen-smtp', () => {
    it('baut Set-UnifiedGroup PS1', () => {
        expect(psEscapeSingle("O'Brien")).toBe("O''Brien");
        const ps1 = buildClassSmtpPs1([{ id: 'g1', name: '1A', smtp: '1a@schule.at' }], 'schule.at');
        expect(ps1).toContain('Set-UnifiedGroup');
        expect(ps1).toContain('1a@schule.at');
    });
});
