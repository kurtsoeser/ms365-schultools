import { describe, it, expect } from 'vitest';
import { normalizeMailNickname } from '../src/shared/utils/mail-nickname.js';

/**
 * Pure Kern der Umbenennen-Logik (DisplayName +1 Stufe / Nickname),
 * gespiegelt aus klassen-umbenennen.js – ohne Graph/DOM.
 */
function computeNewDisplayNamePlusOne(displayName, prefix) {
    const p = String(prefix || '').trim();
    const s = String(displayName || '').trim();
    if (!s) return null;
    const re = p
        ? new RegExp('^' + p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s+(\\d{1,2})([A-Za-zÄÖÜäöüß0-9]*)$', 'i')
        : /^(\d{1,2})([A-Za-zÄÖÜäöüß0-9]*)$/i;
    const m = s.match(re);
    if (!m) return null;
    const grade = parseInt(m[1], 10) + 1;
    const rest = m[2] || '';
    return (p ? p + ' ' : '') + String(grade) + rest;
}

describe('klassen-umbenennen Kernlogik', () => {
    it('Happy-Path: Klasse 1A → Klasse 2A', () => {
        expect(computeNewDisplayNamePlusOne('Klasse 1A', 'Klasse')).toBe('Klasse 2A');
        expect(computeNewDisplayNamePlusOne('Klasse 10HAK', 'Klasse')).toBe('Klasse 11HAK');
    });

    it('Konfliktfall: Muster passt nicht → null', () => {
        expect(computeNewDisplayNamePlusOne('Irgendwas', 'Klasse')).toBe(null);
        expect(computeNewDisplayNamePlusOne('', 'Klasse')).toBe(null);
    });

    it('Nickname mit Umlauten konsistent', () => {
        expect(normalizeMailNickname('Klasse Mädchen')).toBe('klasse-maedchen');
        expect(normalizeMailNickname('Klasse 1A')).toBe('klasse-1a');
    });
});
