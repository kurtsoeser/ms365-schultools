import { describe, it, expect } from 'vitest';
import {
    SESSION_KEY,
    grantAccess,
    isAccessGranted,
    isPinGateEnabled,
    isValidPin,
    isWelcomePath,
    normalizePin,
    resolveReturnUrl,
    revokeAccess,
    safeReturnPath
} from '../src/shared/pin-gate-core.js';

describe('pin-gate-core', () => {
    it('normalizePin: trim', () => {
        expect(normalizePin('  1234  ')).toBe('1234');
        expect(normalizePin(null)).toBe('');
    });

    it('isValidPin: mehrere PINs, case-insensitive', () => {
        expect(isValidPin('ms365-schule', ['MS365-Schule', 'IT-Team'])).toBe(true);
        expect(isValidPin('wrong', ['MS365-Schule'])).toBe(false);
        expect(isValidPin('', ['x'])).toBe(false);
    });

    it('isPinGateEnabled', () => {
        expect(isPinGateEnabled({ enabled: false, pins: ['a'] })).toBe(false);
        expect(isPinGateEnabled({ enabled: true, pins: [] })).toBe(false);
        expect(isPinGateEnabled({ pins: ['a'] })).toBe(true);
    });

    it('Session-Flag', () => {
        const storage = {
            data: {},
            getItem(k) {
                return this.data[k] ?? null;
            },
            setItem(k, v) {
                this.data[k] = v;
            },
            removeItem(k) {
                delete this.data[k];
            }
        };
        expect(isAccessGranted(storage)).toBe(false);
        grantAccess(storage);
        expect(isAccessGranted(storage)).toBe(true);
        expect(storage.getItem(SESSION_KEY)).toBe('1');
        revokeAccess(storage);
        expect(isAccessGranted(storage)).toBe(false);
    });

    it('isWelcomePath', () => {
        expect(isWelcomePath('/repo/welcome.html')).toBe(true);
        expect(isWelcomePath('/index.html')).toBe(false);
    });

    it('safeReturnPath: blockiert welcome und fremde URLs', () => {
        const welcome = 'https://example.com/ms365-schultools/welcome.html';
        expect(safeReturnPath('/ms365-schultools/tools/kursteams.html', 'index.html', welcome)).toBe(
            'tools/kursteams.html'
        );
        expect(safeReturnPath('/welcome.html', 'index.html', welcome)).toBe('index.html');
        expect(safeReturnPath('//evil.example/x', 'index.html', welcome)).toBe('index.html');
        expect(safeReturnPath(null, 'index.html', welcome)).toBe('index.html');
    });

    it('resolveReturnUrl: behält den App-Unterordner (kein doppelter Pfad)', () => {
        const welcome = 'https://example.com/ms365-schultools/welcome.html';
        expect(resolveReturnUrl('/ms365-schultools/', welcome)).toBe(
            'https://example.com/ms365-schultools/'
        );
        expect(resolveReturnUrl('/ms365-schultools/index.html', welcome)).toBe(
            'https://example.com/ms365-schultools/index.html'
        );
        expect(resolveReturnUrl('/ms365-schultools/tools/kursteams.html', welcome)).toBe(
            'https://example.com/ms365-schultools/tools/kursteams.html'
        );
        expect(resolveReturnUrl('tools/kursteams.html', welcome)).toBe(
            'https://example.com/ms365-schultools/tools/kursteams.html'
        );
        expect(resolveReturnUrl(null, welcome)).toBe(
            'https://example.com/ms365-schultools/index.html'
        );
        expect(resolveReturnUrl('https://evil.example/x', welcome)).toBe(
            'https://example.com/ms365-schultools/index.html'
        );
    });
});
