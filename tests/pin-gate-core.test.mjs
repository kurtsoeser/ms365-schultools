import { describe, it, expect } from 'vitest';
import {
    SESSION_KEY,
    grantAccess,
    isAccessGranted,
    isPinGateEnabled,
    isValidPin,
    isWelcomePath,
    normalizePin,
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
        expect(safeReturnPath('/tools/kursteams.html')).toBe('tools/kursteams.html');
        expect(safeReturnPath('/welcome.html')).toBe('index.html');
        expect(safeReturnPath('//evil.example/x')).toBe('index.html');
        expect(safeReturnPath(null)).toBe('index.html');
    });
});
