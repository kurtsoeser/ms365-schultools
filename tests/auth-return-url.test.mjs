import { describe, expect, it } from 'vitest';
import {
    pickAuthReturnUrl,
    pinAuthReturnToThisPage,
    readRequestOrigins
} from '../src/shared/auth-return-url.js';

const ARGE = 'http://localhost:5173/tools/arge-fachgruppen.html';
const NEXTCLOUD = 'http://localhost:5173/tools/nextcloud-teams-migrate.html';
const SCHOOLTOOL = 'http://localhost:5173/ms365-schooltool.html';

function memoryStorage(initial) {
    const map = new Map(Object.entries(initial || {}));
    return {
        get length() {
            return map.size;
        },
        key(i) {
            return Array.from(map.keys())[i] || null;
        },
        getItem(k) {
            return map.has(k) ? map.get(k) : null;
        },
        setItem(k, v) {
            map.set(k, String(v));
        },
        removeItem(k) {
            map.delete(k);
        }
    };
}

describe('pickAuthReturnUrl', () => {
    it('nimmt die Seite, die die Anmeldung gestartet hat, nicht die ältere Session-URL', () => {
        const back = pickAuthReturnUrl({
            pageHref: SCHOOLTOOL + '#code=abc',
            storedHref: NEXTCLOUD,
            requestOrigins: [ARGE]
        });
        expect(back).toBe(ARGE);
    });

    it('fällt auf die gemerkte URL zurück, wenn keine Herkunft da ist', () => {
        expect(
            pickAuthReturnUrl({
                pageHref: SCHOOLTOOL,
                storedHref: ARGE,
                requestOrigins: []
            })
        ).toBe(ARGE);
    });

    it('verwirft fremde Origins und die Redirect-Seite selbst', () => {
        expect(
            pickAuthReturnUrl({
                pageHref: SCHOOLTOOL,
                storedHref: 'https://evil.example/tools/nextcloud-teams-migrate.html',
                requestOrigins: [SCHOOLTOOL, 'javascript:alert(1)']
            })
        ).toBe('');
    });
});

describe('pinAuthReturnToThisPage', () => {
    it('setzt die Rückkehr auf die aktuelle Seite und löst eine fremde MSAL-Herkunft', () => {
        const storage = memoryStorage({
            'ms365-post-login-url': NEXTCLOUD,
            'it-client.request.origin': NEXTCLOUD,
            'it-client.interaction.status': 'interaction_in_progress'
        });
        const win = { location: { href: ARGE }, sessionStorage: storage };
        expect(pinAuthReturnToThisPage(win)).toBe(true);
        expect(storage.getItem('ms365-post-login-url')).toBe(ARGE);
        expect(readRequestOrigins(storage)).toEqual([ARGE]);
        expect(storage.getItem('it-client.interaction.status')).toBe(null);
    });

    it('fasst die Redirect-Seite nicht an', () => {
        const storage = memoryStorage({ 'ms365-post-login-url': NEXTCLOUD });
        const win = { location: { href: SCHOOLTOOL + '#code=1' }, sessionStorage: storage };
        expect(pinAuthReturnToThisPage(win)).toBe(false);
        expect(storage.getItem('ms365-post-login-url')).toBe(NEXTCLOUD);
    });
});
