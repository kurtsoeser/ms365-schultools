import { describe, expect, it } from 'vitest';
import { isDashboardBackLink } from '../src/shared/app-tool-chrome.js';

describe('app-tool-chrome', () => {
    it('erkennt Dashboard-Zurück-Links', () => {
        const a = { getAttribute: (k) => (k === 'href' ? '../index.html' : ''), textContent: ' Dashboard ' };
        expect(isDashboardBackLink(a)).toBe(true);
    });

    it('lehnt andere index-Links ab', () => {
        const a = { getAttribute: (k) => (k === 'href' ? '../index.html#werkzeuge' : ''), textContent: 'Werkzeuge' };
        expect(isDashboardBackLink(a)).toBe(false);
    });
});
