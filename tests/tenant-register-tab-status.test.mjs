import { describe, it, expect } from 'vitest';
import {
    personMatchTabStatus,
    classMatchTabStatus,
    catalogLinkTabStatus
} from '../src/shared/tenant-register-tab-status.js';

describe('tenant-register-tab-status', () => {
    it('personMatchTabStatus: leer = muted', () => {
        expect(personMatchTabStatus({ total: 0, matched: 0, notFound: 0, unchecked: 0 }).kind).toBe('muted');
    });
    it('personMatchTabStatus: alle gematcht = ok', () => {
        expect(personMatchTabStatus({ total: 5, matched: 5, notFound: 0, unchecked: 0 }).kind).toBe('ok');
    });
    it('classMatchTabStatus: Fehler wenn nicht gefunden', () => {
        expect(classMatchTabStatus(3, 1, 2, 0).kind).toBe('error');
    });
    it('catalogLinkTabStatus: teilweise verknüpft = warn', () => {
        expect(catalogLinkTabStatus(4, 2).kind).toBe('warn');
    });
});
