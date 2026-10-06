import { describe, it, expect } from 'vitest';
import {
    DEFAULT_INTRANET_LIST_TITLES,
    resolveIntranetListTitle
} from '../src/shared/intranet-list-title-logic.js';

describe('intranet list titles', () => {
    it('nutzt Standard wenn nichts gespeichert', () => {
        expect(resolveIntranetListTitle('schueler', '')).toBe(DEFAULT_INTRANET_LIST_TITLES.schueler);
        expect(resolveIntranetListTitle('lehrer', '   ')).toBe(DEFAULT_INTRANET_LIST_TITLES.lehrer);
    });

    it('behält benutzerdefinierten Titel', () => {
        expect(resolveIntranetListTitle('klassen', 'Stufen')).toBe('Stufen');
    });
});
