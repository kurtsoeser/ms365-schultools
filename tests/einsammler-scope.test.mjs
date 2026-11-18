import { describe, expect, it } from 'vitest';
import {
    filterAntwortenForHeads,
    filterEinsammlerForViewer,
    isOwnedBy,
    parseAdminMails
} from '../src/tools/einsammler/einsammler-scope.js';

describe('einsammler-scope', () => {
    const list = [
        { id: '1', titel: 'A', erstellerEmail: 'a@schule.at' },
        { id: '2', titel: 'B', erstellerEmail: 'b@schule.at' },
        { id: '3', titel: 'C', erstellerEmail: '' }
    ];

    it('parst Admin-Mails', () => {
        expect(parseAdminMails('x@y.at; z@y.at')).toEqual(['x@y.at', 'z@y.at']);
    });

    it('erkennt Besitz', () => {
        expect(isOwnedBy(list[0], 'a@schule.at')).toBe(true);
        expect(isOwnedBy(list[0], 'b@schule.at')).toBe(false);
        expect(isOwnedBy(list[2], 'a@schule.at')).toBe(false);
    });

    it('filtert für Lehrkraft nur eigene', () => {
        const mine = filterEinsammlerForViewer(list, {
            email: 'a@schule.at',
            admin: false,
            showAll: false
        });
        expect(mine).toHaveLength(1);
        expect(mine[0].id).toBe('1');
    });

    it('Admin mit showAll sieht alle', () => {
        const all = filterEinsammlerForViewer(list, {
            email: 'admin@schule.at',
            admin: true,
            showAll: true
        });
        expect(all).toHaveLength(3);
    });

    it('filtert Antworten zu sichtbaren Köpfen', () => {
        const ants = [
            { id: 'x', einsammlerId: '1' },
            { id: 'y', einsammlerId: '2' }
        ];
        expect(filterAntwortenForHeads(ants, [{ id: '1' }])).toHaveLength(1);
    });
});
