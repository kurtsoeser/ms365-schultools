import { describe, expect, it } from 'vitest';
import { filterAufgaben, canManageAufgabe } from '../src/tools/aufgaben/aufgaben-logic.js';
import { filterMeldungen } from '../src/tools/technik-meldungen/technik-meldungen-logic.js';
import { filterVorgaenge } from '../src/tools/ausgaben/ausgaben-logic.js';
import { createBetriebAdminCache } from '../src/shared/betrieb-admin-entra.js';

describe('betrieb-admin-cache', () => {
    it('setzt und liest Admin-Cache', () => {
        const c = createBetriebAdminCache();
        expect(c.isAdmin('a@x.at')).toBe(false);
        c.set({ email: 'a@x.at', isAdmin: true, sources: ['entra'], demo: false });
        expect(c.isAdmin('a@x.at')).toBe(true);
        expect(c.get().sources).toEqual(['entra']);
    });
});

describe('aufgaben scope', () => {
    const items = [
        {
            id: '1',
            title: 'A',
            status: 'offen',
            zuweisungEmail: 'me@schule.at',
            erstelltVonEmail: 'other@schule.at',
            prioritaet: 'mittel',
            beschreibung: '',
            zuweisungName: 'Me',
            rolle: '',
            kommentar: ''
        },
        {
            id: '2',
            title: 'B',
            status: 'offen',
            zuweisungEmail: 'other@schule.at',
            erstelltVonEmail: 'me@schule.at',
            prioritaet: 'mittel',
            beschreibung: '',
            zuweisungName: 'Other',
            rolle: '',
            kommentar: ''
        },
        {
            id: '3',
            title: 'C',
            status: 'offen',
            zuweisungEmail: 'x@schule.at',
            erstelltVonEmail: 'y@schule.at',
            prioritaet: 'mittel',
            beschreibung: '',
            zuweisungName: 'X',
            rolle: '',
            kommentar: ''
        }
    ];

    it('Lehrkraft sieht ohne showAll nur eigene', () => {
        const list = filterAufgaben(items, {
            scope: '',
            meEmail: 'me@schule.at',
            admin: false,
            showAll: false
        });
        expect(list.map((a) => a.id).sort()).toEqual(['1', '2']);
    });

    it('Admin+showAll sieht alle', () => {
        const list = filterAufgaben(items, {
            scope: '',
            meEmail: 'me@schule.at',
            admin: true,
            showAll: true
        });
        expect(list).toHaveLength(3);
    });

    it('canManageAufgabe für Assignee/Creator/Admin', () => {
        expect(canManageAufgabe(items[0], { email: 'me@schule.at' })).toBe(true);
        expect(canManageAufgabe(items[2], { email: 'me@schule.at' })).toBe(false);
        expect(canManageAufgabe(items[2], { email: 'me@schule.at', admin: true })).toBe(true);
    });
});

describe('technik scope', () => {
    const meldungen = [
        { id: '1', title: 'A', status: 'offen', melderEmail: 'me@schule.at', kategorie: 'beamer', prioritaet: 'mittel', raumCode: 'A', raumName: '', beschreibung: '', melderName: '', kommentar: '' },
        { id: '2', title: 'B', status: 'offen', melderEmail: 'other@schule.at', kategorie: 'beamer', prioritaet: 'mittel', raumCode: 'B', raumName: '', beschreibung: '', melderName: '', kommentar: '' }
    ];

    it('ohne Admin nur eigene Meldungen', () => {
        expect(
            filterMeldungen(meldungen, {
                meEmail: 'me@schule.at',
                admin: false,
                showAll: false
            }).map((m) => m.id)
        ).toEqual(['1']);
    });

    it('Admin Inbox sieht alle', () => {
        expect(
            filterMeldungen(meldungen, {
                meEmail: 'me@schule.at',
                admin: true,
                showAll: true
            })
        ).toHaveLength(2);
    });
});

describe('ausgaben scope', () => {
    const vorgaenge = [
        { id: '1', status: 'ausgegeben', empfaengerEmail: 'me@schule.at', empfaengerName: 'Me', gegenstandId: 'g1', kennung: 'K1', klasseCode: '' },
        { id: '2', status: 'ausgegeben', empfaengerEmail: 'other@schule.at', empfaengerName: 'O', gegenstandId: 'g2', kennung: 'K2', klasseCode: '' }
    ];

    it('Lehrkraft nur eigene Empfänger-Vorgänge', () => {
        expect(
            filterVorgaenge(vorgaenge, {
                meEmail: 'me@schule.at',
                admin: false,
                showAll: false
            }).map((v) => v.id)
        ).toEqual(['1']);
    });

    it('Admin sieht alle', () => {
        expect(
            filterVorgaenge(vorgaenge, {
                meEmail: 'me@schule.at',
                admin: true,
                showAll: true
            })
        ).toHaveLength(2);
    });
});
