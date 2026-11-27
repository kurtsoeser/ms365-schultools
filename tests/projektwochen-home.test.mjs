import { describe, expect, it } from 'vitest';
import { renderApp } from '../src/tools/projektwochen/projektwochen-ui.js';

function paint(state) {
    const root = { innerHTML: '' };
    renderApp(state, root);
    return root.innerHTML;
}

function base(over) {
    return Object.assign(
        {
            loading: false,
            role: 'admin',
            view: 'dashboard',
            bootstrapped: true,
            localDemoOnly: true,
            items: [],
            aktion: {
                aktionId: 'pw',
                title: 'Projektwoche',
                status: 'offen',
                startdatum: '2026-06-15',
                enddatum: '2026-06-19'
            },
            aktionen: [],
            filters: {},
            stammdaten: { teachers: [], classes: [] },
            accountEmail: 'kurt@schule.at',
            attendeeRows: [],
            occupancy: {},
            entraGroupsConfigured: false,
            demoRoleOverride: false,
            planerRoles: [],
            calYear: 2026,
            calMonth: 5
        },
        over
    );
}

const week = [
    {
        angebotId: 'a1',
        itemId: 'a1',
        status: 'beantragt',
        title: 'Zoo',
        tag: 'Mo',
        slot: 'vormittag',
        datum: '2026-06-15',
        kapazitaet: 20,
        lehrerEmail: 'anna@schule.at',
        aktionId: 'pw'
    },
    {
        angebotId: 'a2',
        itemId: 'a2',
        status: 'freigegeben',
        title: 'Kino',
        tag: 'Di',
        slot: 'nachmittag',
        datum: '2026-06-16',
        kapazitaet: 20,
        lehrerEmail: 'anna@schule.at',
        bookingsBookingUrl: 'https://book.example/kino',
        aktionId: 'pw'
    }
];

describe('Projektwochen Zuhause', () => {
    it('Admin sieht die wartenden Anträge und einen Primary', () => {
        const html = paint(base({ items: week }));
        expect(html).toContain('>1</span> Antrag wartet');
        expect(html).toContain('1 Angebot ist frei.');
        expect(html).toContain('data-pw-view="admin">Zur Freigabe');
        expect(html).not.toContain('fb-stat');
        expect(html).not.toContain('pw-filters');
        expect(html).not.toContain('Angebot/Angebote');
    });

    it('Lehrkraft ohne eigene Angebote bekommt Neues Angebot', () => {
        const html = paint(
            base({
                role: 'lehrer',
                items: week
            })
        );
        expect(html).toContain('Noch kein eigenes Angebot');
        expect(html).toContain('data-pw-view="neu">Neues Angebot');
        expect(html).not.toContain('Anträge warten');
        expect(html).not.toContain('fb-btn--ghost" data-pw-view="neu"');
    });

    it('Lehrkraft mit eigenen Angeboten sieht deren Status', () => {
        const html = paint(
            base({
                role: 'lehrer',
                accountEmail: 'anna@schule.at',
                items: week
            })
        );
        expect(html).toContain('2');
        expect(html).toContain('eigene Angebote');
        expect(html).toContain('1 wartet, 1 frei.');
        expect(html).toContain('fb-btn--ghost" data-pw-view="neu">Neues Angebot');
        expect(html).not.toContain('class="fb-btn" data-pw-view="neu"');
    });

    it('Schüler buchen vom Plan, wenn eine Service-Adresse da ist', () => {
        const html = paint(
            base({
                role: 'schueler',
                view: 'plan',
                items: week,
                attendeeRows: [{ angebotId: 'a2', email: 'kurt@schule.at' }]
            })
        );
        expect(html).not.toContain('In Bookings buchen');
        const open = paint(
            base({
                role: 'schueler',
                view: 'plan',
                items: week,
                accountEmail: 'neu@schule.at'
            })
        );
        expect(open).toContain('https://book.example/kino');
        expect(open).toContain(
            '</button><a class="fb-btn fb-btn--ghost" href="https://book.example/kino" target="_blank" rel="noopener">In Bookings buchen</a>'
        );
        expect(open).toContain('pw-seg');
        expect(open).toContain('id="pwBtnPrintMine"');
        expect(open).not.toContain('pw-nav__btn-label">Kalender');
        expect(open).not.toContain('pw-nav__btn-label">Liste');
        expect(open).not.toContain('pw-nav__btn-label">Export');
        expect(open).not.toContain('data-pw-view="neu"');
        expect(open).toContain('id="pwFilterStatus"');
        expect(open).toContain('id="pwFilterQ"');
        expect(open).toContain('Mehr Filter');
        expect(open).toContain('value="beantragt">wartet');
        expect(open).toContain('value="freigegeben">frei');
        expect(open).not.toContain('id="pwFilterTag"');
        expect(open).not.toContain('Filter &amp; Suche');
        const openedFilters = paint(
            base({
                role: 'schueler',
                view: 'plan',
                items: week,
                filtersMore: true
            })
        );
        expect(openedFilters).toContain('id="pwFilterTag"');
        expect(openedFilters).toContain('id="pwFilterLehrer"');
        expect(openedFilters).toContain('Weniger Filter');
    });
});
