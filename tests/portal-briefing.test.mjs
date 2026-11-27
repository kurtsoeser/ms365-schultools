import { describe, expect, it } from 'vitest';
import { buildAufgabenBriefing } from '../src/tools/aufgaben/aufgaben-briefing.js';
import { buildElternbriefBriefing } from '../src/tools/elternbrief/elternbrief-briefing.js';
import { buildFreistellungBriefing } from '../src/tools/freistellung-planer/freistellung-briefing.js';
import { buildSchularbeitenBriefing } from '../src/tools/schularbeiten-planer/schularbeiten-briefing.js';
import { buildSchulaktivitaetenBriefing } from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-briefing.js';
import { buildTechnikBriefing } from '../src/tools/technik-meldungen/technik-meldungen-briefing.js';
import { buildTermineBriefing } from '../src/tools/termine-hub/termine-hub-briefing.js';
import { loadPortalBriefings } from '../src/shared/portal-briefing-load.js';
import { renderPortalBriefingHtml } from '../src/shared/portal-briefing-ui.js';

const NOW = new Date(2026, 10, 23, 8, 0, 0);
const TODAY = '2026-11-23';

describe('Termine-Briefing', () => {
    const termine = [
        { title: 'Wandertag', start: '2026-11-25T09:00:00', end: '2026-11-25T16:00:00', sichtbarkeit: 'oeffentlich' },
        { title: 'Konferenz', start: '2026-11-26T08:00:00', end: '2026-11-26T09:00:00', sichtbarkeit: 'intern' },
        { title: 'Später', start: '2026-12-20T08:00:00', end: '2026-12-20T09:00:00', sichtbarkeit: 'oeffentlich' }
    ];

    it('zeigt Lehrkräften öffentliche und interne Termine der nächsten 14 Tage', () => {
        const card = buildTermineBriefing(termine, { view: 'lehrer', now: NOW });
        expect(card.lines.map((l) => l.label)).toEqual(['Wandertag', 'Konferenz']);
        expect(card.lines[1].meta).toContain('Intern');
        expect(card.emptyText).toBe('');
    });

    it('zeigt Schüler/innen nur öffentliche Termine', () => {
        const card = buildTermineBriefing(termine, { view: 'schueler', now: NOW });
        expect(card.lines.map((l) => l.label)).toEqual(['Wandertag']);
        expect(card.lines[0].meta).not.toContain('Intern');
    });

    it('zeigt Schüler/innen den eigenen Outlook-Kalender', () => {
        const card = buildTermineBriefing(
            [{ title: 'Nachhilfe', start: '2026-11-24T15:00:00', end: '2026-11-24T16:00:00', sichtbarkeit: 'intern', sourceId: 'me' }],
            { view: 'schueler', now: NOW }
        );
        expect(card.lines.map((l) => l.label)).toEqual(['Nachhilfe']);
    });

    it('lässt die Kalenderkachel stehen, wenn nichts ansteht', () => {
        const card = buildTermineBriefing([], { view: 'lehrer', now: NOW });
        expect(card.lines).toEqual([]);
        expect(card.emptyText).toBe('Keine Termine in den nächsten 14 Tagen');
    });
});

describe('Aufgaben-Briefing', () => {
    const aufgaben = [
        { id: '1', title: 'Protokoll', zuweisungEmail: 'kurt@schule.at', status: 'offen', prioritaet: 'hoch', frist: '2026-11-27' },
        { id: '2', title: 'Erledigt', zuweisungEmail: 'kurt@schule.at', status: 'erledigt', prioritaet: 'mittel' },
        { id: '3', title: 'Fremd', zuweisungEmail: 'ada@schule.at', status: 'offen', prioritaet: 'hoch' },
        { id: '4', title: 'Liste', zuweisungEmail: 'kurt@schule.at', status: 'in_arbeit', prioritaet: 'mittel' }
    ];

    it('zeigt nur offene Aufgaben an die eigene E-Mail', () => {
        const card = buildAufgabenBriefing(aufgaben, { email: 'Kurt@schule.at' });
        expect(card.lines.map((l) => l.label)).toEqual(['Protokoll', 'Liste']);
        expect(card.lines[0].meta).toContain('Offen');
        expect(card.lines[0].meta).toContain('27.11.2026');
    });

    it('blendet die Kachel ohne E-Mail aus', () => {
        expect(buildAufgabenBriefing(aufgaben, { email: '' })).toBeNull();
    });

    it('blendet die Kachel aus, wenn nichts offen ist', () => {
        expect(buildAufgabenBriefing([aufgaben[1]], { email: 'kurt@schule.at' })).toBeNull();
    });
});

describe('Schularbeiten-Briefing', () => {
    const items = [
        { titel: 'Mathe', klasseCode: '4A', datum: '2026-12-12', status: 'fixiert', lehrerEmail: 'ada@schule.at' },
        { titel: 'Deutsch', klasseCode: '4A', datum: '2026-12-18', status: 'beantragt', lehrerEmail: 'kurt@schule.at', beantragtVon: 'kurt@schule.at' },
        { titel: 'Physik', klasseCode: '3B', datum: '2026-12-01', status: 'fixiert', lehrerEmail: 'kurt@schule.at' },
        { titel: 'Alt', klasseCode: '4A', datum: '2026-11-01', status: 'fixiert', lehrerEmail: 'kurt@schule.at' },
        { titel: 'Abgelehnt', klasseCode: '4A', datum: '2026-12-02', status: 'abgelehnt', lehrerEmail: 'kurt@schule.at' }
    ];

    it('zeigt Schüler/innen nur fixierte Arbeiten der eigenen Klasse', () => {
        const card = buildSchularbeitenBriefing(items, {
            view: 'schueler',
            email: 'schueler@schule.at',
            klasseCode: '4A',
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Mathe']);
        expect(card.lines[0].meta).toContain('fixiert');
    });

    it('zeigt Lehrkräften eigene beantragte und fixierte Arbeiten', () => {
        const card = buildSchularbeitenBriefing(items, {
            view: 'lehrer',
            email: 'kurt@schule.at',
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Physik', 'Deutsch']);
    });

    it('bleibt ohne Klasse für Schüler/innen unsichtbar', () => {
        expect(
            buildSchularbeitenBriefing(items, { view: 'schueler', email: 'schueler@schule.at', today: TODAY })
        ).toBeNull();
    });
});

describe('Schulaktivitäten-Briefing', () => {
    const items = [
        { aktivitaetId: 'a1', titel: 'Wandertag 3A', status: 'beantragt', startdatum: '2026-12-01', lehrerEmail: 'ada@schule.at', klasseCodes: '3A' },
        { aktivitaetId: 'a2', titel: 'Exkursion 2B', status: 'kv_stellungnahme', startdatum: '2026-12-04', lehrerEmail: 'ada@schule.at', klasseCode: '2B' },
        { aktivitaetId: 'a3', titel: 'Theater', status: 'genehmigt', startdatum: '2026-12-10', lehrerEmail: 'kurt@schule.at', beantragtVon: 'kurt@schule.at', klasseCode: '4A' },
        { aktivitaetId: 'a4', titel: 'Nachreichen', status: 'nachforderung', startdatum: '2026-11-20', lehrerEmail: 'kurt@schule.at', klasseCode: '4A' },
        { aktivitaetId: 'a5', titel: 'Wandertag 4A', status: 'genehmigt', startdatum: '2026-12-03', lehrerEmail: 'ada@schule.at', klasseCode: '4A' }
    ];

    it('zeigt der Direktion offene Entscheidungen und eigene Anträge', () => {
        const card = buildSchulaktivitaetenBriefing(items, {
            view: 'lehrer',
            email: 'kurt@schule.at',
            roles: ['admin', 'lehrer'],
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Wandertag 3A', 'Nachreichen', 'Theater']);
        expect(card.lines[0].meta).toContain('wartet auf Entscheidung');
        expect(card.lines[2].meta).toContain('genehmigt');
    });

    it('zeigt einer Lehrkraft nur eigene Vorgänge', () => {
        const card = buildSchulaktivitaetenBriefing(items, {
            view: 'lehrer',
            email: 'kurt@schule.at',
            roles: ['lehrer'],
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Nachreichen', 'Theater']);
    });

    it('zeigt dem Klassenvorstand Stellungnahmen der eigenen Klassen', () => {
        const card = buildSchulaktivitaetenBriefing(items, {
            view: 'lehrer',
            email: 'kv@schule.at',
            roles: ['kv', 'lehrer'],
            kvKlassen: ['3A'],
            today: TODAY
        });
        expect(card).toBeNull();
    });

    it('nimmt die Stellungnahme der passenden Klasse auf', () => {
        const card = buildSchulaktivitaetenBriefing(items, {
            view: 'lehrer',
            email: 'kv@schule.at',
            roles: ['kv'],
            kvKlassen: ['2B'],
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Exkursion 2B']);
        expect(card.lines[0].meta).toContain('Stellungnahme');
    });

    it('zeigt Schüler/innen genehmigte Aktivitäten der eigenen Klasse', () => {
        const card = buildSchulaktivitaetenBriefing(items, {
            view: 'schueler',
            klasseCode: '4A',
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Wandertag 4A', 'Theater']);
    });

    it('begrenzt die Liste auf fünf Zeilen', () => {
        const many = Array.from({ length: 6 }, (_, i) => ({
            aktivitaetId: 'd' + i,
            titel: 'Antrag ' + i,
            status: 'beantragt',
            startdatum: '2026-12-0' + (i + 1),
            lehrerEmail: 'ada@schule.at'
        }));
        const card = buildSchulaktivitaetenBriefing(many, {
            view: 'lehrer',
            email: 'kurt@schule.at',
            roles: ['admin'],
            today: TODAY
        });
        expect(card.lines).toHaveLength(5);
    });
});

describe('Portal-Briefing Darstellung', () => {
    it('escaped Titel und lässt eine leere Kachel stehen', () => {
        const html = renderPortalBriefingHtml([
            {
                toolId: 'termine-hub',
                title: 'Nächste <Termine>',
                href: 'tools/termine-hub.html',
                lines: [],
                emptyText: 'Keine Termine in den nächsten 14 Tagen'
            }
        ]);
        expect(html).toContain('Nächste &lt;Termine&gt;');
        expect(html).toContain('Keine Termine in den nächsten 14 Tagen');
        expect(html).toContain('Für dich');
        expect(html).toContain('bi-calendar-event');
        expect(html).toContain('dash-portal-briefing__grid');
        expect(html).not.toContain('data-attention');
    });

    it('lädt keine Kachel, wenn das Werkzeug im Portal fehlt', async () => {
        const cards = await loadPortalBriefings({
            view: 'lehrer',
            email: 'kurt@schule.at',
            isVisible: () => false
        });
        expect(cards).toEqual([]);
    });
});

describe('Freistellungs-Briefing', () => {
    const items = [
        {
            itemId: '1',
            schuelerName: 'Anna',
            klasse: '4A',
            beginn: '2026-11-25',
            ende: '2026-11-25',
            status: 'Ausstehend',
            authorEmail: 'schueler@schule.at',
            kvEmail: 'kv@schule.at'
        },
        {
            itemId: '2',
            schuelerName: 'Ben',
            klasse: '2B',
            beginn: '2026-11-26',
            ende: '2026-11-28',
            status: 'Ausstehend',
            authorEmail: 'other@schule.at',
            kvEmail: 'andere@schule.at'
        },
        {
            itemId: '3',
            schuelerName: 'Cara',
            klasse: '4A',
            beginn: '2026-11-24',
            ende: '2026-11-24',
            status: 'Genehmigt',
            authorEmail: 'schueler@schule.at'
        },
        {
            itemId: '4',
            schuelerName: 'Alt',
            klasse: '4A',
            beginn: '2026-11-01',
            ende: '2026-11-02',
            status: 'Genehmigt',
            authorEmail: 'schueler@schule.at'
        },
        {
            itemId: '5',
            schuelerName: 'Dirk',
            klasse: '4A',
            beginn: '2026-11-27',
            status: 'Abgelehnt',
            authorEmail: 'schueler@schule.at'
        }
    ];

    it('zeigt Schüler/innen eigene offene und noch laufende genehmigte Anträge', () => {
        const card = buildFreistellungBriefing(items, {
            view: 'schueler',
            email: 'schueler@schule.at',
            roles: ['direktion'],
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Anna', 'Cara']);
        expect(card.lines[0].meta).toContain('wartet auf Genehmigung');
        expect(card.attention).toBe(true);
    });

    it('zeigt der Direktion nur offene Entscheidungen', () => {
        const card = buildFreistellungBriefing(items, {
            view: 'lehrer',
            email: 'kurt@schule.at',
            roles: ['direktion'],
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Anna', 'Ben']);
        expect(card.lines[0].kicker).toContain('25.11');
        expect(card.lines[0].meta).toBe('4A');
        expect(card.summary).toContain('warten auf Entscheidung');
    });

    it('zeigt dem KV nur die eigene Klasse und die eigene E-Mail', () => {
        const card = buildFreistellungBriefing(items, {
            view: 'lehrer',
            email: 'kv@schule.at',
            roles: ['kv'],
            kvKlassen: ['2B'],
            today: TODAY
        });
        expect(card.lines.map((l) => l.label)).toEqual(['Anna', 'Ben']);
        expect(card.lines.find((l) => l.label === 'Ben').meta).toContain('Stellungnahme');
    });

    it('blendet eine leere KV-Sicht aus', () => {
        expect(
            buildFreistellungBriefing(items, {
                view: 'lehrer',
                email: 'fremd@schule.at',
                roles: ['kv'],
                kvKlassen: [],
                today: TODAY
            })
        ).toBeNull();
    });

    it('blendet die Kachel ohne E-Mail aus', () => {
        expect(buildFreistellungBriefing(items, { view: 'lehrer', email: '', roles: ['direktion'] })).toBeNull();
    });
});

describe('Technik-Briefing', () => {
    const meldungen = [
        { title: 'Beamer 12', melderEmail: 'kurt@schule.at', status: 'offen', prioritaet: 'dringend', raumCode: 'R12' },
        { title: 'Fremd', melderEmail: 'ada@schule.at', status: 'offen', prioritaet: 'hoch', raumCode: 'R1' },
        { title: 'Erledigt', melderEmail: 'kurt@schule.at', status: 'erledigt', prioritaet: 'mittel', raumName: 'Turnsaal' },
        { title: 'Netzwerk', melderEmail: 'kurt@schule.at', status: 'in_arbeit', prioritaet: 'mittel', raumName: 'EDV' }
    ];

    it('zeigt nur eigene offene Meldungen', () => {
        const card = buildTechnikBriefing(meldungen, { email: 'Kurt@schule.at' });
        expect(card.lines.map((l) => l.label)).toEqual(['Beamer 12', 'Netzwerk']);
        expect(card.lines[0].meta).toContain('Dringend');
        expect(card.lines[0].meta).toContain('R12');
        expect(card.attention).toBe(true);
    });

    it('blendet die Kachel aus, wenn nichts offen ist', () => {
        expect(buildTechnikBriefing([meldungen[2]], { email: 'kurt@schule.at' })).toBeNull();
        expect(buildTechnikBriefing(meldungen, { email: '' })).toBeNull();
    });
});

describe('Elternbrief-Briefing', () => {
    const briefe = [
        { id: 'e1', betreff: 'Wandertag', status: 'entwurf', erstellerEmail: 'kurt@schule.at', createdAt: '2026-11-20' },
        { id: 'e2', betreff: 'Fremd', status: 'entwurf', erstellerEmail: 'ada@schule.at', createdAt: '2026-11-21' },
        { id: 'e3', betreff: 'Archiv', status: 'geschlossen', erstellerEmail: 'kurt@schule.at', createdAt: '2026-11-22' },
        { id: 'e4', betreff: 'Skikurs', status: 'gesendet', erstellerEmail: 'kurt@schule.at', createdAt: '2026-11-19' },
        { id: 'e5', betreff: 'Gelesen', status: 'gesendet', erstellerEmail: 'kurt@schule.at', createdAt: '2026-11-18' }
    ];
    const empfaenger = [
        { briefId: 'e4', status: 'offen' },
        { briefId: 'e4', status: 'offen' },
        { briefId: 'e4', status: 'gelesen' },
        { briefId: 'e5', status: 'gelesen' },
        { briefId: 'e5', status: 'gelesen' }
    ];

    it('zeigt eigene Entwürfe und gesendete Briefe mit ungelesenen Empfängern', () => {
        const card = buildElternbriefBriefing(briefe, empfaenger, { email: 'Kurt@schule.at' });
        expect(card.lines.map((l) => l.label)).toEqual(['Wandertag', 'Skikurs']);
        expect(card.lines[0].meta).toBe('Entwurf');
        expect(card.lines[1].meta).toBe('2 von 3 ungelesen');
    });

    it('blendet die Kachel aus, wenn nichts offen ist', () => {
        expect(buildElternbriefBriefing([briefe[4]], empfaenger, { email: 'kurt@schule.at' })).toBeNull();
        expect(buildElternbriefBriefing(briefe, empfaenger, { email: '' })).toBeNull();
    });
});
