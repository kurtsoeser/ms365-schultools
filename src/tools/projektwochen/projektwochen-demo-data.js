/**
 * Demo-Paket Projektwochen – Kurtrocks-Tenant (HAK).
 * Generator für docs/demo-data/projektwochen-demo.json und Seed auf SharePoint.
 * Kein Auto-Seed für andere Schulen: Import/Schreiben nur bewusst (UI oder Script).
 */

import { weekdayLabelDeFromIso } from './projektwochen-logic.js';

export const DEMO_SITE_DEFAULT = 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools';
export const DEMO_SEED_TAG = 'pw-demo-2026';
export const DEMO_AKTION_ID = 'pw-demo-2026';

/** Festwoche Mo–Fr (reproduzierbares JSON). */
export const DEMO_WEEK = {
    startIso: '2026-06-15',
    endIso: '2026-06-19',
    buchungAbIso: '2026-06-01T08:00:00'
};

/** Stammdaten-Vorschlag – Codes/E-Mails wie Schularbeiten-Demo (kurtrocks). */
export const DEMO_STAMMDATEN = {
    schoolName: 'MS365-Schule Kurtrocks',
    domain: 'kurtrocks.onmicrosoft.com',
    subjects: [],
    classes: [
        { code: '1AK', name: '1AK', year: 1 },
        { code: '1BK', name: '1BK', year: 1 },
        { code: '2AK', name: '2AK', year: 2 },
        { code: '2BK', name: '2BK', year: 2 },
        { code: '3AK', name: '3AK', year: 3 },
        { code: '3BK', name: '3BK', year: 3 },
        { code: '4AK', name: '4AK', year: 4 },
        { code: '4BK', name: '4BK', year: 4 },
        { code: '5AK', name: '5AK', year: 5 },
        { code: '5BK', name: '5BK', year: 5 }
    ],
    teachers: [
        { code: 'BAU', name: 'Mag. Julia Bauer', email: 'julia.bauer@kurtrocks.onmicrosoft.com' },
        { code: 'HOF', name: 'Mag. Thomas Hofmann', email: 'thomas.hofmann@kurtrocks.onmicrosoft.com' },
        { code: 'MAY', name: 'Mag. Anna Mayer', email: 'anna.mayer@kurtrocks.onmicrosoft.com' },
        { code: 'SCH', name: 'Dipl.-Ing. Markus Schuster', email: 'markus.schuster@kurtrocks.onmicrosoft.com' },
        { code: 'WEI', name: 'Mag. Eva Weiss', email: 'eva.weiss@kurtrocks.onmicrosoft.com' },
        { code: 'GRA', name: 'Mag. Peter Gruber', email: 'peter.gruber@kurtrocks.onmicrosoft.com' },
        { code: 'KIN', name: 'Mag. Sarah Binder', email: 'sarah.binder@kurtrocks.onmicrosoft.com' },
        { code: 'LEH', name: 'Mag. Christoph Lehmann', email: 'christoph.lehmann@kurtrocks.onmicrosoft.com' }
    ],
    students: [
        { klasse: '1BK', name: 'Demo Schüler 1BK', email: 'schueler.1bk@kurtrocks.onmicrosoft.com' },
        { klasse: '3AK', name: 'Demo Schüler 3AK', email: 'schueler.3ak@kurtrocks.onmicrosoft.com' },
        { klasse: '5BK', name: 'Demo Schüler 5BK', email: 'schueler.5bk@kurtrocks.onmicrosoft.com' }
    ]
};

function teacherByCode(code) {
    return DEMO_STAMMDATEN.teachers.find((t) => t.code === code) || null;
}

function dayIso(dayIndex) {
    const start = new Date(DEMO_WEEK.startIso + 'T12:00:00');
    const d = new Date(start.getFullYear(), start.getMonth(), start.getDate() + dayIndex);
    return (
        d.getFullYear() +
        '-' +
        String(d.getMonth() + 1).padStart(2, '0') +
        '-' +
        String(d.getDate()).padStart(2, '0')
    );
}

/**
 * @returns {object} SPO-Felder PW-Aktionen
 */
export function buildDemoAktionFields() {
    return {
        Title: 'Projektwoche 2026 (Kurtrocks-Demo)',
        AktionId: DEMO_AKTION_ID,
        Startdatum: DEMO_WEEK.startIso,
        Enddatum: DEMO_WEEK.endIso,
        BuchungAbDefault: DEMO_WEEK.buchungAbIso,
        BookingsBusinessId: '',
        BookingsBusinessName: '',
        Status: 'offen',
        Beschreibung:
            DEMO_SEED_TAG +
            ' · Demo-Projektwoche für kurtrocks. SharePoint + App, Bookings-Sync optional danach.'
    };
}

/**
 * Kuratierte Angebote (stabile AngebotId für Upsert).
 * @returns {object[]} SPO-Felder PW-Angebote
 */
export function buildDemoAngebotFields() {
    /** @type {Array<object>} */
    const samples = [
        {
            id: 'ang-demo-zoo',
            title: 'Zoo Schönbrunn',
            kategorie: 'exkursion',
            slot: 'ganztags',
            startzeit: '08:00',
            endzeit: '16:00',
            ort: 'Tiergarten Schönbrunn',
            kapazitaet: 28,
            preisEuro: 14,
            kostenHinweis: 'inkl. Eintritt + U-Bahn',
            status: 'freigegeben',
            lehrerCode: 'BAU',
            dayIndex: 0,
            zielklassen: '1AK;1BK;2AK;2BK',
            beschreibung: 'Führung durch den Zoo inkl. Fütterung. Treffpunkt Schulhof 7:45.',
            hinweisEltern: 'Kostenpflichtig – Einverständnis der Erziehungsberechtigten empfohlen.',
            notizIntern: 'Bus reserviert · ' + DEMO_SEED_TAG
        },
        {
            id: 'ang-demo-excel',
            title: 'Excel-Workshop Fortgeschritten',
            kategorie: 'workshop',
            slot: 'vormittag',
            startzeit: '08:00',
            endzeit: '12:00',
            ort: 'EDV 3',
            kapazitaet: 18,
            preisEuro: 0,
            status: 'beantragt',
            lehrerCode: 'SCH',
            dayIndex: 0,
            zielklassen: '3AK;3BK;4AK;4BK',
            beschreibung: 'Pivot, XLOOKUP, bedingte Formatierung – praxisnah mit Übungsdateien.'
        },
        {
            id: 'ang-demo-kino',
            title: 'Kinobesuch – Wirtschaftsfilm',
            kategorie: 'kultur',
            slot: 'nachmittag',
            startzeit: '13:00',
            endzeit: '17:00',
            ort: 'Cinema Center',
            kapazitaet: 32,
            preisEuro: 9,
            kostenHinweis: 'Ticket',
            status: 'freigegeben',
            lehrerCode: 'MAY',
            dayIndex: 0,
            zielklassen: 'alle',
            hinweisEltern: 'Kostenpflichtig – Einverständnis empfohlen.',
            notizIntern: DEMO_SEED_TAG
        },
        {
            id: 'ang-demo-sport',
            title: 'Sportparcours & Teamspiele',
            kategorie: 'sport',
            slot: 'vormittag',
            startzeit: '08:00',
            endzeit: '12:00',
            ort: 'Turnhalle / Sportplatz',
            kapazitaet: 24,
            preisEuro: 0,
            status: 'freigegeben',
            lehrerCode: 'LEH',
            dayIndex: 1,
            zielklassen: '1AK;1BK;2AK;2BK;3AK;3BK'
        },
        {
            id: 'ang-demo-medien',
            title: 'Workshop Medienkompetenz',
            kategorie: 'workshop',
            slot: 'nachmittag',
            startzeit: '13:00',
            endzeit: '16:30',
            ort: 'Bibliothek',
            kapazitaet: 16,
            preisEuro: 0,
            status: 'entwurf',
            lehrerCode: 'HOF',
            dayIndex: 1,
            zielklassen: 'alle',
            beschreibung: 'Fake News, Quellenkritik, KI-Tools im Alltag.'
        },
        {
            id: 'ang-demo-firma',
            title: 'Betriebsbesichtigung Logistikzentrum',
            kategorie: 'exkursion',
            slot: 'vormittag',
            startzeit: '08:30',
            endzeit: '12:30',
            ort: 'Logistikpark Süd',
            kapazitaet: 20,
            preisEuro: 0,
            status: 'abgelehnt',
            lehrerCode: 'GRA',
            dayIndex: 1,
            zielklassen: '4AK;4BK;5AK;5BK',
            ablehnungsGrund: 'Doppelbelegung Aufsicht / kein zweiter Begleiter',
            notizIntern: DEMO_SEED_TAG
        },
        {
            id: 'ang-demo-erstehilfe',
            title: 'Erste-Hilfe-Auffrischung',
            kategorie: 'workshop',
            slot: 'vormittag',
            startzeit: '08:00',
            endzeit: '12:00',
            ort: 'Mehrzwecksaal',
            kapazitaet: 15,
            preisEuro: 5,
            kostenHinweis: 'Materialpauschale',
            status: 'freigegeben',
            lehrerCode: 'WEI',
            dayIndex: 2,
            zielklassen: 'alle',
            hinweisEltern: 'Geringer Kostenbeitrag – Einverständnis empfohlen.'
        },
        {
            id: 'ang-demo-museum',
            title: 'Technisches Museum Wien',
            kategorie: 'kultur',
            slot: 'ganztags',
            startzeit: '08:15',
            endzeit: '15:30',
            ort: 'Technisches Museum',
            kapazitaet: 26,
            preisEuro: 8,
            status: 'freigegeben',
            lehrerCode: 'SCH',
            dayIndex: 2,
            zielklassen: '2AK;2BK;3AK;3BK',
            hinweisEltern: 'Kostenpflichtig – Einverständnis empfohlen.',
            begleitung: 'BAU'
        },
        {
            id: 'ang-demo-yoga',
            title: 'Bewegung & Entspannung',
            kategorie: 'sport',
            slot: 'nachmittag',
            startzeit: '13:00',
            endzeit: '15:30',
            ort: 'Aula',
            kapazitaet: 20,
            preisEuro: 0,
            status: 'beantragt',
            lehrerCode: 'KIN',
            dayIndex: 2,
            zielklassen: 'alle'
        },
        {
            id: 'ang-demo-bewerbung',
            title: 'Bewerbungstraining & LinkedIn',
            kategorie: 'workshop',
            slot: 'vormittag',
            startzeit: '08:00',
            endzeit: '12:00',
            ort: 'EDV 1',
            kapazitaet: 18,
            preisEuro: 0,
            status: 'freigegeben',
            lehrerCode: 'MAY',
            dayIndex: 3,
            zielklassen: '4AK;4BK;5AK;5BK'
        },
        {
            id: 'ang-demo-rad',
            title: 'Radtour Donaukanal',
            kategorie: 'sport',
            slot: 'nachmittag',
            startzeit: '13:00',
            endzeit: '17:00',
            ort: 'Treffpunkt Schule',
            kapazitaet: 16,
            preisEuro: 0,
            status: 'beantragt',
            lehrerCode: 'LEH',
            dayIndex: 3,
            zielklassen: '1AK;1BK;2AK;2BK',
            hinweisEltern: 'Eigenes Fahrrad + Helm erforderlich.',
            beschreibung: 'Geführte Tour, Pause am Kanal.'
        },
        {
            id: 'ang-demo-theater',
            title: 'Theaterworkshop Impro',
            kategorie: 'kultur',
            slot: 'vormittag',
            startzeit: '08:00',
            endzeit: '12:00',
            ort: 'Aula',
            kapazitaet: 14,
            preisEuro: 0,
            status: 'freigegeben',
            lehrerCode: 'KIN',
            dayIndex: 3,
            zielklassen: 'alle'
        },
        {
            id: 'ang-demo-bank',
            title: 'Exkursion Bank / Finanzen',
            kategorie: 'exkursion',
            slot: 'vormittag',
            startzeit: '08:45',
            endzeit: '12:15',
            ort: 'Partnerbank Innenstadt',
            kapazitaet: 22,
            preisEuro: 0,
            status: 'freigegeben',
            lehrerCode: 'GRA',
            dayIndex: 4,
            zielklassen: '3AK;3BK;4AK;4BK;5AK;5BK'
        },
        {
            id: 'ang-demo-kochen',
            title: 'Gesundes Kochen',
            kategorie: 'workshop',
            slot: 'nachmittag',
            startzeit: '13:00',
            endzeit: '16:30',
            ort: 'Lehrküche',
            kapazitaet: 12,
            preisEuro: 6,
            kostenHinweis: 'Lebensmittel',
            status: 'beantragt',
            lehrerCode: 'WEI',
            dayIndex: 4,
            zielklassen: 'alle',
            hinweisEltern: 'Kostenbeitrag Lebensmittel – Allergien bitte melden.'
        },
        {
            id: 'ang-demo-foto',
            title: 'Smartphone-Fotografie',
            kategorie: 'workshop',
            slot: 'nachmittag',
            startzeit: '13:00',
            endzeit: '16:00',
            ort: 'Hof / EDV 2',
            kapazitaet: 16,
            preisEuro: 0,
            status: 'entwurf',
            lehrerCode: 'HOF',
            dayIndex: 4,
            zielklassen: '1AK;1BK;2AK;2BK;3AK;3BK'
        },
        {
            id: 'ang-demo-escape',
            title: 'Escape-Room Teamchallenge',
            kategorie: 'sonstiges',
            slot: 'nachmittag',
            startzeit: '14:00',
            endzeit: '17:00',
            ort: 'Escape Venue',
            kapazitaet: 20,
            preisEuro: 18,
            status: 'freigegeben',
            lehrerCode: 'BAU',
            dayIndex: 4,
            zielklassen: '3AK;3BK;4AK;4BK;5AK;5BK',
            hinweisEltern: 'Kostenpflichtig – Einverständnis erforderlich.',
            notizIntern: 'Anzahlung geleistet · ' + DEMO_SEED_TAG
        },
        {
            id: 'ang-demo-debate',
            title: 'Debattierclub: Wirtschaft & Ethik',
            kategorie: 'workshop',
            slot: 'vormittag',
            startzeit: '08:00',
            endzeit: '11:30',
            ort: 'Seminarraum 2',
            kapazitaet: 18,
            preisEuro: 0,
            status: 'freigegeben',
            lehrerCode: 'MAY',
            dayIndex: 1,
            zielklassen: '4AK;4BK;5AK;5BK'
        },
        {
            id: 'ang-demo-coding',
            title: 'Mini-Coding mit Python',
            kategorie: 'workshop',
            slot: 'nachmittag',
            startzeit: '13:00',
            endzeit: '16:30',
            ort: 'EDV 3',
            kapazitaet: 16,
            preisEuro: 0,
            status: 'beantragt',
            lehrerCode: 'SCH',
            dayIndex: 3,
            zielklassen: '2AK;2BK;3AK;3BK'
        },
        {
            id: 'ang-demo-wanderung',
            title: 'Wanderung Kahlenberg',
            kategorie: 'sport',
            slot: 'ganztags',
            startzeit: '08:00',
            endzeit: '15:00',
            ort: 'Nußdorf – Kahlenberg',
            kapazitaet: 24,
            preisEuro: 4,
            kostenHinweis: 'Öffentliche Verkehrsmittel',
            status: 'freigegeben',
            lehrerCode: 'LEH',
            dayIndex: 2,
            zielklassen: 'alle',
            begleitung: 'GRA',
            hinweisEltern: 'Festes Schuhwerk, Jause mitnehmen.'
        },
        {
            id: 'ang-demo-podcast',
            title: 'Podcast aufnehmen',
            kategorie: 'workshop',
            slot: 'vormittag',
            startzeit: '08:00',
            endzeit: '12:00',
            ort: 'Medienraum',
            kapazitaet: 12,
            preisEuro: 0,
            status: 'abgelehnt',
            lehrerCode: 'HOF',
            dayIndex: 0,
            zielklassen: '3AK;3BK;4AK;4BK',
            ablehnungsGrund: 'Technikraum belegt – bitte anderen Tag wählen',
            notizIntern: DEMO_SEED_TAG
        }
    ];

    return samples.map((s) => {
        const datum = dayIso(s.dayIndex);
        const t = teacherByCode(s.lehrerCode);
        const freigegeben = s.status === 'freigegeben';
        return {
            Title: s.title,
            AngebotId: s.id,
            AktionId: DEMO_AKTION_ID,
            Beschreibung: s.beschreibung || s.title,
            HinweisEltern: s.hinweisEltern || '',
            Ort: s.ort,
            Treffpunkt: 'Schulhof / Eingang Aula',
            Wochentag: weekdayLabelDeFromIso(datum),
            Datum: datum,
            Slot: s.slot,
            Startzeit: s.startzeit,
            Endzeit: s.endzeit,
            Kapazitaet: s.kapazitaet,
            PreisEuro: s.preisEuro || 0,
            KostenHinweis: s.kostenHinweis || '',
            Zielklassen: s.zielklassen || 'alle',
            LehrerCode: s.lehrerCode,
            LehrerEmail: t ? t.email : '',
            Begleitung: s.begleitung || '',
            Kategorie: s.kategorie,
            Status: s.status,
            BuchungAb: freigegeben ? DEMO_WEEK.buchungAbIso : '',
            AblehnungsGrund: s.ablehnungsGrund || '',
            BookingsServiceId: '',
            BookingsBookingUrl: '',
            SyncStatus: '',
            SyncFehler: '',
            SyncAm: '',
            BeantragtVon: t ? t.email : 'demo@kurtrocks.onmicrosoft.com',
            FreigegebenVon: freigegeben ? 'admin@kurtrocks.onmicrosoft.com' : '',
            FreigegebenAm: freigegeben ? DEMO_WEEK.buchungAbIso : '',
            NotizIntern: s.notizIntern || DEMO_SEED_TAG
        };
    });
}

/**
 * Gesamtpaket für JSON / Seed / lokalen Import.
 */
export function getDemoSeedPackage() {
    const aktion = buildDemoAktionFields();
    const angebote = buildDemoAngebotFields();
    return {
        kind: 'ms365-projektwochen-demo-v1',
        seedTag: DEMO_SEED_TAG,
        siteDefault: DEMO_SITE_DEFAULT,
        stammdaten: DEMO_STAMMDATEN,
        aktion,
        angebote,
        counts: {
            aktionen: 1,
            angebote: angebote.length,
            freigegeben: angebote.filter((a) => a.Status === 'freigegeben').length,
            beantragt: angebote.filter((a) => a.Status === 'beantragt').length,
            teachers: DEMO_STAMMDATEN.teachers.length,
            classes: DEMO_STAMMDATEN.classes.length
        }
    };
}

/**
 * App-State ohne SharePoint (aus Generator oder geparstem JSON).
 * @param {object} [pack]
 */
export function buildLocalDemoState(pack) {
    const p = pack && typeof pack === 'object' ? pack : getDemoSeedPackage();
    const a = p.aktion || buildDemoAktionFields();
    const angeboteFields = Array.isArray(p.angebote) ? p.angebote : buildDemoAngebotFields();
    const aktion = {
        itemId: 'local-aktion-' + String(a.AktionId || DEMO_AKTION_ID),
        title: String(a.Title || ''),
        aktionId: String(a.AktionId || DEMO_AKTION_ID),
        startdatum: String(a.Startdatum || '').slice(0, 10),
        enddatum: String(a.Enddatum || '').slice(0, 10),
        buchungAbDefault: String(a.BuchungAbDefault || ''),
        bookingsBusinessId: String(a.BookingsBusinessId || ''),
        bookingsBusinessName: String(a.BookingsBusinessName || ''),
        status: String(a.Status || 'offen'),
        beschreibung: String(a.Beschreibung || '')
    };
    const angebote = angeboteFields.map((f, idx) => ({
        itemId: 'local-' + String(f.AngebotId || idx),
        angebotId: String(f.AngebotId || ''),
        aktionId: String(f.AktionId || aktion.aktionId),
        title: String(f.Title || ''),
        beschreibung: String(f.Beschreibung || ''),
        hinweisEltern: String(f.HinweisEltern || ''),
        ort: String(f.Ort || ''),
        treffpunkt: String(f.Treffpunkt || ''),
        tag: String(f.Wochentag || ''),
        datum: String(f.Datum || '').slice(0, 10),
        slot: String(f.Slot || 'vormittag'),
        startzeit: String(f.Startzeit || ''),
        endzeit: String(f.Endzeit || ''),
        kapazitaet: Number(f.Kapazitaet) || 0,
        preisEuro: Number(f.PreisEuro) || 0,
        kostenHinweis: String(f.KostenHinweis || ''),
        zielklassen: String(f.Zielklassen || 'alle'),
        lehrerCode: String(f.LehrerCode || ''),
        lehrerEmail: String(f.LehrerEmail || '').toLowerCase(),
        begleitung: String(f.Begleitung || ''),
        kategorie: String(f.Kategorie || 'sonstiges'),
        status: String(f.Status || 'beantragt'),
        buchungAb: String(f.BuchungAb || ''),
        ablehnungsGrund: String(f.AblehnungsGrund || ''),
        bookingsServiceId: String(f.BookingsServiceId || ''),
        bookingsBookingUrl: String(f.BookingsBookingUrl || ''),
        syncStatus: String(f.SyncStatus || ''),
        syncFehler: String(f.SyncFehler || ''),
        syncAm: String(f.SyncAm || ''),
        beantragtVon: String(f.BeantragtVon || ''),
        freigegebenVon: String(f.FreigegebenVon || ''),
        freigegebenAm: String(f.FreigegebenAm || ''),
        notizIntern: String(f.NotizIntern || '')
    }));
    return {
        aktionen: [aktion],
        angebote,
        stammdaten: p.stammdaten || DEMO_STAMMDATEN
    };
}
