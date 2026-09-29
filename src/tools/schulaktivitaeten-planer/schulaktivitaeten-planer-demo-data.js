/**
 * Demo-Paket Schulaktivitäten / Exkursionen – Schuljahr 2026/27 (HAK Kurtrocks).
 * Rein, ohne DOM/fetch – für lokalen Demo-Modus, JSON-Export und SharePoint-Seed.
 */

export const DEMO_SITE_DEFAULT = 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools';
export const DEMO_SCHOOL_YEAR = '2026/27';
export const DEMO_SEED_TAG = 'akt-demo-2026-27';

/** Stammdaten wie Schularbeiten-/Projektwochen-Demo (gleiche Codes/E-Mails). */
export const DEMO_STAMMDATEN = {
    schoolName: 'Demo-HAK Kurtrocks',
    domain: 'kurtrocks.onmicrosoft.com',
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
    ]
};

/** @type {{ Title: string, RegelwerkId: string, MinVorlaufTage: number, MaxGleichzeitigProKlasse: number, Aktiv: boolean }} */
export const DEMO_REGELWERK = {
    Title: 'Standard Schulaktivitäten 2026/27',
    RegelwerkId: 'akt-rw-demo-2627',
    MinVorlaufTage: 7,
    MaxGleichzeitigProKlasse: 1,
    Aktiv: true
};

function teacherEmail(code) {
    const t = DEMO_STAMMDATEN.teachers.find((x) => x.code === code);
    return t ? t.email : '';
}

/**
 * @param {object} p
 * @returns {object} SharePoint-Felder
 */
function akt(p) {
    const start = p.start;
    const end = p.end || start;
    const status = String(p.status || 'beantragt').toLowerCase();
    const decided = status === 'genehmigt' || status === 'abgelehnt';
    const fields = {
        Title: p.title,
        AktivitaetId: p.id,
        Typ: p.typ || 'Exkursion',
        KlasseCode: p.klasse,
        LehrerCode: p.lehrer,
        LehrerEmail: teacherEmail(p.lehrer),
        Begleitung: p.begleitung || '',
        Ort: p.ort || '',
        Startdatum: start,
        Enddatum: end,
        StartZeit: p.startZeit || '',
        EndZeit: p.endZeit || '',
        Status: status,
        Notiz: (p.notiz ? p.notiz + ' · ' : '') + DEMO_SEED_TAG + ' · SJ ' + DEMO_SCHOOL_YEAR,
        AblehnungsGrund: p.ablehnung || '',
        BeantragtVon: teacherEmail(p.lehrer),
        GenehmigtVon: decided ? 'admin@kurtrocks.onmicrosoft.com' : '',
        Verkehrsmittel: p.verkehr || '',
        KostenHinweis: p.kosten || ''
    };
    if (decided) fields.GenehmigtAm = start + 'T10:00:00Z';
    return fields;
}

/**
 * ~36 handkuratierte Aktivitäten über das Schuljahr (keine Klassen-Doppelbelegung am gleichen Tag).
 * @returns {object[]}
 */
export function buildDemoAktivitaeten() {
    return [
        // ——— Herbst 2026 ———
        akt({
            id: 'akt-demo-01',
            title: 'Naturhistorisches Museum',
            typ: 'Exkursion',
            klasse: '1AK',
            lehrer: 'KIN',
            ort: 'Wien, NHM',
            start: '2026-09-24',
            startZeit: '08:30',
            endZeit: '14:00',
            verkehr: 'Bahn ÖBB',
            kosten: 'ca. 18 € / Schüler',
            begleitung: 'MAY',
            status: 'genehmigt',
            notiz: 'Eintritt Gruppenpreis vereinbart'
        }),
        akt({
            id: 'akt-demo-02',
            title: 'Betriebsbesichtigung Raiffeisenbank',
            typ: 'Exkursion',
            klasse: '2AK',
            lehrer: 'HOF',
            ort: 'Krems',
            start: '2026-09-30',
            startZeit: '09:00',
            endZeit: '12:30',
            verkehr: 'Bus',
            kosten: 'kostenlos',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-03',
            title: 'Wandertag Wachau',
            typ: 'Schulaktivitaet',
            klasse: '3AK',
            lehrer: 'WEI',
            ort: 'Dürnstein–Krems',
            start: '2026-10-08',
            startZeit: '08:00',
            endZeit: '16:00',
            verkehr: 'Bahn + zu Fuß',
            kosten: 'ca. 12 €',
            begleitung: 'BAU',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-04',
            title: 'Parlament – Demokratieworkshop',
            typ: 'Exkursion',
            klasse: '4AK',
            lehrer: 'BAU',
            ort: 'Wien, Parlament',
            start: '2026-10-15',
            startZeit: '09:30',
            endZeit: '13:00',
            verkehr: 'Bahn',
            kosten: 'kostenlos',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-05',
            title: 'Englisch Theaterabend',
            typ: 'Veranstaltung',
            klasse: '5AK',
            lehrer: 'MAY',
            ort: 'Wien, English Theatre',
            start: '2026-10-22',
            startZeit: '17:30',
            endZeit: '21:30',
            verkehr: 'Bahn',
            kosten: 'ca. 35 €',
            status: 'genehmigt',
            notiz: 'Abendveranstaltung – Elterninfo versendet'
        }),
        akt({
            id: 'akt-demo-06',
            title: 'Erste-Hilfe-Kurs',
            typ: 'Schulaktivitaet',
            klasse: '1BK',
            lehrer: 'KIN',
            ort: 'Schulaula',
            start: '2026-11-05',
            startZeit: '08:00',
            endZeit: '13:00',
            verkehr: 'vor Ort',
            kosten: 'ca. 25 €',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-07',
            title: 'Kunsthistorisches Museum',
            typ: 'Exkursion',
            klasse: '2BK',
            lehrer: 'BAU',
            ort: 'Wien, KHM',
            start: '2026-11-12',
            startZeit: '08:45',
            endZeit: '14:30',
            verkehr: 'Bahn',
            kosten: 'ca. 15 €',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-08',
            title: 'BWL-Fallstudie bei Spar AG',
            typ: 'Exkursion',
            klasse: '3BK',
            lehrer: 'HOF',
            ort: 'Salzburg (Spar HQ)',
            start: '2026-11-19',
            end: '2026-11-20',
            startZeit: '07:00',
            endZeit: '18:00',
            verkehr: 'Bus',
            kosten: 'ca. 90 € inkl. Übernachtung',
            begleitung: 'GRA',
            status: 'genehmigt',
            notiz: '2 Tage · Hotel gebucht'
        }),
        akt({
            id: 'akt-demo-09',
            title: 'Wiener Börse Besuch',
            typ: 'Exkursion',
            klasse: '4BK',
            lehrer: 'GRA',
            ort: 'Wien, Börse',
            start: '2026-11-26',
            startZeit: '09:00',
            endZeit: '13:00',
            verkehr: 'Bahn',
            kosten: 'kostenlos',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-10',
            title: 'Matura-Infoabend',
            typ: 'Veranstaltung',
            klasse: '5BK',
            lehrer: 'SCH',
            ort: 'Aula',
            start: '2026-12-03',
            startZeit: '18:00',
            endZeit: '20:00',
            verkehr: 'vor Ort',
            kosten: 'kostenlos',
            status: 'genehmigt',
            notiz: 'Eltern + Schüler'
        }),

        // ——— Advent / Winter ———
        akt({
            id: 'akt-demo-11',
            title: 'Adventmarkt Schulchor',
            typ: 'Veranstaltung',
            klasse: '2AK',
            lehrer: 'WEI',
            ort: 'Stadtplatz',
            start: '2026-12-10',
            startZeit: '16:00',
            endZeit: '19:00',
            verkehr: 'zu Fuß',
            kosten: 'kostenlos',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-12',
            title: 'Ski- und Snowboardtag',
            typ: 'Schulaktivitaet',
            klasse: '3AK',
            lehrer: 'SCH',
            ort: 'Hochkar',
            start: '2027-01-14',
            startZeit: '06:30',
            endZeit: '18:00',
            verkehr: 'Reisebus',
            kosten: 'ca. 55 € inkl. Lift',
            begleitung: 'WEI',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-13',
            title: 'Sportwoche Schladming',
            typ: 'Schulaktivitaet',
            klasse: '3BK',
            lehrer: 'SCH',
            ort: 'Schladming',
            start: '2027-01-18',
            end: '2027-01-22',
            startZeit: '07:00',
            endZeit: '17:00',
            verkehr: 'Reisebus',
            kosten: 'ca. 320 €',
            begleitung: 'MAY, WEI',
            status: 'genehmigt',
            notiz: 'Sportwoche 3. Jahrgang'
        }),
        akt({
            id: 'akt-demo-14',
            title: 'Uni Wien – Tag der offenen Tür',
            typ: 'Exkursion',
            klasse: '5AK',
            lehrer: 'BAU',
            ort: 'Wien, Campus',
            start: '2027-01-28',
            startZeit: '09:00',
            endZeit: '15:00',
            verkehr: 'Bahn',
            kosten: 'Fahrkarte',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-15',
            title: 'Workshop Politische Bildung',
            typ: 'Schulaktivitaet',
            klasse: '4AK',
            lehrer: 'BAU',
            ort: 'Klassenraum',
            start: '2027-02-04',
            startZeit: '08:00',
            endZeit: '12:00',
            verkehr: 'vor Ort',
            kosten: 'kostenlos',
            status: 'genehmigt'
        }),

        // ——— Frühjahr 2027 ———
        akt({
            id: 'akt-demo-16',
            title: 'Praxistage Unternehmen',
            typ: 'Schulaktivitaet',
            klasse: '4BK',
            lehrer: 'HOF',
            ort: 'diverse Betriebe',
            start: '2027-03-08',
            end: '2027-03-12',
            verkehr: 'individuell',
            kosten: 'keine',
            status: 'genehmigt',
            notiz: 'Praxistage 4. Jahrgang'
        }),
        akt({
            id: 'akt-demo-17',
            title: 'Technisches Museum',
            typ: 'Exkursion',
            klasse: '1AK',
            lehrer: 'SCH',
            ort: 'Wien, TMW',
            start: '2027-03-17',
            startZeit: '08:30',
            endZeit: '14:00',
            verkehr: 'Bahn',
            kosten: 'ca. 14 €',
            begleitung: 'KIN',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-18',
            title: 'Italienisch Kulturreise Verona',
            typ: 'Exkursion',
            klasse: '4AK',
            lehrer: 'LEH',
            ort: 'Verona',
            start: '2027-03-23',
            end: '2027-03-26',
            startZeit: '06:00',
            endZeit: '20:00',
            verkehr: 'Bus',
            kosten: 'ca. 280 €',
            begleitung: 'MAY',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-19',
            title: 'Umweltbildung Nationalpark Donau-Auen',
            typ: 'Exkursion',
            klasse: '2AK',
            lehrer: 'WEI',
            ort: 'Orth an der Donau',
            start: '2027-04-08',
            startZeit: '08:00',
            endZeit: '15:30',
            verkehr: 'Bus',
            kosten: 'ca. 22 €',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-20',
            title: 'Jobmesse WKO',
            typ: 'Veranstaltung',
            klasse: '5BK',
            lehrer: 'HOF',
            ort: 'St. Pölten, Messe',
            start: '2027-04-15',
            startZeit: '09:00',
            endZeit: '14:00',
            verkehr: 'Bahn',
            kosten: 'Fahrkarte',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-21',
            title: 'Französisch Café littéraire',
            typ: 'Schulaktivitaet',
            klasse: '3AK',
            lehrer: 'LEH',
            ort: 'Schulbibliothek',
            start: '2027-04-22',
            startZeit: '14:00',
            endZeit: '16:00',
            verkehr: 'vor Ort',
            kosten: 'kostenlos',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-22',
            title: 'Bundesmuseum für Wirtschaft',
            typ: 'Exkursion',
            klasse: '2BK',
            lehrer: 'GRA',
            ort: 'Wien',
            start: '2027-04-29',
            startZeit: '09:00',
            endZeit: '13:30',
            verkehr: 'Bahn',
            kosten: 'ca. 10 €',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-23',
            title: 'Bundesliga-Stadionführung',
            typ: 'Exkursion',
            klasse: '1BK',
            lehrer: 'SCH',
            ort: 'Wien, Allianz Stadion',
            start: '2027-05-06',
            startZeit: '10:00',
            endZeit: '13:00',
            verkehr: 'Bahn',
            kosten: 'ca. 12 €',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-24',
            title: 'Projektpräsentation 5. Jahrgang',
            typ: 'Veranstaltung',
            klasse: '5AK',
            lehrer: 'SCH',
            ort: 'Aula',
            start: '2027-05-12',
            startZeit: '09:00',
            endZeit: '12:00',
            verkehr: 'vor Ort',
            kosten: 'kostenlos',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-25',
            title: 'Radausflug Kamptal',
            typ: 'Schulaktivitaet',
            klasse: '3BK',
            lehrer: 'WEI',
            ort: 'Langenlois–Krems',
            start: '2027-05-20',
            startZeit: '08:00',
            endZeit: '15:00',
            verkehr: 'Fahrrad',
            kosten: 'ca. 5 €',
            begleitung: 'KIN',
            status: 'genehmigt'
        }),
        akt({
            id: 'akt-demo-26',
            title: 'Abschlussfeier 5. Klassen',
            typ: 'Veranstaltung',
            klasse: '5BK',
            lehrer: 'BAU',
            ort: 'Festsaal Gemeinde',
            start: '2027-06-18',
            startZeit: '18:00',
            endZeit: '22:00',
            verkehr: 'individuell',
            kosten: 'Beitrag 15 €',
            status: 'genehmigt'
        }),

        // ——— Offene Anträge (Freigabe-Demo) ———
        akt({
            id: 'akt-demo-27',
            title: 'Schönbrunn Zoo + Schloss',
            typ: 'Exkursion',
            klasse: '1AK',
            lehrer: 'KIN',
            ort: 'Wien, Schönbrunn',
            start: '2027-05-27',
            startZeit: '08:00',
            endZeit: '16:00',
            verkehr: 'Bahn',
            kosten: 'ca. 28 €',
            begleitung: 'MAY',
            status: 'beantragt',
            notiz: 'Wunschtermin nach Pfingsten'
        }),
        akt({
            id: 'akt-demo-28',
            title: 'Sparkasse Jugendkonto-Workshop',
            typ: 'Exkursion',
            klasse: '2AK',
            lehrer: 'HOF',
            ort: 'Krems, Filiale',
            start: '2027-06-03',
            startZeit: '09:00',
            endZeit: '11:30',
            verkehr: 'zu Fuß / Bus',
            kosten: 'kostenlos',
            status: 'beantragt'
        }),
        akt({
            id: 'akt-demo-29',
            title: 'Haus der Geschichte Österreich',
            typ: 'Exkursion',
            klasse: '4BK',
            lehrer: 'BAU',
            ort: 'Wien, Hofburg',
            start: '2027-06-09',
            startZeit: '09:00',
            endZeit: '14:00',
            verkehr: 'Bahn',
            kosten: 'ca. 8 €',
            status: 'beantragt'
        }),
        akt({
            id: 'akt-demo-30',
            title: 'IT-Unternehmen Cloud Day',
            typ: 'Exkursion',
            klasse: '4AK',
            lehrer: 'SCH',
            ort: 'Wien, Tech Hub',
            start: '2027-06-10',
            startZeit: '08:30',
            endZeit: '15:00',
            verkehr: 'Bahn',
            kosten: 'kostenlos',
            begleitung: 'GRA',
            status: 'beantragt'
        }),
        akt({
            id: 'akt-demo-31',
            title: 'Elternsprechtag-Abend (Organisation)',
            typ: 'Sonstiges',
            klasse: '3AK',
            lehrer: 'MAY',
            ort: 'Schule',
            start: '2026-11-10',
            startZeit: '16:00',
            endZeit: '19:30',
            verkehr: 'vor Ort',
            kosten: 'keine',
            status: 'beantragt',
            notiz: 'Raumplanung + Aufsicht'
        }),
        akt({
            id: 'akt-demo-32',
            title: 'Spanisch Austausch-Vorbereitung',
            typ: 'Schulaktivitaet',
            klasse: '3BK',
            lehrer: 'LEH',
            ort: 'Sprachlabor',
            start: '2027-02-25',
            startZeit: '14:00',
            endZeit: '16:30',
            verkehr: 'vor Ort',
            kosten: 'kostenlos',
            status: 'beantragt'
        }),

        // ——— Abgelehnt / Grenzfälle ———
        akt({
            id: 'akt-demo-33',
            title: 'Freizeitpark Tagesausflug',
            typ: 'Exkursion',
            klasse: '2BK',
            lehrer: 'KIN',
            ort: 'Prater Wien',
            start: '2026-10-20',
            startZeit: '09:00',
            endZeit: '17:00',
            verkehr: 'Bahn',
            kosten: 'ca. 45 €',
            status: 'abgelehnt',
            ablehnung: 'Zu nah an Schularbeiten-Woche; bitte anderen Termin.'
        }),
        akt({
            id: 'akt-demo-34',
            title: 'Skiurlaub Privatinitiative',
            typ: 'Sonstiges',
            klasse: '5AK',
            lehrer: 'WEI',
            ort: 'Ischgl',
            start: '2027-02-15',
            end: '2027-02-19',
            verkehr: 'Bus',
            kosten: 'ca. 450 €',
            status: 'abgelehnt',
            ablehnung: 'Kein schulischer Mehrwert / Kosten zu hoch.'
        }),
        akt({
            id: 'akt-demo-35',
            title: 'Fotowalk Altstadt',
            typ: 'Schulaktivitaet',
            klasse: '1BK',
            lehrer: 'MAY',
            ort: 'Stadtzentrum',
            start: '2027-06-16',
            startZeit: '08:00',
            endZeit: '11:00',
            verkehr: 'zu Fuß',
            kosten: 'kostenlos',
            status: 'beantragt',
            notiz: 'Optional vor Zeugniskonferenz'
        }),
        akt({
            id: 'akt-demo-36',
            title: 'Diplomarbeit-Präsentationstag',
            typ: 'Veranstaltung',
            klasse: '5AK',
            lehrer: 'GRA',
            ort: 'Aula + Seminarraum',
            start: '2027-04-20',
            startZeit: '08:00',
            endZeit: '16:00',
            verkehr: 'vor Ort',
            kosten: 'kostenlos',
            status: 'genehmigt',
            notiz: 'Gäste: Firmenpartner'
        })
    ];
}

/**
 * Gesamtpaket für Seed / JSON.
 */
export function getDemoSeedPackage() {
    const aktivitaeten = buildDemoAktivitaeten().map((row) => {
        const fields = { ...row };
        if (!fields.GenehmigtAm) delete fields.GenehmigtAm;
        if (!fields.AblehnungsGrund) fields.AblehnungsGrund = '';
        if (!fields.Begleitung) fields.Begleitung = '';
        if (!fields.StartZeit) fields.StartZeit = '';
        if (!fields.EndZeit) fields.EndZeit = '';
        if (!fields.Verkehrsmittel) fields.Verkehrsmittel = '';
        if (!fields.KostenHinweis) fields.KostenHinweis = '';
        if (!fields.GenehmigtVon) fields.GenehmigtVon = '';
        return fields;
    });
    return {
        schoolYear: DEMO_SCHOOL_YEAR,
        seedTag: DEMO_SEED_TAG,
        siteDefault: DEMO_SITE_DEFAULT,
        stammdaten: DEMO_STAMMDATEN,
        regelwerk: DEMO_REGELWERK,
        aktivitaeten,
        counts: {
            aktivitaeten: aktivitaeten.length,
            teachers: DEMO_STAMMDATEN.teachers.length,
            classes: DEMO_STAMMDATEN.classes.length,
            genehmigt: aktivitaeten.filter((a) => a.Status === 'genehmigt').length,
            beantragt: aktivitaeten.filter((a) => a.Status === 'beantragt').length,
            abgelehnt: aktivitaeten.filter((a) => a.Status === 'abgelehnt').length
        }
    };
}

/**
 * Demo-Paket → Planer-State (ohne SharePoint).
 * @param {object} [pack]
 */
export function buildLocalDemoState(pack) {
    const p = pack && typeof pack === 'object' ? pack : getDemoSeedPackage();
    const rows = Array.isArray(p.aktivitaeten) ? p.aktivitaeten : buildDemoAktivitaeten();
    const items = rows.map((f, i) => {
        const start = String(f.Startdatum || '').slice(0, 10);
        return {
            itemId: 'local-' + String(f.AktivitaetId || i),
            aktivitaetId: String(f.AktivitaetId || ''),
            titel: String(f.Title || ''),
            typ: String(f.Typ || 'Exkursion'),
            klasseCode: String(f.KlasseCode || ''),
            lehrerCode: String(f.LehrerCode || ''),
            lehrerEmail: String(f.LehrerEmail || '').toLowerCase(),
            begleitung: String(f.Begleitung || ''),
            ort: String(f.Ort || ''),
            startdatum: start,
            enddatum: String(f.Enddatum || start).slice(0, 10),
            startZeit: String(f.StartZeit || ''),
            endZeit: String(f.EndZeit || ''),
            status: String(f.Status || 'beantragt').toLowerCase(),
            notiz: String(f.Notiz || ''),
            ablehnungsGrund: String(f.AblehnungsGrund || ''),
            beantragtVon: String(f.BeantragtVon || ''),
            genehmigtVon: String(f.GenehmigtVon || ''),
            genehmigtAm: f.GenehmigtAm ? String(f.GenehmigtAm) : '',
            verkehrsmittel: String(f.Verkehrsmittel || ''),
            kostenHinweis: String(f.KostenHinweis || ''),
            schulterminKey: String(f.SchulterminKey || ''),
            _localOnly: true
        };
    });
    const rw = p.regelwerk || DEMO_REGELWERK;
    const rules = {
        itemId: 'local-rw',
        title: String(rw.Title || 'Demo-Regelwerk'),
        regelwerkId: String(rw.RegelwerkId || 'akt-rw-demo'),
        minVorlaufTage: Number(rw.MinVorlaufTage) || 7,
        maxGleichzeitigProKlasse: Number(rw.MaxGleichzeitigProKlasse) || 1,
        aktiv: rw.Aktiv !== false
    };
    return {
        items,
        rules,
        stammdaten: p.stammdaten || DEMO_STAMMDATEN,
        siteDefault: p.siteDefault || DEMO_SITE_DEFAULT,
        schoolYear: p.schoolYear || DEMO_SCHOOL_YEAR,
        counts: p.counts,
        pack: p
    };
}
