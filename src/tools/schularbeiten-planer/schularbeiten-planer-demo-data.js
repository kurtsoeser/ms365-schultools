/**
 * Umfassende Demo-Daten Schularbeiten-Planer – Schuljahr 2026/27 (HAK).
 * Rein, ohne DOM/fetch – für Seed auf SharePoint und Tests.
 */

/** @typedef {{ Title: string, RegelwerkId: string, MaxProTag: number, MaxProWoche: number, AnkuendigungsfristTage: number, SperreVorNotenkonferenzTage: number, Aktiv: boolean }} RegelwerkFields */
/** @typedef {{ Title: string, TerminfensterId: string, Typ: string, Startdatum: string, Enddatum: string, Beschreibung: string }} FensterFields */
/** @typedef {{ Title: string, FachCode: string, Farbe: string, HatSchularbeiten: boolean, ProSemester: number, StandardDauer: number }} FachMetaFields */
/** @typedef {object} SchularbeitFields */

export const DEMO_SITE_DEFAULT = 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools';
export const DEMO_SCHOOL_YEAR = '2026/27';
export const DEMO_SEED_TAG = 'demo-2026-27';

/** Stammdaten-Vorschlag (lokal / Dokumentation) – Codes = SharePoint-Felder. */
export const DEMO_STAMMDATEN = {
    schoolName: 'Demo-HAK Kurtrocks',
    domain: 'kurtrocks.onmicrosoft.com',
    subjects: [
        { code: 'D', name: 'Deutsch' },
        { code: 'E', name: 'Englisch' },
        { code: 'F', name: 'Französisch' },
        { code: 'M', name: 'Mathematik' },
        { code: 'AM', name: 'Angewandte Mathematik' },
        { code: 'BWL', name: 'Betriebswirtschaftslehre' },
        { code: 'UR', name: 'Unternehmensrechnung' },
        { code: 'RW', name: 'Rechnungswesen' },
        { code: 'I', name: 'Italienisch' },
        { code: 'SP', name: 'Spanisch' },
        { code: 'WI', name: 'Wirtschaftsinformatik' },
        { code: 'POE', name: 'Politische Bildung und Recht' }
    ],
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

/** @type {RegelwerkFields} */
export const DEMO_REGELWERK = {
    Title: 'Standard HAK Regelwerk 2026/27',
    RegelwerkId: 'rw-demo-2627',
    MaxProTag: 1,
    MaxProWoche: 2,
    AnkuendigungsfristTage: 7,
    SperreVorNotenkonferenzTage: 7,
    Aktiv: true
};

/** @type {FachMetaFields[]} */
export const DEMO_FACHMETA = [
    { Title: 'Deutsch', FachCode: 'D', Farbe: '#6366f1', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Englisch', FachCode: 'E', Farbe: '#0ea5e9', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Französisch', FachCode: 'F', Farbe: '#ec4899', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Mathematik', FachCode: 'M', Farbe: '#10b981', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Angewandte Mathematik', FachCode: 'AM', Farbe: '#14b8a6', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Betriebswirtschaftslehre', FachCode: 'BWL', Farbe: '#f59e0b', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Unternehmensrechnung', FachCode: 'UR', Farbe: '#f97316', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Rechnungswesen', FachCode: 'RW', Farbe: '#ef4444', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 75 },
    { Title: 'Italienisch', FachCode: 'I', Farbe: '#8b5cf6', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Spanisch', FachCode: 'SP', Farbe: '#e11d48', HatSchularbeiten: true, ProSemester: 2, StandardDauer: 100 },
    { Title: 'Wirtschaftsinformatik', FachCode: 'WI', Farbe: '#64748b', HatSchularbeiten: true, ProSemester: 1, StandardDauer: 75 },
    { Title: 'Politische Bildung und Recht', FachCode: 'POE', Farbe: '#78716c', HatSchularbeiten: false, ProSemester: 0, StandardDauer: 50 }
];

/** @type {FensterFields[]} */
export const DEMO_TERMINFENSTER = [
    {
        Title: 'Herbstferien 2026',
        TerminfensterId: 'tf-herbst-26',
        Typ: 'gesperrt',
        Startdatum: '2026-10-26',
        Enddatum: '2026-10-31',
        Beschreibung: DEMO_SEED_TAG + ' · Herbstferien'
    },
    {
        Title: 'Weihnachtsferien 2026/27',
        TerminfensterId: 'tf-weihnacht-26',
        Typ: 'gesperrt',
        Startdatum: '2026-12-24',
        Enddatum: '2027-01-06',
        Beschreibung: DEMO_SEED_TAG + ' · Weihnachtsferien'
    },
    {
        Title: 'Semesterferien 2027',
        TerminfensterId: 'tf-semester-27',
        Typ: 'gesperrt',
        Startdatum: '2027-02-08',
        Enddatum: '2027-02-13',
        Beschreibung: DEMO_SEED_TAG + ' · Semesterferien'
    },
    {
        Title: 'Osterferien 2027',
        TerminfensterId: 'tf-ostern-27',
        Typ: 'gesperrt',
        Startdatum: '2027-03-29',
        Enddatum: '2027-04-06',
        Beschreibung: DEMO_SEED_TAG + ' · Osterferien'
    },
    {
        Title: 'Pfingstferien 2027',
        TerminfensterId: 'tf-pfingst-27',
        Typ: 'gesperrt',
        Startdatum: '2027-05-17',
        Enddatum: '2027-05-18',
        Beschreibung: DEMO_SEED_TAG + ' · Pfingsten'
    },
    {
        Title: 'Sportwoche 3. Jahrgang',
        TerminfensterId: 'tf-sport-3jg',
        Typ: 'gesperrt',
        Startdatum: '2027-01-18',
        Enddatum: '2027-01-22',
        Beschreibung: DEMO_SEED_TAG + ' · Sportwoche 3AK/3BK'
    },
    {
        Title: 'Praxistage 4. Jahrgang',
        TerminfensterId: 'tf-praxis-4jg',
        Typ: 'gesperrt',
        Startdatum: '2027-03-08',
        Enddatum: '2027-03-12',
        Beschreibung: DEMO_SEED_TAG + ' · Praxistage'
    },
    {
        Title: 'Sperrwoche vor Notenkonferenz WS',
        TerminfensterId: 'tf-nk-ws',
        Typ: 'gesperrt',
        Startdatum: '2027-01-25',
        Enddatum: '2027-01-30',
        Beschreibung: DEMO_SEED_TAG + ' · vor Notenkonferenz Wintersemester'
    },
    {
        Title: 'Sperrwoche vor Notenkonferenz SS',
        TerminfensterId: 'tf-nk-ss',
        Typ: 'gesperrt',
        Startdatum: '2027-06-21',
        Enddatum: '2027-06-25',
        Beschreibung: DEMO_SEED_TAG + ' · vor Notenkonferenz Sommersemester'
    },
    {
        Title: 'Hauptferienbeginn',
        TerminfensterId: 'tf-sommer-27',
        Typ: 'gesperrt',
        Startdatum: '2027-07-03',
        Enddatum: '2027-09-05',
        Beschreibung: DEMO_SEED_TAG + ' · Sommerferien'
    }
];

/**
 * Kern-Schularbeiten (handkuratiert) + generierte Ergänzungen.
 * @returns {SchularbeitFields[]}
 */
export function buildDemoSchularbeiten() {
    /** @type {SchularbeitFields[]} */
    const rows = [
        // ——— Wintersemester: fixierte Termine ———
        sa('sa-d-3ak-ws1', 'Erste Deutsch-Schularbeit: Erörterung', 'D', '3AK', 'BAU', '2026-10-14', 100, 'WS', 'fixiert'),
        sa('sa-m-3ak-ws1', 'Mathematik: Funktionen', 'M', '3AK', 'SCH', '2026-10-21', 100, 'WS', 'fixiert'),
        sa('sa-e-3ak-ws1', 'Englisch: Textinterpretation', 'E', '3AK', 'MAY', '2026-11-04', 100, 'WS', 'fixiert'),
        sa('sa-bwl-3ak-ws1', 'BWL: Unternehmensformen', 'BWL', '3AK', 'HOF', '2026-11-18', 100, 'WS', 'fixiert'),
        sa('sa-d-3bk-ws1', 'Deutsch: Textanalyse', 'D', '3BK', 'BAU', '2026-10-15', 100, 'WS', 'fixiert'),
        sa('sa-m-3bk-ws1', 'Mathematik: Gleichungssysteme', 'M', '3BK', 'SCH', '2026-11-05', 100, 'WS', 'fixiert'),
        sa('sa-ur-4ak-ws1', 'UR: Buchhaltung Grundlagen', 'UR', '4AK', 'GRA', '2026-10-20', 100, 'WS', 'fixiert'),
        sa('sa-rw-4ak-ws1', 'RW: Jahresabschluss', 'RW', '4AK', 'GRA', '2026-11-17', 75, 'WS', 'fixiert'),
        sa('sa-am-5ak-ws1', 'AM: Analysis', 'AM', '5AK', 'SCH', '2026-10-13', 100, 'WS', 'fixiert'),
        sa('sa-bwl-5ak-ws1', 'BWL: Kostenrechnung', 'BWL', '5AK', 'HOF', '2026-11-10', 100, 'WS', 'fixiert'),
        sa('sa-d-5ak-ws1', 'Deutsch: Maturavorbereitung Text', 'D', '5AK', 'BAU', '2026-12-01', 150, 'WS', 'fixiert'),
        sa('sa-e-2ak-ws1', 'Englisch: Listening & Writing', 'E', '2AK', 'MAY', '2026-10-22', 100, 'WS', 'fixiert'),
        sa('sa-m-2ak-ws1', 'Mathematik: Prozentrechnung', 'M', '2AK', 'WEI', '2026-11-12', 100, 'WS', 'fixiert'),
        sa('sa-d-1ak-ws1', 'Deutsch: Erzählung', 'D', '1AK', 'KIN', '2026-11-03', 100, 'WS', 'fixiert'),
        sa('sa-e-1ak-ws1', 'Englisch: Vocabulary Test Essay', 'E', '1AK', 'MAY', '2026-11-24', 100, 'WS', 'fixiert'),
        sa('sa-f-4bk-ws1', 'Französisch: Compréhension', 'F', '4BK', 'LEH', '2026-10-16', 100, 'WS', 'fixiert'),
        sa('sa-wi-4ak-ws1', 'WI: Datenbanken', 'WI', '4AK', 'SCH', '2026-12-03', 75, 'WS', 'fixiert'),

        // ——— Beantragt (offen) ———
        sa('sa-m-3ak-ws2', 'Mathematik: Integralrechnung', 'M', '3AK', 'SCH', '2026-12-09', 100, 'WS', 'beantragt'),
        sa('sa-d-3ak-ws2', 'Deutsch: Textinterpretation Lyrik', 'D', '3AK', 'BAU', '2026-12-15', 100, 'WS', 'beantragt'),
        sa('sa-bwl-4ak-ws2', 'BWL: Marketing-Mix', 'BWL', '4AK', 'HOF', '2026-12-10', 100, 'WS', 'beantragt'),
        sa('sa-e-5ak-ws2', 'Englisch: Academic Writing', 'E', '5AK', 'MAY', '2026-12-08', 100, 'WS', 'beantragt'),
        sa('sa-ur-4bk-ws1', 'UR: Kostenstellenrechnung', 'UR', '4BK', 'GRA', '2027-01-14', 100, 'WS', 'beantragt'),
        sa('sa-am-5bk-ws1', 'AM: Stochastik', 'AM', '5BK', 'WEI', '2027-01-13', 100, 'WS', 'beantragt'),

        // ——— Abgelehnt (Demo) ———
        sa(
            'sa-d-3ak-rej',
            'Deutsch: Textinterpretation (abgelehnt – Frist)',
            'D',
            '3AK',
            'BAU',
            '2026-10-08',
            100,
            'WS',
            'abgelehnt',
            'Ankündigungsfrist unterschritten.'
        ),
        sa(
            'sa-m-2bk-rej',
            'Mathematik: zu nah an Sportwoche-Nachbar',
            'M',
            '2BK',
            'WEI',
            '2026-12-09',
            100,
            'WS',
            'abgelehnt',
            'Bereits max. Schularbeiten in der Kalenderwoche.'
        ),

        // ——— Sommersemester ———
        sa('sa-d-3ak-ss1', 'Deutsch: Erörterung Digitalisierung', 'D', '3AK', 'BAU', '2027-03-02', 100, 'SS', 'fixiert'),
        sa('sa-m-3ak-ss1', 'Mathematik: Trigonometrie', 'M', '3AK', 'SCH', '2027-03-16', 100, 'SS', 'fixiert'),
        sa('sa-e-3ak-ss1', 'Englisch: Literature', 'E', '3AK', 'MAY', '2027-04-13', 100, 'SS', 'fixiert'),
        sa('sa-bwl-3ak-ss1', 'BWL: Personalwirtschaft', 'BWL', '3AK', 'HOF', '2027-04-27', 100, 'SS', 'fixiert'),
        sa('sa-ur-4ak-ss1', 'UR: Bilanzanalyse', 'UR', '4AK', 'GRA', '2027-03-03', 100, 'SS', 'fixiert'),
        sa('sa-rw-4ak-ss1', 'RW: Steuerlehre', 'RW', '4AK', 'GRA', '2027-04-14', 75, 'SS', 'fixiert'),
        sa('sa-am-5ak-ss1', 'AM: Maturavorbereitung', 'AM', '5AK', 'SCH', '2027-03-04', 150, 'SS', 'fixiert'),
        sa('sa-d-5ak-ss1', 'Deutsch: Matura-Probe', 'D', '5AK', 'BAU', '2027-04-15', 150, 'SS', 'fixiert'),
        sa('sa-e-2ak-ss1', 'Englisch: Oral prep written', 'E', '2AK', 'MAY', '2027-03-17', 100, 'SS', 'fixiert'),
        sa('sa-m-1ak-ss1', 'Mathematik: Geometrie', 'M', '1AK', 'WEI', '2027-03-18', 100, 'SS', 'fixiert'),
        sa('sa-f-4bk-ss1', 'Französisch: Production écrite', 'F', '4BK', 'LEH', '2027-04-20', 100, 'SS', 'fixiert'),
        sa('sa-sp-3bk-ss1', 'Spanisch: Comprensión', 'SP', '3BK', 'LEH', '2027-03-23', 100, 'SS', 'fixiert'),
        sa('sa-i-2bk-ss1', 'Italienisch: Grammatica', 'I', '2BK', 'KIN', '2027-04-21', 100, 'SS', 'fixiert'),
        sa('sa-wi-4bk-ss1', 'WI: Tabellenkalkulation Advanced', 'WI', '4BK', 'SCH', '2027-05-04', 75, 'SS', 'fixiert'),

        sa('sa-bwl-5ak-ss2', 'BWL: Investition und Finanzierung', 'BWL', '5AK', 'HOF', '2027-05-11', 100, 'SS', 'beantragt'),
        sa('sa-m-4ak-ss2', 'Mathematik: Folgen und Reihen', 'M', '4AK', 'WEI', '2027-05-12', 100, 'SS', 'beantragt'),
        sa('sa-d-2ak-ss2', 'Deutsch: Erörterung Medien', 'D', '2AK', 'KIN', '2027-05-19', 100, 'SS', 'beantragt'),
        sa('sa-e-4ak-ss2', 'Englisch: Business Correspondence', 'E', '4AK', 'MAY', '2027-06-02', 100, 'SS', 'beantragt')
    ];

    // Weitere Klassen systematisch auffüllen (1BK, 2BK, 4BK, 5BK …)
    const generated = generateExtraRows();
    return rows.concat(generated);
}

function teacherEmail(code) {
    const t = DEMO_STAMMDATEN.teachers.find((x) => x.code === code);
    return t ? t.email : '';
}

/**
 * @param {string} id
 * @param {string} title
 * @param {string} fach
 * @param {string} klasse
 * @param {string} lehrer
 * @param {string} datum
 * @param {number} dauer
 * @param {string} semester
 * @param {string} status
 * @param {string} [ablehnung]
 */
function sa(id, title, fach, klasse, lehrer, datum, dauer, semester, status, ablehnung) {
    const fixed = status === 'fixiert' || status === 'abgelehnt';
    return {
        Title: title,
        SchularbeitId: id,
        FachCode: fach,
        KlasseCode: klasse,
        LehrerCode: lehrer,
        LehrerEmail: teacherEmail(lehrer),
        Datum: datum,
        DauerMinuten: dauer,
        Semester: semester,
        Status: status,
        Notiz: DEMO_SEED_TAG + ' · SJ ' + DEMO_SCHOOL_YEAR,
        AblehnungsGrund: ablehnung || '',
        BeantragtVon: teacherEmail(lehrer),
        FixiertVon: fixed ? 'admin@kurtrocks.onmicrosoft.com' : '',
        FixiertAm: fixed ? datum + 'T10:00:00Z' : undefined
    };
}

function generateExtraRows() {
    /** @type {SchularbeitFields[]} */
    const out = [];
    const plan = [
        { klasse: '1BK', fach: 'D', lehrer: 'KIN', datum: '2026-11-05', sem: 'WS', status: 'fixiert', thema: 'Deutsch: Bericht' },
        { klasse: '1BK', fach: 'M', lehrer: 'WEI', datum: '2026-11-19', sem: 'WS', status: 'fixiert', thema: 'Mathematik: Terme' },
        { klasse: '1BK', fach: 'E', lehrer: 'MAY', datum: '2027-03-24', sem: 'SS', status: 'fixiert', thema: 'Englisch: Reading' },
        { klasse: '2BK', fach: 'D', lehrer: 'BAU', datum: '2026-10-23', sem: 'WS', status: 'fixiert', thema: 'Deutsch: Zusammenfassung' },
        { klasse: '2BK', fach: 'BWL', lehrer: 'HOF', datum: '2026-11-20', sem: 'WS', status: 'fixiert', thema: 'BWL: Märkte' },
        { klasse: '2BK', fach: 'M', lehrer: 'WEI', datum: '2027-03-25', sem: 'SS', status: 'beantragt', thema: 'Mathematik: Gleichungen' },
        { klasse: '4BK', fach: 'BWL', lehrer: 'HOF', datum: '2026-10-27', sem: 'WS', status: 'fixiert', thema: 'BWL: Organisation' },
        { klasse: '4BK', fach: 'E', lehrer: 'MAY', datum: '2026-11-26', sem: 'WS', status: 'fixiert', thema: 'Englisch: Reports' },
        { klasse: '5BK', fach: 'D', lehrer: 'BAU', datum: '2026-10-28', sem: 'WS', status: 'fixiert', thema: 'Deutsch: Erörterung' },
        { klasse: '5BK', fach: 'BWL', lehrer: 'HOF', datum: '2026-11-25', sem: 'WS', status: 'fixiert', thema: 'BWL: Strategie' },
        { klasse: '5BK', fach: 'AM', lehrer: 'SCH', datum: '2027-03-09', sem: 'SS', status: 'fixiert', thema: 'AM: Maturatraining' },
        { klasse: '5BK', fach: 'E', lehrer: 'MAY', datum: '2027-04-22', sem: 'SS', status: 'beantragt', thema: 'Englisch: Matura Mock' },
        { klasse: '3BK', fach: 'BWL', lehrer: 'HOF', datum: '2026-12-02', sem: 'WS', status: 'fixiert', thema: 'BWL: Rechnungswesen-Bezug' },
        { klasse: '3BK', fach: 'E', lehrer: 'MAY', datum: '2027-04-28', sem: 'SS', status: 'fixiert', thema: 'Englisch: Debating prep' },
        { klasse: '1AK', fach: 'BWL', lehrer: 'HOF', datum: '2027-03-10', sem: 'SS', status: 'fixiert', thema: 'BWL: Einführung Betrieb' },
        { klasse: '2AK', fach: 'UR', lehrer: 'GRA', datum: '2027-04-08', sem: 'SS', status: 'beantragt', thema: 'UR: Belege' }
    ];
    plan.forEach((p, i) => {
        out.push(
            sa(
                'sa-gen-' + (i + 1),
                p.thema,
                p.fach,
                p.klasse,
                p.lehrer,
                p.datum,
                p.fach === 'RW' || p.fach === 'WI' ? 75 : 100,
                p.sem,
                p.status
            )
        );
    });
    return out;
}

/**
 * Gesamtpaket für Seed.
 */
export function getDemoSeedPackage() {
    const schularbeiten = buildDemoSchularbeiten().map((row) => {
        const fields = { ...row };
        if (!fields.FixiertAm) delete fields.FixiertAm;
        if (!fields.AblehnungsGrund) fields.AblehnungsGrund = '';
        return fields;
    });
    return {
        schoolYear: DEMO_SCHOOL_YEAR,
        seedTag: DEMO_SEED_TAG,
        siteDefault: DEMO_SITE_DEFAULT,
        stammdaten: DEMO_STAMMDATEN,
        regelwerk: DEMO_REGELWERK,
        fachMeta: DEMO_FACHMETA,
        terminfenster: DEMO_TERMINFENSTER,
        schularbeiten,
        counts: {
            fachMeta: DEMO_FACHMETA.length,
            terminfenster: DEMO_TERMINFENSTER.length,
            schularbeiten: schularbeiten.length,
            teachers: DEMO_STAMMDATEN.teachers.length,
            classes: DEMO_STAMMDATEN.classes.length,
            subjects: DEMO_STAMMDATEN.subjects.length
        }
    };
}
