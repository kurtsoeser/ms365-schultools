/**
 * Lokale Demo-Daten für Freitag-Workshop (ohne SharePoint).
 */
import { buildItemTitle, buildPruefplanKurz } from './srdp-anmeldung-logic.js';

export const DEMO_SITE_DEFAULT = 'https://demo.sharepoint.com/sites/MS365-Schultools/';
/** Haupttermin der Demo – Wizard, Abschlussklassen und Seed-Zeilen müssen übereinstimmen. */
export const DEMO_TERMIN_JAHR = '2027';
export const DEMO_SEED_TAG = 'srdp-demo-' + DEMO_TERMIN_JAHR;

/** IT bleibt angemeldet; Persona filtert als Lehrkraft. */
export const DEMO_TEACHER_PERSONA = {
    email: 'mueller@demo.schule.at',
    name: 'Mag. Müller',
    label: 'Mag. Müller'
};

/**
 * Schüler-Persona der Demo: steht in den Stammdaten, hat aber bewusst noch keine
 * Seed-Anmeldung – sonst zeigt die Schüleransicht „bereits angemeldet“ statt des Formulars.
 */
export const DEMO_STUDENT_PERSONA = {
    email: 'lena.bauer@demo.schule.at',
    name: 'Lena Bauer',
    vorname: 'Lena',
    nachname: 'Bauer',
    klasse: '5AHAK'
};

function isoAdd(days) {
    const d = new Date();
    d.setHours(12, 0, 0, 0);
    d.setDate(d.getDate() + (Number(days) || 0));
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
}

export function buildDemoWizard() {
    return {
        profileId: 'hak',
        jahr: DEMO_TERMIN_JAHR,
        terminJahr: DEMO_TERMIN_JAHR,
        siteUrl: DEMO_SITE_DEFAULT,
        klassen: ['5AHAK', '5BHAK', '5CHAK'],
        lfsText: 'EN\nENWS\nFR\nFRWS\nIT\nITWS\nSP\nSPWS',
        wahlfachText:
            'Religion / Ethik\nKultur\nGeschichte und Internationale Wirtschafts- und Kulturräume\nGeografie und Internationale Wirtschafts- und Kulturräume\nNaturwissenschaften\nRecht\nVolkswirtschaft\nSeminar …',
        seminarText: 'DIGBIZ\nWirtschaftsrecht',
        anmeldeFrist: isoAdd(14),
        haupttermin: isoAdd(45)
    };
}

export function buildDemoStammdaten() {
    return {
        classes: [
            { code: '5AHAK', name: '5AHAK', year: DEMO_TERMIN_JAHR, headEmail: 'kv.ahak@demo.schule.at' },
            { code: '5BHAK', name: '5BHAK', year: DEMO_TERMIN_JAHR, headEmail: 'mueller@demo.schule.at' },
            { code: '5CHAK', name: '5CHAK', year: DEMO_TERMIN_JAHR, headEmail: 'kv.chak@demo.schule.at' }
        ],
        teachers: [
            { code: 'MUEL', name: 'Mag. Müller', email: 'mueller@demo.schule.at' },
            { code: 'SCHM', name: 'Mag. Schmidt', email: 'schmidt@demo.schule.at' },
            { code: 'HOFE', name: 'Mag. Hofer', email: 'hofer@demo.schule.at' },
            { code: 'WAGN', name: 'Mag. Wagner', email: 'wagner@demo.schule.at' },
            { code: 'LEHN', name: 'Mag. Lehner', email: 'lehner@demo.schule.at' },
            { code: 'KVA', name: 'Mag. Berger (KV 5A)', email: 'kv.ahak@demo.schule.at' }
        ],
        subjects: [{ code: 'ENWS' }, { code: 'FRWS' }, { code: 'ITWS' }],
        students: [
            {
                nachname: 'Huber',
                vorname: 'Anna',
                email: 'anna.huber@demo.schule.at',
                klasse: '5AHAK'
            },
            {
                nachname: DEMO_STUDENT_PERSONA.nachname,
                vorname: DEMO_STUDENT_PERSONA.vorname,
                email: DEMO_STUDENT_PERSONA.email,
                klasse: DEMO_STUDENT_PERSONA.klasse
            }
        ]
    };
}

function row(partial) {
    const nachname = partial.nachname;
    const vorname = partial.vorname;
    const variante = partial.variante || 'Variante 1';
    const email = String(partial.createdBy || partial.schuelerEmail || '').trim().toLowerCase();
    return {
        itemId: partial.itemId || 'demo-' + nachname.toLowerCase(),
        title: buildItemTitle(nachname, vorname),
        klasse: partial.klasse,
        nachname,
        vorname,
        titelDiplomarbeit: partial.titelDiplomarbeit || 'Digitale Geschäftsmodelle im Handel',
        betreuungslehrerDA: partial.betreuungslehrerDA || 'Mag. Müller',
        variante,
        pruefplanKurz: buildPruefplanKurz(variante, 'hak'),
        lfs: partial.lfs || 'ENWS',
        lehrerLFS: partial.lehrerLFS || '',
        lehrerBKO: partial.lehrerBKO || 'Mag. Hofer',
        lehrerAM: partial.lehrerAM || '',
        wahlfachMuendlich: partial.wahlfachMuendlich || 'Recht',
        seminar: partial.seminar || '',
        lehrerWahlfach: partial.lehrerWahlfach || 'Mag. Wagner',
        bestaetigung: true,
        nacharbeitErledigt: partial.nacharbeitErledigt === true,
        schulform: 'HAK',
        terminJahr: DEMO_TERMIN_JAHR,
        schuelerEmail: email,
        beantragtVon: email,
        createdBy: email,
        created: partial.created || isoAdd(partial.createdOffset || -3),
        seedTag: DEMO_SEED_TAG,
        webUrl: ''
    };
}

export function buildDemoItems() {
    return [
        row({
            itemId: 'demo-huber',
            nachname: 'Huber',
            vorname: 'Anna',
            klasse: '5AHAK',
            variante: 'Variante 1',
            wahlfachMuendlich: 'Recht',
            lehrerAM: 'Mag. Schmidt',
            lehrerBKO: 'Mag. Hofer; Mag. Wagner',
            createdBy: 'anna.huber@demo.schule.at',
            createdOffset: -1,
            nacharbeitErledigt: true
        }),
        row({
            itemId: 'demo-maier',
            nachname: 'Maier',
            vorname: 'Tom',
            klasse: '5AHAK',
            variante: 'Variante 2',
            lfs: 'FRWS',
            lehrerLFS: 'Mag. Lehner',
            wahlfachMuendlich: 'Seminar …',
            seminar: 'DIGBIZ',
            lehrerWahlfach: 'Mag. Müller',
            betreuungslehrerDA: 'Mag. Schmidt',
            createdBy: 'tom.maier@demo.schule.at',
            createdOffset: -2,
            nacharbeitErledigt: false
        }),
        row({
            itemId: 'demo-leitner',
            nachname: 'Leitner',
            vorname: 'Sara',
            klasse: '5BHAK',
            variante: 'Variante 3',
            lfs: 'ITWS',
            lehrerBKO: 'Mag. Wagner',
            wahlfachMuendlich: 'Naturwissenschaften',
            lehrerWahlfach: 'Mag. Hofer',
            betreuungslehrerDA: 'Mag. Lehner',
            createdBy: 'sara.leitner@demo.schule.at',
            createdOffset: 0,
            nacharbeitErledigt: false
        }),
        row({
            itemId: 'demo-gruber',
            nachname: 'Gruber',
            vorname: 'Jonas',
            klasse: '5BHAK',
            variante: 'Variante 1',
            lfs: 'SPWS',
            lehrerAM: 'Mag. Schmidt',
            wahlfachMuendlich: 'Kultur',
            betreuungslehrerDA: 'Mag. Müller',
            createdBy: 'jonas.gruber@demo.schule.at',
            createdOffset: -4,
            nacharbeitErledigt: true
        }),
        row({
            itemId: 'demo-pichler',
            nachname: 'Pichler',
            vorname: 'Lara',
            klasse: '5CHAK',
            variante: 'Variante 2',
            wahlfachMuendlich: 'Volkswirtschaft',
            lehrerWahlfach: 'Mag. Wagner',
            betreuungslehrerDA: 'Mag. Hofer',
            lehrerLFS: 'Mag. Müller',
            createdBy: 'lara.pichler@demo.schule.at',
            createdOffset: -5,
            nacharbeitErledigt: true
        }),
        row({
            itemId: 'demo-wagner',
            nachname: 'Wagner',
            vorname: 'Felix',
            klasse: '5CHAK',
            variante: 'Variante 1',
            wahlfachMuendlich: 'Geschichte und Internationale Wirtschafts- und Kulturräume',
            lehrerWahlfach: 'Mag. Lehner',
            betreuungslehrerDA: 'Mag. Müller',
            lehrerBKO: 'Mag. Schmidt',
            lehrerAM: 'Mag. Hofer',
            createdBy: 'felix.wagner@demo.schule.at',
            createdOffset: -6,
            nacharbeitErledigt: true
        })
    ];
}

/**
 * ABA-Portal-Excel-Zeilen (wie Export) – für Demo-XLSX und IT-Setup-Muster.
 * Spalten: Schule, Bundesland, Titel, Schüler/in, Klasse, Betreuer/in, …
 */
export function buildDemoAbaPortalExcelRows() {
    return [
        {
            Schule: 'Demo-HAK Linz',
            Bundesland: 'Oberösterreich',
            Titel: 'Nachhaltige Geschäftsmodelle im Einzelhandel',
            'Schüler/in': 'Lena Bauer',
            Klasse: '5AHAK - 2026/27',
            'Betreuer/in': 'Mag. Müller',
            Eingereicht: '10.09.2026',
            Aktualisiert: '12.09.2026',
            Abteilung: '',
            Status: 'Genehmigt durch DIR'
        },
        {
            Schule: 'Demo-HAK Linz',
            Bundesland: 'Oberösterreich',
            Titel: 'Digitale Geschäftsmodelle im Handel',
            'Schüler/in': 'Anna Huber',
            Klasse: '5AHAK - 2026/27',
            'Betreuer/in': 'Mag. Müller',
            Eingereicht: '11.09.2026',
            Aktualisiert: '11.09.2026',
            Abteilung: '',
            Status: 'Genehmigt durch DIR'
        },
        {
            Schule: 'Demo-HAK Linz',
            Bundesland: 'Oberösterreich',
            Titel: 'Buy now, pay later und Konsumfinanzierung',
            'Schüler/in': 'Tom Maier, Sara Leitner',
            Klasse: '5BHAK - 2026/27',
            'Betreuer/in': 'Mag. Schmidt',
            Eingereicht: '15.09.2026',
            Aktualisiert: '16.09.2026',
            Abteilung: '',
            Status: 'Genehmigt durch DIR'
        },
        {
            Schule: 'Demo-HAK Linz',
            Bundesland: 'Oberösterreich',
            Titel: 'Influencer und Finanzentscheidungen von Jugendlichen',
            'Schüler/in': 'Jonas Gruber, Lara Pichler',
            Klasse: '5CHAK - 2026/27',
            'Betreuer/in': 'Mag. Hofer',
            Eingereicht: '20.09.2026',
            Aktualisiert: '21.09.2026',
            Abteilung: '',
            Status: 'Genehmigt durch DIR'
        },
        {
            Schule: 'Demo-HAK Linz',
            Bundesland: 'Oberösterreich',
            Titel: 'Werteorientierte Personalführung in KMU',
            'Schüler/in': 'Felix Wagner',
            Klasse: '5CHAK - 2026/27',
            'Betreuer/in': 'Mag. Wagner',
            Eingereicht: '22.09.2026',
            Aktualisiert: '22.09.2026',
            Abteilung: '',
            Status: 'Genehmigt durch DIR'
        }
    ];
}

/** Demo-Zeilen ABA-Portal (explodiert) für Prefill – inkl. Schülerpersona Lena Bauer. */
export function buildDemoAbschlussarbeiten() {
    const out = [];
    buildDemoAbaPortalExcelRows().forEach((row) => {
        const names = String(row['Schüler/in'] || '')
            .split(/\s*(?:,|;|\bund\b)\s*/i)
            .map((s) => s.trim())
            .filter(Boolean);
        const klasseRoh = String(row.Klasse || '');
        const m = klasseRoh.match(/^(.+?)\s*[-–—]\s*(\d{4})\s*\/\s*(\d{2,4})\s*$/);
        const klasseCode = m ? m[1].trim() : klasseRoh;
        const start = m ? m[2] : '';
        let end = m ? m[3] : '';
        if (end.length === 2 && start) end = start.slice(0, 2) + end;
        names.forEach((schuelerName) => {
            out.push({
                titel: String(row.Titel || ''),
                schuelerName,
                klasseRoh,
                klasseCode,
                schuljahr: start && end ? start + '/' + String(end).slice(-2) : '',
                betreuer: String(row['Betreuer/in'] || ''),
                status: String(row.Status || ''),
                schule: String(row.Schule || ''),
                genehmigt: /genehmigt/i.test(String(row.Status || '')),
                terminJahrHint: end || DEMO_TERMIN_JAHR,
                quelle: 'ABA-Portal'
            });
        });
    });
    return out;
}

/**
 * @param {object} [base]
 */
export function getDemoSeedPackage() {
    return {
        seedTag: DEMO_SEED_TAG,
        wizard: buildDemoWizard(),
        stammdaten: buildDemoStammdaten(),
        items: buildDemoItems(),
        teacherPersona: DEMO_TEACHER_PERSONA,
        studentPersona: DEMO_STUDENT_PERSONA
    };
}

/**
 * @param {object} [base]
 */
export function buildLocalDemoState(base) {
    const pack = getDemoSeedPackage();
    const prev = base && typeof base === 'object' ? base : {};
    return Object.assign({}, prev, {
        localDemoOnly: true,
        siteUrl: DEMO_SITE_DEFAULT,
        listId: 'demo-list',
        siteId: 'demo-site',
        listWebUrl: DEMO_SITE_DEFAULT,
        wizard: pack.wizard,
        stammdaten: pack.stammdaten,
        items: pack.items,
        abschlussarbeiten: buildDemoAbschlussarbeiten(),
        abschlussarbeitHits: [],
        abschlussarbeitenMeta: {
            listTitle: 'Demo ABA-Portal',
            count: buildDemoAbschlussarbeiten().length,
            lastImportAt: new Date().toISOString()
        },
        // URL-Flag ?demoRole=1 nicht überschreiben – sonst fehlen die Rollen-Chips.
        demoRoleOverride: !!prev.demoRoleOverride,
        demoPersonaEmail: '',
        demoPersonaName: '',
        planerRoles: ['schueler', 'lehrer', 'admin'],
        error: '',
        softError: '',
        stammdatenLive: {
            ok: true,
            error: '',
            classCount: pack.stammdaten.classes.length,
            teacherCount: pack.stammdaten.teachers.length,
            missing: [],
            source: 'demo'
        },
        calMode: 'month',
        calWeekAnchor: isoAdd(0),
        calYear: new Date().getFullYear(),
        calMonth: new Date().getMonth() + 1
    });
}

/**
 * Nur Seed-Zeilen entfernen (Reset).
 * @param {object[]} items
 */
export function isDemoSeedItem(item) {
    if (!item) return false;
    if (String(item.seedTag || '') === DEMO_SEED_TAG) return true;
    const id = String(item.itemId || '');
    return id.startsWith('demo-');
}

export function stripDemoSeedItems(items) {
    return (Array.isArray(items) ? items : []).filter((it) => !isDemoSeedItem(it));
}
