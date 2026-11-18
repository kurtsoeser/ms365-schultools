/**
 * Lokale Demo-Daten für Freitag-Workshop (ohne SharePoint).
 */
import { buildItemTitle, buildPruefplanKurz } from './srdp-anmeldung-logic.js';

export const DEMO_SITE_DEFAULT = 'https://demo.sharepoint.com/sites/MS365-Schultools/';
export const DEMO_SEED_TAG = 'srdp-demo-2026';

/** IT bleibt angemeldet; Persona filtert als Lehrkraft. */
export const DEMO_TEACHER_PERSONA = {
    email: 'mueller@demo.schule.at',
    name: 'Mag. Müller',
    label: 'Mag. Müller'
};

export const DEMO_STUDENT_PERSONA = {
    email: 'anna.huber@demo.schule.at',
    name: 'Anna Huber'
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
        jahr: '2026',
        terminJahr: '2026',
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
            { code: '5AHAK', name: '5AHAK', year: '2026', headEmail: 'kv.ahak@demo.schule.at' },
            { code: '5BHAK', name: '5BHAK', year: '2026', headEmail: 'mueller@demo.schule.at' },
            { code: '5CHAK', name: '5CHAK', year: '2026', headEmail: 'kv.chak@demo.schule.at' }
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
        lehrerLFS: partial.lehrerLFS || 'Mag. Schmidt',
        lehrerBKO: partial.lehrerBKO || 'Mag. Hofer',
        wahlfachMuendlich: partial.wahlfachMuendlich || 'Recht',
        seminar: partial.seminar || '',
        lehrerWahlfach: partial.lehrerWahlfach || 'Mag. Wagner',
        bestaetigung: true,
        nacharbeitErledigt: partial.nacharbeitErledigt === true,
        schulform: 'HAK',
        terminJahr: '2026',
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
            lehrerLFS: 'Mag. Müller',
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
            lehrerLFS: 'Mag. Schmidt',
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
            createdBy: 'felix.wagner@demo.schule.at',
            createdOffset: -6,
            nacharbeitErledigt: true
        })
    ];
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
    return Object.assign({}, base || {}, {
        localDemoOnly: true,
        siteUrl: DEMO_SITE_DEFAULT,
        listId: 'demo-list',
        siteId: 'demo-site',
        listWebUrl: DEMO_SITE_DEFAULT,
        wizard: pack.wizard,
        stammdaten: pack.stammdaten,
        items: pack.items,
        demoRoleOverride: false,
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
