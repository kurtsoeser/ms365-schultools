/**
 * Demo-Paket Freistellungen – Schuljahr 2026/27 (HAK Kurtrocks).
 * Lokal + SharePoint-Seed (Tag zum gezielten Zurücksetzen).
 */

import { approvalPath, inclusiveDayCount, toIsoDateOnly } from './freistellung-planer-logic.js';

export const DEMO_SITE_DEFAULT = 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools';
export const DEMO_SCHOOL_YEAR = '2026/27';
export const DEMO_SEED_TAG = 'fr-demo-2026-27';

/** Stammdaten kompatibel zu Schulaktivitäten-Demo; Klassen mit KV-E-Mail. */
export const DEMO_STAMMDATEN = {
    schoolName: 'Demo-HAK Kurtrocks',
    domain: 'kurtrocks.onmicrosoft.com',
    classes: [
        {
            code: '1AK',
            name: '1AK',
            year: 1,
            headName: 'Mag. Sarah Binder',
            headEmail: 'sarah.binder@kurtrocks.onmicrosoft.com'
        },
        {
            code: '1BK',
            name: '1BK',
            year: 1,
            headName: 'Mag. Christoph Lehmann',
            headEmail: 'christoph.lehmann@kurtrocks.onmicrosoft.com'
        },
        {
            code: '2AK',
            name: '2AK',
            year: 2,
            headName: 'Mag. Thomas Hofmann',
            headEmail: 'thomas.hofmann@kurtrocks.onmicrosoft.com'
        },
        {
            code: '2BK',
            name: '2BK',
            year: 2,
            headName: 'Mag. Eva Weiss',
            headEmail: 'eva.weiss@kurtrocks.onmicrosoft.com'
        },
        {
            code: '3AK',
            name: '3AK',
            year: 3,
            headName: 'Mag. Anna Mayer',
            headEmail: 'anna.mayer@kurtrocks.onmicrosoft.com'
        },
        {
            code: '3BK',
            name: '3BK',
            year: 3,
            headName: 'Mag. Peter Gruber',
            headEmail: 'peter.gruber@kurtrocks.onmicrosoft.com'
        },
        {
            code: '4AK',
            name: '4AK',
            year: 4,
            headName: 'Mag. Julia Bauer',
            headEmail: 'julia.bauer@kurtrocks.onmicrosoft.com'
        },
        {
            code: '4BK',
            name: '4BK',
            year: 4,
            headName: 'Dipl.-Ing. Markus Schuster',
            headEmail: 'markus.schuster@kurtrocks.onmicrosoft.com'
        },
        {
            code: '5AK',
            name: '5AK',
            year: 5,
            headName: 'Mag. Sarah Binder',
            headEmail: 'sarah.binder@kurtrocks.onmicrosoft.com'
        },
        {
            code: '5BK',
            name: '5BK',
            year: 5,
            headName: 'Mag. Thomas Hofmann',
            headEmail: 'thomas.hofmann@kurtrocks.onmicrosoft.com'
        }
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

function kvFor(klasse) {
    const c = DEMO_STAMMDATEN.classes.find((x) => x.code === klasse);
    return {
        email: c ? String(c.headEmail || '').toLowerCase() : '',
        name: c ? c.headName || '' : ''
    };
}

/**
 * @param {object} p
 * @returns {object} SharePoint fields (+ _demoId / _kvEmail für Seed)
 */
function fr(p) {
    const kv = kvFor(p.klasse);
    const id = p.id;
    const beschreibung =
        String(p.beschreibung || '').trim() +
        ' · ' +
        DEMO_SEED_TAG +
        ' · id:' +
        id +
        ' · SJ ' +
        DEMO_SCHOOL_YEAR;
    return {
        Title: p.name + ' (' + p.klasse + ')',
        Beginn: p.beginn,
        Ende: p.ende || p.beginn,
        Status: p.status || 'Ausstehend',
        Klasse: p.klasse,
        Kategorie: p.kategorie || 'Sonstiges',
        Beschreibung: beschreibung,
        Bemerkungen: p.bemerkungen || '',
        _demoId: id,
        _kvEmail: kv.email,
        _kvName: kv.name,
        _authorEmail: p.author || String(p.name || 'schueler').toLowerCase().replace(/\s+/g, '.') + '@kurtrocks.onmicrosoft.com'
    };
}

/**
 * 25 Freistellungen über SJ 2026/27 (mix Status, 1-Tag / mehrtägig, Kategorien).
 * @returns {object[]}
 */
export function buildDemoFreistellungen() {
    return [
        // ——— Herbst 2026 ———
        fr({
            id: 'fr-demo-01',
            name: 'Lena Berger',
            klasse: '1AK',
            beginn: '2026-09-18',
            status: 'Genehmigt',
            kategorie: 'Ärztlicher Termin',
            beschreibung: 'Zahnarztkontrolle vormittags',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-02',
            name: 'Jonas Hofer',
            klasse: '2AK',
            beginn: '2026-09-22',
            ende: '2026-09-24',
            status: 'Genehmigt',
            kategorie: 'Familiäre Angelegenheit',
            beschreibung: 'Familienfeier im Ausland',
            bemerkungen: 'KV + Direktion: genehmigt'
        }),
        fr({
            id: 'fr-demo-03',
            name: 'Mia Steiner',
            klasse: '3AK',
            beginn: '2026-09-29',
            status: 'Abgelehnt',
            kategorie: 'Sonstiges',
            beschreibung: 'Privater Kurzurlaub',
            bemerkungen: 'Direktion: nicht genehmigt – Unterrichtszeit'
        }),
        fr({
            id: 'fr-demo-04',
            name: 'Noah Wagner',
            klasse: '4AK',
            beginn: '2026-10-06',
            status: 'Genehmigt',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Schnuppertag Steuerberatung Mayer',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-05',
            name: 'Emma Leitner',
            klasse: '5AK',
            beginn: '2026-10-12',
            ende: '2026-10-14',
            status: 'Genehmigt',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Aufnahmeverfahren FH',
            bemerkungen: 'KV + Direktion: ok'
        }),
        fr({
            id: 'fr-demo-06',
            name: 'Luca Bauer',
            klasse: '1BK',
            beginn: '2026-10-20',
            status: 'Ausstehend',
            kategorie: 'Ärztlicher Termin',
            beschreibung: 'Orthopädie-Kontrolle'
        }),
        fr({
            id: 'fr-demo-07',
            name: 'Sophie Gruber',
            klasse: '2BK',
            beginn: '2026-10-27',
            ende: '2026-10-28',
            status: 'Ausstehend',
            kategorie: 'Familiäre Angelegenheit',
            beschreibung: 'Verwandtenbesuch Beerdigung'
        }),
        fr({
            id: 'fr-demo-08',
            name: 'Paul Moser',
            klasse: '3BK',
            beginn: '2026-11-05',
            status: 'Genehmigt',
            kategorie: 'Ärztlicher Termin',
            beschreibung: 'Augenarzt',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-09',
            name: 'Anna Winkler',
            klasse: '4BK',
            beginn: '2026-11-12',
            ende: '2026-11-13',
            status: 'Genehmigt',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Assessment Center Bank',
            bemerkungen: 'KV + Direktion: genehmigt'
        }),
        fr({
            id: 'fr-demo-10',
            name: 'Felix Haas',
            klasse: '5BK',
            beginn: '2026-11-19',
            status: 'Abgelehnt',
            kategorie: 'Sonstiges',
            beschreibung: 'Konzertbesuch',
            bemerkungen: 'KV: abgelehnt'
        }),
        // ——— Winter 2026/27 ———
        fr({
            id: 'fr-demo-11',
            name: 'Laura Fischer',
            klasse: '1AK',
            beginn: '2026-12-03',
            status: 'Genehmigt',
            kategorie: 'Ärztlicher Termin',
            beschreibung: 'Kieferorthopädie',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-12',
            name: 'David Huber',
            klasse: '2AK',
            beginn: '2026-12-10',
            ende: '2026-12-11',
            status: 'Ausstehend',
            kategorie: 'Familiäre Angelegenheit',
            beschreibung: 'Hochzeit der Schwester'
        }),
        fr({
            id: 'fr-demo-13',
            name: 'Julia Auer',
            klasse: '3AK',
            beginn: '2027-01-14',
            status: 'Genehmigt',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Schnuppertag Magazin Verlag',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-14',
            name: 'Tobias Eder',
            klasse: '4AK',
            beginn: '2027-01-20',
            ende: '2027-01-22',
            status: 'Ausstehend',
            kategorie: 'Ärztlicher Termin',
            beschreibung: 'Stationäre Untersuchung Spital'
        }),
        fr({
            id: 'fr-demo-15',
            name: 'Sarah König',
            klasse: '5AK',
            beginn: '2027-01-28',
            status: 'Genehmigt',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Vorstellungsgespräch Wien',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-16',
            name: 'Max Reiter',
            klasse: '1BK',
            beginn: '2027-02-04',
            status: 'Ausstehend',
            kategorie: 'Sonstiges',
            beschreibung: 'Führerscheinprüfung Theorie'
        }),
        fr({
            id: 'fr-demo-17',
            name: 'Nina Schmid',
            klasse: '2BK',
            beginn: '2027-02-11',
            ende: '2027-02-12',
            status: 'Genehmigt',
            kategorie: 'Familiäre Angelegenheit',
            beschreibung: 'Pflege Angehörige',
            bemerkungen: 'KV + Direktion: genehmigt'
        }),
        fr({
            id: 'fr-demo-18',
            name: 'Simon Maier',
            klasse: '3BK',
            beginn: '2027-02-18',
            status: 'Abgelehnt',
            kategorie: 'Sonstiges',
            beschreibung: 'Skiurlaub Familie',
            bemerkungen: 'Direktion: abgelehnt – Schularbeitenwoche'
        }),
        // ——— Frühjahr / Sommer 2027 ———
        fr({
            id: 'fr-demo-19',
            name: 'Lisa Pichler',
            klasse: '4BK',
            beginn: '2027-03-04',
            status: 'Genehmigt',
            kategorie: 'Ärztlicher Termin',
            beschreibung: 'Physiotherapie-Block',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-20',
            name: 'Benjamin Ortner',
            klasse: '5BK',
            beginn: '2027-03-11',
            ende: '2027-03-13',
            status: 'Ausstehend',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Matura-Vorbereitungskurs auswärts'
        }),
        fr({
            id: 'fr-demo-21',
            name: 'Klara Fuchs',
            klasse: '1AK',
            beginn: '2027-03-25',
            status: 'Genehmigt',
            kategorie: 'Familiäre Angelegenheit',
            beschreibung: 'Behördentermin mit Eltern',
            bemerkungen: 'KV: ok'
        }),
        fr({
            id: 'fr-demo-22',
            name: 'Elias Brandstetter',
            klasse: '2AK',
            beginn: '2027-04-08',
            status: 'Ausstehend',
            kategorie: 'Ärztlicher Termin',
            beschreibung: 'MRT-Termin'
        }),
        fr({
            id: 'fr-demo-23',
            name: 'Hannah Wallner',
            klasse: '3AK',
            beginn: '2027-04-15',
            ende: '2027-04-16',
            status: 'Genehmigt',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Lehrlingsmesse + Bewerbungsgespräch',
            bemerkungen: 'KV + Direktion: ok'
        }),
        fr({
            id: 'fr-demo-24',
            name: 'Marcel Steiner',
            klasse: '4AK',
            beginn: '2027-05-06',
            status: 'Ausstehend',
            kategorie: 'Sonstiges',
            beschreibung: 'Musikwettbewerb Landesfinale'
        }),
        fr({
            id: 'fr-demo-25',
            name: 'Victoria Lang',
            klasse: '5AK',
            beginn: '2027-05-20',
            ende: '2027-05-21',
            status: 'Genehmigt',
            kategorie: 'Familiäre Angelegenheit',
            beschreibung: 'Abschlussfeier Geschwister',
            bemerkungen: 'KV + Direktion: genehmigt'
        }),
        fr({
            id: 'fr-demo-26',
            name: 'Fabian Moser',
            klasse: '5BK',
            beginn: '2027-06-03',
            status: 'Ausstehend',
            kategorie: 'Bewerbung / Schnuppertag',
            beschreibung: 'Jobinterview IT-Firma Linz'
        })
    ];
}

/**
 * Demo-ID aus Beschreibung extrahieren.
 * @param {string} beschreibung
 */
export function extractDemoId(beschreibung) {
    const m = /id:(fr-(?:demo|test)-\d+)/i.exec(String(beschreibung || ''));
    return m ? m[1] : '';
}

/**
 * @param {object} fields
 * @param {string} [seedTag]
 */
export function isDemoFreistellungFields(fields, seedTag) {
    const tag = String(seedTag || DEMO_SEED_TAG);
    const text = String((fields && fields.Beschreibung) || '') + ' ' + String((fields && fields.Bemerkungen) || '');
    return text.indexOf(tag) !== -1 || !!extractDemoId(fields && fields.Beschreibung);
}

/**
 * SharePoint-Felder ohne interne Hilfskeys.
 * @param {object} row
 */
export function toSharePointFields(row) {
    const fields = {
        Title: row.Title,
        Beginn: row.Beginn,
        Ende: row.Ende,
        Status: row.Status,
        Klasse: row.Klasse,
        Kategorie: row.Kategorie,
        Beschreibung: row.Beschreibung,
        Bemerkungen: row.Bemerkungen || ''
    };
    if (row.GenehmigtVonKV) fields.GenehmigtVonKV = row.GenehmigtVonKV;
    if (row.GenehmigtAmKV) fields.GenehmigtAmKV = row.GenehmigtAmKV;
    if (row.GenehmigtVonDirektion) fields.GenehmigtVonDirektion = row.GenehmigtVonDirektion;
    if (row.GenehmigtAmDirektion) fields.GenehmigtAmDirektion = row.GenehmigtAmDirektion;
    if (row.AbgelehntVon) fields.AbgelehntVon = row.AbgelehntVon;
    if (row.AbgelehntAm) fields.AbgelehntAm = row.AbgelehntAm;
    if (row.KlassenvorstandLookupId) {
        fields.KlassenvorstandLookupId = String(row.KlassenvorstandLookupId);
    }
    return fields;
}

/**
 * UI-Items aus Demo-Feldern (lokal).
 * @param {object} [opts]
 */
/**
 * Planer-Items aus importiertem Demo-/Test-Paket (JSON).
 * @param {object} pack
 * @param {object} [opts]
 */
export function itemsFromDemoPack(pack, opts) {
    const o = opts || {};
    const rows = (pack && pack.freistellungen) || [];
    return rows.map((row, idx) => {
        const beginn = toIsoDateOnly(row.Beginn) || '';
        const ende = toIsoDateOnly(row.Ende) || beginn;
        const path = approvalPath(beginn, ende);
        const nameMatch = /^(.+?)\s*\(/.exec(row.Title || '');
        return {
            itemId: 'local-' + (row._demoId || 'row-' + idx),
            titel: row.Title,
            schuelerName: nameMatch ? nameMatch[1].trim() : row.Title,
            klasse: row.Klasse,
            beginn,
            ende,
            status: row.Status || 'Ausstehend',
            kategorie: row.Kategorie,
            beschreibung: row.Beschreibung,
            bemerkungen: row.Bemerkungen || '',
            kvEmail: row._kvEmail || '',
            kvName: row._kvName || '',
            authorEmail: row._authorEmail || o.accountEmail || '',
            beantragtVon: row._authorEmail || o.accountEmail || '',
            dayCount: inclusiveDayCount(beginn, ende),
            multiDay: path.multiDay,
            approvalLabel: path.label,
            demoId: row._demoId || extractDemoId(row.Beschreibung),
            genehmigtVonKv: row.GenehmigtVonKV || '',
            genehmigtAmKv: toIsoDateOnly(row.GenehmigtAmKV) || '',
            genehmigtVonDirektion: row.GenehmigtVonDirektion || '',
            genehmigtAmDirektion: toIsoDateOnly(row.GenehmigtAmDirektion) || '',
            abgelehntVon: row.AbgelehntVon || '',
            abgelehntAm: toIsoDateOnly(row.AbgelehntAm) || ''
        };
    });
}

export function buildLocalDemoItems(opts) {
    const o = opts || {};
    const accountEmail = String(o.accountEmail || '').toLowerCase();
    const accountName = String(o.accountName || '').trim();
    const rows = buildDemoFreistellungen();

    // Optional: ersten offenen Antrag auf aktuellen User mappen (Demo-Rolle Schüler)
    if (accountEmail && accountName) {
        const firstOpen = rows.find((r) => r.Status === 'Ausstehend');
        if (firstOpen) {
            firstOpen.Title = accountName + ' (' + firstOpen.Klasse + ')';
            firstOpen._authorEmail = accountEmail;
            firstOpen.Beschreibung = firstOpen.Beschreibung.replace(
                /^[^·]+/,
                'Persönlicher Demo-Antrag '
            );
        }
    }

    return rows.map((row, idx) => {
        const beginn = row.Beginn;
        const ende = row.Ende || beginn;
        const path = approvalPath(beginn, ende);
        const nameMatch = /^(.+?)\s*\(/.exec(row.Title || '');
        return {
            itemId: 'local-' + (row._demoId || idx),
            titel: row.Title,
            schuelerName: nameMatch ? nameMatch[1].trim() : row.Title,
            klasse: row.Klasse,
            beginn,
            ende,
            status: row.Status,
            kategorie: row.Kategorie,
            beschreibung: row.Beschreibung,
            bemerkungen: row.Bemerkungen || '',
            kvEmail: row._kvEmail || '',
            kvName: row._kvName || '',
            authorEmail: row._authorEmail || '',
            beantragtVon: row._authorEmail || '',
            dayCount: inclusiveDayCount(beginn, ende),
            multiDay: path.multiDay,
            approvalLabel: path.label,
            demoId: row._demoId || ''
        };
    });
}

export function getDemoSeedPackage() {
    const freistellungen = buildDemoFreistellungen();
    return {
        version: 1,
        schoolYear: DEMO_SCHOOL_YEAR,
        seedTag: DEMO_SEED_TAG,
        siteUrl: DEMO_SITE_DEFAULT,
        stammdaten: DEMO_STAMMDATEN,
        freistellungen,
        counts: {
            gesamt: freistellungen.length,
            ausstehend: freistellungen.filter((r) => r.Status === 'Ausstehend').length,
            genehmigt: freistellungen.filter((r) => r.Status === 'Genehmigt').length,
            abgelehnt: freistellungen.filter((r) => r.Status === 'Abgelehnt').length
        }
    };
}

/**
 * @param {string|object} raw
 */
export function parseDemoImportJson(raw) {
    let data = raw;
    if (typeof raw === 'string') {
        try {
            data = JSON.parse(raw);
        } catch {
            throw new Error('Keine gültige JSON-Datei.');
        }
    }
    if (!data || typeof data !== 'object') throw new Error('Leeres Demo-Paket.');
    if (!Array.isArray(data.freistellungen) || !data.freistellungen.length) {
        throw new Error('Feld „freistellungen“ fehlt oder ist leer.');
    }
    if (!data.stammdaten) data.stammdaten = DEMO_STAMMDATEN;
    if (!data.seedTag) data.seedTag = DEMO_SEED_TAG;
    if (!data.schoolYear) data.schoolYear = DEMO_SCHOOL_YEAR;
    return data;
}
