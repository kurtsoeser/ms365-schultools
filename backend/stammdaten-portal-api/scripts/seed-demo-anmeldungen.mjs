/**
 * Legt Demo-Anmeldungen im Live-Gateway/SharePoint an (enroll → unlock → save/submit).
 *
 * Usage (PowerShell):
 *   $env:STAMMDATEN_ADMIN_KEY = (Get-Content .\.admin-key.local.txt -Raw).Trim()
 *   node scripts/seed-demo-anmeldungen.mjs
 */
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const API = (process.env.STAMMDATEN_API_BASE || 'https://func-ms365-ssp-dev.azurewebsites.net/api').replace(
    /\/$/,
    ''
);

function loadAdminKey() {
    if (process.env.STAMMDATEN_ADMIN_KEY) return String(process.env.STAMMDATEN_ADMIN_KEY).trim();
    try {
        return readFileSync(join(__dirname, '../.admin-key.local.txt'), 'utf8').trim();
    } catch {
        throw new Error('STAMMDATEN_ADMIN_KEY fehlt (Env oder .admin-key.local.txt).');
    }
}

function makeValidSvnr(ddmmjj) {
    const weights = [3, 7, 9, 5, 8, 4, 2, 1, 6];
    const birth = String(ddmmjj).padStart(6, '0').slice(0, 6);
    for (let serial = 100; serial < 999; serial += 1) {
        const head = String(serial).padStart(3, '0');
        const nine = [...head, ...birth].map(Number);
        let sum = 0;
        for (let i = 0; i < 9; i += 1) sum += nine[i] * weights[i];
        const rem = sum % 11;
        if (rem === 10) continue;
        return `${head}${rem}${birth}`;
    }
    throw new Error(`SVNR für ${ddmmjj} nicht gefunden`);
}

/** Geburtsdatum ISO → TTMMJJ für SVNR */
function isoToSvnrDate(iso) {
    const [y, m, d] = iso.split('-');
    return `${d}${m}${y.slice(2)}`;
}

const DEMOS = [
    {
        familienname: 'Mustermann',
        vorname: 'Max',
        geschlecht: 'männlich',
        geburtsdatum: '2012-05-15',
        geburtsort: 'Steyr',
        religion: 'röm.-kath.',
        strasse: 'Ennskai',
        hausnummer: '12',
        plz: '4400',
        ort: 'Steyr',
        mutter: { vorname: 'Maria', familienname: 'Mustermann', mail: 'maria.mustermann@example.com', mobil: '+436641001001' },
        vater: { vorname: 'Thomas', familienname: 'Mustermann', mail: 'thomas.mustermann@example.com', mobil: '+436641001002' }
    },
    {
        familienname: 'Huber',
        vorname: 'Anna',
        geschlecht: 'weiblich',
        geburtsdatum: '2011-03-03',
        geburtsort: 'Linz',
        religion: 'röm.-kath.',
        strasse: 'Hauptplatz',
        hausnummer: '4',
        plz: '4020',
        ort: 'Linz',
        mutter: { vorname: 'Eva', familienname: 'Huber', mail: 'eva.huber@example.com', mobil: '+436641002001' },
        vater: { vorname: 'Josef', familienname: 'Huber', mail: 'josef.huber@example.com', mobil: '+436641002002' }
    },
    {
        familienname: 'Gruber',
        vorname: 'Lukas',
        geschlecht: 'männlich',
        geburtsdatum: '2014-10-22',
        geburtsort: 'Wels',
        religion: 'evang. A.B.',
        strasse: 'Ringstraße',
        hausnummer: '8a',
        plz: '4600',
        ort: 'Wels',
        mutter: { vorname: 'Sabine', familienname: 'Gruber', mail: 'sabine.gruber@example.com', mobil: '+436641003001' },
        vater: { vorname: 'Michael', familienname: 'Gruber', mail: 'michael.gruber@example.com', mobil: '+436641003002' }
    },
    {
        familienname: 'Wagner',
        vorname: 'Sophie',
        geschlecht: 'weiblich',
        geburtsdatum: '2015-09-08',
        geburtsort: 'Steyr',
        religion: 'o.B.',
        strasse: 'Blumauergasse',
        hausnummer: '3',
        plz: '4400',
        ort: 'Steyr',
        mutter: { vorname: 'Nina', familienname: 'Wagner', mail: 'nina.wagner@example.com', mobil: '+436641004001' },
        vater: { vorname: 'Andreas', familienname: 'Wagner', mail: 'andreas.wagner@example.com', mobil: '+436641004002' }
    },
    {
        familienname: 'Leitner',
        vorname: 'Jonas',
        geschlecht: 'männlich',
        geburtsdatum: '2012-12-19',
        geburtsort: 'Enns',
        religion: 'röm.-kath.',
        strasse: 'Wiener Straße',
        hausnummer: '21',
        plz: '4470',
        ort: 'Enns',
        mutter: { vorname: 'Christina', familienname: 'Leitner', mail: 'christina.leitner@example.com', mobil: '+436641005001' },
        vater: { vorname: 'Peter', familienname: 'Leitner', mail: 'peter.leitner@example.com', mobil: '+436641005002' }
    },
    {
        familienname: 'Bauer',
        vorname: 'Emma',
        geschlecht: 'weiblich',
        geburtsdatum: '2010-06-27',
        geburtsort: 'Linz',
        religion: 'islam. (IGGÖ)',
        strasse: 'Landstraße',
        hausnummer: '55',
        plz: '4020',
        ort: 'Linz',
        mutter: { vorname: 'Ayse', familienname: 'Bauer', mail: 'ayse.bauer@example.com', mobil: '+436641006001' },
        vater: { vorname: 'Mehmet', familienname: 'Bauer', mail: 'mehmet.bauer@example.com', mobil: '+436641006002' }
    },
    {
        familienname: 'Steiner',
        vorname: 'Noah',
        geschlecht: 'männlich',
        geburtsdatum: '2013-08-11',
        geburtsort: 'Steyr',
        religion: 'röm.-kath.',
        strasse: 'Dürerstraße',
        hausnummer: '7',
        plz: '4400',
        ort: 'Steyr',
        mutter: { vorname: 'Lisa', familienname: 'Steiner', mail: 'lisa.steiner@example.com', mobil: '+436641007001' },
        vater: { vorname: 'Daniel', familienname: 'Steiner', mail: 'daniel.steiner@example.com', mobil: '+436641007002' }
    },
    {
        familienname: 'Mayer',
        vorname: 'Lara',
        geschlecht: 'weiblich',
        geburtsdatum: '2014-04-05',
        geburtsort: 'Garsten',
        religion: 'röm.-kath.',
        strasse: 'Kirchenplatz',
        hausnummer: '2',
        plz: '4451',
        ort: 'Garsten',
        mutter: { vorname: 'Monika', familienname: 'Mayer', mail: 'monika.mayer@example.com', mobil: '+436641008001' },
        vater: { vorname: 'Franz', familienname: 'Mayer', mail: 'franz.mayer@example.com', mobil: '+436641008002' }
    },
    {
        familienname: 'Pichler',
        vorname: 'Tim',
        geschlecht: 'männlich',
        geburtsdatum: '2015-09-30',
        geburtsort: 'Sierning',
        religion: 'evang. A.B.',
        strasse: 'Schulstraße',
        hausnummer: '9',
        plz: '4522',
        ort: 'Sierning',
        mutter: { vorname: 'Claudia', familienname: 'Pichler', mail: 'claudia.pichler@example.com', mobil: '+436641009001' },
        vater: { vorname: 'Martin', familienname: 'Pichler', mail: 'martin.pichler@example.com', mobil: '+436641009002' }
    },
    {
        familienname: 'Hofer',
        vorname: 'Mia',
        geschlecht: 'weiblich',
        geburtsdatum: '2011-07-16',
        geburtsort: 'Steyr',
        religion: 'o.B.',
        strasse: 'Stadtplatz',
        hausnummer: '18',
        plz: '4400',
        ort: 'Steyr',
        mutter: { vorname: 'Julia', familienname: 'Hofer', mail: 'julia.hofer@example.com', mobil: '+436641010001' },
        vater: { vorname: 'Stefan', familienname: 'Hofer', mail: 'stefan.hofer@example.com', mobil: '+436641010002' }
    }
];

function buildPayload(d) {
    const svnr = makeValidSvnr(isoToSvnrDate(d.geburtsdatum));
    return {
        schueler: {
            familienname: d.familienname,
            vorname: d.vorname,
            weitereVornamen: '',
            geschlecht: d.geschlecht,
            geburtsdatum: d.geburtsdatum,
            geburtsort: d.geburtsort,
            geburtsstaat: 'Österreich',
            staatsbuergerschaft: 'Österreich',
            religion: d.religion,
            alltagssprache: 'Deutsch',
            erstsprache: 'Deutsch',
            svnr,
            mail: `${d.vorname.toLowerCase()}.${d.familienname.toLowerCase()}@schueler.example.at`,
            mobil: ''
        },
        wohnadresse: {
            land: 'A',
            plz: d.plz,
            ort: d.ort,
            gemeindenummer: '',
            strasse: d.strasse,
            hausnummer: d.hausnummer
        },
        erziehungsberechtigte: [
            {
                adressart: 'Mutter',
                anrede: 'Frau',
                vorname: d.mutter.vorname,
                familienname: d.mutter.familienname,
                mobil: d.mutter.mobil,
                mail: d.mutter.mail,
                erziehungsberechtigt: 'ja',
                zahlungspflichtig: 'ja',
                entscheideAn: 'ja',
                land: 'A',
                plz: d.plz,
                ort: d.ort,
                strasse: d.strasse,
                hausnummer: d.hausnummer
            },
            {
                adressart: 'Vater',
                anrede: 'Herr',
                vorname: d.vater.vorname,
                familienname: d.vater.familienname,
                mobil: d.vater.mobil,
                mail: d.vater.mail,
                erziehungsberechtigt: 'ja',
                zahlungspflichtig: 'ja',
                entscheideAn: 'ja',
                land: 'A',
                plz: d.plz,
                ort: d.ort,
                strasse: d.strasse,
                hausnummer: d.hausnummer
            }
        ],
        sameAddressAsStudent: true
    };
}

async function jsonFetch(url, opts = {}) {
    const res = await fetch(url, opts);
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || data.message || `HTTP ${res.status} ${url}`);
    return data;
}

async function main() {
    const adminKey = loadAdminKey();
    const rows = [];
    console.log(`API: ${API}`);
    console.log(`Lege ${DEMOS.length} Demo-Anmeldungen an …\n`);

    for (const d of DEMOS) {
        const title = `${d.familienname}, ${d.vorname}`;
        const payload = buildPayload(d);
        const enrolled = await jsonFetch(`${API}/ssp-enroll`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'x-admin-key': adminKey },
            body: JSON.stringify({ title, payload })
        });
        const unlocked = await jsonFetch(`${API}/unlock`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: enrolled.id, pin: enrolled.pin })
        });
        const saved = await jsonFetch(`${API}/save`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                Authorization: `Bearer ${unlocked.token}`
            },
            body: JSON.stringify({ payload, submit: true })
        });
        const formUrl = `http://127.0.0.1:5188/tools/schueler-stammdaten-formular.html?id=${enrolled.id}&schule=ms365schule&mode=remote&api=${encodeURIComponent(API)}`;
        rows.push({
            name: title,
            id: enrolled.id,
            pin: enrolled.pin,
            status: saved.record?.status || '?',
            svnr: payload.schueler.svnr,
            formUrl
        });
        console.log(`✓ ${title}  PIN ${enrolled.pin}  ${saved.record?.status}`);
    }

    console.log('\n--- Demo-Übersicht (PIN getrennt vom Link) ---\n');
    console.log('| Name | PIN | GUID |');
    console.log('|---|---|---|');
    for (const r of rows) {
        console.log(`| ${r.name} | \`${r.pin}\` | \`${r.id}\` |`);
    }
    console.log('\nPINs und Links nicht committen. Seed-Skript: scripts/seed-demo-anmeldungen.mjs');
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
