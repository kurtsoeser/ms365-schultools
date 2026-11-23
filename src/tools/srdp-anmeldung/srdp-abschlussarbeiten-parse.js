/**
 * ABA-Portal Excel → normalisierte Abschlussarbeiten (Diplomarbeit / ABA/VWA).
 */

const TITLE_KEYS = ['titel', 'title', 'thema', 'arbeitstitel'];
const STUDENT_KEYS = [
    'schueler/in',
    'schuler/in',
    'schueler',
    'schuler',
    'schüler/in',
    'schüler',
    'student',
    'name'
];
const CLASS_KEYS = ['klasse', 'class', 'kl.'];
const SUPERVISOR_KEYS = ['betreuer/in', 'betreuer', 'betreuung', 'betreuungslehrer', 'supervisor'];
const STATUS_KEYS = ['status', 'zustand'];
const SCHOOL_KEYS = ['schule', 'school'];

/**
 * @param {string} s
 */
export function normalizePersonKey(s) {
    return String(s || '')
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/ß/g, 'ss')
        .replace(/[^a-z0-9]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * @param {string} header
 */
function normHeader(header) {
    return String(header || '')
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/ß/g, 'ss')
        .replace(/\s+/g, ' ');
}

/**
 * @param {Record<string, unknown>} row
 * @param {string[]} candidates
 */
function getField(row, candidates) {
    if (!row || typeof row !== 'object') return '';
    const map = new Map();
    Object.keys(row).forEach((k) => map.set(normHeader(k), row[k]));
    for (let i = 0; i < candidates.length; i++) {
        const v = map.get(normHeader(candidates[i]));
        if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
}

/**
 * „Tim Breitfuß, Oskar Kofler und Michael Mayr“ → Einzelpersonen.
 * @param {string} raw
 * @returns {string[]}
 */
export function splitStudentNames(raw) {
    const text = String(raw || '').trim();
    if (!text) return [];
    return text
        .split(/\s*(?:,|;|\bund\b|\b&\b)\s*/i)
        .map((s) => s.trim())
        .filter(Boolean);
}

/**
 * „5chk - 2026/27“ → { klasseCode, schuljahr, terminJahrHint }
 * @param {string} raw
 */
export function parseAbaKlasseField(raw) {
    const s = String(raw || '').trim();
    if (!s) return { klasseCode: '', schuljahr: '', terminJahrHint: '' };
    const m = s.match(/^(.+?)\s*[-–—]\s*(\d{4})\s*\/\s*(\d{2,4})\s*$/);
    if (m) {
        const start = m[2];
        let end = m[3];
        if (end.length === 2) end = start.slice(0, 2) + end;
        return {
            klasseCode: m[1].trim(),
            schuljahr: start + '/' + end.slice(-2),
            terminJahrHint: end
        };
    }
    const onlyYear = s.match(/(\d{4})\s*\/\s*(\d{2,4})/);
    if (onlyYear) {
        const start = onlyYear[1];
        let end = onlyYear[2];
        if (end.length === 2) end = start.slice(0, 2) + end;
        return {
            klasseCode: s.replace(onlyYear[0], '').replace(/[-–—]/g, '').trim(),
            schuljahr: start + '/' + end.slice(-2),
            terminJahrHint: end
        };
    }
    return { klasseCode: s, schuljahr: '', terminJahrHint: '' };
}

/**
 * @param {object[]} jsonRows sheet_to_json rows
 * @param {{ terminJahr?: string|number, preferGenehmigt?: boolean }} [opts]
 */
export function parseAbaPortalRows(jsonRows, opts) {
    const o = opts || {};
    const terminJahr = String(o.terminJahr || '').trim();
    const preferGenehmigt = o.preferGenehmigt !== false;
    /** @type {Array<{
     *   titel: string,
     *   schuelerName: string,
     *   klasseRoh: string,
     *   klasseCode: string,
     *   schuljahr: string,
     *   betreuer: string,
     *   status: string,
     *   schule: string,
     *   genehmigt: boolean,
     *   terminJahrHint: string
     * }>} */
    const out = [];
    let skippedYear = 0;
    let skippedEmpty = 0;
    (Array.isArray(jsonRows) ? jsonRows : []).forEach((row) => {
        const titel = getField(row, TITLE_KEYS);
        const studentsRaw = getField(row, STUDENT_KEYS);
        const klasseRoh = getField(row, CLASS_KEYS);
        const betreuer = getField(row, SUPERVISOR_KEYS);
        const status = getField(row, STATUS_KEYS);
        const schule = getField(row, SCHOOL_KEYS);
        if (!titel && !studentsRaw) {
            skippedEmpty += 1;
            return;
        }
        const klasseInfo = parseAbaKlasseField(klasseRoh);
        if (terminJahr && klasseInfo.terminJahrHint) {
            const hint = String(klasseInfo.terminJahrHint);
            const startHint = klasseInfo.schuljahr
                ? String(klasseInfo.schuljahr).slice(0, 4)
                : String(Number(hint) - 1);
            // Terminjahr = Maturajahr (Ende Schuljahr) oder Startjahr im Feld
            if (hint !== terminJahr && startHint !== terminJahr) {
                skippedYear += 1;
                return;
            }
        }
        const genehmigt = /genehmigt/i.test(status);
        if (preferGenehmigt && status && !genehmigt) {
            // trotzdem importieren, aber markieren
        }
        const names = splitStudentNames(studentsRaw);
        if (!names.length) {
            skippedEmpty += 1;
            return;
        }
        names.forEach((schuelerName) => {
            out.push({
                titel: titel || '',
                schuelerName,
                klasseRoh,
                klasseCode: klasseInfo.klasseCode,
                schuljahr: klasseInfo.schuljahr,
                betreuer,
                status,
                schule,
                genehmigt,
                terminJahrHint: klasseInfo.terminJahrHint
            });
        });
    });
    return {
        rows: out,
        summary: {
            persons: out.length,
            genehmigt: out.filter((r) => r.genehmigt).length,
            skippedYear,
            skippedEmpty
        }
    };
}

/**
 * @param {string} vorname
 * @param {string} nachname
 */
export function personMatchKeys(vorname, nachname) {
    const v = String(vorname || '').trim();
    const n = String(nachname || '').trim();
    const keys = [];
    if (v && n) {
        keys.push(normalizePersonKey(v + ' ' + n));
        keys.push(normalizePersonKey(n + ' ' + v));
        keys.push(normalizePersonKey(n + ', ' + v));
    }
    if (n) keys.push(normalizePersonKey(n));
    return [...new Set(keys.filter(Boolean))];
}

/**
 * @param {Array<{ schuelerName?: string, klasseCode?: string, titel?: string, betreuer?: string }>} rows
 * @param {{ vorname?: string, nachname?: string, klasse?: string }} person
 */
export function findAbschlussarbeitenForPerson(rows, person) {
    const list = Array.isArray(rows) ? rows : [];
    const keys = personMatchKeys(person && person.vorname, person && person.nachname);
    if (!keys.length) return [];
    const klasseWant = normalizePersonKey(person && person.klasse).replace(/\s+/g, '');
    const hits = list.filter((row) => {
        const nameKey = normalizePersonKey(row && row.schuelerName);
        if (!nameKey || keys.indexOf(nameKey) === -1) return false;
        if (klasseWant) {
            const kc = normalizePersonKey(row && row.klasseCode).replace(/\s+/g, '');
            if (kc && kc !== klasseWant && kc.indexOf(klasseWant) === -1 && klasseWant.indexOf(kc) === -1) {
                return false;
            }
        }
        return true;
    });
    return hits;
}

/**
 * Betreuer-Name an Lehrer-Choice-Liste mappen.
 * @param {string} betreuer
 * @param {string[]} lehrerChoices
 */
export function mapBetreuerToLehrerChoice(betreuer, lehrerChoices) {
    const want = normalizePersonKey(betreuer);
    if (!want) return '';
    const list = Array.isArray(lehrerChoices) ? lehrerChoices : [];
    const exact = list.find((l) => normalizePersonKey(l) === want);
    if (exact) return exact;
    const partial = list.find((l) => {
        const k = normalizePersonKey(l);
        return k.indexOf(want) !== -1 || want.indexOf(k) !== -1;
    });
    return partial || String(betreuer || '').trim();
}

/**
 * Prefill-Felder für Wizard je Profil.
 * @param {object} hit
 * @param {string} [profileId]
 * @param {string[]} [lehrerChoices]
 */
export function abschlussarbeitPrefillFields(hit, profileId, lehrerChoices) {
    const titel = String((hit && hit.titel) || '').trim();
    const betreuer = mapBetreuerToLehrerChoice((hit && hit.betreuer) || '', lehrerChoices || []);
    const id = String(profileId || 'hak')
        .trim()
        .toLowerCase();
    if (id === 'ahs') {
        return {
            hatABA: true,
            titelABA: titel,
            betreuungslehrerABA: betreuer,
            titelDiplomarbeit: '',
            betreuungslehrerDA: '',
            abaPrefill: true
        };
    }
    return {
        hatABA: false,
        titelDiplomarbeit: titel,
        betreuungslehrerDA: betreuer,
        titelABA: '',
        betreuungslehrerABA: '',
        abaPrefill: true
    };
}
