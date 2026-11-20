/**
 * Zentrale sRDP-Termine laut https://www.matura.gv.at/pruefungstermine
 * Stand der eingebetteten Seite: Schuljahr 2025/26 und 2026/27.
 * Eingebettet ist der letzte bekannte Stand. Beim Öffnen wird die Seite neu gelesen.
 */

export const MATURA_TERMIN_URL = 'https://www.matura.gv.at/pruefungstermine';

/** @type {Record<string, number>} */
const MONTHS = {
    jänner: 1,
    jaenner: 1,
    januar: 1,
    februar: 2,
    märz: 3,
    maerz: 3,
    april: 4,
    mai: 5,
    juni: 6,
    juli: 7,
    august: 8,
    september: 9,
    oktober: 10,
    november: 11,
    dezember: 12
};

const MONTH_NAMES = [
    '',
    'Jänner',
    'Februar',
    'März',
    'April',
    'Mai',
    'Juni',
    'Juli',
    'August',
    'September',
    'Oktober',
    'November',
    'Dezember'
];

const DATE_RE = /^(\d{1,2})\.\s+([A-Za-zÄÖÜäöü]+)\s+(20\d{2})$/;
const TERM_RE = /^(?:#{1,3}\s*)?(Haupttermin|Herbsttermin|Wintertermin)\s+(20\d{2})\b/i;
const YEAR_RE = /Schuljahr\s+(\d{4}\s*\/\s*\d{2})/i;

export const MATURA_PAGE_TEXT = `
Schuljahr 2025/26

Wintertermin 2027
Deutsch (AHS/BHS/BRP)
11. Jänner 2027
Spanisch, Slowenisch, Kroatisch, Ungarisch (AHS/BHS) als Unterrichtssprache
15. Jänner 2027
Englisch (AHS/BHS/BRP)
12. Jänner 2027
Französisch (AHS/BHS)
14. Jänner 2027
Italienisch (AHS/BHS)
18. Jänner 2027
Latein (AHS)
08. Jänner 2027
Griechisch (AHS)
19. Jänner 2027
(angewandte) Mathematik (AHS/BHS/BRP)
13. Jänner 2027
mündliche Kompensationsprüfung (AHS/BHS/BRP)
28. Jänner 2027

Schuljahr 2026/27

Haupttermin 2027
Deutsch (AHS/BHS/BRP)
05. Mai 2027
Spanisch, Slowenisch, Kroatisch, Ungarisch (AHS/BHS) als Unterrichtssprache
14. Mai 2027
Englisch (AHS/BHS/BRP)
11. Mai 2027
Französisch (AHS/BHS)
13. Mai 2027
Italienisch (AHS/BHS)
07. Mai 2027
Latein (AHS)
12. Mai 2027
Griechisch (AHS)
21. Mai 2027
(angewandte) Mathematik (AHS/BHS/BRP)
10. Mai 2027
mündliche Kompensationsprüfung (AHS/BHS/BRP)
01. Juni 2027
02. Juni 2027

Herbsttermin 2027
Deutsch (AHS/BHS/BRP)
16. September 2027
Spanisch, Slowenisch, Kroatisch, Ungarisch (AHS/BHS) als Unterrichtssprache
22. September 2027
Englisch (AHS/BHS/BRP)
17. September 2027
Französisch (AHS/BHS)
20. September 2027
Italienisch (AHS/BHS)
23. September 2027
Latein (AHS)
27. September 2027
Griechisch (AHS)
15. September 2027
(angewandte) Mathematik (AHS/BHS/BRP)
21. September 2027
mündliche Kompensationsprüfung (AHS/BHS/BRP)
12. Oktober 2027

Wintertermin 2028
Deutsch (AHS/BHS/BRP)
11. Jänner 2028
Spanisch, Slowenisch, Kroatisch, Ungarisch (AHS/BHS) als Unterrichtssprache
17. Jänner 2028
Englisch (AHS/BHS/BRP)
12. Jänner 2028
Französisch (AHS/BHS)
14. Jänner 2028
Italienisch (AHS/BHS)
18. Jänner 2028
Latein (AHS)
10. Jänner 2028
Griechisch (AHS)
19. Jänner 2028
(angewandte) Mathematik (AHS/BHS/BRP)
13. Jänner 2028
mündliche Kompensationsprüfung (AHS/BHS/BRP)
01. Februar 2028
`;

/**
 * @param {string} value
 * @returns {string} ISO-Datum oder leer
 */
export function parseGermanDate(value) {
    const m = String(value || '')
        .trim()
        .match(DATE_RE);
    if (!m) return '';
    const month = MONTHS[m[2].toLowerCase()];
    if (!month) return '';
    return m[3] + '-' + String(month).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
}

/** @param {string} value */
export function normalizeIsoDate(value) {
    const match = String(value || '')
        .trim()
        .match(/(20\d{2})-(\d{2})-(\d{2})/);
    if (!match) return '';
    const month = Number(match[2]);
    const day = Number(match[3]);
    if (month < 1 || month > 12 || day < 1 || day > 31) return '';
    return match[1] + '-' + match[2] + '-' + match[3];
}

/**
 * Datum aus der versteckten Listenspalte BfkSchriftlich.
 * @param {{ description?: string, defaultValue?: { value?: string } }|null|undefined} column
 */
export function bfkIsoFromColumn(column) {
    if (!column) return '';
    const chunks = [column.description, column.defaultValue && column.defaultValue.value];
    for (let i = 0; i < chunks.length; i++) {
        const iso = normalizeIsoDate(chunks[i]);
        if (iso) return iso;
    }
    return '';
}

export function formatGermanDate(iso) {
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return '';
    const name = MONTH_NAMES[Number(m[2])] || '';
    if (!name) return '';
    return Number(m[3]) + '. ' + name + ' ' + m[1];
}

/**
 * @param {string} label
 */
export function fachKeyFromLabel(label) {
    const s = String(label || '').toLowerCase();
    if (s.includes('kompensation')) return 'kompensation';
    if (s.includes('deutsch')) return 'deutsch';
    if (s.includes('englisch')) return 'englisch';
    if (s.includes('französ') || s.includes('franzoes')) return 'franzoesisch';
    if (s.includes('italien')) return 'italienisch';
    if (s.includes('spanisch') || s.includes('slowen') || s.includes('kroat') || s.includes('ungar')) {
        return 'spanisch';
    }
    if (s.includes('latein')) return 'latein';
    if (s.includes('griech')) return 'griechisch';
    if (s.includes('mathematik')) return 'mathematik';
    return '';
}

/**
 * @param {string} text
 * @returns {Array<{ schuljahr: string, art: string, jahr: string, faecher: Array<{ fach: string, key: string, daten: string[] }> }>}
 */
export function parseMaturaTerminPage(text) {
    const lines = String(text || '')
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);
    /** @type {Array<{ schuljahr: string, art: string, jahr: string, faecher: Array<{ fach: string, key: string, daten: string[] }> }>} */
    const out = [];
    let schuljahr = '';
    /** @type {{ schuljahr: string, art: string, jahr: string, faecher: Array<{ fach: string, key: string, daten: string[] }> }|null} */
    let current = null;
    let pending = '';

    const pushDate = (iso) => {
        if (!current || !pending || !iso) return;
        let row = current.faecher.find((f) => f.fach === pending);
        if (!row) {
            row = { fach: pending, key: fachKeyFromLabel(pending), daten: [] };
            current.faecher.push(row);
        }
        if (!row.daten.includes(iso)) row.daten.push(iso);
    };

    lines.forEach((line) => {
        const year = line.match(YEAR_RE);
        if (year) {
            schuljahr = year[1].replace(/\s+/g, '');
            pending = '';
            return;
        }
        const term = line.match(TERM_RE);
        if (term) {
            current = {
                schuljahr,
                art: term[1].toLowerCase(),
                jahr: term[2],
                faecher: []
            };
            out.push(current);
            pending = '';
            return;
        }
        if (!current) return;
        if (/^quelle:/i.test(line) || /inhalt aufklappen/i.test(line)) {
            pending = '';
            return;
        }
        const iso = parseGermanDate(line);
        if (iso) {
            pushDate(iso);
            return;
        }
        pending = line;
    });
    return out;
}

export const PRUEFUNGSTERMINE = parseMaturaTerminPage(MATURA_PAGE_TEXT);

/**
 * Haupttermin des Jahres, Herbst desselben Jahres, Winter des Folgejahres.
 * @param {string|number} terminJahr
 * @param {typeof PRUEFUNGSTERMINE} [catalog]
 */
export function periodenForTerminJahr(terminJahr, catalog) {
    const jahr = String(terminJahr || '').trim();
    const list = Array.isArray(catalog) ? catalog : getActivePruefungstermine();
    const next = /^\d{4}$/.test(jahr) ? String(Number(jahr) + 1) : '';
    const find = (art, year) => list.find((row) => row.art === art && row.jahr === year) || null;
    const haupt = find('haupttermin', jahr);
    const herbst = find('herbsttermin', jahr);
    const winter = next ? find('wintertermin', next) : null;
    return {
        jahr,
        schuljahr: (haupt && haupt.schuljahr) || (herbst && herbst.schuljahr) || '',
        perioden: [
            { id: 'haupt', label: 'Haupttermin', neben: '', block: haupt },
            { id: 'herbst', label: 'Herbsttermin', neben: '1. Nebentermin', block: herbst },
            { id: 'winter', label: 'Wintertermin', neben: '2. Nebentermin', block: winter }
        ].filter((row) => row.block && row.block.faecher.length)
    };
}

/**
 * @param {string} lfs
 */
export function lfsToFachKey(lfs) {
    const s = String(lfs || '')
        .trim()
        .toLowerCase();
    if (!s) return '';
    if (/englisch|^en\b|^enws|^en[\s-]/.test(s) || s === 'en' || s.startsWith('enws') || s.startsWith('en ')) {
        return 'englisch';
    }
    if (/franz|^fr\b|^frws/.test(s) || s === 'fr' || s.startsWith('frws')) return 'franzoesisch';
    if (/italien|^it\b|^itws/.test(s) || s === 'it' || s.startsWith('itws')) return 'italienisch';
    if (/spanisch|slowen|kroat|ungar|^sp\b|^spws/.test(s) || s === 'sp' || s.startsWith('spws')) {
        return 'spanisch';
    }
    if (/latein|^lat\b/.test(s)) return 'latein';
    if (/griech/.test(s)) return 'griechisch';
    return '';
}

/**
 * Prüfungsteil der Variante → zentrales Fach.
 * BFK, BKO, Wahlfach und Seminar haben keinen Termin auf matura.gv.at.
 * @param {string} part
 * @param {string} [lfs]
 */
export function examFachKey(part, lfs) {
    const p = String(part || '')
        .trim()
        .toLowerCase();
    if (!p) return '';
    if (/(bfk|bko|wahlfach|seminar|fachkolloquium|fachtheorie|schwerpunkt|bw-rw|\bbwr\b)/.test(p)) {
        return '';
    }
    if (p === 'd' || p.startsWith('d ') || p.startsWith('d(') || p.includes('deutsch')) return 'deutsch';
    if (
        p === 'am' ||
        p.startsWith('am ') ||
        p.startsWith('am(') ||
        p === 'm' ||
        p.startsWith('m ') ||
        p.startsWith('m(') ||
        p.includes('mathematik')
    ) {
        return 'mathematik';
    }
    if (p.includes('englisch')) return 'englisch';
    if (p.includes('französ') || p.includes('franz')) return 'franzoesisch';
    if (p.includes('italien')) return 'italienisch';
    if (p.includes('spanisch')) return 'spanisch';
    if (p.includes('latein')) return 'latein';
    if (p.includes('griech')) return 'griechisch';
    if (p.includes('lfs') || p.includes('fremdsprache')) return lfsToFachKey(lfs) || 'lfs-offen';
    return '';
}

const FACH_LABEL = {
    deutsch: 'Deutsch',
    englisch: 'Englisch',
    franzoesisch: 'Französisch',
    italienisch: 'Italienisch',
    spanisch: 'Spanisch',
    latein: 'Latein',
    griechisch: 'Griechisch',
    mathematik: 'Mathematik',
    kompensation: 'Mündliche Kompensationsprüfung'
};

/**
 * Datum des Haupttermins für einen Prüfungsteil.
 * @param {string} part
 * @param {string} [lfs]
 * @param {string|number} [terminJahr]
 * @param {{ bfkIso?: string }} [school]
 */
export function datumForExamPart(part, lfs, terminJahr, school) {
    const key = examFachKey(part, lfs);
    if (key === 'lfs-offen') {
        return {
            key: '',
            fach: 'Lebende Fremdsprache',
            iso: '',
            label: '',
            hinweis: 'Sprache noch wählen',
            zentral: true
        };
    }
    if (!key) {
        const p = String(part || '').toLowerCase();
        const bfkIso = normalizeIsoDate(school && school.bfkIso);
        if (p.includes('bfk') && bfkIso) {
            return {
                key: 'bfk',
                fach: 'BFK schriftlich',
                iso: bfkIso,
                label: formatGermanDate(bfkIso),
                hinweis: '',
                zentral: false
            };
        }
        return {
            key: '',
            fach: '',
            iso: '',
            label: '',
            hinweis: 'Termin legt die Schule fest',
            zentral: false
        };
    }
    const bundle = periodenForTerminJahr(terminJahr);
    const haupt = bundle.perioden.find((row) => row.id === 'haupt');
    const fach = haupt && haupt.block.faecher.find((row) => row.key === key);
    if (!fach || !fach.daten.length) {
        return {
            key,
            fach: FACH_LABEL[key] || key,
            iso: '',
            label: '',
            hinweis: 'Kein veröffentlichter Termin',
            zentral: true
        };
    }
    const iso = fach.daten[0];
    return {
        key,
        fach: FACH_LABEL[key] || fach.fach,
        iso,
        label: formatGermanDate(iso),
        hinweis: '',
        zentral: true
    };
}

const CACHE_KEY = 'ms365-srdp-matura-termine-v1';
const CACHE_MAX_AGE_MS = 12 * 60 * 60 * 1000;

/** @type {typeof PRUEFUNGSTERMINE|null} */
let liveCatalog = null;

export function getActivePruefungstermine() {
    return liveCatalog && liveCatalog.length ? liveCatalog : PRUEFUNGSTERMINE;
}

/**
 * HTML der Matura-Seite in Zeilen, die der Parser versteht.
 * @param {string} html
 */
export function htmlToMaturaText(html) {
    let s = String(html || '');
    s = s.replace(/<script[\s\S]*?<\/script>/gi, '\n');
    s = s.replace(/<style[\s\S]*?<\/style>/gi, '\n');
    s = s.replace(/<br\s*\/?>/gi, '\n');
    s = s.replace(/<\/(?:p|h[1-6]|tr|li|div|span|td|th)>/gi, '\n');
    s = s.replace(/<[^>]+>/g, ' ');
    s = s
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/&auml;/gi, 'ä')
        .replace(/&ouml;/gi, 'ö')
        .replace(/&uuml;/gi, 'ü')
        .replace(/&Auml;/g, 'Ä')
        .replace(/&Ouml;/g, 'Ö')
        .replace(/&Uuml;/g, 'Ü')
        .replace(/&#(\d+);/g, (_, n) => {
            const code = Number(n);
            return Number.isFinite(code) ? String.fromCharCode(code) : '';
        });
    return s
        .split(/\r?\n/)
        .map((line) => line.replace(/\s+/g, ' ').trim())
        .filter(Boolean)
        .join('\n');
}

/**
 * Live-Termine ergänzen den mitgelieferten Stand. Gleiches Jahr gewinnt die Seite.
 * @param {typeof PRUEFUNGSTERMINE} primary
 * @param {typeof PRUEFUNGSTERMINE} fallback
 */
export function mergePruefungstermine(primary, fallback) {
    const map = new Map();
    (Array.isArray(fallback) ? fallback : []).forEach((row) => {
        if (!row || !row.art || !row.jahr) return;
        map.set(row.art + '|' + row.jahr, row);
    });
    (Array.isArray(primary) ? primary : []).forEach((row) => {
        if (!row || !row.art || !row.jahr || !(row.faecher || []).length) return;
        map.set(row.art + '|' + row.jahr, row);
    });
    return [...map.values()];
}

function readTerminCache() {
    try {
        if (typeof localStorage === 'undefined') return null;
        const raw = localStorage.getItem(CACHE_KEY);
        if (!raw) return null;
        const data = JSON.parse(raw);
        const at = Date.parse(data && data.fetchedAt);
        if (!Number.isFinite(at) || Date.now() - at > CACHE_MAX_AGE_MS) return null;
        const catalog = Array.isArray(data.catalog) ? data.catalog : [];
        if (!catalog.length) return null;
        return { catalog, fetchedAt: data.fetchedAt, source: 'cache' };
    } catch {
        return null;
    }
}

function cacheCoversYear(catalog, terminJahr) {
    const jahr = String(terminJahr || '').trim();
    if (!/^\d{4}$/.test(jahr)) return true;
    return (catalog || []).some(
        (row) =>
            row &&
            (row.art === 'haupttermin' || row.art === 'herbsttermin') &&
            row.jahr === jahr &&
            (row.faecher || []).length
    );
}

function writeTerminCache(catalog, fetchedAt) {
    try {
        if (typeof localStorage === 'undefined') return;
        localStorage.setItem(
            CACHE_KEY,
            JSON.stringify({ catalog, fetchedAt: fetchedAt || new Date().toISOString() })
        );
    } catch {
        /* Speicher voll oder gesperrt */
    }
}

function licenseApiTerminUrl() {
    try {
        const cfg = typeof window !== 'undefined' ? window.MS365_LICENSE_API : null;
        const base = String((cfg && cfg.baseUrl) || '').replace(/\/+$/, '');
        if (!base) return '';
        return base + '/matura-termine';
    } catch {
        return '';
    }
}

/**
 * @param {typeof fetch} fetchFn
 * @param {string} url
 * @param {Record<string, string>} [headers]
 */
async function readTerminPayload(fetchFn, url, headers) {
    const res = await fetchFn(url, { headers: headers || { Accept: 'text/html,application/json' } });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const type = String(res.headers.get('content-type') || '');
    if (type.includes('json')) {
        const data = await res.json();
        const html = String((data && (data.html || data.text)) || '');
        if (!html.trim()) throw new Error('Leere Antwort');
        return { text: html.trim().startsWith('<') ? htmlToMaturaText(html) : html, fetchedAt: data.fetchedAt || '' };
    }
    const html = await res.text();
    return { text: htmlToMaturaText(html), fetchedAt: '' };
}

let refreshFlight = null;

/**
 * Liest https://www.matura.gv.at/pruefungstermine neu.
 * Der Browser darf die Seite nicht direkt lesen, deshalb zuerst die License-API.
 * @param {{ fetch?: typeof fetch, force?: boolean }} [opts]
 */
export async function refreshPruefungstermineFromWeb(opts) {
    if (refreshFlight && !(opts && opts.force)) return refreshFlight;
    const run = (async () => {
        const wanted = String((opts && opts.terminJahr) || '').trim();
        const cached = opts && opts.force ? null : readTerminCache();
        if (cached && cacheCoversYear(cached.catalog, wanted)) {
            liveCatalog = mergePruefungstermine(cached.catalog, PRUEFUNGSTERMINE);
            return { source: 'cache', fetchedAt: cached.fetchedAt, catalog: liveCatalog };
        }
        const fetchFn = (opts && opts.fetch) || (typeof fetch === 'function' ? fetch : null);
        if (!fetchFn) {
            return { source: 'snapshot', fetchedAt: '', catalog: getActivePruefungstermine() };
        }
        const apiUrl = licenseApiTerminUrl();
        const headers = { Accept: 'application/json, text/html' };
        try {
            const cfg = typeof window !== 'undefined' ? window.MS365_LICENSE_API : null;
            const key = String((cfg && cfg.functionKey) || '').trim();
            if (key) headers['x-functions-key'] = key;
        } catch {
            /* ignore */
        }
        /** @type {{ text: string, fetchedAt: string }|null} */
        let payload = null;
        if (apiUrl) {
            try {
                payload = await readTerminPayload(fetchFn, apiUrl, headers);
            } catch {
                payload = null;
            }
        }
        if (!payload) {
            try {
                payload = await readTerminPayload(fetchFn, MATURA_TERMIN_URL, { Accept: 'text/html' });
            } catch {
                payload = null;
            }
        }
        if (!payload) {
            return { source: 'snapshot', fetchedAt: '', catalog: getActivePruefungstermine() };
        }
        const parsed = parseMaturaTerminPage(payload.text);
        if (!parsed.length) {
            return { source: 'snapshot', fetchedAt: '', catalog: getActivePruefungstermine() };
        }
        const fetchedAt = payload.fetchedAt || new Date().toISOString();
        liveCatalog = mergePruefungstermine(parsed, PRUEFUNGSTERMINE);
        writeTerminCache(parsed, fetchedAt);
        return { source: 'live', fetchedAt, catalog: liveCatalog };
    })();
    refreshFlight = run;
    try {
        return await run;
    } finally {
        if (refreshFlight === run) refreshFlight = null;
    }
}
