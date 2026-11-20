/**
 * Reine sRDP/sRP-Anmeldungs-Logik (kein DOM/Graph) – Vitest.
 */
import {
    VARIANT_LABELS,
    getProfile,
    PROFILE_HAK
} from './srdp-anmeldung-schema.js';
import { datumForExamPart } from './srdp-pruefungstermine.js';

/**
 * @param {string|number|null|undefined} nachname
 * @param {string|number|null|undefined} vorname
 */
export function buildItemTitle(nachname, vorname) {
    const n = String(nachname == null ? '' : nachname).trim();
    const v = String(vorname == null ? '' : vorname).trim();
    if (n && v) return n + ', ' + v;
    return n || v || '';
}

/**
 * @param {string|number|null|undefined} variante
 * @param {string} [profileId]
 */
export function resolveVariant(variante, profileId) {
    const profile = getProfile(profileId);
    const raw = String(variante == null ? '' : variante).trim();
    if (!raw) return null;
    const digit = raw.match(/([123])\s*$/);
    const key = digit ? digit[1] : '';
    const found = profile.variants.find(
        (v) => v.key === key || v.label.toLowerCase() === raw.toLowerCase()
    );
    return found || null;
}

/**
 * @param {string|number|null|undefined} variante
 * @param {string} [profileId]
 */
export function buildPruefplanKurz(variante, profileId) {
    const v = resolveVariant(variante, profileId);
    if (!v) return '';
    return (
        'schriftlich: ' +
        v.schriftlich.join(', ') +
        ' · mündlich: ' +
        v.muendlich.join(', ')
    );
}

/**
 * Gewählte Prüfungsteile getrennt nach schriftlich / mündlich, mit den Angaben aus der Anmeldung.
 * @param {object} item
 * @param {string} [profileId]
 * @returns {{ schriftlich: { part: string, detail: string }[], muendlich: { part: string, detail: string }[] }}
 */
export function buildPruefplanDetail(item, profileId, opts) {
    const it = item || {};
    const jahr = (opts && opts.terminJahr) || it.terminJahr || '';
    const v = resolveVariant(it.variante, profileId);
    const detailFor = (part) => {
        const p = String(part || '').trim().toLowerCase();
        if (p === 'd' || p.startsWith('d ') || p.startsWith('d(')) {
            return [it.titelDiplomarbeit || it.titelABA, it.betreuungslehrerDA || it.betreuungslehrerABA]
                .map((x) => String(x || '').trim())
                .filter(Boolean)
                .join(' · ');
        }
        if (p.indexOf('lfs') !== -1) {
            return [it.lfs, it.lehrerLFS]
                .map((x) => String(x || '').trim())
                .filter(Boolean)
                .join(' · ');
        }
        if (p.indexOf('bko') !== -1) return String(it.lehrerBKO || '').trim();
        if (p.indexOf('wahlfach') !== -1) {
            const wf = String(it.wahlfachMuendlich || it.wahlfach || '').trim();
            const sem = isSeminarWahlfach(wf) ? String(it.seminar || '').trim() : '';
            return [wf, sem, it.lehrerWahlfach]
                .map((x) => String(x || '').trim())
                .filter(Boolean)
                .join(' · ');
        }
        return '';
    };
    const mapParts = (parts) =>
        (Array.isArray(parts) ? parts : []).map((part) => {
            const name = String(part || '').trim();
            const when = jahr
                ? datumForExamPart(name, it.lfs, jahr, {
                      bfkIso: (opts && (opts.bfkDatum || opts.bfkIso)) || ''
                  })
                : null;
            return {
                part: name,
                detail: detailFor(part),
                datum: when && when.label ? when.label : '',
                datumHinweis: when && !when.label ? when.hinweis : '',
                datumIso: when && when.iso ? when.iso : '',
                fachKey: when && when.key ? when.key : '',
                fach: when && when.fach ? when.fach : ''
            };
        });
    if (!v) {
        return { schriftlich: [], muendlich: [] };
    }
    return {
        schriftlich: mapParts(v.schriftlich),
        muendlich: mapParts(v.muendlich)
    };
}

/**
 * @param {string|number|null|undefined} wahlfach
 */
export function isSeminarWahlfach(wahlfach) {
    const w = String(wahlfach == null ? '' : wahlfach).trim().toLowerCase();
    if (!w) return false;
    return w === 'seminar' || w.indexOf('seminar') === 0;
}

/**
 * @param {{
 *   profileId?: string,
 *   variante?: string,
 *   wahlfach?: string,
 *   seminar?: string,
 *   lfs?: string,
 *   lehrerLfs?: string,
 *   nachname?: string,
 *   vorname?: string,
 *   bestaetigung?: boolean
 * }} draft
 */
export function validateAnmeldungDraft(draft) {
    const d = draft || {};
    const profile = getProfile(d.profileId);
    /** @type {string[]} */
    const errors = [];
    const variante = resolveVariant(d.variante, profile.id);
    if (!variante) errors.push('Variante fehlt oder ist ungültig.');
    if (!String(d.nachname || '').trim()) errors.push('Nachname fehlt.');
    if (!String(d.vorname || '').trim()) errors.push('Vorname fehlt.');
    if (profile.extraColumnNames.indexOf('LFS') !== -1) {
        if (!String(d.lfs || '').trim()) errors.push((profile.lfsFieldLabel || 'LFS') + ' fehlt.');
        if (!String(d.lehrerLfs || '').trim()) errors.push('LehrerIn LFS fehlt.');
    }
    if (isSeminarWahlfach(d.wahlfach) && !String(d.seminar || '').trim()) {
        errors.push('Seminar-Bezeichnung fehlt (Wahlfach Seminar).');
    }
    if (d.bestaetigung !== true && d.bestaetigung !== 'Ja' && d.bestaetigung !== 'ja') {
        errors.push('Bestätigung der Anmeldung fehlt.');
    }
    return {
        ok: errors.length === 0,
        errors,
        title: buildItemTitle(d.nachname, d.vorname),
        pruefplanKurz: variante ? buildPruefplanKurz(variante.key, profile.id) : '',
        varianteLabel: variante ? variante.label : '',
        profileId: profile.id
    };
}

/**
 * Abschlussklassen für den Termin – Jahr-Match bevorzugt, sonst Fallback
 * (Klassen ohne Jahr / alle Codes), damit das Schüler-Dropdown nie leer bleibt.
 * @param {Array<{ code?: string, name?: string, year?: string|number }>} classes
 * @param {string|number} terminJahr
 * @returns {string[]}
 */
export function filterAbschlussklassen(classes, terminJahr) {
    const y = String(terminJahr || '').trim();
    const list = Array.isArray(classes) ? classes : [];
    const codeOf = (c) => {
        const code = String((c && c.code) || '').trim();
        const name = String((c && c.name) || '').trim();
        return code || name;
    };
    if (/^\d{4}$/.test(y)) {
        const byYear = list
            .filter((c) => String((c && c.year) || '').trim() === y)
            .map(codeOf)
            .filter(Boolean);
        if (byYear.length) return [...new Set(byYear)];
    }
    const withoutYear = list
        .filter((c) => !String((c && c.year) || '').trim())
        .map(codeOf)
        .filter(Boolean);
    if (withoutYear.length) return [...new Set(withoutYear)];
    return [...new Set(list.map(codeOf).filter(Boolean))];
}

/**
 * @param {Array<{ code?: string, name?: string }>} subjects
 * @param {string[]} [hints]
 */
export function suggestLfsChoices(subjects, hints) {
    const hintList = Array.isArray(hints) && hints.length ? hints : PROFILE_HAK.lfsHints;
    const fromSubjects = [];
    (Array.isArray(subjects) ? subjects : []).forEach((s) => {
        const code = String((s && s.code) || '').trim().toUpperCase();
        if (!code) return;
        if (/(WS|LFS|EN|FR|IT|SP|RU|LAT)/i.test(code) || hintList.indexOf(code) !== -1) {
            fromSubjects.push(code);
        }
    });
    const seen = new Set();
    const out = [];
    fromSubjects.concat(hintList).forEach((c) => {
        const k = String(c).toUpperCase();
        if (seen.has(k)) return;
        seen.add(k);
        out.push(k);
    });
    return out;
}

/**
 * @param {Array<{ code?: string, name?: string, email?: string }>} teachers
 */
export function teacherChoiceLabels(teachers) {
    const out = [];
    const seen = new Set();
    (Array.isArray(teachers) ? teachers : []).forEach((t) => {
        const name = String((t && t.name) || '').trim();
        const code = String((t && t.code) || '').trim();
        const label = name || code;
        if (!label) return;
        const key = label.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(label);
    });
    out.sort((a, b) => a.localeCompare(b, 'de'));
    return out;
}

/**
 * @param {string} text
 */
export function parseChoiceLines(text) {
    const seen = new Set();
    const out = [];
    String(text || '')
        .split(/\r?\n/)
        .forEach((line) => {
            const s = line.trim();
            if (!s || s.startsWith('#')) return;
            const key = s.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push(s);
        });
    return out;
}

const DEMO_LFS_CODES = ['en', 'enws', 'fr', 'frws', 'it', 'itws', 'sp', 'spws'];

/** Unveränderte Vorschlagsliste EN / ENWS / FR … – nicht die Auswahl aus Schritt 3. */
export function looksLikeDemoLfs(value) {
    const list = (Array.isArray(value) ? value : parseChoiceLines(value))
        .map((s) => String(s || '').trim().toLowerCase())
        .filter(Boolean)
        .sort();
    if (list.length !== DEMO_LFS_CODES.length) return false;
    return list.every((code, i) => code === DEMO_LFS_CODES[i]);
}

/**
 * Was Schritt 3 zuletzt gespeichert hat.
 * Das Formular schreibt `lfs` / `wahlfach` / `seminar`, die Anmeldung liest `lfsText`.
 * @param {object} wiz
 * @param {string} textKey
 * @param {string} mirrorKey
 */
function looksLikeStockSeminare(value) {
    const list = (Array.isArray(value) ? value : parseChoiceLines(value))
        .map((s) => String(s || '').trim().toLowerCase())
        .filter(Boolean)
        .sort();
    return list.length === 2 && list[0] === 'digbiz' && list[1] === 'wirtschaftsrecht';
}

function isStockChoiceCatalog(lines) {
    return looksLikeDemoLfs(lines) || looksLikeStockSeminare(lines);
}

export function wizardConfiguredChoices(wiz, textKey, mirrorKey) {
    const src = wiz || {};
    const mirror = src[mirrorKey];
    const fromMirror =
        typeof mirror === 'string'
            ? parseChoiceLines(mirror)
            : Array.isArray(mirror)
              ? mirror.map((s) => String(s || '').trim()).filter(Boolean)
              : [];
    const fromText = parseChoiceLines(src[textKey] || '');
    if (
        fromMirror.length &&
        fromText.length &&
        isStockChoiceCatalog(fromMirror) &&
        !isStockChoiceCatalog(fromText)
    ) {
        return fromText;
    }
    if (
        fromText.length &&
        fromMirror.length &&
        isStockChoiceCatalog(fromText) &&
        !isStockChoiceCatalog(fromMirror)
    ) {
        return fromMirror;
    }
    if (fromMirror.length) return fromMirror;
    return fromText;
}

/**
 * Lehrer-Rollen an einer Anmeldung (Anzeige-Label-Match, case-insensitive).
 * @param {object} item
 * @param {string} teacherLabel
 * @returns {Array<{ role: string, label: string }>}
 */
export function teacherRolesOnItem(item, teacherLabel) {
    const want = String(teacherLabel || '')
        .trim()
        .toLowerCase();
    if (!want || !item) return [];
    /** @type {Array<{ key: string, role: string, label: string }>} */
    const fields = [
        { key: 'betreuungslehrerDA', role: 'da', label: 'Diplomarbeit' },
        { key: 'betreuungslehrerABA', role: 'aba', label: 'ABA / VWA' },
        { key: 'lehrerLFS', role: 'lfs', label: 'LFS' },
        { key: 'lehrerBKO', role: 'bko', label: 'BKO mündlich' },
        { key: 'lehrerWahlfach', role: 'wahlfach', label: 'Wahlfach mündlich' },
        { key: 'lehrerSchwerpunkt', role: 'schwerpunkt', label: 'Schwerpunkt' },
        { key: 'lehrerFachkolloquium', role: 'fachkolloquium', label: 'Fachkolloquium' },
        { key: 'lehrerFachtheorie', role: 'fachtheorie', label: 'Fachtheorie' },
        { key: 'lehrerVierteKlausur', role: 'vierte', label: '4. Klausur' },
        { key: 'lehrerMuendlich', role: 'muendlich', label: 'Mündlich' }
    ];
    const out = [];
    for (let i = 0; i < fields.length; i++) {
        const f = fields[i];
        const val = String(item[f.key] || '').trim().toLowerCase();
        if (val && val === want) out.push({ role: f.role, label: f.label });
    }
    return out;
}

/**
 * @param {object[]} items
 * @param {string} teacherLabel
 */
export function filterItemsForTeacher(items, teacherLabel) {
    const list = Array.isArray(items) ? items : [];
    return list.filter((it) => teacherRolesOnItem(it, teacherLabel).length > 0);
}

/** Klasse vergleichbar machen: „5A“ und „5a“ sind dieselbe Klasse. */
export function normalizeKlasseCode(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
}

function studentRowKlasse(row) {
    return String((row && (row.klasse || row.klasseCode || row.classCode || row.class)) || '').trim();
}

function studentEmailHits(row, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!row || !em) return false;
    return [row.email, row.mail, row.upn, row.userPrincipalName].some(
        (v) =>
            String(v || '')
                .trim()
                .toLowerCase() === em
    );
}

/** „DEMO Klasse 5A“ nennt die Klasse 5A, „15A“ und „5AHAK“ nicht. */
export function groupNameNamesClass(displayName, code) {
    const name = String(displayName || '').toLowerCase();
    const c = normalizeKlasseCode(code);
    if (!name || !c) return false;
    const escaped = c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return new RegExp('(^|[^a-z0-9])' + escaped + '([^a-z0-9]|$)').test(name);
}

/** „Jahrgangsgruppe DEMO Klasse 5A / Abschluss 2027“ → 5A. */
export function klasseCodeFromGroupName(displayName) {
    const name = String(displayName || '');
    const m = name.match(/(?:klasse|class)\s+([0-9]{1,2}[a-zA-Z]{1,8})\b/i);
    return m ? String(m[1]).trim() : '';
}

function classEmails(row) {
    const raw = []
        .concat((row && row.studentEmails) || [], (row && row.students) || [])
        .map((entry) => {
            if (entry && typeof entry === 'object') {
                return entry.email || entry.mail || entry.upn || '';
            }
            return entry;
        });
    return raw
        .map((v) =>
            String(v || '')
                .trim()
                .toLowerCase()
        )
        .filter(Boolean);
}

/**
 * Klasse des angemeldeten Schülers.
 * Reihenfolge: Schülerliste (Zeile mit Klasse gewinnt, auch aus einem anderen Schuljahr),
 * E-Mails an der Klasse, dann die Microsoft-365-Gruppe, in der das Konto Mitglied ist.
 * @param {{ email?: string, students?: object[], classes?: object[], classLinks?: object[], memberGroups?: object[], preferredCodes?: string[] }} [input]
 */
export function resolveStudentAssignedKlasse(input) {
    const ctx = input || {};
    const em = String(ctx.email || '')
        .trim()
        .toLowerCase();
    const students = Array.isArray(ctx.students) ? ctx.students : [];
    if (em) {
        let hit = '';
        students.forEach((s) => {
            if (hit || !studentEmailHits(s, em)) return;
            const k = studentRowKlasse(s);
            if (k) hit = k;
        });
        if (hit) return hit;
    }
    const classes = Array.isArray(ctx.classes) ? ctx.classes : [];
    if (em) {
        for (let i = 0; i < classes.length; i++) {
            const c = classes[i];
            if (!classEmails(c).includes(em)) continue;
            const code = String((c && (c.code || c.name)) || '').trim();
            if (code) return code;
        }
    }
    const groups = Array.isArray(ctx.memberGroups) ? ctx.memberGroups : [];
    const links = Array.isArray(ctx.classLinks) ? ctx.classLinks : [];
    const preferred = (Array.isArray(ctx.preferredCodes) ? ctx.preferredCodes : [])
        .map((c) => normalizeKlasseCode(c))
        .filter(Boolean);
    /** @type {string[]} */
    const named = [];
    const remember = (code) => {
        const c = String(code || '').trim();
        if (c) named.push(c);
    };
    groups.forEach((g) => {
        const id = String((g && (g.id || g.groupId)) || '')
            .trim()
            .toLowerCase();
        const mail = String((g && g.mail) || '')
            .trim()
            .toLowerCase();
        const dn = String((g && (g.displayName || g.name)) || '').trim();
        const pools = classes.concat(links);
        for (let i = 0; i < pools.length; i++) {
            const c = pools[i];
            const gid = String((c && (c.m365GroupId || c.graphGroupId || c.groupId)) || '')
                .trim()
                .toLowerCase();
            const gmail = String((c && (c.m365GroupMail || c.mail)) || '')
                .trim()
                .toLowerCase();
            const gname = String((c && (c.m365GroupName || c.displayName)) || '')
                .trim()
                .toLowerCase();
            const code = String((c && (c.code || c.classCode || c.name)) || '').trim();
            if (!code) continue;
            if ((gid && id && gid === id) || (gmail && mail && gmail === mail)) {
                remember(code);
                continue;
            }
            if (gname && dn && gname === dn.toLowerCase()) remember(code);
            else if (groupNameNamesClass(dn, code)) remember(code);
        }
        remember(klasseCodeFromGroupName(dn));
    });
    if (!named.length) return '';
    if (preferred.length) {
        const pref = named.find((c) => preferred.includes(normalizeKlasseCode(c)));
        if (pref) return pref;
    }
    return named[0];
}

/**
 * Haupttermin, den die App öffnen soll.
 * Gespeichertes Jahr gilt, außer es ist nur das laufende Kalenderjahr
 * (das setzt der Browser als Default). Ab August ist der nächste Frühjahrs-Termin gemeint.
 * @param {Date|string|number} [now]
 * @param {string|number} [savedJahr]
 */
export function preferredSrdpTerminJahr(now, savedJahr) {
    const d = now instanceof Date && !Number.isNaN(now.getTime()) ? now : new Date(now || Date.now());
    const cal = d.getFullYear();
    const upcoming = d.getMonth() >= 7 ? cal + 1 : cal;
    const saved = String(savedJahr || '').trim();
    if (/^\d{4}$/.test(saved) && Number(saved) !== cal) return saved;
    return String(upcoming);
}

/**
 * Jahre, in denen eine Anmeldeliste liegen kann. Bevorzugtes Jahr zuerst.
 * @param {Date|string|number} [now]
 * @param {string|number} [savedJahr]
 * @returns {string[]}
 */
export function srdpYearsToProbe(now, savedJahr) {
    const d = now instanceof Date && !Number.isNaN(now.getTime()) ? now : new Date(now || Date.now());
    const cal = d.getFullYear();
    const preferred = preferredSrdpTerminJahr(d, savedJahr);
    const years = [preferred, String(cal + 1), String(cal), String(cal - 1)];
    const saved = String(savedJahr || '').trim();
    if (/^\d{4}$/.test(saved)) years.push(saved);
    return [...new Set(years)];
}

/**
 * @param {object[]} items
 * @param {string[]} klasseCodes
 */
export function filterItemsForKlassen(items, klasseCodes) {
    const set = new Set(
        (Array.isArray(klasseCodes) ? klasseCodes : [])
            .map((c) => normalizeKlasseCode(c))
            .filter(Boolean)
    );
    if (!set.size) return [];
    return (Array.isArray(items) ? items : []).filter((it) =>
        set.has(normalizeKlasseCode(it && it.klasse))
    );
}

/**
 * KPIs / Auswertungs-Buckets für Admin.
 * @param {object[]} items
 */
export function computeAnmeldungStats(items) {
    const list = Array.isArray(items) ? items : [];
    /** @type {Record<string, number>} */
    const byKlasse = {};
    /** @type {Record<string, number>} */
    const byVariante = {};
    /** @type {Record<string, number>} */
    const byWahlfach = {};
    let mitSeminar = 0;
    list.forEach((it) => {
        const k = String(it.klasse || '–').trim() || '–';
        const v = String(it.variante || '–').trim() || '–';
        const w = String(it.wahlfachMuendlich || it.wahlfach || '–').trim() || '–';
        byKlasse[k] = (byKlasse[k] || 0) + 1;
        byVariante[v] = (byVariante[v] || 0) + 1;
        byWahlfach[w] = (byWahlfach[w] || 0) + 1;
        if (isSeminarWahlfach(w) || String(it.seminar || '').trim()) mitSeminar += 1;
    });
    const sortBag = (obj) =>
        Object.keys(obj)
            .sort((a, b) => a.localeCompare(b, 'de'))
            .map((key) => ({ key, count: obj[key] }));
    return {
        total: list.length,
        byKlasse: sortBag(byKlasse),
        byVariante: sortBag(byVariante),
        byWahlfach: sortBag(byWahlfach),
        mitSeminar
    };
}

/**
 * Gruppiert Anmeldungen für Liste und Auswertung.
 * @param {object[]} items
 * @param {'klasse'|'variante'|'wahlfach'|string} by
 * @returns {{ key: string, items: object[] }[]}
 */
export function groupItemsBy(items, by) {
    const mode = by === 'variante' || by === 'wahlfach' || by === 'klasse' ? by : '';
    if (!mode) return [];
    const keyOf = (it) => {
        if (mode === 'variante') return String(it.variante || '').trim() || 'Ohne Variante';
        if (mode === 'wahlfach') {
            return String(it.wahlfachMuendlich || it.wahlfach || '').trim() || 'Ohne Wahlfach';
        }
        return String(it.klasse || '').trim() || 'Ohne Klasse';
    };
    /** @type {Map<string, object[]>} */
    const map = new Map();
    (Array.isArray(items) ? items : []).forEach((it) => {
        const key = keyOf(it);
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(it);
    });
    return [...map.entries()]
        .sort((a, b) => a[0].localeCompare(b[0], 'de'))
        .map(([key, rows]) => ({ key, items: rows }));
}

/**
 * Zusammenfassungstexte für Schüler-Erfolgsseite.
 * @param {object} item
 * @param {string} [profileId]
 */
export function buildAnmeldungSummary(item, profileId) {
    const it = item || {};
    const pruefplan = buildPruefplanKurz(it.variante, profileId) || String(it.pruefplanKurz || '').trim();
    /** @type {{ label: string, value: string }[]} */
    const rows = [
        { label: 'Name', value: buildItemTitle(it.nachname, it.vorname) || String(it.title || '').trim() },
        { label: 'Klasse', value: String(it.klasse || '').trim() },
        { label: 'Variante', value: String(it.variante || '').trim() },
        { label: 'Prüfplan', value: pruefplan },
        {
            label: 'Diplomarbeit / ABA',
            value: String(it.titelDiplomarbeit || it.titelABA || '').trim()
        },
        {
            label: 'Betreuung',
            value: String(it.betreuungslehrerDA || it.betreuungslehrerABA || '').trim()
        },
        { label: 'LFS', value: String(it.lfs || '').trim() },
        { label: 'LehrerIn LFS', value: String(it.lehrerLFS || '').trim() },
        { label: 'LehrerIn BKO', value: String(it.lehrerBKO || '').trim() },
        {
            label: 'Wahlfach mündlich',
            value: String(it.wahlfachMuendlich || it.wahlfach || '').trim()
        },
        { label: 'Seminar', value: String(it.seminar || '').trim() },
        { label: 'LehrerIn Wahlfach', value: String(it.lehrerWahlfach || '').trim() }
    ].filter((r) => r.value);
    return { rows, pruefplan, title: rows[0] ? rows[0].value : '' };
}

/**
 * Mehrfachfilter: `;` / `,` / Zeilen → Codes.
 * @param {string|string[]|null|undefined} value
 * @returns {string[]}
 */
export function parseCodeList(value) {
    if (Array.isArray(value)) {
        return value.map((s) => String(s || '').trim()).filter(Boolean);
    }
    return String(value || '')
        .split(/[;,\n]+/)
        .map((s) => s.trim())
        .filter(Boolean);
}

/**
 * @param {string[]|string|null|undefined} value
 */
export function joinCodeList(value) {
    if (Array.isArray(value)) {
        return value
            .map((s) => String(s || '').trim())
            .filter(Boolean)
            .join('; ');
    }
    return parseCodeList(value).join('; ');
}

/**
 * Anzeigename → Vorname / Nachname (MSAL meist „Vorname Nachname“).
 * @param {string} displayName
 * @returns {{ vorname: string, nachname: string }}
 */
export function splitPersonName(displayName) {
    const raw = String(displayName || '').trim();
    if (!raw) return { vorname: '', nachname: '' };
    if (raw.indexOf(',') >= 0) {
        const parts = raw.split(',').map((p) => p.trim()).filter(Boolean);
        return {
            nachname: parts[0] || '',
            vorname: parts.slice(1).join(' ')
        };
    }
    const parts = raw.split(/\s+/).filter(Boolean);
    if (parts.length === 1) return { vorname: '', nachname: parts[0] };
    return {
        vorname: parts[0],
        nachname: parts.slice(1).join(' ')
    };
}

/**
 * Ownership für Schüler-Anmeldung (E-Mail Pflicht, damit „Meine Anmeldung“ bleibt).
 * @param {object} draft
 * @param {{ accountEmail?: string, accountName?: string, students?: object[] }} ctx
 */
export function resolveSrdpOwnership(draft, ctx) {
    const d = draft && typeof draft === 'object' ? draft : {};
    const students = Array.isArray(ctx && ctx.students) ? ctx.students : [];
    const account = String((ctx && ctx.accountEmail) || '')
        .trim()
        .toLowerCase();

    let schuelerEmail = String(d.schuelerEmail || d.beantragtVon || d.createdBy || '')
        .trim()
        .toLowerCase();
    if (!schuelerEmail) schuelerEmail = account;

    const hits = schuelerEmail ? students.filter((s) => studentEmailHits(s, schuelerEmail)) : [];
    const byEmail = hits.find((s) => studentRowKlasse(s)) || hits[0] || null;

    const fromAccount = splitPersonName((ctx && ctx.accountName) || '');
    const nachname = String(
        d.nachname ||
            (byEmail && (byEmail.nachname || byEmail.lastName)) ||
            fromAccount.nachname ||
            ''
    ).trim();
    const vorname = String(
        d.vorname ||
            (byEmail && (byEmail.vorname || byEmail.firstName)) ||
            fromAccount.vorname ||
            ''
    ).trim();
    const klasse = String(
        d.klasse ||
            studentRowKlasse(byEmail) ||
            resolveStudentAssignedKlasse({
                email: schuelerEmail,
                students,
                classes: (ctx && ctx.classes) || [],
                classLinks: (ctx && ctx.classLinks) || [],
                memberGroups: (ctx && ctx.memberGroups) || [],
                preferredCodes: (ctx && ctx.preferredCodes) || []
            })
    ).trim();

    return {
        schuelerEmail,
        beantragtVon: schuelerEmail,
        nachname,
        vorname,
        klasse
    };
}

/**
 * @param {{ schuelerEmail?: string, beantragtVon?: string }} ownership
 */
export function assertSrdpOwnership(ownership) {
    const email = String((ownership && (ownership.schuelerEmail || ownership.beantragtVon)) || '')
        .trim()
        .toLowerCase();
    if (!email || email.indexOf('@') < 0) {
        throw new Error(
            'Schüler-E-Mail fehlt (Anmeldung nicht erkannt). Bitte anmelden und erneut absenden.'
        );
    }
}

/**
 * Anmeldefenster offen? Ohne Frist = offen (Setup noch nicht gesetzt).
 * Mit `anmeldeFrist` (ISO-Datum): offen solange heute ≤ Frist.
 * @param {{ anmeldeFrist?: string, anmeldeVon?: string }|null} options
 * @param {Date|string} [now]
 * @returns {boolean}
 */
export function isSrdpAnmeldeFensterOffen(options, now = new Date()) {
    const o = options || {};
    const bis = String(o.anmeldeFrist || '')
        .trim()
        .slice(0, 10);
    if (!bis || !/^\d{4}-\d{2}-\d{2}$/.test(bis)) return true;
    const von = String(o.anmeldeVon || '')
        .trim()
        .slice(0, 10);
    const today =
        now instanceof Date && !Number.isNaN(now.getTime())
            ? `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
                  now.getDate()
              ).padStart(2, '0')}`
            : String(now || '')
                  .trim()
                  .slice(0, 10);
    if (!today || !/^\d{4}-\d{2}-\d{2}$/.test(today)) return true;
    if (von && /^\d{4}-\d{2}-\d{2}$/.test(von) && today < von) return false;
    return today <= bis;
}

/**
 * Bereits vorhandene Anmeldung derselben Schüler-E-Mail.
 * @param {object[]} items
 * @param {string} email
 */
export function findAnmeldungBySchuelerEmail(items, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return null;
    return (
        (Array.isArray(items) ? items : []).find((it) => {
            const owned = String(it.schuelerEmail || it.beantragtVon || it.createdBy || '')
                .trim()
                .toLowerCase();
            return owned && owned === em;
        }) || null
    );
}

/**
 * Bestätigung (Schüler) ≠ organisatorische Nacharbeit (Sekretariat/Prüfungsliste).
 * @param {object[]} items
 */
export function countNacharbeitPending(items) {
    return (Array.isArray(items) ? items : []).filter((it) => {
        if (!it || !it.bestaetigung) return false;
        return !it.nacharbeitErledigt;
    }).length;
}

/**
 * Klassen ohne Anmeldung (Admin-Aufmerksamkeit).
 * @param {object[]} items
 * @param {string[]} klassen
 */
export function klassenOhneAnmeldung(items, klassen) {
    const want = (Array.isArray(klassen) ? klassen : []).map((k) => String(k || '').trim()).filter(Boolean);
    const have = new Set(
        (Array.isArray(items) ? items : []).map((it) => String(it.klasse || '').trim()).filter(Boolean)
    );
    return want.filter((k) => !have.has(k));
}

export { VARIANT_LABELS };
export { PROFILE_HAK };
