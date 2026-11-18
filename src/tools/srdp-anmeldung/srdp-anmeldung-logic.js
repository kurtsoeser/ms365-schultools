/**
 * Reine sRDP/sRP-Anmeldungs-Logik (kein DOM/Graph) – Vitest.
 */
import {
    VARIANT_LABELS,
    getProfile,
    PROFILE_HAK
} from './srdp-anmeldung-schema.js';

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
 * @param {Array<{ code?: string, name?: string, year?: string|number }>} classes
 * @param {string|number} terminJahr
 */
export function filterAbschlussklassen(classes, terminJahr) {
    const y = String(terminJahr || '').trim();
    const list = Array.isArray(classes) ? classes : [];
    if (!/^\d{4}$/.test(y)) return [];
    return list
        .filter((c) => String((c && c.year) || '').trim() === y)
        .map((c) => {
            const code = String((c && c.code) || '').trim();
            const name = String((c && c.name) || '').trim();
            return code || name;
        })
        .filter(Boolean);
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

/**
 * @param {object[]} items
 * @param {string[]} klasseCodes
 */
export function filterItemsForKlassen(items, klasseCodes) {
    const set = new Set(
        (Array.isArray(klasseCodes) ? klasseCodes : [])
            .map((c) => String(c || '').trim())
            .filter(Boolean)
    );
    if (!set.size) return [];
    return (Array.isArray(items) ? items : []).filter((it) => set.has(String(it.klasse || '').trim()));
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
 * Ownership für Schüler-Anmeldung (E-Mail Pflicht, damit „Meine Anmeldung“ bleibt).
 * @param {object} draft
 * @param {{ accountEmail?: string, students?: object[] }} ctx
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

    const byEmail =
        schuelerEmail &&
        students.find((s) => {
            const em = String((s && s.email) || '')
                .trim()
                .toLowerCase();
            const upn = String((s && s.upn) || '')
                .trim()
                .toLowerCase();
            return em === schuelerEmail || upn === schuelerEmail;
        });

    const nachname = String(d.nachname || (byEmail && (byEmail.nachname || byEmail.lastName)) || '').trim();
    const vorname = String(d.vorname || (byEmail && (byEmail.vorname || byEmail.firstName)) || '').trim();
    const klasse = String(
        d.klasse || (byEmail && (byEmail.klasse || byEmail.klasseCode || byEmail.classCode)) || ''
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
