/**
 * Domänenlogik Kollegiumsumfragen.
 */

import {
    ANTWORT_STATUS,
    FRAGE_TYP,
    STATUS,
    ZIELGRUPPE,
    normalizeAntwort,
    normalizeUmfrage,
    toIsoDateOnly
} from './kollegium-umfragen-model.js';

/**
 * @param {object} draft
 */
export function validateUmfrage(draft) {
    const errors = [];
    const u = normalizeUmfrage(draft);
    if (!u.titel) errors.push('Titel fehlt.');
    if (!Object.values(ZIELGRUPPE).includes(u.zielgruppe)) errors.push('Zielgruppe ungültig.');
    if (u.zielgruppe === ZIELGRUPPE.ABTEILUNG && !u.abteilung) {
        errors.push('Abteilung fehlt.');
    }
    if (u.zielgruppe === ZIELGRUPPE.MANUELL && !parseMails(u.empfaengerMails).length) {
        errors.push('Mindestens eine Empfänger-Mail nötig.');
    }
    if (!u.fragen.length) errors.push('Mindestens eine Frage nötig.');
    for (const f of u.fragen) {
        if (!f.text) errors.push('Frage ohne Text.');
        if (
            (f.typ === FRAGE_TYP.SINGLE || f.typ === FRAGE_TYP.MULTI) &&
            (!f.optionen || f.optionen.length < 2)
        ) {
            errors.push('Auswahlfrage braucht mind. 2 Optionen: ' + (f.text || f.id));
        }
    }
    return { ok: errors.length === 0, errors };
}

/**
 * @param {string} text
 * @returns {string[]}
 */
export function parseMails(text) {
    return String(text || '')
        .split(/[\s,;]+/)
        .map((s) => s.trim().toLowerCase())
        .filter((s) => s.includes('@'));
}

/**
 * KV-/Abteilungs-/Lehrer-Kontext aus Stammdaten (Klassenvorstand-Mails).
 * @param {string} [email]
 * @returns {{ email: string, isKv: boolean, isTeacher: boolean, abteilung: string, kvKlassen: string[] }}
 */
export function resolveTeacherContext(email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    /** @type {string[]} */
    const kvKlassen = [];
    let isKv = false;
    let isTeacher = false;
    let abteilung = '';
    if (!em) return { email: '', isKv: false, isTeacher: false, abteilung: '', kvKlassen };

    const considerClass = (c) => {
        if (!c || typeof c !== 'object') return;
        const head = String(c.headEmail || c.klassenvorstandEmail || c.kvEmail || '')
            .trim()
            .toLowerCase();
        if (head === em) {
            isKv = true;
            isTeacher = true;
            const code = String(c.code || c.name || '').trim();
            if (code) kvKlassen.push(code);
        }
        if (!abteilung && (c.abteilung || c.department)) {
            /* nur setzen wenn Person der Klasse zugeordnet – hier optional leer */
        }
    };

    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const core = window.ms365TenantSettingsLoad();
            const data = (core && core.data) || core || {};
            for (const c of Array.isArray(data.classes) ? data.classes : []) considerClass(c);
            for (const t of Array.isArray(data.teachers) ? data.teachers : []) {
                const mail = String((t && (t.email || t.mail)) || '')
                    .trim()
                    .toLowerCase();
                if (mail !== em) continue;
                isTeacher = true;
                abteilung = String((t && (t.abteilung || t.department || t.fachgruppe)) || '').trim();
                if (t && (t.isKv === true || /kv|klassenvorstand/i.test(String(t.role || '')))) {
                    isKv = true;
                }
            }
        }
    } catch {
        /* ignore */
    }
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365StammdatenCanonical &&
            typeof window.ms365StammdatenCanonical.listAllClasses === 'function'
        ) {
            for (const c of window.ms365StammdatenCanonical.listAllClasses() || []) considerClass(c);
        }
    } catch {
        /* ignore */
    }
    return { email: em, isKv, isTeacher, abteilung, kvKlassen };
}

/**
 * Darf diese Lehrkraft die Umfrage sehen/beantworten?
 * @param {object} umfrage
 * @param {{ email?: string, isKv?: boolean, isTeacher?: boolean, abteilung?: string }} teacher
 */
export function teacherInAudience(umfrage, teacher) {
    const u = normalizeUmfrage(umfrage);
    if (u.status !== STATUS.AKTIV && u.status !== STATUS.GESCHLOSSEN) return false;
    const email = String((teacher && teacher.email) || '')
        .trim()
        .toLowerCase();
    if (!email) return false;

    const ctx = { ...resolveTeacherContext(email), ...teacher };

    if (u.zielgruppe === ZIELGRUPPE.MANUELL) {
        return parseMails(u.empfaengerMails).includes(email);
    }
    if (u.zielgruppe === ZIELGRUPPE.KV) {
        // Wenn Mails gepflegt: nur diese; sonst Stammdaten-KV (headEmail) / Entra-KV
        const mails = parseMails(u.empfaengerMails);
        if (mails.length) return mails.includes(email);
        return !!(ctx && ctx.isKv === true);
    }
    if (u.zielgruppe === ZIELGRUPPE.ABTEILUNG) {
        const mails = parseMails(u.empfaengerMails);
        if (mails.length) return mails.includes(email);
        const abt = String((ctx && ctx.abteilung) || '')
            .trim()
            .toLowerCase();
        return !!abt && abt === String(u.abteilung || '').toLowerCase();
    }
    // alle_lehrkraefte – Mail-Whitelist ODER nachweisbare Lehrkraft (Stammdaten/Entra), nie „jeder MSAL-User“
    const mails = parseMails(u.empfaengerMails);
    if (mails.length) return mails.includes(email);
    return !!(ctx && ctx.isTeacher === true);
}

/**
 * @param {object} umfrage
 * @param {object} antwortDraft
 */
/**
 * Umfrage noch offen (Status + optionales Fälligkeitsdatum)?
 * @param {object} umfrage
 * @param {Date|string} [now]
 */
export function umfrageNochOffen(umfrage, now = new Date()) {
    const u = normalizeUmfrage(umfrage);
    if (u.status !== STATUS.AKTIV) return false;
    const due = toIsoDateOnly(u.faelligAm);
    if (!due) return true;
    const today =
        now instanceof Date && !Number.isNaN(now.getTime())
            ? toIsoDateOnly(now)
            : toIsoDateOnly(now);
    return !(today && today > due);
}

export function validateAntwortPayload(umfrage, antwortDraft) {
    const errors = [];
    const u = normalizeUmfrage(umfrage);
    const a = normalizeAntwort(antwortDraft);
    if (!a.umfrageId) errors.push('Umfrage-ID fehlt.');
    if (!a.lehrerEmail) errors.push('Anmeldung (E-Mail) fehlt.');
    if (!umfrageNochOffen(u)) {
        errors.push('Die Umfrage ist geschlossen oder die Frist ist abgelaufen.');
    }
    const werte = a.werte || {};
    for (const f of u.fragen) {
        if (!f.pflicht) continue;
        const v = werte[f.id];
        if (f.typ === FRAGE_TYP.DATEI) {
            const meta = v && typeof v === 'object' ? v : null;
            const name = meta && meta.name ? String(meta.name) : String(v || '');
            if (!name && !a.anhangName) {
                errors.push('Datei fehlt: ' + f.text);
            }
            continue;
        }
        if (v == null || v === '') {
            errors.push('Pflichtfrage: ' + f.text);
            continue;
        }
        if (f.typ === FRAGE_TYP.MULTI && (!Array.isArray(v) || !v.length)) {
            errors.push('Pflichtfrage: ' + f.text);
        }
    }
    return { ok: errors.length === 0, errors };
}

/**
 * @param {object} umfrage
 * @param {object[]} antworten
 */
export function progressStats(umfrage, antworten) {
    const id = String((umfrage && umfrage.id) || '').trim();
    const rows = (antworten || []).filter((a) => a && a.umfrageId === id);
    const total = rows.length;
    const abgegeben = rows.filter((a) => a.status === ANTWORT_STATUS.ABGEGEBEN).length;
    const offen = total - abgegeben;
    const pct = total ? Math.round((abgegeben / total) * 100) : 0;
    return { total, offen, abgegeben, pct };
}

/**
 * Auswertung je Frage.
 * @param {object} umfrage
 * @param {object[]} antworten
 * @param {{ anonym?: boolean }} [opts]
 */
export function buildAuswertung(umfrage, antworten, opts = {}) {
    const u = normalizeUmfrage(umfrage);
    const rows = (antworten || []).filter(
        (a) => a && a.umfrageId === u.id && a.status === ANTWORT_STATUS.ABGEGEBEN
    );
    const anonym = opts.anonym != null ? !!opts.anonym : !!u.anonym;

    const fragen = u.fragen.map((f) => {
        /** @type {{ label: string, count: number }[]} */
        const buckets = [];
        /** @type {string[]} */
        const texte = [];
        /** @type {number[]} */
        const skalen = [];

        const bump = (label) => {
            const key = String(label || '').trim() || '(leer)';
            const found = buckets.find((b) => b.label === key);
            if (found) found.count += 1;
            else buckets.push({ label: key, count: 1 });
        };

        for (const r of rows) {
            const v = r.werte && r.werte[f.id];
            if (f.typ === FRAGE_TYP.SINGLE) bump(v);
            else if (f.typ === FRAGE_TYP.MULTI) {
                const arr = Array.isArray(v) ? v : v ? [v] : [];
                for (const x of arr) bump(x);
            } else if (f.typ === FRAGE_TYP.SKALA) {
                const n = Number(v);
                if (!Number.isNaN(n)) {
                    skalen.push(n);
                    bump(String(n));
                }
            } else if (f.typ === FRAGE_TYP.TEXT) {
                const t = String(v || '').trim();
                if (t) texte.push(t);
            } else if (f.typ === FRAGE_TYP.DATEI) {
                const name =
                    v && typeof v === 'object'
                        ? String(v.name || '').trim()
                        : String(v || '').trim();
                if (name) {
                    texte.push(name);
                    bump('Datei vorhanden');
                } else bump('(keine Datei)');
            }
        }

        buckets.sort((a, b) => b.count - a.count);
        const avg =
            skalen.length > 0
                ? Math.round((skalen.reduce((s, n) => s + n, 0) / skalen.length) * 10) / 10
                : null;

        return {
            frageId: f.id,
            text: f.text,
            typ: f.typ,
            buckets,
            texte: anonym ? texte : texte,
            avg,
            antwortenCount: rows.length
        };
    });

    return {
        umfrageId: u.id,
        titel: u.titel,
        anonym,
        abgegeben: rows.length,
        fragen,
        respondents: anonym
            ? []
            : rows.map((r) => ({
                  email: r.lehrerEmail,
                  name: r.lehrerName,
                  abgegebenAm: r.abgegebenAm
              }))
    };
}

/**
 * CSV-Export Auswertung (flach).
 * @param {ReturnType<typeof buildAuswertung>} auswertung
 */
export function auswertungToCsv(auswertung) {
    const lines = ['Frage;Typ;Option/Wert;Anzahl'];
    for (const f of auswertung.fragen || []) {
        if (f.typ === FRAGE_TYP.TEXT) {
            for (const t of f.texte || []) {
                lines.push(
                    [csv(f.text), csv(f.typ), csv(t), '1'].join(';')
                );
            }
        } else {
            for (const b of f.buckets || []) {
                lines.push([csv(f.text), csv(f.typ), csv(b.label), String(b.count)].join(';'));
            }
            if (f.avg != null) {
                lines.push([csv(f.text), csv(f.typ), csv('Durchschnitt'), String(f.avg)].join(';'));
            }
        }
    }
    return lines.join('\n');
}

function csv(v) {
    const s = String(v ?? '');
    if (/[;"\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
}

/**
 * Empfängerzeilen aus Stammdaten-ähnlichem Text: Name;E-Mail[;KV|Abteilung]
 * @param {string} text
 */
export function parseLehrerListe(text) {
    const lines = String(text || '')
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
    /** @type {Array<{ name: string, email: string, isKv: boolean, abteilung: string }>} */
    const out = [];
    for (const line of lines) {
        const parts = line.split(/[;|\t]/).map((p) => p.trim());
        let name = '';
        let email = '';
        let isKv = false;
        let abteilung = '';
        for (const p of parts) {
            if (p.includes('@')) email = p.toLowerCase();
            else if (/^kv$/i.test(p) || /klassenvorstand/i.test(p)) isKv = true;
            else if (!name) name = p;
            else abteilung = p;
        }
        if (!email && name.includes('@')) {
            email = name.toLowerCase();
            name = '';
        }
        if (!email) continue;
        out.push({ name, email, isKv, abteilung });
    }
    return out;
}
