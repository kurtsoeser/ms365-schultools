/**
 * Reine Geschäftslogik: GUID/PIN, SVNR, Status, Payload-Normalisierung.
 * Läuft in Browser und Node (Vitest) ohne DOM.
 */

export const STATUS = Object.freeze({
    VORERFASST: 'Vorerfasst',
    IN_BEARBEITUNG: 'In Bearbeitung',
    VOLLSTAENDIG: 'Vollständig',
    /** Sekretariats-Stempel: freigegeben für Sokrates-Export */
    GEPRUEFT: 'Geprüft',
    IMPORTIERT: 'Importiert'
});

export const STATUS_LIST = Object.freeze(Object.values(STATUS));

/** Eltern dürfen Speichern/Absenden nur vor dem Sekretariats-Stempel. */
export const PARENT_EDITABLE_STATUSES = Object.freeze([
    STATUS.VORERFASST,
    STATUS.IN_BEARBEITUNG,
    STATUS.VOLLSTAENDIG
]);

/**
 * Erlaubte Admin-Statusübergänge (kein Skip, kein stiller Rücksprung nach Geprüft/Importiert).
 * Schlüssel = aktueller Status, Wert = erlaubte Ziel-Status.
 */
export const ALLOWED_ADMIN_TRANSITIONS = Object.freeze({
    [STATUS.VORERFASST]: Object.freeze([STATUS.IN_BEARBEITUNG]),
    [STATUS.IN_BEARBEITUNG]: Object.freeze([STATUS.VOLLSTAENDIG]),
    [STATUS.VOLLSTAENDIG]: Object.freeze([STATUS.GEPRUEFT]),
    [STATUS.GEPRUEFT]: Object.freeze([STATUS.IMPORTIERT]),
    [STATUS.IMPORTIERT]: Object.freeze([])
});

const SVNR_WEIGHTS = [3, 7, 9, 5, 8, 4, 2, 1, 6];

/** @returns {string} UUID v4 */
export function createUniqueGuid() {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return crypto.randomUUID();
    }
    const bytes = new Uint8Array(16);
    if (typeof crypto !== 'undefined' && crypto.getRandomValues) {
        crypto.getRandomValues(bytes);
    } else {
        for (let i = 0; i < 16; i += 1) bytes[i] = Math.floor(Math.random() * 256);
    }
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = [...bytes].map((b) => b.toString(16).padStart(2, '0')).join('');
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

/** 6-stelliger numerischer PIN */
export function createAccessPin() {
    const n = new Uint32Array(1);
    if (typeof crypto !== 'undefined' && crypto.getRandomValues) {
        crypto.getRandomValues(n);
    } else {
        n[0] = Math.floor(Math.random() * 0xffffffff);
    }
    return String(100000 + (n[0] % 900000));
}

/**
 * @param {string} pin
 * @param {string} pepper
 * @returns {Promise<string>} hex SHA-256
 */
export async function hashPin(pin, pepper = '') {
    const normalized = String(pin || '').trim();
    const material = `${String(pepper || '')}:${normalized}`;
    const data = new TextEncoder().encode(material);
    const digest = await crypto.subtle.digest('SHA-256', data);
    return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

/**
 * @param {string} pin
 * @param {string} expectedHash
 * @param {string} pepper
 */
export async function verifyPin(pin, expectedHash, pepper = '') {
    if (!expectedHash) return false;
    const actual = await hashPin(pin, pepper);
    return timingSafeEqualHex(actual, String(expectedHash));
}

/** @param {string} a @param {string} b */
function timingSafeEqualHex(a, b) {
    if (a.length !== b.length) return false;
    let diff = 0;
    for (let i = 0; i < a.length; i += 1) {
        diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
    }
    return diff === 0;
}

/** @param {unknown} value */
export function normalizeSvnr(value) {
    return String(value || '').replace(/\D/g, '').slice(0, 10);
}

/**
 * Österreichische SVNR: 10 Ziffern.
 * Stellen 1–3 Laufnummer, Stelle 4 Prüfziffer, 5–10 Geburtsdatum TTMMJJ.
 * Gewichte 3,7,9,5,8,4,2,1,6 auf Stellen 1,2,3,5,6,7,8,9,10; Summe mod 11 = Prüfziffer
 * (Rest 10 → Nummer ungültig).
 * @param {unknown} value
 * @returns {{ ok: boolean, svnr: string, kurz: string, error?: string }}
 */
export function validateSvnr(value) {
    const svnr = normalizeSvnr(value);
    if (!svnr) return { ok: false, svnr: '', kurz: '', error: 'SVNR fehlt.' };
    if (svnr.length !== 10) {
        return { ok: false, svnr, kurz: svnr.slice(0, 4), error: 'SVNR muss genau 10 Ziffern haben.' };
    }
    const digits = [...svnr].map(Number);
    // Stellen 1,2,3,5,6,7,8,9,10 (Index 0,1,2,4,5,6,7,8,9)
    const forCheck = [digits[0], digits[1], digits[2], digits[4], digits[5], digits[6], digits[7], digits[8], digits[9]];
    let sum = 0;
    for (let i = 0; i < 9; i += 1) sum += forCheck[i] * SVNR_WEIGHTS[i];
    const rem = sum % 11;
    if (rem === 10) {
        return { ok: false, svnr, kurz: svnr.slice(0, 4), error: 'SVNR-Prüfziffer ungültig.' };
    }
    if (rem !== digits[3]) {
        return { ok: false, svnr, kurz: svnr.slice(0, 4), error: 'SVNR-Prüfziffer stimmt nicht.' };
    }
    return { ok: true, svnr, kurz: svnr.slice(0, 4) };
}

/** @param {unknown} status */
export function normalizeStatus(status) {
    const s = String(status || '').trim();
    return STATUS_LIST.includes(s) ? s : STATUS.VORERFASST;
}

/** @param {unknown} status */
export function isParentEditLocked(status) {
    return !PARENT_EDITABLE_STATUSES.includes(normalizeStatus(status));
}

/** @param {unknown} status */
export function canMarkGeprueft(status) {
    return normalizeStatus(status) === STATUS.VOLLSTAENDIG;
}

/** @param {unknown} status */
export function canMarkImported(status) {
    return normalizeStatus(status) === STATUS.GEPRUEFT;
}

/**
 * @param {unknown} fromStatus
 * @param {unknown} toStatus
 * @returns {{ ok: boolean, code?: string, message?: string, from: string, to: string }}
 */
export function canAdminTransition(fromStatus, toStatus) {
    const from = normalizeStatus(fromStatus);
    const to = normalizeStatus(toStatus);
    if (from === to) {
        return { ok: true, from, to };
    }
    const allowed = ALLOWED_ADMIN_TRANSITIONS[from] || [];
    if (allowed.includes(to)) {
        return { ok: true, from, to };
    }
    return {
        ok: false,
        code: 'invalid_transition',
        message: `Statuswechsel „${from}“ → „${to}“ ist nicht erlaubt.`,
        from,
        to
    };
}

/**
 * Sekretariats-Stempel „Geprüft“ inkl. Ownership.
 * @param {object} record
 * @param {{ actorEmail?: string, actorName?: string, now?: string }} [opts]
 */
export function applyGeprueftStamp(record, opts = {}) {
    if (!record || !canMarkGeprueft(record.status)) {
        return {
            ok: false,
            code: 'invalid_transition',
            message: 'Nur Datensätze mit Status „Vollständig“ können geprüft werden.',
            record
        };
    }
    const now = opts.now || new Date().toISOString();
    const actorEmail = String(opts.actorEmail || '').trim();
    const actorName = String(opts.actorName || '').trim();
    return {
        ok: true,
        record: {
            ...record,
            status: STATUS.GEPRUEFT,
            reviewedAt: now,
            reviewedBy: actorEmail || actorName || 'sekretariat',
            reviewedByName: actorName || '',
            updatedAt: now
        }
    };
}

/**
 * Nach Sokrates-Export: Geprüft → Importiert.
 * @param {object} record
 * @param {{ now?: string }} [opts]
 */
export function applyImportedStamp(record, opts = {}) {
    if (!record || !canMarkImported(record.status)) {
        return {
            ok: false,
            code: 'invalid_transition',
            message: 'Nur Datensätze mit Status „Geprüft“ können als importiert markiert werden.',
            record
        };
    }
    const now = opts.now || new Date().toISOString();
    return {
        ok: true,
        record: {
            ...record,
            status: STATUS.IMPORTIERT,
            updatedAt: now
        }
    };
}

/**
 * @param {object} [raw]
 * @returns {object}
 */
export function createEmptyPayload(raw = {}) {
    return {
        schueler: {
            familienname: '',
            vorname: '',
            weitereVornamen: '',
            geschlecht: '',
            geburtsdatum: '',
            geburtsort: '',
            geburtsstaat: '',
            staatsbuergerschaft: 'Österreich',
            religion: '',
            alltagssprache: 'Deutsch',
            erstsprache: 'Deutsch',
            svnr: '',
            mail: '',
            mail2: '',
            mobil: '',
            tel2: '',
            tel3: '',
            ...(raw.schueler || {})
        },
        wohnadresse: {
            land: 'A',
            plz: '',
            ort: '',
            gemeindenummer: '',
            strasse: '',
            hausnummer: '',
            ...(raw.wohnadresse || {})
        },
        erziehungsberechtigte: [
            emptyGuardian({ adressart: 'Mutter', anrede: 'Frau', ...(raw.erziehungsberechtigte?.[0] || {}) }),
            emptyGuardian({ adressart: 'Vater', anrede: 'Herr', ...(raw.erziehungsberechtigte?.[1] || {}) })
        ],
        sameAddressAsStudent: raw.sameAddressAsStudent !== false
    };
}

function emptyGuardian(overrides = {}) {
    return {
        adressart: '',
        anrede: '',
        titel: '',
        akadGrad: '',
        vorname: '',
        vorname2: '',
        familienname: '',
        akadGradNach: '',
        geburtsdatum: '',
        land: 'A',
        plz: '',
        ort: '',
        gemeindenummer: '',
        strasse: '',
        hausnummer: '',
        mobil: '',
        tel2: '',
        tel3: '',
        mail: '',
        mail2: '',
        erziehungsberechtigt: 'ja',
        zahlungspflichtig: 'ja',
        entscheideAn: 'ja',
        ...overrides
    };
}

/**
 * @param {unknown} raw
 */
export function normalizePayload(raw) {
    const base = createEmptyPayload(raw && typeof raw === 'object' ? raw : {});
    const sv = validateSvnr(base.schueler.svnr);
    if (sv.ok) base.schueler.svnr = sv.svnr;
    else base.schueler.svnr = normalizeSvnr(base.schueler.svnr);
    if (base.sameAddressAsStudent) {
        for (const g of base.erziehungsberechtigte) {
            g.land = base.wohnadresse.land;
            g.plz = base.wohnadresse.plz;
            g.ort = base.wohnadresse.ort;
            g.gemeindenummer = base.wohnadresse.gemeindenummer;
            g.strasse = base.wohnadresse.strasse;
            g.hausnummer = base.wohnadresse.hausnummer;
        }
    }
    return base;
}

/**
 * @param {object} payload
 * @param {Array<{ id: string, required?: boolean, visible?: boolean }>} fields
 * @param {{ mode?: 'draft' | 'submit' }} [opts]
 */
export function validatePayloadAgainstProfile(payload, fields, opts = {}) {
    const mode = opts.mode || 'draft';
    const data = normalizePayload(payload);
    const errors = [];

    if (mode === 'draft') {
        return { ok: true, errors, payload: data };
    }

    for (const field of fields || []) {
        if (!field || field.visible === false || !field.required) continue;
        const value = readFieldValue(data, field.id);
        if (field.id === 'schueler.svnr') {
            const sv = validateSvnr(value);
            if (!sv.ok) errors.push({ id: field.id, message: sv.error || 'SVNR ungültig.' });
            continue;
        }
        if (isEmptyValue(value)) {
            errors.push({ id: field.id, message: 'Pflichtfeld.' });
        }
    }

    return { ok: errors.length === 0, errors, payload: data };
}

/** @param {object} data @param {string} path */
export function readFieldValue(data, path) {
    const parts = String(path || '').split('.');
    let cur = data;
    for (const p of parts) {
        if (cur == null) return '';
        const idx = /^\d+$/.test(p) ? Number(p) : null;
        cur = idx != null ? cur[idx] : cur[p];
    }
    return cur == null ? '' : cur;
}

function isEmptyValue(value) {
    return String(value ?? '').trim() === '';
}

/**
 * @param {object} record
 */
export function publicRecordView(record) {
    if (!record) return null;
    return {
        id: record.uniqueGuid || record.id,
        status: normalizeStatus(record.status),
        title: record.title || '',
        payload: normalizePayload(record.payload),
        updatedAt: record.updatedAt || null,
        reviewedAt: record.reviewedAt || null,
        reviewedBy: record.reviewedBy || '',
        reviewedByName: record.reviewedByName || '',
        parentEditLocked: isParentEditLocked(record.status)
    };
}

/**
 * Admin-Datensatz anlegen (PIN-Klartext nur einmal zurückgeben).
 * @param {{ title?: string, payload?: object, pepper?: string }} opts
 */
export async function createEnrollmentRecord(opts = {}) {
    const uniqueGuid = createUniqueGuid();
    const accessPin = createAccessPin();
    const accessPinHash = await hashPin(accessPin, opts.pepper || '');
    const payload = normalizePayload(opts.payload);
    const title =
        String(opts.title || '').trim() ||
        [payload.schueler.familienname, payload.schueler.vorname].filter(Boolean).join(', ') ||
        'Neue Anmeldung';
    const now = new Date().toISOString();
    return {
        record: {
            uniqueGuid,
            accessPinHash,
            status: STATUS.VORERFASST,
            title,
            payload,
            failedAttempts: 0,
            lockedUntil: null,
            createdAt: now,
            updatedAt: now
        },
        accessPin
    };
}

/**
 * @param {object} record
 * @param {string} pin
 * @param {{ pepper?: string, maxAttempts?: number, lockMinutes?: number, now?: number }} [opts]
 */
export async function attemptUnlock(record, pin, opts = {}) {
    const maxAttempts = opts.maxAttempts ?? 5;
    const lockMinutes = opts.lockMinutes ?? 15;
    const now = opts.now ?? Date.now();
    if (!record) {
        return { ok: false, code: 'invalid', message: 'Ungültiger Zugang.' };
    }
    if (record.lockedUntil && new Date(record.lockedUntil).getTime() > now) {
        return { ok: false, code: 'locked', message: 'Zugang vorübergehend gesperrt. Bitte später erneut versuchen.' };
    }
    const valid = await verifyPin(pin, record.accessPinHash, opts.pepper || '');
    if (!valid) {
        const failed = Number(record.failedAttempts || 0) + 1;
        const next = { ...record, failedAttempts: failed, updatedAt: new Date(now).toISOString() };
        if (failed >= maxAttempts) {
            next.lockedUntil = new Date(now + lockMinutes * 60_000).toISOString();
            next.failedAttempts = 0;
        }
        return { ok: false, code: 'invalid', message: 'Ungültiger Zugang.', record: next };
    }
    return {
        ok: true,
        record: {
            ...record,
            failedAttempts: 0,
            lockedUntil: null,
            updatedAt: new Date(now).toISOString()
        }
    };
}
