'use strict';

const STATUS = Object.freeze({
    VORERFASST: 'Vorerfasst',
    IN_BEARBEITUNG: 'In Bearbeitung',
    VOLLSTAENDIG: 'Vollständig',
    GEPRUEFT: 'Geprüft',
    IMPORTIERT: 'Importiert'
});

const STATUS_LIST = Object.freeze(Object.values(STATUS));

const PARENT_EDITABLE = Object.freeze([
    STATUS.VORERFASST,
    STATUS.IN_BEARBEITUNG,
    STATUS.VOLLSTAENDIG
]);

const ALLOWED_ADMIN_TRANSITIONS = Object.freeze({
    [STATUS.VORERFASST]: Object.freeze([STATUS.IN_BEARBEITUNG]),
    [STATUS.IN_BEARBEITUNG]: Object.freeze([STATUS.VOLLSTAENDIG]),
    [STATUS.VOLLSTAENDIG]: Object.freeze([STATUS.GEPRUEFT]),
    [STATUS.GEPRUEFT]: Object.freeze([STATUS.IMPORTIERT]),
    [STATUS.IMPORTIERT]: Object.freeze([])
});

function normalizeStatus(status) {
    const s = String(status || '').trim();
    return STATUS_LIST.includes(s) ? s : STATUS.VORERFASST;
}

function isParentEditLocked(status) {
    return !PARENT_EDITABLE.includes(normalizeStatus(status));
}

function canAdminTransition(fromStatus, toStatus) {
    const from = normalizeStatus(fromStatus);
    const to = normalizeStatus(toStatus);
    if (from === to) return { ok: true, from, to };
    const allowed = ALLOWED_ADMIN_TRANSITIONS[from] || [];
    if (allowed.includes(to)) return { ok: true, from, to };
    return {
        ok: false,
        code: 'invalid_transition',
        message: `Statuswechsel „${from}“ → „${to}“ ist nicht erlaubt.`,
        from,
        to
    };
}

/**
 * @param {object} record
 * @param {{ actorEmail?: string, actorName?: string, now?: string }} [opts]
 */
function applyGeprueftStamp(record, opts = {}) {
    if (!record || normalizeStatus(record.status) !== STATUS.VOLLSTAENDIG) {
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

function applyImportedStamp(record, opts = {}) {
    if (!record || normalizeStatus(record.status) !== STATUS.GEPRUEFT) {
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

module.exports = {
    STATUS,
    STATUS_LIST,
    isParentEditLocked,
    canAdminTransition,
    applyGeprueftStamp,
    applyImportedStamp,
    normalizeStatus
};
