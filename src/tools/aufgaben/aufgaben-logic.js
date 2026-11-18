/**
 * Domänenlogik Aufgaben (DOM-frei, Familie betrieb).
 */

import {
    PRIORITAET,
    ROLLE,
    STATUS,
    STATUS_LABEL,
    normalizeAufgabe,
    toIsoDateOnly
} from './aufgaben-model.js';

/**
 * @param {object} draft
 */
export function validateAufgabe(draft) {
    const errors = [];
    const a = normalizeAufgabe(draft);
    if (!a.title) errors.push('Titel fehlt.');
    if (!a.zuweisungName && !a.zuweisungEmail) {
        errors.push('Zuweisung (Name oder E-Mail) fehlt.');
    }
    if (a.rolle && !Object.values(ROLLE).includes(a.rolle)) {
        errors.push('Rolle ungültig.');
    }
    if (!Object.values(PRIORITAET).includes(a.prioritaet)) {
        errors.push('Priorität ungültig.');
    }
    if (!Object.values(STATUS).includes(a.status)) {
        errors.push('Status ungültig.');
    }
    return { ok: errors.length === 0, errors, item: a };
}

/**
 * @param {object[]} aufgaben
 * @param {{ scope?: string, status?: string, q?: string, meEmail?: string, admin?: boolean, showAll?: boolean }} filters
 */
export function filterAufgaben(aufgaben, filters = {}) {
    const scope = String(filters.scope || '')
        .trim()
        .toLowerCase();
    const status = String(filters.status || '')
        .trim()
        .toLowerCase();
    const q = String(filters.q || '')
        .trim()
        .toLowerCase();
    const me = String(filters.meEmail || '')
        .trim()
        .toLowerCase();
    const canSeeAll = !!(filters.admin && filters.showAll);

    return (aufgaben || []).filter((a) => {
        if (!a) return false;
        if (status && a.status !== status) return false;
        if (scope === 'mir') {
            if (!me || a.zuweisungEmail !== me) return false;
        } else if (scope === 'von_mir') {
            if (!me || a.erstelltVonEmail !== me) return false;
        } else if (scope === 'offen') {
            if (a.status !== STATUS.OFFEN && a.status !== STATUS.IN_ARBEIT) return false;
            if (!canSeeAll) {
                if (!me || (a.zuweisungEmail !== me && a.erstelltVonEmail !== me)) return false;
            }
        } else if (!scope) {
            if (!canSeeAll) {
                if (!me || (a.zuweisungEmail !== me && a.erstelltVonEmail !== me)) return false;
            }
        }
        if (q) {
            const hay = (
                a.title +
                ' ' +
                a.beschreibung +
                ' ' +
                a.zuweisungName +
                ' ' +
                a.zuweisungEmail +
                ' ' +
                a.rolle +
                ' ' +
                a.kommentar
            ).toLowerCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });
}

/**
 * Darf Status ändern / kommentieren?
 * Melder (Ersteller) und Bearbeiter (Zuweisung) sowie Admin.
 * @param {object} aufgabe
 * @param {{ email?: string, admin?: boolean }} viewer
 */
export function canManageAufgabe(aufgabe, viewer) {
    if (viewer && viewer.admin) return true;
    const me = String((viewer && viewer.email) || '')
        .trim()
        .toLowerCase();
    if (!me || !aufgabe) return false;
    return (
        String(aufgabe.zuweisungEmail || '').toLowerCase() === me ||
        String(aufgabe.erstelltVonEmail || '').toLowerCase() === me
    );
}

/**
 * Status-Pipeline-Aktionen (offen → in_arbeit → erledigt/abgelehnt).
 * @param {string} status
 * @returns {{ to: string, label: string, kind?: string }[]}
 */
export function statusPipelineActions(status) {
    const st = String(status || '')
        .trim()
        .toLowerCase();
    if (st === STATUS.ERLEDIGT || st === STATUS.ABGELEHNT) {
        return [{ to: STATUS.OFFEN, label: 'Wieder öffnen' }];
    }
    if (st === STATUS.IN_ARBEIT) {
        return [
            { to: STATUS.ERLEDIGT, label: 'Erledigt', kind: 'success' },
            { to: STATUS.ABGELEHNT, label: 'Ablehnen' },
            { to: STATUS.OFFEN, label: 'Zurück auf offen' }
        ];
    }
    return [
        { to: STATUS.IN_ARBEIT, label: 'In Arbeit' },
        { to: STATUS.ERLEDIGT, label: 'Erledigt', kind: 'success' },
        { to: STATUS.ABGELEHNT, label: 'Ablehnen' }
    ];
}

/**
 * @param {object[]} aufgaben
 * @param {string} [meEmail]
 */
export function aufgabenStats(aufgaben, meEmail) {
    const list = aufgaben || [];
    const offen = list.filter((a) => a.status === STATUS.OFFEN || a.status === STATUS.IN_ARBEIT);
    const me = String(meEmail || '')
        .trim()
        .toLowerCase();
    const meineOffen = me ? offen.filter((a) => a.zuweisungEmail === me) : [];
    const heute = toIsoDateOnly(new Date());
    const ueberfaellig = offen.filter((a) => a.frist && a.frist < heute).length;
    const dringend = offen.filter((a) => a.prioritaet === PRIORITAET.DRINGEND).length;
    return {
        total: list.length,
        offen: offen.length,
        inArbeit: list.filter((a) => a.status === STATUS.IN_ARBEIT).length,
        meineOffen: meineOffen.length,
        erledigt: list.filter((a) => a.status === STATUS.ERLEDIGT).length,
        ueberfaellig,
        dringend
    };
}

/**
 * Sortierung: dringend zuerst, dann Frist, dann Titel.
 * @param {object[]} aufgaben
 */
export function sortAufgaben(aufgaben) {
    const prioRank = { dringend: 0, hoch: 1, mittel: 2, niedrig: 3 };
    return [...(aufgaben || [])].sort((a, b) => {
        const pa = prioRank[a.prioritaet] ?? 9;
        const pb = prioRank[b.prioritaet] ?? 9;
        if (pa !== pb) return pa - pb;
        const fa = a.frist || '9999';
        const fb = b.frist || '9999';
        if (fa !== fb) return fa < fb ? -1 : 1;
        return String(a.title || '').localeCompare(String(b.title || ''), 'de');
    });
}

/**
 * @param {string} status
 */
export function statusLabel(status) {
    return STATUS_LABEL[status] || status || '–';
}
