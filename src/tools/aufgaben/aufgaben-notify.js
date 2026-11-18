/**
 * Benachrichtigungen Aufgaben: mailto (sofort) + Hinweise für PA-Flow.
 * Kein Graph Mail.Send – bewusst clientseitig / Power Automate.
 */

import { STATUS, STATUS_LABEL, normalizeAufgabe } from './aufgaben-model.js';

/**
 * @param {string} to
 * @param {string} subject
 * @param {string} body
 */
export function buildMailtoHref(to, subject, body) {
    const em = String(to || '')
        .trim()
        .toLowerCase();
    if (!em || !em.includes('@')) return '';
    return (
        'mailto:' +
        em +
        '?subject=' +
        encodeURIComponent(subject || '') +
        '&body=' +
        encodeURIComponent(body || '')
    );
}

/**
 * Mail an Zugewiesene: neue Aufgabe.
 * @param {object} aufgabe
 * @param {{ fromName?: string }} [opts]
 */
export function buildZuweisungMailto(aufgabe, opts = {}) {
    const a = normalizeAufgabe(aufgabe);
    const to = a.zuweisungEmail;
    if (!to) return '';
    const from = String(opts.fromName || a.erstelltVon || '').trim();
    const subject = 'Aufgabe zugewiesen: ' + (a.title || 'ohne Titel');
    const body =
        'Guten Tag' +
        (a.zuweisungName ? ' ' + a.zuweisungName : '') +
        ',\n\n' +
        'Ihnen wurde eine Aufgabe zugewiesen:\n\n' +
        'Titel: ' +
        (a.title || '–') +
        '\n' +
        (a.beschreibung ? 'Beschreibung: ' + a.beschreibung + '\n' : '') +
        (a.frist ? 'Frist: ' + a.frist + '\n' : '') +
        'Priorität: ' +
        (a.prioritaet || '–') +
        '\n' +
        (from ? '\nVon: ' + from + '\n' : '') +
        '\nBitte im Schul-Portal unter „Aufgaben“ öffnen und Status setzen.\n';
    return buildMailtoHref(to, subject, body);
}

/**
 * Mail an Ersteller: erledigt / abgelehnt.
 * @param {object} aufgabe
 * @param {{ actorName?: string }} [opts]
 */
export function buildStatusMailto(aufgabe, opts = {}) {
    const a = normalizeAufgabe(aufgabe);
    if (a.status !== STATUS.ERLEDIGT && a.status !== STATUS.ABGELEHNT) return '';
    const to = a.erstelltVonEmail;
    if (!to) return '';
    const actor = String(opts.actorName || a.bearbeitetVonName || a.zuweisungName || '').trim();
    const label = STATUS_LABEL[a.status] || a.status;
    const subject = 'Aufgabe ' + label.toLowerCase() + ': ' + (a.title || 'ohne Titel');
    const body =
        'Guten Tag' +
        (a.erstelltVon ? ' ' + a.erstelltVon : '') +
        ',\n\n' +
        'Die Aufgabe „' +
        (a.title || '–') +
        '“ ist jetzt: ' +
        label +
        '.\n' +
        (actor ? 'Bearbeitet von: ' + actor + '\n' : '') +
        (a.kommentar ? '\nKommentar:\n' + a.kommentar + '\n' : '') +
        '\nDetails im Schul-Portal unter „Aufgaben“.\n';
    return buildMailtoHref(to, subject, body);
}

/**
 * Öffnet mailto im Browser (kein Popup-Blocker bei User-Geste).
 * @param {string} href
 * @returns {boolean}
 */
export function openMailto(href) {
    const h = String(href || '').trim();
    if (!h || h.indexOf('mailto:') !== 0) return false;
    try {
        window.location.href = h;
        return true;
    } catch {
        return false;
    }
}
