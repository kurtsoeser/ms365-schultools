/**
 * Audit-Spalten für Freistellungs-Statusänderungen (Planer-Fallback / Graph-PATCH).
 */
import { toIsoDateOnly } from './freistellung-planer-logic.js';

/**
 * @param {string} [displayName]
 * @param {string} [email]
 */
export function formatFreistellungAuditorLabel(displayName, email) {
    const name = String(displayName || '').trim();
    const em = String(email || '').trim().toLowerCase();
    if (name && em) return name + ' <' + em + '>';
    return name || em || '';
}

/**
 * @param {object} opts
 * @param {string} opts.status Genehmigt | Abgelehnt | …
 * @param {'kv'|'direktion'|string} [opts.role]
 * @param {string} [opts.actorName]
 * @param {string} [opts.actorEmail]
 * @param {string} [opts.today] yyyy-MM-dd
 */
export function buildFreistellungStatusFields(opts) {
    const o = opts || {};
    const status = String(o.status || '').trim();
    const role = String(o.role || 'kv').toLowerCase();
    const today = toIsoDateOnly(o.today || new Date()) || '';
    const actor = formatFreistellungAuditorLabel(o.actorName, o.actorEmail);
    /** @type {Record<string, string>} */
    const fields = { Status: status };

    if (status === 'Abgelehnt') {
        if (actor) fields.AbgelehntVon = actor;
        if (today) fields.AbgelehntAm = today;
        return fields;
    }

    if (status === 'Genehmigt') {
        if (role === 'direktion') {
            if (actor) fields.GenehmigtVonDirektion = actor;
            if (today) fields.GenehmigtAmDirektion = today;
        } else {
            if (actor) fields.GenehmigtVonKV = actor;
            if (today) fields.GenehmigtAmKV = today;
        }
        return fields;
    }

    return fields;
}
