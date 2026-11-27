/**
 * Portal-Startseite: eigene Elternbriefe, die noch Arbeit oder ungelesene Empfänger haben.
 */
import { DASHBOARD_TOOL_LINKS } from '../../shared/dashboard-audience-catalog.js';
import { progressStats } from './elternbrief-logic.js';
import { STATUS, STATUS_LABEL } from './elternbrief-model.js';

export const ELTERNBRIEF_BRIEFING_MAX = 5;

/**
 * @param {number} n
 */
function offenLabel(n) {
    return n === 1 ? '1 offen' : n + ' offen';
}

/**
 * @param {object} brief
 * @param {object[]} empfaenger
 * @returns {{ rank: number, meta: string }|null}
 */
function openState(brief, empfaenger) {
    const status = String(brief && brief.status ? brief.status : '')
        .trim()
        .toLowerCase();
    if (status === STATUS.ENTWURF) {
        return { rank: 0, meta: STATUS_LABEL.entwurf };
    }
    if (status === STATUS.BEREIT) {
        return { rank: 1, meta: STATUS_LABEL.bereit };
    }
    if (status !== STATUS.GESENDET) return null;
    const stats = progressStats(brief, empfaenger);
    const unread = stats.offen + stats.abgelaufen;
    if (!stats.total || unread <= 0) return null;
    return {
        rank: 2,
        meta: unread + ' von ' + stats.total + ' ungelesen'
    };
}

/**
 * @param {object[]} briefe
 * @param {object[]} empfaenger
 * @param {{ email?: string, max?: number }} [opts]
 * @returns {object|null}
 */
export function buildElternbriefBriefing(briefe, empfaenger, opts) {
    const o = opts || {};
    const email = String(o.email || '')
        .trim()
        .toLowerCase();
    if (!email) return null;
    const max = Number(o.max) > 0 ? Number(o.max) : ELTERNBRIEF_BRIEFING_MAX;
    const rows = (Array.isArray(briefe) ? briefe : [])
        .map((brief) => {
            const owner = String(brief && brief.erstellerEmail ? brief.erstellerEmail : '')
                .trim()
                .toLowerCase();
            if (!owner || owner !== email) return null;
            const state = openState(brief, empfaenger);
            if (!state) return null;
            return { brief, ...state };
        })
        .filter(Boolean)
        .sort((a, b) => {
            if (a.rank !== b.rank) return a.rank - b.rank;
            return String(b.brief.createdAt || '').localeCompare(String(a.brief.createdAt || ''));
        });
    if (!rows.length) return null;
    const lines = rows.slice(0, max).map((row) => ({
        label: String(row.brief.betreff || '').trim() || 'Elternbrief',
        meta: row.meta
    }));
    return {
        toolId: 'elternbrief',
        title: 'Elternbriefe',
        href: DASHBOARD_TOOL_LINKS.elternbrief || 'tools/elternbrief.html',
        lines,
        summary: offenLabel(lines.length),
        attention: true,
        emptyText: ''
    };
}
