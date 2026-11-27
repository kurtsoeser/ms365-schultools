/**
 * Portal-Startseite: eigene offene Technik-Meldungen.
 * Die Schul-Inbox bleibt im Werkzeug.
 */
import { DASHBOARD_TOOL_LINKS } from '../../shared/dashboard-audience-catalog.js';
import { filterMeldungen, sortMeldungen } from './technik-meldungen-logic.js';
import { PRIORITAET, PRIORITAET_LABEL, STATUS_LABEL } from './technik-meldungen-model.js';

export const TECHNIK_BRIEFING_MAX = 5;

/**
 * @param {number} n
 */
function offenLabel(n) {
    return n === 1 ? '1 offen' : n + ' offen';
}

/**
 * @param {object[]} meldungen
 * @param {{ email?: string, max?: number }} [opts]
 * @returns {object|null}
 */
export function buildTechnikBriefing(meldungen, opts) {
    const o = opts || {};
    const email = String(o.email || '')
        .trim()
        .toLowerCase();
    if (!email) return null;
    const max = Number(o.max) > 0 ? Number(o.max) : TECHNIK_BRIEFING_MAX;
    const mine = sortMeldungen(
        filterMeldungen(meldungen, { meEmail: email, nurOffen: true })
    );
    if (!mine.length) return null;
    const lines = mine.slice(0, max).map((m) => {
        const room = String(m.raumName || m.raumCode || '').trim();
        const status = STATUS_LABEL[m.status] || '';
        const urgent = m.prioritaet === PRIORITAET.DRINGEND ? PRIORITAET_LABEL.dringend : '';
        return {
            label: String(m.title || '').trim() || 'Meldung',
            meta: [urgent, status, room].filter(Boolean).join(' · ')
        };
    });
    return {
        toolId: 'technik-meldungen',
        title: 'Technik-Meldungen',
        href: DASHBOARD_TOOL_LINKS['technik-meldungen'] || 'tools/technik-meldungen.html',
        lines,
        summary: offenLabel(lines.length),
        attention: true,
        emptyText: ''
    };
}
