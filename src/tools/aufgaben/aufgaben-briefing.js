/**
 * Portal-Startseite: offene Aufgaben, die dieser Person zugewiesen sind.
 */
import { DASHBOARD_TOOL_LINKS } from '../../shared/dashboard-audience-catalog.js';
import { filterAufgaben, sortAufgaben, statusLabel } from './aufgaben-logic.js';
import { STATUS } from './aufgaben-model.js';

export const AUFGABEN_BRIEFING_MAX = 5;

/**
 * @param {string} iso
 */
function deDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
    if (!m) return '';
    return m[3] + '.' + m[2] + '.' + m[1];
}

/**
 * @param {object[]} aufgaben
 * @param {{ email?: string, max?: number }} [opts]
 * @returns {object|null}
 */
export function buildAufgabenBriefing(aufgaben, opts) {
    const o = opts || {};
    const email = String(o.email || '')
        .trim()
        .toLowerCase();
    if (!email) return null;
    const max = Number(o.max) > 0 ? Number(o.max) : AUFGABEN_BRIEFING_MAX;
    const mine = sortAufgaben(
        filterAufgaben(aufgaben, { scope: 'mir', meEmail: email })
    ).filter((a) => a && (a.status === STATUS.OFFEN || a.status === STATUS.IN_ARBEIT));
    if (!mine.length) return null;
    const lines = mine.slice(0, max).map((a) => {
        const frist = deDate(a.frist);
        return {
            label: a.title || 'Aufgabe',
            meta: [statusLabel(a.status), frist ? 'bis ' + frist : ''].filter(Boolean).join(' · ')
        };
    });
    return {
        toolId: 'aufgaben',
        title: 'Aufgaben',
        href: DASHBOARD_TOOL_LINKS.aufgaben || 'tools/aufgaben.html',
        lines,
        summary: lines.length === 1 ? '1 offen' : lines.length + ' offen',
        attention: true,
        emptyText: ''
    };
}
