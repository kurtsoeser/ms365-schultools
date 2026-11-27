/**
 * Portal-Startseite: nächste Schularbeiten dieser Person.
 * Lehrkräfte sehen eigene Einträge. Schüler/innen sehen fixierte Termine der eigenen Klasse.
 * Die Schul-Inbox bleibt im Werkzeug.
 */
import { DASHBOARD_TOOL_LINKS } from '../../shared/dashboard-audience-catalog.js';
import { schularbeitDisplayTitle } from './schularbeiten-planer-logic.js';
import { filterSchularbeiten, formatDeDate } from './schularbeiten-planer-state.js';

export const SCHULARBEITEN_BRIEFING_MAX = 5;

/**
 * @param {Date|string|number} [now]
 */
export function briefingTodayIso(now) {
    const d = now instanceof Date ? now : now ? new Date(now) : new Date();
    if (Number.isNaN(d.getTime())) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
}

/**
 * @param {object[]} items
 * @param {{ view?: string, email?: string, klasseCode?: string, teacherCode?: string, schuljahr?: string, today?: string, now?: Date, max?: number }} [opts]
 * @returns {object|null}
 */
export function buildSchularbeitenBriefing(items, opts) {
    const o = opts || {};
    const view = o.view === 'schueler' ? 'schueler' : 'lehrer';
    const email = String(o.email || '')
        .trim()
        .toLowerCase();
    const klasseCode = String(o.klasseCode || '').trim();
    const teacherCode = String(o.teacherCode || '').trim();
    const today = String(o.today || briefingTodayIso(o.now) || '').trim();
    const max = Number(o.max) > 0 ? Number(o.max) : SCHULARBEITEN_BRIEFING_MAX;
    if (view === 'schueler' && !klasseCode) return null;
    if (view !== 'schueler' && !email && !teacherCode) return null;

    const normalized = (Array.isArray(items) ? items : []).map((sa) =>
        Object.assign({}, sa, {
            lehrerEmail: String((sa && sa.lehrerEmail) || '')
                .trim()
                .toLowerCase(),
            beantragtVon: String((sa && sa.beantragtVon) || '')
                .trim()
                .toLowerCase(),
            status: String((sa && sa.status) || '')
                .trim()
                .toLowerCase()
        })
    );
    const scope =
        view === 'schueler'
            ? {
                  role: 'schueler',
                  accountEmail: email,
                  studentMatch: { klasse: klasseCode },
                  schuljahr: o.schuljahr || ''
              }
            : {
                  role: 'lehrer',
                  accountEmail: email,
                  onlyMine: true,
                  teacherMatch: teacherCode ? { code: teacherCode, email: email } : null,
                  schuljahr: o.schuljahr || ''
              };
    const rows = filterSchularbeiten(normalized, {}, scope)
        .filter((sa) => {
            const datum = String(sa.datum || '');
            if (!datum || (today && datum < today)) return false;
            const st = String(sa.status || '').toLowerCase();
            if (view === 'lehrer' && st !== 'beantragt' && st !== 'fixiert') return false;
            return true;
        })
        .sort((a, b) => String(a.datum || '').localeCompare(String(b.datum || '')));
    if (!rows.length) return null;
    const lines = rows.slice(0, max).map((sa) => ({
        label: schularbeitDisplayTitle(sa),
        meta: [formatDeDate(sa.datum), sa.klasseCode, sa.status === 'fixiert' ? 'fixiert' : 'beantragt']
            .filter(Boolean)
            .join(' · ')
    }));
    return {
        toolId: 'schularbeiten-planer',
        title: 'Schularbeiten',
        href: DASHBOARD_TOOL_LINKS['schularbeiten-planer'] || 'tools/schularbeiten-planer.html',
        lines,
        summary: lines.length === 1 ? '1 anstehend' : lines.length + ' anstehend',
        attention: lines.some((line) => String(line.meta).includes('beantragt')),
        emptyText: ''
    };
}
