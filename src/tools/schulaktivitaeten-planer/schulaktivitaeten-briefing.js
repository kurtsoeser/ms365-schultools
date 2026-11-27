/**
 * Portal-Startseite: eigene Schulaktivitäten und offene Entscheidungen dieser App-Rolle.
 */
import { DASHBOARD_TOOL_LINKS } from '../../shared/dashboard-audience-catalog.js';
import { klasseCodesOf } from './schulaktivitaeten-planer-logic.js';
import { formatDeDate } from './schulaktivitaeten-planer-state.js';

export const AKT_BRIEFING_MAX = 5;

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
 * @param {object} it
 * @param {string} email
 */
function isOwn(it, email) {
    if (!email) return false;
    const von = String((it && it.beantragtVon) || '')
        .trim()
        .toLowerCase();
    const lehrer = String((it && it.lehrerEmail) || '')
        .trim()
        .toLowerCase();
    return von === email || lehrer === email;
}

/**
 * @param {object} it
 * @param {string[]} codes
 */
function touchesClass(it, codes) {
    const wanted = new Set((codes || []).map((c) => String(c || '').trim().toUpperCase()).filter(Boolean));
    if (!wanted.size) return false;
    return klasseCodesOf(it).some((c) => wanted.has(String(c || '').trim().toUpperCase()));
}

/**
 * @param {object} it
 */
function itemKey(it) {
    const id = String((it && (it.aktivitaetId || it.itemId)) || '').trim();
    if (id) return id;
    return [it && it.titel, it && it.startdatum, it && it.status].join('|');
}

/**
 * @param {object[]} items
 * @param {{
 *   view?: string,
 *   email?: string,
 *   roles?: string[],
 *   kvKlassen?: string[],
 *   klasseCode?: string,
 *   today?: string,
 *   now?: Date,
 *   max?: number
 * }} [opts]
 * @returns {object|null}
 */
export function buildSchulaktivitaetenBriefing(items, opts) {
    const o = opts || {};
    const view = o.view === 'schueler' ? 'schueler' : 'lehrer';
    const email = String(o.email || '')
        .trim()
        .toLowerCase();
    const roles = (Array.isArray(o.roles) ? o.roles : [])
        .map((r) => String(r || '').trim().toLowerCase())
        .filter(Boolean);
    const today = String(o.today || briefingTodayIso(o.now) || '').trim();
    const max = Number(o.max) > 0 ? Number(o.max) : AKT_BRIEFING_MAX;
    const list = Array.isArray(items) ? items : [];
    /** @type {Map<string, { it: object, meta: string, rank: number }>} */
    const picked = new Map();

    function add(it, meta, rank) {
        if (!it) return;
        const key = itemKey(it);
        const prev = picked.get(key);
        if (prev && prev.rank <= rank) return;
        picked.set(key, { it, meta, rank });
    }

    if (view === 'schueler') {
        const klasse = String(o.klasseCode || '').trim();
        if (!klasse) return null;
        list.forEach((it) => {
            const st = String(it && it.status ? it.status : '').toLowerCase();
            const start = String((it && it.startdatum) || '');
            if (st !== 'genehmigt') return;
            if (today && start && start < today) return;
            if (!touchesClass(it, [klasse])) return;
            add(it, 'genehmigt · deine Klasse', 0);
        });
    } else {
        if (!email && !roles.length) return null;
        if (roles.includes('admin')) {
            list.forEach((it) => {
                if (String(it && it.status ? it.status : '').toLowerCase() === 'beantragt') {
                    add(it, 'wartet auf Entscheidung', 0);
                }
            });
        }
        if (roles.includes('kv')) {
            const classes = Array.isArray(o.kvKlassen) ? o.kvKlassen : [];
            list.forEach((it) => {
                if (String(it && it.status ? it.status : '').toLowerCase() !== 'kv_stellungnahme') return;
                if (!touchesClass(it, classes)) return;
                add(it, 'Stellungnahme', 1);
            });
        }
        list.forEach((it) => {
            if (!isOwn(it, email)) return;
            const st = String(it && it.status ? it.status : '').toLowerCase();
            const start = String((it && it.startdatum) || '');
            if (st === 'nachforderung') add(it, 'Ergänzung nötig', 2);
            else if (st === 'beantragt' || st === 'kv_stellungnahme') add(it, 'wartet auf Genehmigung', 3);
            else if (st === 'genehmigt' && (!today || !start || start >= today)) add(it, 'genehmigt', 4);
        });
    }

    const rows = Array.from(picked.values()).sort((a, b) => {
        if (a.rank !== b.rank) return a.rank - b.rank;
        return String(a.it.startdatum || '').localeCompare(String(b.it.startdatum || ''));
    });
    if (!rows.length) return null;
    const visible = rows.slice(0, max);
    const attention = visible.some((row) => row.rank <= 1);
    const n = visible.length;
    return {
        toolId: 'schulaktivitaeten-planer',
        title: 'Schulaktivitäten',
        href: DASHBOARD_TOOL_LINKS['schulaktivitaeten-planer'] || 'tools/schulaktivitaeten-planer.html',
        lines: visible.map((row) => ({
            label: String(row.it.titel || '').trim() || 'Aktivität',
            meta: [formatDeDate(row.it.startdatum), row.meta].filter((part) => part && part !== '–').join(' · ')
        })),
        summary: attention ? (n === 1 ? '1 offen' : n + ' offen') : n === 1 ? '1 anstehend' : n + ' anstehend',
        attention,
        emptyText: ''
    };
}
