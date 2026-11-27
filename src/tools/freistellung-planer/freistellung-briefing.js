/**
 * Portal-Startseite: eigene Freistellungen und offene Entscheidungen dieser App-Rolle.
 */
import { DASHBOARD_TOOL_LINKS } from '../../shared/dashboard-audience-catalog.js';
import { filterFreistellungen } from './freistellung-planer-logic.js';

export const FREISTELLUNG_BRIEFING_MAX = 5;

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
 * @param {unknown} value
 */
function isoDay(value) {
    const m = /^(\d{4}-\d{2}-\d{2})/.exec(String(value || ''));
    return m ? m[1] : '';
}

/**
 * @param {object} it
 */
function shortRange(it) {
    const start = isoDay(it && it.beginn);
    const end = isoDay(it && it.ende) || start;
    if (!start) return '';
    const a = String(Number(start.slice(8, 10))) + '.' + String(Number(start.slice(5, 7))) + '.';
    if (!end || end === start) return a;
    const sameMonth = start.slice(0, 7) === end.slice(0, 7);
    const bDay = String(Number(end.slice(8, 10)));
    const b = sameMonth
        ? bDay + '.' + String(Number(end.slice(5, 7))) + '.'
        : bDay + '.' + String(Number(end.slice(5, 7))) + '.';
    const aDay = String(Number(start.slice(8, 10)));
    return (sameMonth ? aDay : a.replace(/\.$/, '')) + '.–' + b;
}

/**
 * @param {object} it
 */
function itemKey(it) {
    const id = String((it && it.itemId) || '').trim();
    if (id) return id;
    return [it && it.schuelerName, it && it.beginn, it && it.klasse].join('|');
}

/**
 * @param {object} it
 * @param {string} today
 */
function stillAhead(it, today) {
    const end = isoDay(it && (it.ende || it.beginn));
    if (!today || !end) return true;
    return end >= today;
}

/**
 * @param {number} n
 * @param {string} one
 * @param {string} many
 */
function countLabel(n, one, many) {
    return n === 1 ? '1 ' + one : n + ' ' + many;
}

/**
 * @param {object[]} items
 * @param {{
 *   view?: string,
 *   email?: string,
 *   roles?: string[],
 *   kvKlassen?: string[],
 *   today?: string,
 *   now?: Date,
 *   max?: number
 * }} [opts]
 * @returns {object|null}
 */
export function buildFreistellungBriefing(items, opts) {
    const o = opts || {};
    const view = o.view === 'schueler' ? 'schueler' : 'lehrer';
    const email = String(o.email || '')
        .trim()
        .toLowerCase();
    if (!email) return null;
    const roles = (Array.isArray(o.roles) ? o.roles : [])
        .map((r) => String(r || '').trim().toLowerCase())
        .filter(Boolean);
    const today = String(o.today || briefingTodayIso(o.now) || '').trim();
    const max = Number(o.max) > 0 ? Number(o.max) : FREISTELLUNG_BRIEFING_MAX;
    const list = Array.isArray(items) ? items : [];
    const isDirektion = view !== 'schueler' && roles.some((r) => r === 'direktion' || r === 'admin');
    const isKv = view !== 'schueler' && !isDirektion && roles.includes('kv');

    /** @type {Map<string, { it: object, meta: string, rank: number }>} */
    const picked = new Map();
    function add(it, phrase, rank) {
        if (!it) return;
        const key = itemKey(it);
        const prev = picked.get(key);
        if (prev && prev.rank <= rank) return;
        picked.set(key, { it, meta: phrase, rank });
    }

    if (isDirektion) {
        list.forEach((it) => {
            if (String(it && it.status ? it.status : '').trim().toLowerCase() !== 'ausstehend') return;
            add(it, 'wartet auf Entscheidung', 0);
        });
    } else if (isKv) {
        const classes = Array.isArray(o.kvKlassen) ? o.kvKlassen : [];
        filterFreistellungen(list, { status: 'Ausstehend' }, {
            onlyKv: true,
            accountEmail: email,
            kvClassCodes: new Set(classes)
        }).forEach((it) => add(it, 'Stellungnahme', 0));
        filterFreistellungen(list, { status: 'Ausstehend' }, {
            onlyMine: true,
            accountEmail: email
        }).forEach((it) => add(it, 'wartet auf Genehmigung', 1));
    } else {
        filterFreistellungen(list, { status: 'Ausstehend' }, {
            onlyMine: true,
            accountEmail: email
        }).forEach((it) => add(it, 'wartet auf Genehmigung', 0));
        filterFreistellungen(list, { status: 'Genehmigt' }, {
            onlyMine: true,
            accountEmail: email
        }).forEach((it) => {
            if (!stillAhead(it, today)) return;
            add(it, 'genehmigt', 1);
        });
    }

    const rows = Array.from(picked.values()).sort((a, b) => {
        if (a.rank !== b.rank) return a.rank - b.rank;
        return isoDay(a.it.beginn).localeCompare(isoDay(b.it.beginn));
    });
    if (!rows.length) return null;
    const lines = rows.slice(0, max).map((row) => ({
        kicker: shortRange(row.it),
        label: String(row.it.schuelerName || row.it.titel || '').trim() || 'Freistellung',
        meta: [row.it.klasse, isDirektion ? '' : row.meta].filter(Boolean).join(' · ')
    }));
    const attention = rows.some((row) => row.rank === 0);
    const summary = isDirektion
        ? countLabel(lines.length, 'wartet auf Entscheidung', 'warten auf Entscheidung')
        : attention
          ? countLabel(lines.length, 'offen', 'offen')
          : countLabel(lines.length, 'genehmigt', 'genehmigt');
    return {
        toolId: 'freistellung-planer',
        title: 'Freistellungen',
        href: DASHBOARD_TOOL_LINKS['freistellung-planer'] || 'tools/freistellung-planer.html',
        lines,
        summary,
        attention,
        emptyText: ''
    };
}
