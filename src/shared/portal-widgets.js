/**
 * Macht aus den geladenen Werkzeug-Daten die auswählbaren Dashboard-Kacheln.
 */
import { buildTermineBriefing, buildTermineHeute, buildTermineWeek } from '../tools/termine-hub/termine-hub-briefing.js';
import { DASHBOARD_TOOL_LINKS } from './dashboard-audience-catalog.js';
import { PORTAL_WIDGETS } from './portal-widget-catalog.js';

export const PORTAL_WIDGET_PENDING = 'pending';

/**
 * @param {import('./portal-widget-catalog.js').PORTAL_WIDGETS[number]} widget
 * @param {string} emptyText
 */
function emptyWidget(widget, emptyText) {
    return {
        widgetId: widget.id,
        toolId: widget.toolId,
        title: widget.title,
        href: DASHBOARD_TOOL_LINKS[widget.toolId] || '#',
        presentation: widget.presentation || 'list',
        lines: [],
        days: [],
        stat: '0',
        statLabel: emptyText,
        summary: '',
        attention: false,
        emptyText
    };
}

/**
 * @param {object|null} card
 * @param {import('./portal-widget-catalog.js').PORTAL_WIDGETS[number]} widget
 * @param {string} emptyText
 */
function asList(card, widget, emptyText) {
    const base = card || emptyWidget(widget, emptyText);
    return {
        ...base,
        widgetId: widget.id,
        toolId: widget.toolId,
        title: widget.title,
        presentation: 'list',
        emptyText: base.lines && base.lines.length ? '' : base.emptyText || emptyText
    };
}

/**
 * @param {object|null} card
 * @param {import('./portal-widget-catalog.js').PORTAL_WIDGETS[number]} widget
 * @param {string} emptyText
 */
function asStat(card, widget, emptyText) {
    const n = card && Array.isArray(card.lines) ? card.lines.length : 0;
    return {
        widgetId: widget.id,
        toolId: widget.toolId,
        title: widget.title,
        href: (card && card.href) || DASHBOARD_TOOL_LINKS[widget.toolId] || '#',
        presentation: 'stat',
        stat: String(n),
        statLabel: n && card && card.summary ? card.summary : emptyText,
        lines: [],
        summary: '',
        attention: !!(n && card && card.attention),
        emptyText: ''
    };
}

/**
 * @param {{
 *   view?: string,
 *   termine?: object[]|'pending'|undefined,
 *   aktivitaeten?: object|null|'pending'|undefined,
 *   freistellung?: object|null|'pending'|undefined,
 *   aufgaben?: object|null|'pending'|undefined,
 *   schularbeiten?: object|null|'pending'|undefined,
 *   technik?: object|null|'pending'|undefined,
 *   elternbrief?: object|null|'pending'|undefined
 * }} bag
 */
export function expandPortalWidgets(bag) {
    const data = bag || {};
    const view = data.view === 'schueler' ? 'schueler' : 'lehrer';
    /** @type {object[]} */
    const cards = [];
    const widget = (id) => PORTAL_WIDGETS.find((item) => item.id === id);

    if (data.termine !== undefined && data.termine !== PORTAL_WIDGET_PENDING) {
        const events = Array.isArray(data.termine) ? data.termine : [];
        const now = data.now instanceof Date ? data.now : new Date();
        const liste = buildTermineBriefing(events, { view, max: 8, now });
        const heute = buildTermineHeute(events, { view, now });
        const woche = buildTermineWeek(events, { view, now });
        cards.push(asList(heute, widget('termine-heute'), 'Heute nichts im Kalender'));
        cards.push(asList(liste, widget('termine-liste'), 'Keine Termine in den nächsten 14 Tagen'));
        cards.push({ ...woche, widgetId: 'termine-woche', title: 'Diese Woche', presentation: 'week' });
        cards.push(asStat(liste, widget('termine-zahl'), 'Keine in 14 Tagen'));
    }

    const lists = [
        ['aktivitaeten', 'aktivitaeten-liste', 'aktivitaeten-zahl', 'Nichts offen'],
        ['freistellung', 'freistellung-liste', 'freistellung-zahl', 'Nichts offen'],
        ['aufgaben', 'aufgaben-liste', 'aufgaben-zahl', 'Nichts offen'],
        ['schularbeiten', 'schularbeiten-liste', 'schularbeiten-zahl', 'Nichts anstehend'],
        ['technik', 'technik-liste', 'technik-zahl', 'Nichts offen'],
        ['elternbrief', 'elternbrief-liste', 'elternbrief-zahl', 'Nichts offen']
    ];
    lists.forEach(([key, listId, statId, emptyText]) => {
        if (data[key] === undefined || data[key] === PORTAL_WIDGET_PENDING) return;
        const card = data[key] && typeof data[key] === 'object' ? data[key] : null;
        cards.push(asList(card, widget(listId), emptyText));
        cards.push(asStat(card, widget(statId), emptyText));
    });
    return cards.filter(Boolean);
}
