/**
 * Zahlen & Fakten auf Aufgaben-Kacheln („Was möchten Sie tun?“) – angebunden an Register & Hygiene.
 */
import { taskRowToolId } from './dashboard-task-tool-copy.js';
import { resolveAggregateTaskRowFacts, resolveToolTaskRowFacts } from './dashboard-task-row-facts-logic.js';

export {
    classTeamsLinkedCounts,
    hygieneTargetNumericHint,
    registerSnapshotLine,
    resolveAggregateTaskRowFacts,
    resolveToolTaskRowFacts
} from './dashboard-task-row-facts-logic.js';

/**
 * @param {HTMLElement} row
 * @param {{
 *   container: object|null,
 *   settings: object|null,
 *   hygieneById: Record<string, string>|null,
 *   show: boolean
 * }} ctx
 */
export function applyTaskRowFacts(row, ctx) {
    if (!row || !ctx || !ctx.show) {
        if (row) {
            row.removeAttribute('data-dash-fact-desc');
            row.removeAttribute('data-hygiene-hint');
            row.removeAttribute('data-hygiene-tone');
            row.classList.remove('dash-task-row--live-metric');
        }
        return;
    }
    const hygiene = typeof window !== 'undefined' ? window.ms365MembershipHygiene : null;
    const settings = ctx.settings;
    const container = ctx.container;
    const byId = ctx.hygieneById || {};
    const links =
        container && container.setup && Array.isArray(container.setup.catalogLinks)
            ? container.setup.catalogLinks
            : [];

    const hygieneId = String(row.getAttribute('data-dash-hygiene-id') || '').trim();
    const agg = String(row.getAttribute('data-dash-hygiene-aggregate') || '').trim();
    const toolId = taskRowToolId(row);
    const href = row.getAttribute('href') || '';

    let facts = null;
    if (hygieneId || agg) {
        facts = resolveAggregateTaskRowFacts({
            hygieneId: hygieneId,
            aggregate: agg,
            container: container,
            settings: settings,
            hygieneById: byId,
            hygieneApi: hygiene,
            links: links
        });
    }
    if (!facts) {
        facts = resolveToolTaskRowFacts({
            toolId: toolId,
            href: href,
            container: container,
            settings: settings,
            hygieneApi: hygiene
        });
    }
    if (!facts) {
        row.removeAttribute('data-dash-fact-desc');
        row.removeAttribute('data-hygiene-hint');
        row.classList.remove('dash-task-row--live-metric');
        return;
    }

    row.classList.toggle('dash-task-row--live-metric', !!facts.desc);
    if (facts.tone) row.setAttribute('data-hygiene-tone', facts.tone);
    else row.removeAttribute('data-hygiene-tone');
    if (facts.chip) row.setAttribute('data-hygiene-hint', facts.chip);
    else row.removeAttribute('data-hygiene-hint');
    if (facts.desc) row.setAttribute('data-dash-fact-desc', facts.desc);
    else row.removeAttribute('data-dash-fact-desc');
}

/**
 * @param {{
 *   container: object|null,
 *   settings: object|null,
 *   hygieneById: Record<string, string>|null,
 *   show: boolean,
 *   root?: ParentNode
 * }} opts
 */
export function applyAllTaskRowFacts(opts) {
    const root = opts.root || document;
    const ctx = {
        container: opts.container,
        settings: opts.settings,
        hygieneById: opts.hygieneById,
        show: !!opts.show
    };
    root.querySelectorAll('.dash-task-links--stack .dash-task-row').forEach(function (row) {
        applyTaskRowFacts(row, ctx);
    });
}

if (typeof window !== 'undefined') {
    window.ms365ApplyTaskRowFacts = applyTaskRowFacts;
    window.ms365ApplyAllTaskRowFacts = applyAllTaskRowFacts;
}
