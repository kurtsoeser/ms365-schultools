/**
 * Planung: gemischte Verwaltungs-Sammelgruppe → Schulleitung + Verwaltung (Personal).
 */
import { diffMemberships, normEmailList } from '../../shared/membership-reconcile.js';

/**
 * @param {{
 *   verwaltungGroupId?: string|null,
 *   schulleitungGroupId?: string|null,
 *   completedAt?: string|null,
 *   skippedAt?: string|null,
 *   schulleitungEmails?: string[],
 *   verwaltungEmails?: string[],
 *   graphMembersVerwaltung?: string[]|null
 * }} input
 */
export function assessVerwaltungSplitMigration(input) {
    const o = input && typeof input === 'object' ? input : {};
    if (o.completedAt) {
        return { phase: 'done', showBanner: false, reason: 'completed' };
    }
    const vwG = String(o.verwaltungGroupId || '').trim();
    const slG = String(o.schulleitungGroupId || '').trim();
    const sl = normEmailList(o.schulleitungEmails || []);
    const vw = normEmailList(o.verwaltungEmails || []);
    if (!vwG) {
        return { phase: 'none', showBanner: false, reason: 'no_verwaltung_group' };
    }
    const graphVw = o.graphMembersVerwaltung == null ? null : normEmailList(o.graphMembersVerwaltung);
    const slInVwGraph = graphVw ? sl.filter(function (em) { return graphVw.indexOf(em) >= 0; }) : [];
    const needsSplit =
        !slG ||
        (graphVw && sl.length && slInVwGraph.length > 0) ||
        (graphVw && sl.length && graphVw.length > vw.length + 2);
    if (!needsSplit && slG) {
        return { phase: 'none', showBanner: false, reason: 'already_split' };
    }
    if (!slG) {
        return {
            phase: 'need_schulleitung_group',
            showBanner: !o.skippedAt,
            reason: 'missing_schulleitung_group',
            schulleitungCount: sl.length,
            verwaltungCount: vw.length,
            schulleitungInVerwaltungGroup: slInVwGraph.length
        };
    }
    return {
        phase: 'ready',
        showBanner: !o.skippedAt,
        reason: slInVwGraph.length ? 'schulleitung_in_verwaltung_group' : 'schulleitung_group_new',
        schulleitungCount: sl.length,
        verwaltungCount: vw.length,
        schulleitungInVerwaltungGroup: slInVwGraph.length
    };
}

/**
 * @param {string[]} schulleitungEmails
 * @param {string[]} verwaltungEmails
 * @param {string[]} graphMembersVerwaltung
 * @param {string[]} graphMembersSchulleitung
 */
export function buildVerwaltungSplitExecutionPlan(
    schulleitungEmails,
    verwaltungEmails,
    graphMembersVerwaltung,
    graphMembersSchulleitung
) {
    const sl = normEmailList(schulleitungEmails);
    const vw = normEmailList(verwaltungEmails);
    const gVw = normEmailList(graphMembersVerwaltung);
    const gSl = normEmailList(graphMembersSchulleitung);
    const slPlan = diffMemberships(sl, gSl);
    const vwPlan = diffMemberships(vw, gVw);
    const schulleitungStillInVw = sl.filter(function (em) {
        return gVw.indexOf(em) >= 0;
    });
    return {
        schulleitung: { join: slPlan.onlyLocal, leave: slPlan.onlyGraph },
        verwaltung: { join: vwPlan.onlyLocal, leave: vwPlan.onlyGraph },
        schulleitungStillInVerwaltungGroup: schulleitungStillInVw
    };
}
