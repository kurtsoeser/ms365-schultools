/**
 * Schul-Kontext vor Rollenprüfung: IT-Bibliothek-Pull, Site-URL, Planer-Gruppen von SharePoint.
 */
import {
    resolveFreistellungSiteUrl,
    persistSiteUrl,
    loadSetupCfg,
    persistSetupFields,
    isLikelySharePointTenantRoot
} from './freistellung-planer-state.js';
import { syncPlannerPermissionsFromSite } from './freistellung-planer-remote-config.js';
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { tryResolveFrListId } from './freistellung-planer-graph.js';
import { discoverFreistellungPlanerContext } from './freistellung-planer-site-discover.js';

function isLoggedIn() {
    try {
        return typeof window.ms365AuthIsLoggedIn === 'function' && !!window.ms365AuthIsLoggedIn();
    } catch {
        return false;
    }
}

/**
 * Wartet auf Auto-Sync (Pull nach Login), ohne die Planer-Seite neu zu laden.
 * @param {{ reloadOnApply?: boolean, force?: boolean, timeoutMs?: number }} [opts]
 */
export async function waitForStammdatenSpoSessionPull(opts) {
    const options = Object.assign({ reloadOnApply: false, timeoutMs: 25000 }, opts || {});
    if (!isLoggedIn()) return { skipped: true, reason: 'not-logged-in' };
    const auto = typeof window !== 'undefined' ? window.ms365StammdatenSpoAutoSync : null;
    if (!auto || typeof auto.runSessionPull !== 'function') {
        return { skipped: true, reason: 'no-autosync' };
    }
    let result;
    try {
        result = await auto.runSessionPull({
            reloadOnApply: options.reloadOnApply === true,
            force: options.force === true
        });
    } catch (e) {
        result = { error: e && e.message ? e.message : String(e) };
    }
    const deadline = Date.now() + options.timeoutMs;
    while (Date.now() < deadline) {
        const st = typeof auto.getStatus === 'function' ? auto.getStatus() : null;
        if (!st || st.phase !== 'pulling') break;
        await new Promise(function (r) {
            setTimeout(r, 80);
        });
    }
    return result || { skipped: true };
}

/**
 * Richtige Team-Site finden, wenn nur Mandanten-Stammweb im Browser steht.
 * @param {{ siteUrl?: string, listId?: string, listName?: string }} [hints]
 * @returns {Promise<{ siteUrl: string, listId: string }|null>}
 */
export async function resolveFreistellungPlanerSiteAndList(hints) {
    const h = hints || {};
    const setup = loadSetupCfg();
    const listName =
        String(h.listName || setup.listName || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;
    let site = resolveFreistellungSiteUrl(h.siteUrl || '');
    let listId = String(h.listId || setup.listId || '').trim();

    if (site && isLikelySharePointTenantRoot(site)) {
        const setupSite = String(setup.siteUrl || '').trim();
        if (setupSite && !isLikelySharePointTenantRoot(setupSite)) {
            site = setupSite;
        }
    }

    let ok = '';
    if (site) {
        ok = await tryResolveFrListId(site, { listName, listId });
    }
    if (ok) {
        return { siteUrl: site.replace(/\/$/, ''), listId: ok || listId };
    }

    const found = await discoverFreistellungPlanerContext({
        preferSiteUrl: site,
        listId,
        listName
    });
    if (!found) return null;

    persistSiteUrl(found.siteUrl);
    persistSetupFields({
        siteUrl: found.siteUrl,
        listId: found.listId,
        listName: found.listName || listName
    });
    return { siteUrl: found.siteUrl, listId: found.listId };
}

/**
 * @param {{ siteUrl?: string, listId?: string, listName?: string }} [hints]
 */
export async function pullPlannerGroupsFromSharePoint(hints) {
    const resolved = await resolveFreistellungPlanerSiteAndList(hints);
    if (!resolved) return { skipped: true, reason: 'no-site' };
    const site = resolved.siteUrl;
    const listId = resolved.listId;
    const setup = loadSetupCfg();
    const listName =
        String((hints && hints.listName) || setup.listName || LIST_TITLE_DEFAULT).trim() ||
        LIST_TITLE_DEFAULT;
    try {
        await syncPlannerPermissionsFromSite(site, listId || undefined, {
            listName,
            listId: listId || setup.listId || ''
        });
        return { ok: true, siteUrl: site, listId: listId };
    } catch (e) {
        return {
            skipped: true,
            reason: 'sync-failed',
            error: e && e.message ? e.message : String(e)
        };
    }
}

/**
 * Nach Login: Stammdaten-Pull, Site, Planer-Gruppen (Listen-Beschreibung / Site Assets).
 * @param {{ siteUrl?: string, listId?: string, listName?: string }} [hints]
 */
export async function ensureFreistellungPlanerSchoolContext(hints) {
    if (!isLoggedIn()) return { skipped: true, reason: 'not-logged-in' };
    await waitForStammdatenSpoSessionPull({ reloadOnApply: false });
    return pullPlannerGroupsFromSharePoint(hints);
}
