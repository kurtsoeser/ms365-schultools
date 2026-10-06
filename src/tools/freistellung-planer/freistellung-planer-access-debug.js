/**
 * Schrittweise Diagnose: Warum hat dieses Konto (keinen) Planer-Zugriff?
 * Kein DOM – testbar.
 */
import {
    loadEffectivePermissionsConfig,
    entraGroupsConfigured,
    normalizePermissionsConfig
} from './freistellung-planer-permissions.js';
import {
    probePlanerEntraMembership,
    shouldGrantSchuelerViaDashboardAudience
} from './freistellung-planer-entra-role.js';
import { accountIsPlannerUserInList } from './freistellung-planer-direktion-users.js';
import {
    resolveFreistellungSiteUrl,
    loadSetupCfg,
    isLikelyFreistellungStaffAccount
} from './freistellung-planer-state.js';
import {
    fetchPlannerPermissionsFromList,
    fetchPlannerPermissionsFromSite
} from './freistellung-planer-remote-config.js';
import { resolveFrContext, probeFreistellungListRead } from './freistellung-planer-graph.js';
import { resolveFreistellungPlanerSiteAndList } from './freistellung-planer-bootstrap.js';
import { isLikelySharePointTenantRoot } from './freistellung-planer-state.js';
import { LIST_TITLE_DEFAULT } from './freistellung-planer-schema.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import { dashboardAudienceGroupsConfigured } from '../../shared/dashboard-audience-groups-store.js';
import { isReady as isItLibrarySyncReady } from '../../shared/stammdaten-sharepoint-sync-api.js';

/** @typedef {'ok'|'warn'|'fail'|'skip'|'info'} DbgStatus */

/**
 * @param {string} id
 * @param {string} title
 * @param {DbgStatus} status
 * @param {string} [detail]
 */
function step(id, title, status, detail) {
    return {
        id,
        title,
        status,
        detail: detail ? String(detail) : ''
    };
}

export function formatGuidForDebug(guid) {
    const g = String(guid || '').trim();
    if (!g) return '(leer)';
    if (g.length < 12) return g;
    return g.slice(0, 8) + '…' + g.slice(-4);
}

export function isFreistellungAccessDebugEnabled() {
    try {
        const p = new URLSearchParams(typeof window !== 'undefined' ? window.location.search || '' : '');
        return p.get('accessDebug') === '1';
    } catch {
        return false;
    }
}

function isLoggedIn() {
    try {
        return typeof window.ms365AuthIsLoggedIn === 'function' && !!window.ms365AuthIsLoggedIn();
    } catch {
        return false;
    }
}

/**
 * @param {object} state Planer-State (accountEmail, siteUrl, listId, planerAccessDenied, …)
 * @param {{ refreshRemote?: boolean }} [opts]
 */
export async function runFreistellungAccessDiagnostics(state, opts) {
    const options = opts || {};
    /** @type {ReturnType<typeof step>[]} */
    const steps = [];
    const s = state || {};
    const accountEmail = String(s.accountEmail || '').trim().toLowerCase();

    if (!isLoggedIn()) {
        steps.push(step('login', 'Microsoft-Anmeldung', 'fail', 'Nicht angemeldet – zuerst mit Schul-Konto anmelden.'));
        return pack(steps, s);
    }
    steps.push(
        step(
            'login',
            'Microsoft-Anmeldung',
            'ok',
            accountEmail ? 'Angemeldet als ' + accountEmail : 'Angemeldet (E-Mail nicht ermittelt)'
        )
    );

    const perms = loadEffectivePermissionsConfig();
    const entraCfg = entraGroupsConfigured(perms);
    if (!entraCfg) {
        steps.push(
            step(
                'local-groups',
                'Planer-Gruppen-IDs im Browser',
                'fail',
                'Keine vollständigen Entra-Gruppen-IDs lokal (Schüler/KV/Direktion). ' +
                    'IT: Setup „Gruppen speichern“ – Schüler brauchen keine IT-Bibliothek, nur Marker in der Listen-Beschreibung oder SiteAssets-JSON.'
            )
        );
    } else {
        const c = normalizePermissionsConfig(perms);
        steps.push(
            step(
                'local-groups',
                'Planer-Gruppen-IDs im Browser',
                'ok',
                'Schüler ' +
                    formatGuidForDebug(c.groupSchuelerId) +
                    ', KV ' +
                    formatGuidForDebug(c.groupKvId) +
                    ', Direktion ' +
                    formatGuidForDebug(c.groupDirektionId)
            )
        );
    }

    const aud = loadSchoolAudienceGroups();
    const dashAud = dashboardAudienceGroupsConfigured();
    steps.push(
        step(
            'stammdaten-audience',
            'Stammdaten Schüler-Sammelgruppe',
            aud.groupSchuelerId ? 'ok' : 'warn',
            aud.groupSchuelerId
                ? 'ID ' + formatGuidForDebug(aud.groupSchuelerId)
                : 'Keine kanonische Schüler-Gruppe in den Stammdaten lokal – optional, wenn Planer-Gruppe reicht.'
        )
    );
    steps.push(
        step(
            'dashboard-audience',
            'Dashboard-Persona-Gruppen',
            dashAud ? 'info' : 'skip',
            dashAud
                ? 'Konfiguriert – Fallback „Schüler“ möglich, wenn Entra-Planer-Gruppen fehlschlagen.'
                : 'Nicht konfiguriert – kein Dashboard-Fallback.'
        )
    );

    const setup = loadSetupCfg();
    let site = resolveFreistellungSiteUrl(s.siteUrl || '');
    let listId = String(s.listId || setup.listId || '').trim();
    const listName = String(s.listName || setup.listName || LIST_TITLE_DEFAULT).trim() || LIST_TITLE_DEFAULT;

    const resolved = await resolveFreistellungPlanerSiteAndList({
        siteUrl: site,
        listId,
        listName
    });
    if (resolved) {
        site = resolved.siteUrl;
        listId = resolved.listId || listId;
        const viaNote =
            isLikelySharePointTenantRoot(resolveFreistellungSiteUrl('')) && !isLikelySharePointTenantRoot(site)
                ? ' (Team-Site per Suche gefunden – nicht Mandanten-Stammweb)'
                : '';
        steps.push(
            step(
                'site-url',
                'SharePoint-Site (Freistellungen)',
                'ok',
                site + viaNote + (listId ? '\nListen-ID: ' + listId : '')
            )
        );
    } else if (!site) {
        steps.push(
            step(
                'site-url',
                'SharePoint-Site (Freistellungen)',
                'fail',
                'Keine Site-URL und keine Team-Site mit Liste „' +
                    listName +
                    '“ gefunden. IT: Setup-Site speichern oder Schüler braucht Zugriff auf MS365-Schultools.'
            )
        );
    } else {
        steps.push(
            step(
                'site-url',
                'SharePoint-Site (Freistellungen)',
                'fail',
                site +
                    (isLikelySharePointTenantRoot(site)
                        ? ' ist nur die Mandanten-Stammweb – die Liste liegt vermutlich unter /sites/MS365-Schultools.'
                        : ' – Liste „' + listName + '“ hier nicht auffindbar.')
            )
        );
    }

    let itReady = false;
    try {
        itReady = isItLibrarySyncReady();
    } catch {
        itReady = false;
    }
    steps.push(
        step(
            'it-library',
            'IT-Bibliothek (Stammdaten-Sync)',
            itReady ? 'info' : 'skip',
            itReady
                ? 'Verknüpft – optional für Schüler; enthält Schul-Backup (nur IT lesen).'
                : 'Nicht verknüpft / nicht lesbar – für Schüler normal. Gruppen-IDs müssen von der Freistellungsliste kommen.'
        )
    );

    let remoteFromFile = null;
    if (site && listId) {
        try {
            const remoteFromList = await fetchPlannerPermissionsFromList(site, listId);
            const okRemote = entraGroupsConfigured(remoteFromList || {});
            steps.push(
                step(
                    'remote-list-desc',
                    'Gruppen aus Listen-Beschreibung',
                    okRemote ? 'ok' : 'fail',
                    okRemote
                        ? 'Schüler-ID ' + formatGuidForDebug((remoteFromList && remoteFromList.groupSchuelerId) || '')
                        : 'Kein gültiger MS365_FR_GROUPS-Marker oder keine Schüler-Gruppen-ID – IT im Setup speichern.'
                )
            );
        } catch (e) {
            steps.push(
                step(
                    'remote-list-desc',
                    'Gruppen aus Listen-Beschreibung',
                    'fail',
                    (e && e.message ? e.message : String(e)) +
                        ' – Setup: „Gruppen speichern“ / Berechtigungen; Marker MS365_FR_GROUPS in der Beschreibung.'
                )
            );
        }
    } else if (site && !listId) {
        steps.push(
            step(
                'remote-list-desc',
                'Gruppen aus Listen-Beschreibung',
                'warn',
                'Listen-ID unbekannt – Liste „' + listName + '“ nicht aufgelöst (Recht oder falscher Site-URL).'
            )
        );
    }
    if (site) {
        try {
            const packed = await fetchPlannerPermissionsFromSite(site);
            remoteFromFile = packed && packed.permissions ? packed.permissions : null;
        } catch {
            remoteFromFile = null;
        }
    }
    if (site && remoteFromFile && entraGroupsConfigured(remoteFromFile)) {
        steps.push(
            step(
                'remote-site-file',
                'Gruppen aus SiteAssets-Datei',
                'ok',
                'ms365/freistellung-planer-groups.json lesbar; Schüler-ID ' +
                    formatGuidForDebug(remoteFromFile.groupSchuelerId)
            )
        );
    }

    const entraProbe = await probePlanerEntraMembership(perms);
    if (entraProbe.error) {
        steps.push(
            step(
                'entra-membership',
                'Mitgliedschaft Schüler-Entra-Gruppe',
                'fail',
                'Graph-Prüfung fehlgeschlagen: ' + entraProbe.error
            )
        );
    } else if (!entraProbe.allIds.length) {
        steps.push(
            step(
                'entra-membership',
                'Mitgliedschaft Schüler-Entra-Gruppe',
                'skip',
                'Keine Gruppen-IDs zum Prüfen – zuerst Schritt „Planer-Gruppen-IDs“ beheben.'
            )
        );
    } else if (entraProbe.memberSchueler) {
        steps.push(
            step(
                'entra-membership',
                'Mitgliedschaft Schüler-Entra-Gruppe',
                'ok',
                'Konto ist in mindestens einer Schüler-Gruppe (' +
                    entraProbe.schuelerIds.map(formatGuidForDebug).join(', ') +
                    '). Rollen: ' +
                    (entraProbe.roles.length ? entraProbe.roles.join(', ') : '–')
            )
        );
    } else {
        steps.push(
            step(
                'entra-membership',
                'Mitgliedschaft Schüler-Entra-Gruppe',
                'fail',
                'Konto ist in keiner konfigurierten Schüler-Gruppe. Prüfen: Entra-Mitgliedschaft für ' +
                    accountEmail +
                    '; geprüfte IDs: ' +
                    entraProbe.schuelerIds.map(formatGuidForDebug).join(', ')
            )
        );
    }

    const scope = {
        studentMatch: s.studentMatch,
        kvMatch: s.kvMatch,
        direktionMatch: s.direktionMatch,
        accountEmail
    };
    if (accountIsPlannerUserInList(accountEmail, perms.schuelerUsers)) {
        steps.push(step('schueler-user', 'Einzelperson im Setup (schuelerUsers)', 'ok', 'Konto steht in der Schüler-Einzelpersonenliste.'));
    } else {
        steps.push(step('schueler-user', 'Einzelperson im Setup (schuelerUsers)', 'skip', 'Nicht als Einzelperson eingetragen (Sammelgruppe reicht).'));
    }

    let dashGrant = false;
    try {
        dashGrant = await shouldGrantSchuelerViaDashboardAudience(scope);
    } catch {
        dashGrant = false;
    }
    steps.push(
        step(
            'dashboard-fallback',
            'Dashboard-Schüler-Fallback',
            dashGrant ? 'ok' : 'skip',
            dashGrant ? 'Würde Rolle Schüler über Dashboard-Audience vergeben.' : 'Greift nicht (oder Persona nicht Schüler).'
        )
    );

    if (s.studentMatch) {
        steps.push(step('stammdaten-student', 'Schüler in Stammdaten (E-Mail)', 'ok', 'E-Mail in Schüler-Stammdaten gefunden.'));
    } else {
        steps.push(
            step(
                'stammdaten-student',
                'Schüler in Stammdaten (E-Mail)',
                'info',
                'Kein Treffer in lokalen Stammdaten – für Schüler oft OK, wenn Entra-Gruppe + Liste passen.'
            )
        );
    }

    const staffLike = isLikelyFreistellungStaffAccount(s);
    steps.push(
        step(
            'staff-guard',
            'Lehrkraft-/KV-Sperre',
            staffLike ? 'warn' : 'ok',
            staffLike
                ? 'Konto wirkt wie Lehrkraft/Verwaltung – automatische Schüler-Rolle nur über Liste eingeschränkt.'
                : 'Kein Lehrkraft-Signal – Listen-Fallback für Schüler möglich.'
        )
    );

    let listRead = false;
    if (site && listId) {
        try {
            const ctx = await resolveFrContext(site, { listName, listId });
            listRead = await probeFreistellungListRead(ctx);
        } catch {
            listRead = false;
        }
    }
    if (!site) {
        steps.push(step('list-read', 'SharePoint-Liste lesen', 'skip', 'Keine Site-URL.'));
    } else if (listRead) {
        steps.push(step('list-read', 'SharePoint-Liste lesen', 'ok', 'Liste „' + listName + '“ ist für dieses Konto lesbar (Anträge laden/speichern möglich).'));
    } else {
        steps.push(
            step(
                'list-read',
                'SharePoint-Liste lesen',
                'fail',
                'Liste nicht lesbar – im Setup „Berechtigungen“ mit Schüler-Gruppe setzen; ggf. falsche Site/Liste.'
            )
        );
    }

    const appRole = s.planerAccessDenied
        ? 'fail'
        : s.role === 'schueler' || (s.planerRoles || []).includes('schueler')
          ? 'ok'
          : 'warn';
    steps.push(
        step(
            'app-role',
            'App-Entscheidung (aktueller Tab)',
            appRole,
            s.planerAccessDenied
                ? 'Zugriff verweigert. ' + (s.roleHintPublic || s.roleHintStaff || s.roleHint || '')
                : 'Rolle „' +
                  (s.role || '–') +
                  '“; Quellen: ' +
                  JSON.stringify(s.planerRoleSources || {})
        )
    );

    if (options.refreshRemote && site) {
        steps.push(
            step('refresh-note', 'Hinweis', 'info', 'Remote erneut abgefragt (ohne localStorage zu überschreiben).')
        );
    }

    return pack(steps, s);
}

/**
 * @param {ReturnType<typeof step>[]} steps
 * @param {object} state
 */
function pack(steps, state) {
    const fails = steps.filter((x) => x.status === 'fail');
    let summary = '';
    if (!steps.find((x) => x.id === 'login') || steps.find((x) => x.id === 'login' && x.status === 'fail')) {
        summary = 'Zuerst anmelden.';
    } else if (fails.length === 0) {
        summary = state.planerAccessDenied
            ? 'Alle geprüften Schritte OK, aber die App verweigert noch – „Diagnose erneut“ nach Aktualisieren oder IT mit diesem Bericht.'
            : 'Zugriff wirkt freigegeben.';
    } else {
        summary =
            'Erster kritischer Schritt: „' +
            fails[0].title +
            '“ – ' +
            (fails[0].detail || '').slice(0, 200);
    }
    return {
        at: new Date().toISOString(),
        accountEmail: String(state.accountEmail || ''),
        planerAccessDenied: !!state.planerAccessDenied,
        role: state.role || '',
        steps,
        summary
    };
}

export function formatDiagnosticsReport(report) {
    const r = report || { steps: [], summary: '' };
    const lines = [
        'Freistellungs-Planer Zugriffsdiagnose',
        'Zeit: ' + (r.at || ''),
        'Konto: ' + (r.accountEmail || ''),
        'App-Rolle: ' + (r.role || '–') + (r.planerAccessDenied ? ' (verweigert)' : ''),
        'Kurz: ' + (r.summary || ''),
        '',
        'Schritte:'
    ];
    (r.steps || []).forEach(function (st, i) {
        lines.push(
            (i + 1) + '. [' + st.status + '] ' + st.title + (st.detail ? '\n   ' + st.detail : '')
        );
    });
    return lines.join('\n');
}
