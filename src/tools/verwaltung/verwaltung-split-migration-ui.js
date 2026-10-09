/**
 * UI: Assistent Aufteilen der legacy Verwaltungs-Sammelgruppe.
 */
import {
    assessVerwaltungSplitMigration,
    buildVerwaltungSplitExecutionPlan
} from './verwaltung-split-migration-logic.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {{
 *   getMigrationState: () => object,
 *   getSplitEmails: () => { schulleitung: string[], verwaltung: string[] },
 *   getMatchedGroupIds: () => { schulleitung: string|null, verwaltung: string|null },
 *   getGraphToken: () => Promise<string>,
 *   graphUnifiedGroups: object,
 *   patchMigration: (patch: object) => void,
 *   ensureOwners: (token: string, gid: string) => Promise<unknown>,
 *   toast: (msg: string) => void,
 *   dlgConfirm: (msg: string, opts?: object) => Promise<boolean>,
 *   onFocusSchulleitung: () => void,
 *   onAfterSuccess: () => void
 * }} deps
 */
export function mountVerwaltungSplitMigration(deps) {
    const host = document.getElementById('vwSplitMigrationBanner');
    if (!host || !deps) return { refresh: function () {} };

    let graphVw = null;
    let graphSl = null;
    let busy = false;

    async function loadGraphMembers(gid) {
        if (!gid) return [];
        const token = await deps.getGraphToken();
        const gug = deps.graphUnifiedGroups;
        if (!gug || typeof gug.fetchGroupMembers !== 'function') return [];
        const mem = await gug.fetchGroupMembers(token, gid);
        return (mem.items || [])
            .map(function (m) {
                return String((m && (m.mail || m.userPrincipalName)) || '').trim().toLowerCase();
            })
            .filter(function (em) {
                return em.indexOf('@') !== -1;
            });
    }

    function render() {
        if (busy) return;
        const state = deps.getMigrationState();
        const ids = deps.getMatchedGroupIds();
        const split = deps.getSplitEmails();
        const assessment = assessVerwaltungSplitMigration({
            verwaltungGroupId: ids.verwaltung,
            schulleitungGroupId: ids.schulleitung,
            completedAt: state.completedAt,
            skippedAt: state.skippedAt,
            schulleitungEmails: split.schulleitung,
            verwaltungEmails: split.verwaltung,
            graphMembersVerwaltung: graphVw
        });
        if (!assessment.showBanner) {
            host.hidden = true;
            host.replaceChildren();
            return;
        }
        host.hidden = false;
        const phase = assessment.phase;
        let body = '';
        if (phase === 'need_schulleitung_group') {
            body =
                '<p><strong>Schritt 1:</strong> Links <em>Schulleitung</em> wählen und eine Microsoft-365-Gruppe matchen oder anlegen. ' +
                'Danach können Sie die bisherige gemischte Verwaltungsgruppe aufteilen.</p>' +
                '<p class="muted">Stammdaten: ' +
                String(assessment.schulleitungCount) +
                ' Schulleitung · ' +
                String(assessment.verwaltungCount) +
                ' Verwaltung (Personal).</p>';
        } else {
            body =
                '<p>Die Verwaltungs-Sammelgruppe enthält vermutlich noch <strong>Schulleitung</strong> und <strong>Personal</strong> gemeinsam. ' +
                'Der Assistent synchronisiert die Schulleitung in ihre eigene Gruppe und bereinigt die Verwaltungsgruppe auf Personal.</p>';
            if (graphVw) {
                body +=
                    '<p class="muted">In der Verwaltungsgruppe online: ' +
                    String(graphVw.length) +
                    ' Mitglieder · davon Schulleitung in dieser Gruppe: ' +
                    String(assessment.schulleitungInVerwaltungGroup || 0) +
                    '.</p>';
            } else {
                body += '<p class="muted">Vorschau: Graph-Mitglieder der Verwaltungsgruppe laden.</p>';
            }
        }
        host.innerHTML =
            '<div class="vw-split-banner__inner">' +
            '<div class="vw-split-banner__text">' +
            '<h3 class="vw-split-banner__title"><i class="bi bi-signpost-split" aria-hidden="true"></i> Sammelgruppen aufteilen</h3>' +
            body +
            '</div>' +
            '<div class="vw-split-banner__actions">' +
            (phase === 'need_schulleitung_group'
                ? '<button type="button" class="btn btn-primary" id="vwSplitGoSchulleitung">Schulleitung-Gruppe einrichten</button>'
                : '<button type="button" class="btn" id="vwSplitPreview">Vorschau laden</button>' +
                  '<button type="button" class="btn btn-primary" id="vwSplitRun"' +
                  (graphVw ? '' : ' disabled') +
                  '>Aufteilen &amp; synchronisieren</button>') +
            '<button type="button" class="btn" id="vwSplitSkip">Später</button>' +
            '</div></div>';

        const goSl = document.getElementById('vwSplitGoSchulleitung');
        if (goSl) {
            goSl.addEventListener('click', function () {
                deps.onFocusSchulleitung();
            });
        }
        const btnPreview = document.getElementById('vwSplitPreview');
        if (btnPreview) {
            btnPreview.addEventListener('click', function () {
                void preview();
            });
        }
        const btnRun = document.getElementById('vwSplitRun');
        if (btnRun) {
            btnRun.addEventListener('click', function () {
                void runSplit();
            });
        }
        const btnSkip = document.getElementById('vwSplitSkip');
        if (btnSkip) {
            btnSkip.addEventListener('click', function () {
                deps.patchMigration({ skippedAt: new Date().toISOString() });
                render();
            });
        }
    }

    async function preview() {
        const ids = deps.getMatchedGroupIds();
        if (!ids.verwaltung) {
            deps.toast('Keine Verwaltungsgruppe gematcht.');
            return;
        }
        busy = true;
        try {
            graphVw = await loadGraphMembers(ids.verwaltung);
            graphSl = ids.schulleitung ? await loadGraphMembers(ids.schulleitung) : [];
        } catch (e) {
            deps.toast('Graph: ' + (e.message || e));
        } finally {
            busy = false;
            render();
        }
    }

    async function runSplit() {
        const ids = deps.getMatchedGroupIds();
        const split = deps.getSplitEmails();
        if (!ids.verwaltung || !ids.schulleitung) {
            deps.toast('Bitte Verwaltung- und Schulleitung-Gruppe matchen.');
            return;
        }
        if (!graphVw) {
            await preview();
            if (!graphVw) return;
        }
        const plan = buildVerwaltungSplitExecutionPlan(
            split.schulleitung,
            split.verwaltung,
            graphVw,
            graphSl || []
        );
        const summary =
            'Schulleitung-Gruppe: +' +
            plan.schulleitung.join.length +
            ' / −' +
            plan.schulleitung.leave.length +
            '. Verwaltungsgruppe: +' +
            plan.verwaltung.join.length +
            ' / −' +
            plan.verwaltung.leave.length +
            ' (inkl. Entfernen der Schulleitung aus der Personal-Gruppe).';
        const ok = await deps.dlgConfirm(
            summary + '\n\nFortfahren?',
            { title: 'Sammelgruppen aufteilen', okText: 'Ausführen' }
        );
        if (!ok) return;
        busy = true;
        const gug = deps.graphUnifiedGroups;
        try {
            const token = await deps.getGraphToken();
            const labelSl = 'Schulleitung';
            const labelVw = 'Verwaltung (Personal)';
            if (plan.schulleitung.join.length) {
                await gug.syncEmailsToGroup(token, ids.schulleitung, plan.schulleitung.join, labelSl, null);
            }
            if (plan.schulleitung.leave.length && typeof gug.removeEmailsFromGroup === 'function') {
                await gug.removeEmailsFromGroup(token, ids.schulleitung, plan.schulleitung.leave, labelSl, null);
            }
            if (plan.verwaltung.join.length) {
                await gug.syncEmailsToGroup(token, ids.verwaltung, plan.verwaltung.join, labelVw, null);
            }
            if (plan.verwaltung.leave.length && typeof gug.removeEmailsFromGroup === 'function') {
                await gug.removeEmailsFromGroup(token, ids.verwaltung, plan.verwaltung.leave, labelVw, null);
            }
            if (typeof deps.ensureOwners === 'function') {
                await deps.ensureOwners(token, ids.schulleitung);
                await deps.ensureOwners(token, ids.verwaltung);
            }
            deps.patchMigration({ completedAt: new Date().toISOString(), skippedAt: null });
            deps.toast('Aufteilen abgeschlossen.');
            graphVw = null;
            graphSl = null;
            deps.onAfterSuccess();
            render();
        } catch (e) {
            deps.toast('Fehler: ' + (e.message || e));
        } finally {
            busy = false;
        }
    }

    return {
        refresh: function () {
            render();
        },
        resetPreview: function () {
            graphVw = null;
            graphSl = null;
        }
    };
}

if (typeof window !== 'undefined') {
    window.ms365VerwaltungSplitMigrationUi = { mountVerwaltungSplitMigration };
}
