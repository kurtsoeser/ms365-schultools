/**
 * Teams-Antrag UI – Antrag, meine Liste, IT-Inbox, Einstellungen.
 */
import {
    fetchGroupById,
    fetchGroupUnifiedAllowedGroupId,
    readGruppenerstellungLocal,
    resolveFreigeberAccess,
    searchSecurityGroups
} from './teams-antrag-freigeber.js';
import { approveAndCreateTeam } from './teams-antrag-graph.js';
import {
    ART,
    ART_LABEL,
    STATUS,
    STATUS_LABEL,
    filterAntraege,
    openQueue,
    suggestDisplayName,
    suggestMailNickname
} from './teams-antrag-logic.js';
import { loadTeamsAntragSettings, saveTeamsAntragSettings } from './teams-antrag-settings.js';
import {
    clearAll,
    getAntrag,
    loadDemoData,
    loadState,
    patchItFields,
    seedDemoIfEmpty,
    setAntragStatus,
    upsertAntrag
} from './teams-antrag-store.js';
import { getDashboardPersonas } from '../../shared/dashboard-persona-session.js';
import { readDemoRoleOverrideFromUrl } from '../../shared/app-role-switcher-policy.js';

const VIEWS = [
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2' },
    { id: 'neu', label: 'Neuer Antrag', icon: 'bi-plus-circle' },
    { id: 'meine', label: 'Meine Anträge', icon: 'bi-person' },
    { id: 'inbox', label: 'Offen (IT)', icon: 'bi-inbox' },
    { id: 'alle', label: 'Alle Anträge', icon: 'bi-list-ul' },
    { id: 'einstellungen', label: 'Einstellungen', icon: 'bi-gear' }
];

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function accountInfo() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            return window.ms365AuthGetAccountInfo() || null;
        }
    } catch {
        /* ignore */
    }
    return null;
}

function currentEmail() {
    const a = accountInfo();
    return String((a && (a.username || a.upn)) || '')
        .trim()
        .toLowerCase();
}

function currentName() {
    const a = accountInfo();
    return String((a && a.name) || '').trim();
}

function pillClass(status) {
    if (status === STATUS.ANGELEGT || status === STATUS.GENEHMIGT) return 'ta-pill ta-pill--ok';
    if (status === STATUS.ABGELEHNT || status === STATUS.ANLAGE_FEHLGESCHLAGEN) return 'ta-pill ta-pill--err';
    if (status === STATUS.BEANTRAGT || status === STATUS.IN_PRUEFUNG || status === STATUS.ZURUECKGESTELLT) {
        return 'ta-pill ta-pill--warn';
    }
    return 'ta-pill';
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    /** @type {{ view: string, message: {type:string,text:string}|null, detailId: string, filters: object, log: string, freigeber: object|null, freigeberLoading: boolean, groupHits: object[] }} */
    const ui = {
        view: 'uebersicht',
        message: null,
        detailId: '',
        filters: { q: '', status: '' },
        log: '',
        freigeber: null,
        freigeberLoading: false,
        groupHits: []
    };

    /** @type {import('../../shared/dashboard-audience-resolve.js').DashboardPersonas|null} */
    let personas = null;
    const demoUi = readDemoRoleOverrideFromUrl();
    if (demoUi) seedDemoIfEmpty();

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    function settings() {
        return loadTeamsAntragSettings();
    }

    /** Echtes Freigeber-Recht – Setup-Hinweis öffnet keine Inbox. */
    function canApprove() {
        return !!(ui.freigeber && ui.freigeber.ok);
    }

    function showInboxNav() {
        return canApprove();
    }

    function canManageSettings() {
        return canApprove() || !!(personas && personas.isIt);
    }

    function isOwnAntrag(a) {
        const mail = currentEmail();
        if (!mail || !a) return false;
        return String(a.antragstellerEmail || '')
            .trim()
            .toLowerCase() === mail;
    }

    async function refreshFreigeber(opts) {
        const silent = !!(opts && opts.silent);
        ui.freigeberLoading = true;
        if (!silent) render();
        try {
            ui.freigeber = await resolveFreigeberAccess({
                email: currentEmail(),
                settings: settings()
            });
        } catch (e) {
            ui.freigeber = {
                ok: false,
                reasons: [],
                groupId: '',
                groupName: '',
                source: '',
                setupNeeded: true,
                error: e && e.message ? e.message : String(e)
            };
        }
        ui.freigeberLoading = false;
        render();
    }

    function render() {
        root.innerHTML =
            '<div class="ta-shell">' +
            renderNav() +
            '<div class="ta-main">' +
            renderMessage() +
            renderView() +
            '</div></div>';
        bind();
    }

    function renderNav() {
        const queue = openQueue(loadState().antraege).length;
        const showInbox = showInboxNav();
        return (
            '<nav class="ta-nav" aria-label="Teams-Antrag">' +
            '<div class="ta-nav__brand"><i class="bi bi-people-fill" aria-hidden="true"></i><div><strong>Teams-Antrag</strong><span>kontrolliert anlegen</span></div></div>' +
            '<div class="ta-nav__list">' +
            VIEWS.filter((v) => {
                    if (v.id === 'inbox' || v.id === 'alle') return showInbox;
                    if (v.id === 'einstellungen') return canManageSettings();
                    return true;
                })
                .map((v) => {
                    const badge =
                        v.id === 'inbox' && queue
                            ? '<span class="ta-nav__badge">' + queue + '</span>'
                            : v.id === 'alle'
                              ? '<span class="ta-nav__badge">' + loadState().antraege.length + '</span>'
                              : '';
                    return (
                        '<button type="button" class="ta-nav__btn' +
                        (ui.view === v.id ? ' is-active' : '') +
                        '" data-view="' +
                        v.id +
                        '"><i class="bi ' +
                        v.icon +
                        '" aria-hidden="true"></i>' +
                        esc(v.label) +
                        badge +
                        '</button>'
                    );
                })
                .join('') +
            '</div></nav>'
        );
    }

    function renderMessage() {
        if (!ui.message) return '';
        const cls =
            ui.message.type === 'ok'
                ? 'ta-msg ta-msg--ok'
                : ui.message.type === 'warn'
                  ? 'ta-msg ta-msg--warn'
                  : 'ta-msg ta-msg--err';
        return '<div class="' + cls + '">' + esc(ui.message.text) + '</div>';
    }

    function renderView() {
        if (ui.view === 'neu') return renderNeu();
        if (ui.view === 'meine') return renderListe({ onlyMine: true, itMode: false });
        if (ui.view === 'alle') {
            if (!canApprove()) return renderFreigeberGate('Alle Anträge') || renderUebersicht();
            return renderAlleAntraege();
        }
        if (ui.view === 'inbox') {
            if (!canApprove()) return renderFreigeberGate('Offen (IT)') || renderUebersicht();
            return renderInbox();
        }
        if (ui.view === 'einstellungen') return renderEinstellungen();
        return renderUebersicht();
    }

    function renderUebersicht() {
        const all = loadState().antraege;
        const mine = currentEmail()
            ? all.filter((a) => a.antragstellerEmail === currentEmail())
            : [];
        const queue = openQueue(all);
        const fb = ui.freigeber;
        const freigeberOk = canApprove();
        const freigeberLine = ui.freigeberLoading
            ? '<p class="ta-hint">Prüfe Freigabe-Berechtigung …</p>'
            : freigeberOk
              ? '<p class="ta-msg ta-msg--ok">Sie dürfen Anträge freigeben' +
                (fb.reasons.length ? ' (' + esc(fb.reasons.join(', ')) + ')' : '') +
                '.</p>'
              : fb && fb.setupNeeded
                ? canManageSettings()
                    ? '<p class="ta-msg ta-msg--warn">Freigeber noch nicht eingerichtet – unter <strong>Einstellungen</strong> die Gruppe „Wer darf Teams anlegen?“ zuordnen.</p>'
                    : '<p class="ta-msg ta-msg--warn">Freigeber noch nicht eingerichtet. Bitte Schul-IT: Gruppe unter Einstellungen zuordnen.</p>'
                : '<p class="ta-hint">Freigabe nur für IT / Mitglieder der Freigeber-Gruppe. Sie sehen hier nur Ihre eigenen Anträge.</p>';
        return (
            '<h2>Übersicht</h2>' +
            '<p class="ta-hint">Antrag mit schulischem Microsoft-Konto. <strong>Anlegen/Freigeben</strong> nur durch die IT – idealerweise dieselbe Sicherheitsgruppe wie unter „Wer darf Teams anlegen?“.</p>' +
            freigeberLine +
            '<div class="ta-grid">' +
            (freigeberOk
                ? '<div class="ta-stat"><strong>' +
                  all.length +
                  '</strong><span>Anträge gesamt</span></div>'
                : '') +
            '<div class="ta-stat"><strong>' +
            mine.length +
            '</strong><span>Meine</span></div>' +
            (freigeberOk
                ? '<div class="ta-stat"><strong>' +
                  queue.length +
                  '</strong><span>Offen für IT</span></div>'
                : '') +
            '</div>' +
            '<div class="ta-actions">' +
            '<button type="button" class="btn btn-success" data-go="neu"><i class="bi bi-plus-lg"></i>Neuer Antrag</button>' +
            (showInboxNav()
                ? '<button type="button" class="btn" data-go="inbox"><i class="bi bi-inbox"></i>Offen (IT)</button>' +
                  '<button type="button" class="btn btn-success" data-go="alle"><i class="bi bi-list-ul"></i>Alle Anträge</button>'
                : '') +
            (canManageSettings()
                ? '<button type="button" class="btn" data-go="einstellungen"><i class="bi bi-gear"></i>Freigeber einrichten</button>'
                : '') +
            (freigeberOk
                ? '<a class="btn" href="weitere-teams-gruppen.html"><i class="bi bi-magic"></i>Weitere Teams (Bulk)</a>' +
                  '<a class="btn" href="gruppenerstellung.html"><i class="bi bi-shield-lock"></i>Wer darf Teams anlegen?</a>'
                : '') +
            '</div>' +
            (!currentEmail()
                ? '<p class="ta-hint">Bitte oben anmelden, damit Anträge Ihrer Person zugeordnet werden.</p>'
                : '')
        );
    }

    function renderNeu() {
        const artOpts = Object.keys(ART_LABEL)
            .map(
                (k) =>
                    '<option value="' +
                    k +
                    '">' +
                    esc(ART_LABEL[k]) +
                    '</option>'
            )
            .join('');
        return (
            '<h2>Neuer Antrag</h2>' +
            '<p class="ta-hint">' +
            esc(settings().nomenklaturHinweis) +
            '</p>' +
            '<form id="taNeuForm">' +
            '<div class="ta-field"><label for="taTitle">Titel</label><input id="taTitle" name="title" required maxlength="120" placeholder="z. B. Robotik-AG"></div>' +
            '<div class="ta-field"><label for="taArt">Art</label><select id="taArt" name="art">' +
            artOpts +
            '</select></div>' +
            '<div id="taDiplomHint" class="ta-msg ta-msg--warn" hidden>Für Diplomarbeit / VWA bitte das <a href="diplomarbeiten.html">Diplom-/VWA-Tool</a> verwenden – dort gibt es Dach-Teams pro Maturajahrgang.</div>' +
            '<div class="ta-field"><label for="taZweck">Zweck / Begründung</label><textarea id="taZweck" name="zweck" required placeholder="Wozu wird das Team benötigt?"></textarea></div>' +
            '<div class="ta-field"><label for="taName">Namensvorschlag</label><input id="taName" name="namensvorschlag" maxlength="120"></div>' +
            '<div class="ta-field"><label for="taJahr">Schuljahr</label><input id="taJahr" name="schuljahr" placeholder="2026/27"></div>' +
            '<div class="ta-field"><label for="taOwners">Gewünschte Besitzer (Mails, eine pro Zeile)</label><textarea id="taOwners" name="besitzerMails"></textarea></div>' +
            '<div class="ta-field"><label for="taMembers">Gewünschte Mitglieder (optional)</label><textarea id="taMembers" name="mitgliederMails"></textarea></div>' +
            '<div class="ta-actions">' +
            '<button type="submit" class="btn btn-success" data-submit="beantragen"><i class="bi bi-send"></i>Beantragen</button>' +
            '<button type="submit" class="btn" data-submit="entwurf">Als Entwurf speichern</button>' +
            '</div></form>'
        );
    }

    /**
     * @param {{ onlyMine?: boolean, itMode?: boolean, title?: string, intro?: string }} opts
     */
    function renderListe(opts) {
        const onlyMine = !!(opts && opts.onlyMine);
        const itMode = !!(opts && opts.itMode);
        const rows = filterAntraege(loadState().antraege, {
            q: ui.filters.q,
            status: ui.filters.status,
            mineEmail: currentEmail(),
            onlyMine
        });
        const statusOpts =
            '<option value="">Alle Status</option>' +
            Object.keys(STATUS_LABEL)
                .map(
                    (k) =>
                        '<option value="' +
                        k +
                        '"' +
                        (ui.filters.status === k ? ' selected' : '') +
                        '>' +
                        esc(STATUS_LABEL[k]) +
                        '</option>'
                )
                .join('');
        const title = (opts && opts.title) || (onlyMine ? 'Meine Anträge' : 'Alle Anträge');
        const intro =
            (opts && opts.intro) ||
            (onlyMine
                ? ''
                : '<p class="ta-hint">Gesamtübersicht für die IT: alle beantragten, genehmigten, abgelehnten und angelegten Teams – filterbar nach Status und Suche.</p>');
        return (
            '<h2>' +
            esc(title) +
            '</h2>' +
            intro +
            '<div class="ta-filters">' +
            '<div class="ta-field"><label for="taQ">Suche</label><input id="taQ" value="' +
            esc(ui.filters.q) +
            '"></div>' +
            '<div class="ta-field"><label for="taStatus">Status</label><select id="taStatus">' +
            statusOpts +
            '</select></div>' +
            '<button type="button" class="btn" id="taFilterBtn">Filtern</button>' +
            (itMode
                ? '<button type="button" class="btn" data-go="inbox"><i class="bi bi-inbox"></i>Nur offene</button>'
                : '') +
            '</div>' +
            '<p class="muted">' +
            rows.length +
            ' von ' +
            loadState().antraege.length +
            ' Anträgen</p>' +
            renderTable(rows, { itMode }) +
            (ui.detailId ? renderDetail(ui.detailId, itMode) : '')
        );
    }

    function renderFreigeberGate(title) {
        const fb = ui.freigeber;
        if (ui.freigeberLoading) {
            return '<h2>' + esc(title) + '</h2><p class="ta-hint">Berechtigung wird geprüft …</p>';
        }
        if (fb && fb.setupNeeded && !fb.ok) {
            return (
                '<h2>' +
                esc(title) +
                '</h2>' +
                '<p class="ta-msg ta-msg--warn">Noch keine Freigeber-Gruppe zugeordnet. Bitte unter Einstellungen die Sicherheitsgruppe aus „Wer darf Teams anlegen?“ übernehmen.</p>' +
                '<div class="ta-actions"><button type="button" class="btn btn-success" data-go="einstellungen">Zu Einstellungen</button>' +
                '<a class="btn" href="gruppenerstellung.html">Wer darf Teams anlegen?</a></div>'
            );
        }
        if (!fb || !fb.ok) {
            return (
                '<h2>' +
                esc(title) +
                '</h2>' +
                '<p class="ta-hint">Nur Freigeber (IT-Gruppe / Extra-Mails / Schul-IT) sehen alle Anträge.</p>' +
                '<div class="ta-actions"><button type="button" class="btn" data-go="einstellungen">Einstellungen</button></div>'
            );
        }
        return '';
    }

    function renderAlleAntraege() {
        const gate = renderFreigeberGate('Alle Anträge');
        if (gate) return gate;
        return renderListe({
            onlyMine: false,
            itMode: true,
            title: 'Alle Anträge',
            intro:
                '<p class="ta-hint"><strong>IT-Gesamtübersicht:</strong> alle beantragten Teams – offen, genehmigt, angelegt, abgelehnt. Eintrag öffnen zum Prüfen oder Nachbearbeiten.</p>' +
                '<div class="ta-actions" style="margin-bottom:10px">' +
                '<button type="button" class="btn" data-go="inbox"><i class="bi bi-inbox"></i>Nur offene Warteschlange</button>' +
                '<a class="btn" href="weitere-teams-gruppen.html"><i class="bi bi-magic"></i>Bulk: Weitere Teams</a>' +
                '</div>'
        });
    }

    function renderInbox() {
        const gate = renderFreigeberGate('Offen (IT)');
        if (gate) return gate;
        const rows = openQueue(loadState().antraege);
        return (
            '<h2>Offen (IT)</h2>' +
            '<p class="ta-hint">Nur noch offene Anträge (beantragt / in Prüfung / fehlgeschlagen). Die <strong>Gesamtübersicht</strong> aller Anträge finden Sie unter „Alle Anträge“.</p>' +
            '<div class="ta-actions" style="margin-bottom:10px">' +
            '<button type="button" class="btn btn-success" data-go="alle"><i class="bi bi-list-ul"></i>Alle Anträge anzeigen</button>' +
            '<a class="btn" href="weitere-teams-gruppen.html"><i class="bi bi-magic"></i>Bulk: Weitere Teams</a>' +
            '<a class="btn" href="gruppenerstellung.html"><i class="bi bi-shield-lock"></i>Wer darf Teams anlegen?</a>' +
            '</div>' +
            '<p class="muted">' +
            rows.length +
            ' offene Anträge</p>' +
            renderTable(rows, { itMode: true }) +
            (ui.detailId ? renderDetail(ui.detailId, true) : '') +
            (ui.log
                ? '<details open style="margin-top:12px"><summary>Protokoll</summary><pre class="tm-log" style="white-space:pre-wrap">' +
                  esc(ui.log) +
                  '</pre></details>'
                : '')
        );
    }

    /**
     * @param {object[]} rows
     * @param {{ itMode?: boolean }} [opts]
     */
    function renderTable(rows, opts) {
        const itMode = !!(opts && opts.itMode);
        if (!rows.length) {
            return '<p class="muted">Keine Anträge.</p>';
        }
        return (
            '<div class="ta-table-wrap"><table class="ta-table"><thead><tr>' +
            '<th>Titel</th><th>Art</th><th>Antragsteller</th>' +
            (itMode ? '<th>Schuljahr</th><th>Finaler Name</th>' : '') +
            '<th>Status</th><th></th>' +
            '</tr></thead><tbody>' +
            rows
                .map((a) => {
                    return (
                        '<tr>' +
                        '<td>' +
                        esc(a.title) +
                        '</td><td>' +
                        esc(ART_LABEL[a.art] || a.art) +
                        '</td><td>' +
                        esc(a.antragstellerName || a.antragstellerEmail || '–') +
                        '</td>' +
                        (itMode
                            ? '<td>' +
                              esc(a.schuljahr || '–') +
                              '</td><td>' +
                              esc(a.finalDisplayName || a.namensvorschlag || '–') +
                              '</td>'
                            : '') +
                        '<td><span class="' +
                        pillClass(a.status) +
                        '">' +
                        esc(STATUS_LABEL[a.status] || a.status) +
                        '</span></td>' +
                        '<td><button type="button" class="btn btn-sm" data-detail="' +
                        esc(a.id) +
                        '">Öffnen</button></td></tr>'
                    );
                })
                .join('') +
            '</tbody></table></div>'
        );
    }

    function renderDetail(id, itMode) {
        const a = getAntrag(id);
        if (!a) return '';
        if (!itMode && !isOwnAntrag(a) && !canApprove()) {
            return '<p class="ta-msg ta-msg--warn">Nur eigene Anträge sind sichtbar.</p>';
        }
        const display = a.finalDisplayName || suggestDisplayName(a);
        const nick = a.finalMailNickname || suggestMailNickname(a);
        let body =
            '<div class="ta-detail" data-antrag-id="' +
            esc(a.id) +
            '"><h3>' +
            esc(a.title) +
            '</h3>' +
            '<p><span class="' +
            pillClass(a.status) +
            '">' +
            esc(STATUS_LABEL[a.status] || a.status) +
            '</span> · ' +
            esc(ART_LABEL[a.art] || a.art) +
            '</p>' +
            '<p>' +
            esc(a.zweck) +
            '</p>' +
            '<p class="muted">Antragsteller: ' +
            esc(a.antragstellerName || '–') +
            ' &lt;' +
            esc(a.antragstellerEmail || '') +
            '&gt;</p>';

        if (a.teamWebUrl) {
            body +=
                '<p><a class="btn btn-sm" href="' +
                esc(a.teamWebUrl) +
                '" target="_blank" rel="noopener">Team öffnen</a></p>';
        }
        if (a.lastError) {
            body += '<p class="ta-msg ta-msg--err">' + esc(a.lastError) + '</p>';
        }

        if (itMode && a.art !== ART.DIPLOM_VWA) {
            body +=
                '<div class="ta-field"><label for="taFinalName">Finaler Anzeigename</label><input id="taFinalName" value="' +
                esc(display) +
                '"></div>' +
                '<div class="ta-field"><label for="taFinalNick">Finaler Alias</label><input id="taFinalNick" value="' +
                esc(nick) +
                '"></div>' +
                '<div class="ta-field"><label for="taItNotiz">IT-Notiz</label><textarea id="taItNotiz">' +
                esc(a.notizIt) +
                '</textarea></div>' +
                '<div class="ta-field"><label for="taAblehnung">Ablehnungsgrund (falls Ablehnung)</label><input id="taAblehnung" value="' +
                esc(a.ablehnungsgrund) +
                '"></div>' +
                '<div class="ta-actions">' +
                '<button type="button" class="btn" data-it="pruefung">In Prüfung</button>' +
                '<button type="button" class="btn btn-success" data-it="genehmigen"><i class="bi bi-check2"></i>Genehmigen &amp; anlegen</button>' +
                '<button type="button" class="btn" data-it="zurueck">Zurückstellen</button>' +
                '<button type="button" class="btn" data-it="ablehnen">Ablehnen</button>' +
                '</div>';
        } else if (a.art === ART.DIPLOM_VWA) {
            body +=
                '<p class="ta-msg ta-msg--warn">Bitte im <a href="diplomarbeiten.html">Diplom-/VWA-Tool</a> anlegen.</p>';
        }

        body += '</div>';
        return body;
    }

    function renderEinstellungen() {
        if (!canManageSettings()) {
            return (
                '<h2>Einstellungen</h2>' +
                '<p class="ta-msg ta-msg--warn">Freigeber-Einstellungen nur für Freigeber bzw. Schul-IT.</p>'
            );
        }
        return renderEinstellungenBody();
    }

    function renderEinstellungenBody() {
        const s = settings();
        const fb = ui.freigeber;
        const groupLabel =
            s.freigeberGroupName ||
            (fb && fb.groupName) ||
            (s.freigeberGroupId ? s.freigeberGroupId : '– noch nicht gesetzt –');
        const sourceHint =
            fb && fb.source === 'group-unified'
                ? 'Aktuell aus der Mandanten-Richtlinie Group.Unified gelesen.'
                : fb && fb.source === 'gruppenerstellung-local'
                  ? 'Vorschlag aus dem lokalen Stand von „Wer darf Teams anlegen?“.'
                  : fb && fb.source === 'settings'
                    ? 'Fest in diesen Tool-Einstellungen gespeichert.'
                    : 'Noch keine Gruppe erkannt – bitte zuordnen.';
        const hits = (ui.groupHits || [])
            .map(function (g) {
                const hint =
                    g.securityEnabled && !g.mailEnabled
                        ? 'reine Sicherheitsgruppe'
                        : g.mailEnabled
                          ? 'mail-aktiviert'
                          : 'Gruppe';
                return (
                    '<button type="button" class="btn btn-sm" data-pick-group="' +
                    esc(g.id) +
                    '" data-pick-name="' +
                    esc(g.displayName) +
                    '">' +
                    esc(g.displayName) +
                    ' <span class="muted">(' +
                    esc(hint) +
                    ')</span></button>'
                );
            })
            .join(' ');
        return (
            '<h2>Wer darf Anträge freigeben?</h2>' +
            '<p class="ta-hint">Grundsätzlich nur die <strong>IT</strong> – dieselbe Personengruppe, die im Tenant Teams anlegen darf. ' +
            'Im Tool <a href="gruppenerstellung.html">Wer darf Teams anlegen?</a> ist das eine <strong>reine Sicherheitsgruppe</strong> ' +
            '(nicht mail-aktiviert; typischer Name z. B. GruppenErsteller). Diese Gruppe hier zuordnen.</p>' +
            '<div class="ta-detail">' +
            '<p><strong>Aktuelle Freigeber-Gruppe:</strong> ' +
            esc(groupLabel) +
            (s.freigeberGroupId
                ? ' <code style="font-size:0.8em">' + esc(s.freigeberGroupId) + '</code>'
                : '') +
            '</p>' +
            '<p class="muted">' +
            esc(sourceHint) +
            '</p>' +
            '<div class="ta-actions">' +
            '<button type="button" class="btn btn-success" id="taImportGp"><i class="bi bi-download"></i>Aus „Wer darf Teams anlegen?“ übernehmen</button>' +
            '<button type="button" class="btn" id="taImportUnified"><i class="bi bi-cloud-download"></i>Aus Mandant (Group.Unified) lesen</button>' +
            '<button type="button" class="btn" id="taClearGroup">Gruppe entfernen</button>' +
            '<a class="btn" href="gruppenerstellung.html" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i>Gruppe in Entra pflegen</a>' +
            '</div></div>' +
            '<div class="ta-field" style="margin-top:14px"><label for="taGroupSearch">Sicherheitsgruppe suchen (Anzeigename)</label>' +
            '<div class="ta-actions" style="align-items:end">' +
            '<input id="taGroupSearch" placeholder="z. B. GruppenErsteller" style="max-width:280px">' +
            '<button type="button" class="btn" id="taGroupSearchBtn">Suchen</button></div>' +
            (hits ? '<div class="ta-actions" style="margin-top:8px">' + hits + '</div>' : '') +
            '</div>' +
            '<div class="ta-field"><label for="taGroupId">Oder Objekt-ID der Gruppe (GUID)</label>' +
            '<input id="taGroupId" value="' +
            esc(s.freigeberGroupId) +
            '" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" spellcheck="false"></div>' +
            '<div class="ta-field"><label for="taFreigeber">Zusätzliche einzelne Freigeber (optional, eine Mail pro Zeile)</label>' +
            '<textarea id="taFreigeber" placeholder="nur falls jemand freigeben soll, der nicht in der Gruppe ist">' +
            esc(s.freigeberMails.join('\n')) +
            '</textarea></div>' +
            '<div class="ta-field"><label for="taNomen">Nomenklatur-Hinweis (für Antragsteller)</label>' +
            '<textarea id="taNomen">' +
            esc(s.nomenklaturHinweis) +
            '</textarea></div>' +
            '<label class="checkbox-label" style="display:flex;gap:8px;margin:8px 0"><input type="checkbox" id="taSchuelerOk"' +
            (s.schuelerDuertenBeantragen ? ' checked' : '') +
            '><span>Schüler:innen dürfen beantragen</span></label>' +
            '<div class="ta-actions">' +
            '<button type="button" class="btn btn-success" id="taSaveSettings">Speichern</button>' +
            '<a class="btn" href="weitere-teams-gruppen.html"><i class="bi bi-magic"></i>Weitere Teams (Bulk)</a>' +
            (demoUi
                ? '<button type="button" class="btn" id="taDemo">Demo laden</button>' +
                  '<button type="button" class="btn" id="taClear">Lokal leeren</button>'
                : '') +
            '</div>' +
            '<p class="ta-hint">Schul-IT laut Lizenzkontakt sieht die Inbox zusätzlich. Speicherung lokal im Browser; SharePoint-Liste optional.' +
            (demoUi ? ' Demo-Aktionen aktiv (?demoRole=1).' : '') +
            '</p>'
        );
    }

    function bind() {
        root.querySelectorAll('[data-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = String(btn.getAttribute('data-view') || 'uebersicht');
                ui.detailId = '';
                ui.message = null;
                render();
            });
        });
        root.querySelectorAll('[data-go]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = String(btn.getAttribute('data-go') || 'uebersicht');
                render();
            });
        });

        const art = root.querySelector('#taArt');
        if (art) {
            const hint = root.querySelector('#taDiplomHint');
            const sync = () => {
                if (hint) hint.hidden = art.value !== ART.DIPLOM_VWA;
            };
            art.addEventListener('change', sync);
            sync();
        }

        const form = root.querySelector('#taNeuForm');
        if (form) {
            form.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const submitter = ev.submitter;
                const mode =
                    (submitter && submitter.getAttribute('data-submit')) || 'beantragen';
                const fd = new FormData(form);
                const draft = {
                    title: fd.get('title'),
                    art: fd.get('art'),
                    zweck: fd.get('zweck'),
                    namensvorschlag: fd.get('namensvorschlag'),
                    schuljahr: fd.get('schuljahr'),
                    besitzerMails: fd.get('besitzerMails'),
                    mitgliederMails: fd.get('mitgliederMails'),
                    antragstellerEmail: currentEmail(),
                    antragstellerName: currentName(),
                    status: mode === 'entwurf' ? STATUS.ENTWURF : STATUS.BEANTRAGT,
                    quelle: 'ui'
                };
                if (draft.art === ART.DIPLOM_VWA) {
                    setMessage('warn', 'Bitte das Diplom-/VWA-Tool verwenden.');
                    render();
                    return;
                }
                if (
                    !settings().schuelerDuertenBeantragen &&
                    personas &&
                    personas.isSchueler &&
                    !personas.isLehrer &&
                    !personas.isIt &&
                    !canApprove()
                ) {
                    setMessage('err', 'Schüler:innen dürfen derzeit keine Teams beantragen.');
                    render();
                    return;
                }
                const res = upsertAntrag(draft, { requireApplicant: mode !== 'entwurf' });
                if (!res.ok) {
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                setMessage(
                    'ok',
                    mode === 'entwurf' ? 'Entwurf gespeichert.' : 'Antrag eingereicht.'
                );
                ui.view = 'meine';
                ui.detailId = res.item.id;
                render();
            });
        }

        const filterBtn = root.querySelector('#taFilterBtn');
        if (filterBtn) {
            filterBtn.addEventListener('click', () => {
                const q = root.querySelector('#taQ');
                const st = root.querySelector('#taStatus');
                ui.filters.q = q ? q.value : '';
                ui.filters.status = st ? st.value : '';
                render();
            });
        }

        root.querySelectorAll('[data-detail]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.detailId = String(btn.getAttribute('data-detail') || '');
                render();
            });
        });

        root.querySelectorAll('[data-it]').forEach((btn) => {
            btn.addEventListener('click', () => {
                void handleIt(String(btn.getAttribute('data-it') || ''));
            });
        });

        const saveSet = root.querySelector('#taSaveSettings');
        if (saveSet) {
            saveSet.addEventListener('click', () => {
                if (!canManageSettings()) {
                    setMessage('err', 'Keine Berechtigung für Einstellungen.');
                    render();
                    return;
                }
                void (async () => {
                    const freigeber = root.querySelector('#taFreigeber');
                    const nomen = root.querySelector('#taNomen');
                    const sch = root.querySelector('#taSchuelerOk');
                    const gid = root.querySelector('#taGroupId');
                    const cur = settings();
                    const nextId = gid ? String(gid.value || '').trim() : cur.freigeberGroupId;
                    let nextName = cur.freigeberGroupName;
                    if (nextId && nextId !== cur.freigeberGroupId) {
                        const meta = await fetchGroupById(nextId);
                        nextName = (meta && meta.displayName) || '';
                    }
                    if (!nextId) nextName = '';
                    saveTeamsAntragSettings({
                        freigeberMails: freigeber ? freigeber.value.split(/[\n;,]+/) : [],
                        nomenklaturHinweis: nomen ? nomen.value : '',
                        schuelerDuertenBeantragen: sch ? !!sch.checked : true,
                        freigeberGroupId: nextId,
                        freigeberGroupName: nextName
                    });
                    setMessage('ok', 'Einstellungen gespeichert.');
                    void refreshFreigeber();
                })();
            });
        }

        const importGp = root.querySelector('#taImportGp');
        if (importGp) {
            importGp.addEventListener('click', () => {
                if (!canManageSettings()) {
                    setMessage('err', 'Keine Berechtigung für Einstellungen.');
                    render();
                    return;
                }
                const local = readGruppenerstellungLocal();
                if (!local) {
                    setMessage(
                        'warn',
                        'Kein lokaler Stand gefunden. Bitte zuerst unter „Wer darf Teams anlegen?“ eine Gruppe wählen/speichern – oder aus dem Mandanten lesen.'
                    );
                    render();
                    return;
                }
                saveTeamsAntragSettings({
                    freigeberGroupId: local.groupId,
                    freigeberGroupName: local.groupDisplayName || 'GruppenErsteller'
                });
                setMessage('ok', 'Gruppe übernommen: ' + (local.groupDisplayName || local.groupId));
                void refreshFreigeber();
            });
        }

        const importUni = root.querySelector('#taImportUnified');
        if (importUni) {
            importUni.addEventListener('click', () => {
                if (!canManageSettings()) {
                    setMessage('err', 'Keine Berechtigung für Einstellungen.');
                    render();
                    return;
                }
                void (async () => {
                    setMessage('warn', 'Lese Group.Unified …');
                    render();
                    const fromDir = await fetchGroupUnifiedAllowedGroupId();
                    if (!fromDir || !fromDir.groupId) {
                        setMessage(
                            'err',
                            'Keine GroupCreationAllowedGroupId im Mandanten gefunden. Zuerst „Wer darf Teams anlegen?“ setzen.'
                        );
                        render();
                        return;
                    }
                    const meta = await fetchGroupById(fromDir.groupId);
                    saveTeamsAntragSettings({
                        freigeberGroupId: fromDir.groupId,
                        freigeberGroupName: (meta && meta.displayName) || ''
                    });
                    setMessage(
                        'ok',
                        'Mandanten-Gruppe übernommen' +
                            (meta && meta.displayName ? ': ' + meta.displayName : '') +
                            '.'
                    );
                    void refreshFreigeber();
                })();
            });
        }

        const clearGroup = root.querySelector('#taClearGroup');
        if (clearGroup) {
            clearGroup.addEventListener('click', () => {
                if (!canManageSettings()) {
                    setMessage('err', 'Keine Berechtigung für Einstellungen.');
                    render();
                    return;
                }
                saveTeamsAntragSettings({ freigeberGroupId: '', freigeberGroupName: '' });
                setMessage('ok', 'Freigeber-Gruppe entfernt.');
                void refreshFreigeber();
            });
        }

        const searchBtn = root.querySelector('#taGroupSearchBtn');
        if (searchBtn) {
            searchBtn.addEventListener('click', () => {
                void (async () => {
                    const qEl = root.querySelector('#taGroupSearch');
                    const q = qEl ? qEl.value : '';
                    try {
                        ui.groupHits = await searchSecurityGroups(q);
                        if (!ui.groupHits.length) {
                            setMessage('warn', 'Keine Gruppen gefunden.');
                        } else {
                            setMessage('ok', ui.groupHits.length + ' Treffer – bitte auswählen.');
                        }
                    } catch (e) {
                        setMessage('err', e && e.message ? e.message : String(e));
                        ui.groupHits = [];
                    }
                    render();
                })();
            });
        }

        root.querySelectorAll('[data-pick-group]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!canManageSettings()) {
                    setMessage('err', 'Keine Berechtigung für Einstellungen.');
                    render();
                    return;
                }
                const id = String(btn.getAttribute('data-pick-group') || '');
                const name = String(btn.getAttribute('data-pick-name') || '');
                saveTeamsAntragSettings({ freigeberGroupId: id, freigeberGroupName: name });
                ui.groupHits = [];
                setMessage('ok', 'Gruppe gewählt: ' + name);
                void refreshFreigeber();
            });
        });

        const demo = root.querySelector('#taDemo');
        if (demo) {
            demo.addEventListener('click', () => {
                if (!demoUi || !canManageSettings()) {
                    setMessage('err', 'Demo nur mit ?demoRole=1 und Freigeber-/IT-Recht.');
                    render();
                    return;
                }
                loadDemoData();
                setMessage('ok', 'Demo geladen.');
                render();
            });
        }
        const clear = root.querySelector('#taClear');
        if (clear) {
            clear.addEventListener('click', () => {
                if (!demoUi || !canManageSettings()) {
                    setMessage('err', 'Leeren nur mit ?demoRole=1 und Freigeber-/IT-Recht.');
                    render();
                    return;
                }
                if (!window.confirm('Lokale Anträge wirklich löschen?')) return;
                clearAll();
                setMessage('ok', 'Lokal geleert.');
                render();
            });
        }
    }

    async function handleIt(action) {
        if (!ui.freigeber || !ui.freigeber.ok) {
            setMessage('err', 'Keine Freigabe-Berechtigung.');
            render();
            return;
        }
        const id = ui.detailId;
        if (!id) return;
        const nameEl = root.querySelector('#taFinalName');
        const nickEl = root.querySelector('#taFinalNick');
        const notizEl = root.querySelector('#taItNotiz');
        const ablEl = root.querySelector('#taAblehnung');
        patchItFields(id, {
            finalDisplayName: nameEl ? nameEl.value : '',
            finalMailNickname: nickEl ? nickEl.value : '',
            notizIt: notizEl ? notizEl.value : '',
            ablehnungsgrund: ablEl ? ablEl.value : ''
        });

        if (action === 'pruefung') {
            setAntragStatus(id, STATUS.IN_PRUEFUNG);
            setMessage('ok', 'Status: In Prüfung.');
            render();
            return;
        }
        if (action === 'zurueck') {
            setAntragStatus(id, STATUS.ZURUECKGESTELLT);
            setMessage('warn', 'Zurückgestellt.');
            render();
            return;
        }
        if (action === 'ablehnen') {
            const reason = ablEl ? ablEl.value.trim() : '';
            if (!reason) {
                setMessage('err', 'Ablehnungsgrund angeben.');
                render();
                return;
            }
            setAntragStatus(id, STATUS.ABGELEHNT, { ablehnungsgrund: reason });
            setMessage('ok', 'Abgelehnt.');
            render();
            return;
        }
        if (action === 'genehmigen') {
            ui.log = '';
            setMessage('warn', 'Lege Team an …');
            render();
            try {
                await approveAndCreateTeam(id, {
                    onLog: (m) => {
                        ui.log += (ui.log ? '\n' : '') + m;
                    }
                });
                setMessage('ok', 'Team angelegt.');
            } catch (e) {
                setMessage('err', e && e.message ? e.message : String(e));
            }
            render();
        }
    }

    render();
    void getDashboardPersonas()
        .then((p) => {
            personas = p || null;
            render();
        })
        .catch(() => {
            personas = null;
        });
    void refreshFreigeber({ silent: true });
}

export default { mountApp };
