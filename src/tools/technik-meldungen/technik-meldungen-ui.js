/**
 * Technik-Meldungen UI – Melder vs. Admin-Inbox, Calm UX (Familie betrieb).
 */

import {
    canChangeStatus,
    canCommentOnMeldung,
    filterMeldungen,
    meldungenStats,
    sortMeldungen,
    statusPipelineActions
} from './technik-meldungen-logic.js';
import {
    DEMO_ME_EMAIL,
    DEMO_ME_NAME,
    KATEGORIE,
    KATEGORIE_LABEL,
    PRIORITAET,
    PRIORITAET_LABEL,
    STATUS,
    STATUS_LABEL
} from './technik-meldungen-model.js';
import {
    appendKommentar,
    clearAll,
    clearDemoSeed,
    loadDemoData,
    loadState,
    removeMeldung,
    setMeldungStatus,
    upsertMeldung
} from './technik-meldungen-store.js';
import {
    canShowTechnikRoleSwitcher,
    getTechnikAccount,
    getTechnikAdminAccessCache,
    getTechnikDemoRoleOverride,
    isTechnikAdmin,
    isTechnikDemoRoleUiEnabled,
    resolveTechnikAdmin,
    setTechnikDemoRoleOverride
} from './technik-meldungen-entra-role.js';
import { resolveTechnikSiteUrl } from './technik-meldungen-settings.js';
import {
    getTechnikStammdatenMeta,
    hydrateTechnikStammdaten
} from './technik-meldungen-stammdaten.js';
import {
    getTechnikSyncMeta,
    pullTechnikFromSharePoint,
    pushTechnikDelete,
    pushTechnikMeldung
} from './technik-meldungen-sync.js';
import { bindSystemTabClicks, renderSystemAdminPage } from '../../shared/system-admin-shell.js';

const MELDER_VIEWS = [
    { id: 'uebersicht', label: 'Meine Meldungen', icon: 'bi-inbox' },
    { id: 'neu', label: 'Neue Meldung', icon: 'bi-plus-circle' },
    { id: 'erledigt', label: 'Erledigt', icon: 'bi-check2-circle' },
    { id: 'zugaenge', label: 'Zugänge', icon: 'bi-wifi' }
];

const ADMIN_EXTRA = [
    { id: 'inbox', label: 'IT-Inbox', icon: 'bi-inboxes' },
    { id: 'system', label: 'Systemadministration', icon: 'bi-gear' }
];

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatDeDate(iso) {
    const s = String(iso || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s || '–';
    const [y, m, d] = s.split('-');
    return d + '.' + m + '.' + y;
}

function statusPill(st) {
    if (st === STATUS.ERLEDIGT) return 'ag-pill--ok';
    if (st === STATUS.ABGELEHNT) return 'ag-pill--bad';
    if (st === STATUS.IN_ARBEIT) return 'ag-pill--offen';
    return 'ag-pill--offen';
}

function prioPill(p) {
    if (p === PRIORITAET.DRINGEND || p === PRIORITAET.HOCH) return 'ag-pill--bad';
    if (p === PRIORITAET.MITTEL) return 'ag-pill--offen';
    return 'ag-pill';
}

function readRaumPrefill() {
    try {
        const p = new URLSearchParams(window.location.search || '');
        return String(p.get('raum') || p.get('raumCode') || '').trim();
    } catch {
        return '';
    }
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    const acc0 = getTechnikAccount();
    /** @type {{
     *   view: string,
     *   message: {type:string,text:string}|null,
     *   meEmail: string,
     *   meName: string,
     *   showAll: boolean,
     *   q: string,
     *   kategorie: string,
     *   raumPrefill: string
     * }} */
    const ui = {
        view: 'uebersicht',
        systemTab: 'site',
        message: null,
        meEmail: acc0.email || DEMO_ME_EMAIL,
        meName: acc0.name || (acc0.email ? acc0.email : DEMO_ME_NAME),
        showAll: false,
        q: '',
        kategorie: '',
        raumPrefill: readRaumPrefill()
    };

    if (ui.raumPrefill) ui.view = 'neu';

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    async function afterLocalWrite(item, okMsg) {
        const push = await pushTechnikMeldung(item);
        if (push.ok) {
            setMessage('ok', okMsg || 'In SharePoint gespeichert.');
        } else if (push.skipped) {
            setMessage('ok', okMsg || 'Gespeichert (Demo lokal).');
        } else {
            setMessage(
                'warn',
                (okMsg ? okMsg + ' ' : '') +
                    (push.error || 'SharePoint-Sync fehlgeschlagen – lokal gespeichert.')
            );
        }
        render();
        return push;
    }

    function state() {
        return loadState();
    }

    function syncAccount() {
        const acc = getTechnikAccount();
        if (acc.loggedIn) {
            ui.meEmail = acc.email;
            ui.meName = acc.name || acc.email;
        }
    }

    function admin() {
        return isTechnikAdmin(ui.meEmail);
    }

    function actor() {
        return { actorName: ui.meName, actorEmail: ui.meEmail };
    }

    function listFilters(extra) {
        return {
            meEmail: ui.meEmail,
            admin: admin(),
            showAll: ui.showAll,
            q: ui.q,
            kategorie: ui.kategorie,
            ...(extra || {})
        };
    }

    function navViews() {
        const isAdm = admin();
        if (!isAdm) return MELDER_VIEWS.slice();
        return [
            { id: 'inbox', label: 'IT-Inbox', icon: 'bi-inboxes' },
            ...MELDER_VIEWS.filter((v) => v.id !== 'uebersicht'),
            { id: 'uebersicht', label: 'Meine', icon: 'bi-person' },
            ...ADMIN_EXTRA.filter((v) => v.id === 'system')
        ];
    }

    function render() {
        syncAccount();
        if (!admin()) {
            ui.showAll = false;
            if (ui.view === 'inbox' || ui.view === 'system') ui.view = 'uebersicht';
        }
        root.innerHTML =
            '<div class="ag-shell">' +
            renderNav() +
            '<div class="ag-main">' +
            renderMessage() +
            renderView() +
            '</div></div>';
        bind();
    }

    function renderNav() {
        const isAdm = admin();
        const role = getTechnikAdminAccessCache();
        const views = navViews();
        const buttons = views
            .map(
                (v) =>
                    '<button type="button" class="ag-nav__btn' +
                    (ui.view === v.id ? ' is-active' : '') +
                    '" data-ag-view="' +
                    esc(v.id) +
                    '"><i class="bi ' +
                    esc(v.icon) +
                    '" aria-hidden="true"></i>' +
                    esc(v.label) +
                    '</button>'
            )
            .join('');

        let roleChips = '';
        if (canShowTechnikRoleSwitcher()) {
            const cur = getTechnikDemoRoleOverride() || (isAdm ? 'admin' : 'melder');
            roleChips =
                '<div class="tm-role-chips" role="group" aria-label="Demo-Rolle">' +
                '<button type="button" class="btn' +
                (cur === 'melder' ? ' btn-success' : '') +
                '" data-tm-demo-role="melder">Melder</button>' +
                '<button type="button" class="btn' +
                (cur === 'admin' ? ' btn-success' : '') +
                '" data-tm-demo-role="admin">Admin</button>' +
                '</div>';
        }

        return (
            '<nav class="ag-nav" aria-label="Technik-Meldungen">' +
            '<div class="ag-nav__brand">' +
            '<i class="bi bi-tools" aria-hidden="true"></i>' +
            '<div><strong>Technik</strong><span>Melden · Status · IT</span></div>' +
            '</div>' +
            '<div class="ag-nav__list">' +
            buttons +
            '</div>' +
            '<p class="ag-nav__hint">' +
            (isAdm
                ? 'Admin-Inbox: Status setzen und auswerten.'
                : 'Du meldest Probleme; die IT bearbeitet den Status.') +
            (isAdm
                ? ' · Admin' +
                  (role.demo || isTechnikDemoRoleUiEnabled()
                      ? ' (Demo)'
                      : role.sources.length
                        ? ' (' + esc(role.sources.join(', ')) + ')'
                        : '')
                : '') +
            '</p>' +
            roleChips +
            '</nav>'
        );
    }

    function renderMessage() {
        if (!ui.message) return '';
        const cls =
            ui.message.type === 'ok'
                ? 'ag-msg--ok'
                : ui.message.type === 'warn'
                  ? 'ag-msg--warn'
                  : 'ag-msg--err';
        return '<p class="ag-msg ' + cls + '" role="status">' + esc(ui.message.text) + '</p>';
    }

    function renderPipelineLegend() {
        const steps = [
            { id: STATUS.OFFEN, label: 'Offen' },
            { id: STATUS.IN_ARBEIT, label: 'In Arbeit' },
            { id: STATUS.ERLEDIGT, label: 'Erledigt' }
        ];
        return (
            '<ul class="tm-pipeline" aria-label="Status-Pipeline">' +
            steps
                .map(
                    (s, i) =>
                        (i ? '<li class="tm-pipeline__arrow" aria-hidden="true">→</li>' : '') +
                        '<li class="tm-pipeline__step">' +
                        esc(s.label) +
                        '</li>'
                )
                .join('') +
            '</ul>'
        );
    }

    function renderView() {
        if (ui.view === 'neu') return renderNeu();
        if (ui.view === 'erledigt') return renderErledigt();
        if (ui.view === 'zugaenge') return renderZugaenge();
        if (ui.view === 'inbox') return renderInbox();
        if (ui.view === 'system') return renderSystem();
        return renderUebersicht();
    }

    function card(label, value) {
        return (
            '<div class="ag-card"><p class="ag-card__label">' +
            esc(label) +
            '</p><p class="ag-card__value">' +
            esc(value) +
            '</p></div>'
        );
    }

    function renderCalmBanner(opts) {
        const eyebrow = opts.eyebrow || '';
        const text = opts.text || '';
        const cta = opts.cta || '';
        const kind = opts.kind === 'ops' ? 'tm-calm-banner--ops' : 'tm-calm-banner--action';
        return (
            '<section class="tm-calm-banner ' +
            kind +
            '" role="status">' +
            '<div class="tm-calm-banner__text">' +
            (eyebrow ? '<p class="tm-calm-banner__eyebrow">' + esc(eyebrow) + '</p>' : '') +
            '<p>' +
            esc(text) +
            '</p></div>' +
            (cta || '') +
            '</section>'
        );
    }

    function quietEmpty(text) {
        return (
            '<p class="tm-quiet-status" role="status"><i class="bi bi-check-circle" aria-hidden="true"></i><span>' +
            esc(text) +
            '</span></p>'
        );
    }

    function renderFilters() {
        const katOpts =
            '<option value="">Alle Kategorien</option>' +
            Object.keys(KATEGORIE_LABEL)
                .map(
                    (k) =>
                        '<option value="' +
                        esc(k) +
                        '"' +
                        (ui.kategorie === k ? ' selected' : '') +
                        '>' +
                        esc(KATEGORIE_LABEL[k]) +
                        '</option>'
                )
                .join('');
        return (
            '<div class="tm-filter-row">' +
            '<input type="search" id="tmSearch" placeholder="Suchen …" value="' +
            esc(ui.q) +
            '">' +
            '<select id="tmKat">' +
            katOpts +
            '</select>' +
            (admin()
                ? '<label class="ag-item__meta"><input type="checkbox" id="tmShowAll"' +
                  (ui.showAll ? ' checked' : '') +
                  '> Inbox (alle)</label>'
                : '') +
            '</div>'
        );
    }

    function renderUebersicht() {
        const s = state();
        const visible = filterMeldungen(s.meldungen, listFilters({}));
        const st = meldungenStats(visible);
        const offen = sortMeldungen(filterMeldungen(s.meldungen, listFilters({ nurOffen: true })));
        const banner =
            st.offen > 0
                ? renderCalmBanner({
                      eyebrow: 'Was ist offen?',
                      text:
                          st.offen === 1
                              ? '1 offene Meldung wartet auf Feedback der IT.'
                              : st.offen + ' offene Meldungen – Priorität und Status siehst du unten.',
                      cta:
                          '<button type="button" class="btn btn-success" data-ag-go="neu">Neue Meldung</button>'
                  })
                : renderCalmBanner({
                      eyebrow: 'Alles ruhig',
                      text: 'Keine offenen eigenen Meldungen. Bei Defekten einfach neu melden.',
                      cta:
                          '<button type="button" class="btn btn-success" data-ag-go="neu">Technik melden</button>'
                  });

        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Meine Meldungen</h2>' +
            renderSyncHint() +
            banner +
            renderPipelineLegend() +
            '<div class="ag-cards">' +
            card('Offen', st.offen) +
            card('In Arbeit', st.inArbeit) +
            card('Dringend', st.dringend) +
            card('Erledigt', st.erledigt) +
            '</div>' +
            renderFilters() +
            '<p class="ag-item__meta" style="margin-bottom:10px;">Angemeldet: <strong>' +
            esc(ui.meName) +
            '</strong> (' +
            esc(ui.meEmail) +
            ') · nur eigene</p>' +
            '<div class="ag-list">' +
            (offen.map(renderItem).join('') || quietEmpty('Keine offenen Technik-Meldungen.')) +
            '</div>'
        );
    }

    function renderInbox() {
        ui.showAll = true;
        const s = state();
        const visible = sortMeldungen(
            filterMeldungen(s.meldungen, listFilters({ showAll: true, nurOffen: true }))
        );
        const st = meldungenStats(
            filterMeldungen(s.meldungen, { admin: true, showAll: true, meEmail: ui.meEmail })
        );
        const banner =
            st.offen > 0
                ? renderCalmBanner({
                      kind: 'ops',
                      eyebrow: 'IT-Inbox',
                      text:
                          st.dringend > 0
                              ? st.dringend +
                                ' dringend · ' +
                                st.offen +
                                ' offen – Status setzen und Kommentar hinterlassen.'
                              : st.offen + ' offene Tickets in der Pipeline.'
                  })
                : renderCalmBanner({
                      eyebrow: 'Inbox leer',
                      text: 'Aktuell keine offenen Technik-Tickets.'
                  });

        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">IT-Inbox</h2>' +
            renderSyncHint() +
            banner +
            renderPipelineLegend() +
            '<div class="ag-cards">' +
            card('Offen', st.offen) +
            card('In Arbeit', st.inArbeit) +
            card('Dringend', st.dringend) +
            card('Gesamt', st.total) +
            '</div>' +
            renderFilters() +
            '<div class="ag-list">' +
            (visible.map(renderItem).join('') || quietEmpty('Keine offenen Tickets.')) +
            '</div>'
        );
    }

    function renderErledigt() {
        const list = sortMeldungen(
            filterMeldungen(state().meldungen, listFilters({ status: STATUS.ERLEDIGT }))
        );
        const abgelehnt = sortMeldungen(
            filterMeldungen(state().meldungen, listFilters({ status: STATUS.ABGELEHNT }))
        );
        const rows = list.concat(abgelehnt);
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Erledigt / Abgelehnt</h2>' +
            renderFilters() +
            '<div class="ag-list">' +
            (rows.map(renderItem).join('') || quietEmpty('Noch keine erledigten Meldungen.')) +
            '</div>'
        );
    }

    function renderNeu() {
        const katOpts = Object.keys(KATEGORIE_LABEL)
            .map(
                (k) =>
                    '<option value="' +
                    esc(k) +
                    '"' +
                    (k === KATEGORIE.BEAMER ? ' selected' : '') +
                    '>' +
                    esc(KATEGORIE_LABEL[k]) +
                    '</option>'
            )
            .join('');
        const prioOpts = Object.keys(PRIORITAET_LABEL)
            .map(
                (k) =>
                    '<option value="' +
                    esc(k) +
                    '"' +
                    (k === PRIORITAET.MITTEL ? ' selected' : '') +
                    '>' +
                    esc(PRIORITAET_LABEL[k]) +
                    '</option>'
            )
            .join('');
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Neue Fehlermeldung</h2>' +
            renderCalmBanner({
                eyebrow: 'Kurz melden',
                text: 'Titel, Raum und Kategorie reichen – die IT übernimmt Status und Nacharbeit.'
            }) +
            '<form class="ag-form" id="tmNeuForm">' +
            '<label>Titel<input name="title" required maxlength="200" placeholder="Beamer flackert"></label>' +
            '<label>Beschreibung<textarea name="beschreibung" rows="3" placeholder="Was genau funktioniert nicht?"></textarea></label>' +
            '<div class="ag-form__row">' +
            '<label>Raum-Code<input name="raumCode" required placeholder="PHY1" value="' +
            esc(ui.raumPrefill) +
            '"></label>' +
            '<label>Raum-Name<input name="raumName" placeholder="Physiksaal"></label>' +
            '</div>' +
            '<div class="ag-form__row">' +
            '<label>Kategorie<select name="kategorie">' +
            katOpts +
            '</select></label>' +
            '<label>Priorität<select name="prioritaet">' +
            prioOpts +
            '</select></label>' +
            '</div>' +
            '<label>Kommentar (optional)<textarea name="kommentar" rows="2"></textarea></label>' +
            '<div class="ag-form__actions">' +
            '<button type="submit" class="btn btn-success">Melden</button>' +
            '<button type="button" class="btn" data-ag-go="uebersicht">Abbrechen</button>' +
            '</div>' +
            '</form>'
        );
    }

    function renderZugaenge() {
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Zugänge / Hinweise</h2>' +
            '<article class="ag-item">' +
            '<h3 class="ag-item__title"><i class="bi bi-wifi" aria-hidden="true"></i> WLAN &amp; Login</h3>' +
            '<p class="ag-item__meta">Dieses Tool speichert <strong>keine Passwörter</strong>. Die IT hinterlegt hier nur öffentlich nutzbare Hinweise.</p>' +
            '<p class="ag-item__meta" style="margin-top:10px;"><span class="ag-pill">Hinweis</span> WLAN-SSID: <em>Schule-WLAN</em> – Zugangsdaten beim IT-Team oder über das Schulportal.</p>' +
            '<p class="ag-item__meta"><span class="ag-pill">Hinweis</span> Login: Microsoft 365 (Schulkonto) – Passwort-Reset über Self-Service / IT.</p>' +
            '<p class="ag-item__meta"><span class="ag-pill">Hinweis</span> Support: <em>it@schule.demo</em> · Mo–Fr 07:30–15:00</p>' +
            '</article>' +
            '<div class="ag-toolbar" style="margin-top:14px;">' +
            '<button type="button" class="btn btn-success" data-ag-go="neu">Technik melden</button>' +
            '<button type="button" class="btn" data-ag-go="uebersicht">Zu meinen Meldungen</button>' +
            '</div>'
        );
    }

    function renderSyncHint() {
        const sync = getTechnikSyncMeta();
        if (sync.source === 'live' && sync.ok) return '';
        if (!sync.siteUrl && sync.source !== 'demo') {
            return renderCalmBanner({
                kind: 'ops',
                eyebrow: 'SharePoint',
                text: 'Noch keine Site – IT richtet die Liste ein. Bis dahin lokal/Demo.'
            });
        }
        if (sync.source === 'demo' || sync.source === 'local') {
            return renderCalmBanner({
                kind: 'ops',
                eyebrow: 'Lokal / Demo',
                text:
                    sync.error ||
                    'SharePoint nicht verbunden – Meldungen bleiben lokal, bis die Liste greift.'
            });
        }
        return '';
    }

    function renderSystem() {
        const meta = getTechnikStammdatenMeta();
        const sync = getTechnikSyncMeta();
        const site = resolveTechnikSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
        const s = state();
        const syncLine =
            'Sync: ' +
            sync.source +
            (sync.listTitle ? ' · Liste „' + sync.listTitle + '“' : '') +
            (sync.lastPullAt ? ' · zuletzt ' + formatDeDate(sync.lastPullAt) : '') +
            (sync.connected ? ' · verbunden' : ' · nicht verbunden') +
            (sync.error ? ' · ' + sync.error : '');
        const stammLine =
            'Stammdaten: ' +
            (meta.source || '–') +
            (meta.classCount || meta.teacherCount
                ? ' · ' + meta.classCount + ' Klassen · ' + meta.teacherCount + ' Lehrkräfte'
                : '') +
            (meta.error ? ' · ' + meta.error : '');
        return renderSystemAdminPage({
            prefix: 'tm',
            activeTab: ui.systemTab,
            dataAttr: 'data-tm-sys-tab',
            ariaLabel: 'Systemadministration Technik-Meldungen',
            panels: [
                {
                    id: 'site',
                    step: '1',
                    title: 'Datenquelle',
                    leadHtml: 'SharePoint-Liste der Meldungen. Die Status-Pipeline bleibt die Inbox.',
                    body:
                        '<p><code>' +
                        esc(site) +
                        '</code></p><p class="muted">' +
                        esc(syncLine) +
                        '</p><p class="muted">Lokal: ' +
                        (s.meldungen || []).length +
                        ' Meldungen' +
                        (s.localDemoOnly ? ' · Demo local-first' : '') +
                        ' · mit spoItemId: ' +
                        (s.meldungen || []).filter((m) => m.spoItemId).length +
                        '</p><div class="ms-sys-actions">' +
                        '<button type="button" class="btn ms-sys-btn-primary" data-ag-pull>Von SharePoint laden</button>' +
                        '<a class="btn ms-sys-btn-secondary" href="sharepoint-liste-technik-meldungen.html"><i class="bi bi-list-ul"></i> Listen einrichten</a>' +
                        '</div>'
                },
                {
                    id: 'perms',
                    step: '2',
                    title: 'Berechtigungen',
                    leadHtml: 'Melder und IT/Admin trennt Entra. Listenrechte setzt der Assistent.',
                    body:
                        '<div class="ms-sys-links"><a class="btn ms-sys-btn-secondary" href="sharepoint-liste-technik-meldungen.html"><i class="bi bi-shield-check"></i>Listen-Assistent</a></div>'
                },
                {
                    id: 'kalender',
                    step: '3',
                    title: 'Kalender',
                    leadHtml: 'Technik-Meldungen haben keinen Outlook-Feed. Der Raumplan verlinkt hierher.',
                    body:
                        '<div class="ms-sys-actions"><a class="btn ms-sys-btn-secondary" href="raumplan.html"><i class="bi bi-map"></i>Raumplan</a></div>'
                },
                {
                    id: 'wartung',
                    step: '4',
                    title: 'Wartung',
                    leadHtml: esc(stammLine),
                    body:
                        '<div class="ms-sys-actions">' +
                        '<button type="button" class="btn ms-sys-btn-secondary" data-ag-demo>Demo laden</button>' +
                        '<a class="btn ms-sys-btn-ghost" href="../tenant.html"><i class="bi bi-gear"></i>Stammdaten</a>' +
                        '<a class="btn ms-sys-btn-ghost" href="../hilfe.html"><i class="bi bi-life-preserver"></i>Hilfe</a>' +
                        '</div>'
                },
                {
                    id: 'danger',
                    title: 'Danger Zone',
                    danger: true,
                    leadHtml: 'Demo-Seed und lokaler Cache. SharePoint-Zeilen bleiben, bis sie dort gelöscht werden.',
                    body:
                        '<div class="ms-sys-actions">' +
                        '<button type="button" class="btn ms-sys-btn-danger" data-ag-clear-seed>Demo-Seed entfernen</button>' +
                        '<button type="button" class="btn ms-sys-btn-danger" data-ag-clear>Lokal leeren</button>' +
                        '</div>'
                }
            ]
        });
    }

    function renderItem(m) {
        const isAdm = admin();
        const ctx = { admin: isAdm, meEmail: ui.meEmail };
        const actions = canChangeStatus(m, ctx)
            ? statusPipelineActions(m.status)
                  .map(
                      (a) =>
                          '<button type="button" class="btn' +
                          (a.kind === 'success' ? ' btn-success' : '') +
                          '" data-tm-status="' +
                          esc(m.id) +
                          '" data-tm-to="' +
                          esc(a.to) +
                          '">' +
                          esc(a.label) +
                          '</button>'
                  )
                  .join('') +
              (m.status === STATUS.ERLEDIGT || m.status === STATUS.ABGELEHNT
                  ? '<button type="button" class="btn" data-tm-del="' +
                    esc(m.id) +
                    '">Löschen</button>'
                  : '')
            : '';

        return (
            '<article class="ag-item">' +
            '<h3 class="ag-item__title">' +
            esc(m.title) +
            ' <span class="ag-pill ' +
            statusPill(m.status) +
            '">' +
            esc(STATUS_LABEL[m.status] || m.status) +
            '</span> <span class="ag-pill ' +
            prioPill(m.prioritaet) +
            '">' +
            esc(PRIORITAET_LABEL[m.prioritaet] || m.prioritaet) +
            '</span> <span class="ag-pill">' +
            esc(KATEGORIE_LABEL[m.kategorie] || m.kategorie) +
            '</span></h3>' +
            '<p class="ag-item__meta">' +
            esc(m.raumCode || '') +
            (m.raumName ? ' · ' + esc(m.raumName) : '') +
            ' · gemeldet ' +
            esc(formatDeDate(m.gemeldetAm)) +
            ' von ' +
            esc(m.melderName || m.melderEmail || '–') +
            (m.erledigtAm ? ' · erledigt ' + esc(formatDeDate(m.erledigtAm)) : '') +
            (m.bearbeitetVonName || m.bearbeitetVonEmail
                ? ' · zuletzt ' + esc(m.bearbeitetVonName || m.bearbeitetVonEmail)
                : '') +
            '</p>' +
            (m.beschreibung
                ? '<p class="ag-item__meta">' + esc(m.beschreibung) + '</p>'
                : '') +
            (m.kommentar
                ? '<p class="ag-item__meta" style="white-space:pre-wrap;">' +
                  esc(m.kommentar) +
                  '</p>'
                : '') +
            '<div class="ag-item__actions">' +
            actions +
            (canCommentOnMeldung(m, ctx)
                ? '<button type="button" class="btn" data-tm-kommentar="' +
                  esc(m.id) +
                  '">Kommentar</button>'
                : '') +
            '</div></article>'
        );
    }

    function bind() {
        root.querySelectorAll('[data-ag-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ag-view') || 'uebersicht';
                ui.message = null;
                if (ui.view === 'inbox') ui.showAll = true;
                render();
            });
        });
        bindSystemTabClicks(root, {
            dataAttr: 'data-tm-sys-tab',
            onSelect: (tab) => {
                ui.systemTab = tab;
                render();
            }
        });
        root.querySelectorAll('[data-ag-go]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ag-go') || 'uebersicht';
                render();
            });
        });

        root.querySelector('#tmShowAll')?.addEventListener('change', (ev) => {
            ui.showAll = !!ev.target.checked && admin();
            render();
        });
        root.querySelector('#tmSearch')?.addEventListener('change', (ev) => {
            ui.q = String(ev.target.value || '').trim();
            render();
        });
        root.querySelector('#tmKat')?.addEventListener('change', (ev) => {
            ui.kategorie = String(ev.target.value || '').trim();
            render();
        });

        root.querySelectorAll('[data-tm-demo-role]').forEach((btn) => {
            btn.addEventListener('click', () => {
                setTechnikDemoRoleOverride(btn.getAttribute('data-tm-demo-role'));
                render();
            });
        });

        root.querySelector('#tmNeuForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const fd = new FormData(ev.target);
            const res = upsertMeldung(
                {
                    title: String(fd.get('title') || ''),
                    beschreibung: String(fd.get('beschreibung') || ''),
                    raumCode: String(fd.get('raumCode') || ''),
                    raumName: String(fd.get('raumName') || ''),
                    kategorie: String(fd.get('kategorie') || KATEGORIE.SONSTIGES),
                    prioritaet: String(fd.get('prioritaet') || PRIORITAET.MITTEL),
                    status: STATUS.OFFEN,
                    melderName: ui.meName,
                    melderEmail: ui.meEmail,
                    kommentar: String(fd.get('kommentar') || '')
                },
                actor()
            );
            if (!res.ok) {
                setMessage('err', (res.errors || []).join(' '));
                render();
                return;
            }
            ui.view = admin() ? 'inbox' : 'uebersicht';
            ui.raumPrefill = '';
            void afterLocalWrite(res.item, 'Meldung gespeichert.');
        });

        root.querySelectorAll('[data-tm-status]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!admin()) {
                    setMessage('err', 'Status nur für IT/Admin.');
                    render();
                    return;
                }
                const id = btn.getAttribute('data-tm-status') || '';
                const to = btn.getAttribute('data-tm-to') || STATUS.OFFEN;
                const res = setMeldungStatus(id, to, actor());
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                void afterLocalWrite(res.item, 'Status: ' + (STATUS_LABEL[to] || to));
            });
        });

        root.querySelectorAll('[data-tm-kommentar]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-tm-kommentar') || '';
                const item = state().meldungen.find((m) => m.id === id);
                if (!canCommentOnMeldung(item, { admin: admin(), meEmail: ui.meEmail })) return;
                const text = window.prompt('Kommentar hinzufügen:');
                if (text == null) return;
                const res = appendKommentar(id, text, actor());
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                void afterLocalWrite(res.item, 'Kommentar gespeichert.');
            });
        });

        root.querySelectorAll('[data-tm-del]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!admin()) return;
                if (!window.confirm('Meldung löschen?')) return;
                const id = btn.getAttribute('data-tm-del') || '';
                const item = state().meldungen.find((m) => m.id === id);
                removeMeldung(id);
                void pushTechnikDelete(item || {}).then((del) => {
                    setMessage(
                        del.ok ? 'ok' : 'warn',
                        del.ok
                            ? 'Gelöscht.'
                            : 'Lokal gelöscht. ' + (del.error || 'SharePoint-Löschen fehlgeschlagen.')
                    );
                    render();
                });
            });
        });

        root.querySelector('[data-ag-pull]')?.addEventListener('click', () => {
            if (!admin()) return;
            setMessage('ok', 'Lade von SharePoint …');
            render();
            void pullTechnikFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                setMessage(
                    r.ok ? 'ok' : 'warn',
                    r.ok
                        ? 'SharePoint geladen (' + (r.state.meldungen || []).length + ' Meldungen).'
                        : r.softError || r.meta.error || 'Pull fehlgeschlagen.'
                );
                render();
            });
        });

        root.querySelector('[data-ag-demo]')?.addEventListener('click', () => {
            if (!admin()) return;
            loadDemoData();
            setMessage('ok', 'Demo geladen (local-first, ohne SPO-Push der Seeds).');
            ui.view = 'inbox';
            ui.showAll = true;
            render();
        });

        root.querySelector('[data-ag-clear-seed]')?.addEventListener('click', () => {
            if (!admin()) return;
            clearDemoSeed();
            setMessage('ok', 'Demo-Seed entfernt.');
            render();
        });

        root.querySelector('[data-ag-clear]')?.addEventListener('click', () => {
            if (!admin()) return;
            if (!window.confirm('Lokale Meldungen leeren?')) return;
            clearAll();
            setMessage('ok', 'Geleert.');
            render();
        });
    }

    render();
    void resolveTechnikAdmin(ui.meEmail).then(() => {
        if (!admin()) ui.showAll = false;
        render();
    });
    void pullTechnikFromSharePoint().then((r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (r.ok) setMessage('ok', 'Mit SharePoint verbunden.');
        render();
    });
    void hydrateTechnikStammdaten().then(() => {
        if (ui.view === 'system') render();
    });
}
