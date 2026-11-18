/**
 * Aufgaben UI – Melder vs. Admin, SPO-Sync, Vorlagen, Lehrer-Picker, Mailto (Familie betrieb).
 */

import {
    aufgabenStats,
    canManageAufgabe,
    filterAufgaben,
    sortAufgaben,
    statusPipelineActions
} from './aufgaben-logic.js';
import {
    DEMO_ME_EMAIL,
    DEMO_ME_NAME,
    PRIORITAET,
    PRIORITAET_LABEL,
    ROLLE_LABEL,
    STATUS,
    STATUS_LABEL,
    toIsoDateOnly
} from './aufgaben-model.js';
import {
    appendKommentar,
    clearAll,
    clearDemoSeed,
    loadDemoData,
    loadState,
    removeAufgabe,
    seedDemoIfEmpty,
    setAufgabeStatus,
    upsertAufgabe
} from './aufgaben-store.js';
import {
    canShowAufgabenRoleSwitcher,
    getAufgabenAccount,
    getAufgabenAdminAccessCache,
    getAufgabenDemoRoleOverride,
    isAufgabenAdmin,
    isAufgabenDemoRoleUiEnabled,
    resolveAufgabenAdmin,
    setAufgabenDemoRoleOverride
} from './aufgaben-entra-role.js';
import { resolveAufgabenSiteUrl } from './aufgaben-settings.js';
import {
    getAufgabenStammdatenMeta,
    getAufgabenTeachers,
    hydrateAufgabenStammdaten
} from './aufgaben-stammdaten.js';
import {
    getAufgabenSyncMeta,
    pullAufgabenFromSharePoint,
    pushAufgabe,
    pushAufgabeDelete
} from './aufgaben-sync.js';
import {
    AUFGABEN_VORLAGEN,
    draftFromAufgabeVorlage
} from './aufgaben-templates.js';
import {
    buildStatusMailto,
    buildZuweisungMailto,
    openMailto
} from './aufgaben-notify.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';

const MELDER_VIEWS = [
    { id: 'uebersicht', label: 'Meine Aufgaben', icon: 'bi-inbox' },
    { id: 'liste', label: 'Liste', icon: 'bi-list-task' },
    { id: 'neu', label: 'Neue Aufgabe', icon: 'bi-plus-circle' },
    { id: 'erledigt', label: 'Erledigt', icon: 'bi-check2-circle' }
];

const ADMIN_EXTRA = [
    { id: 'alle', label: 'Alle offen', icon: 'bi-inboxes' },
    { id: 'system', label: 'System', icon: 'bi-gear' }
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

function wantsDemoSeed() {
    try {
        const p = new URLSearchParams(window.location.search || '');
        return p.get('demo') === '1';
    } catch {
        return false;
    }
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    const acc0 = getAufgabenAccount();
    /** @type {{
     *   view: string,
     *   message: {type:string,text:string}|null,
     *   filters: { scope: string, status: string, q: string },
     *   meEmail: string,
     *   meName: string,
     *   showAll: boolean,
     *   draft: { title: string, beschreibung: string, rolle: string, prioritaet: string, zuweisungName: string, zuweisungEmail: string, frist: string, kommentar: string, notify: boolean }
     * }} */
    const ui = {
        view: 'uebersicht',
        message: null,
        filters: { scope: 'mir', status: '', q: '' },
        meEmail: acc0.email || DEMO_ME_EMAIL,
        meName: acc0.name || (acc0.email ? acc0.email : DEMO_ME_NAME),
        showAll: false,
        draft: emptyDraft()
    };

    if (wantsDemoSeed()) {
        seedDemoIfEmpty();
    }

    function emptyDraft() {
        return {
            title: '',
            beschreibung: '',
            rolle: '',
            prioritaet: PRIORITAET.MITTEL,
            zuweisungName: '',
            zuweisungEmail: '',
            frist: '',
            kommentar: '',
            notify: true
        };
    }

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    async function afterLocalWrite(item, okMsg, notifyKind) {
        const push = await pushAufgabe(item);
        let msg = okMsg || 'Gespeichert.';
        if (push.ok && !push.skipped) msg = (okMsg || 'Gespeichert') + ' · SharePoint ok.';
        else if (push.skipped) msg = okMsg || 'Gespeichert (Demo lokal).';
        else
            msg =
                (okMsg ? okMsg + ' ' : '') +
                (push.error || 'SharePoint-Sync fehlgeschlagen – lokal gespeichert.');

        const saved = push.item || item;
        if (notifyKind === 'zuweisung' && ui.draft.notify !== false) {
            const href = buildZuweisungMailto(saved, { fromName: ui.meName });
            if (href) openMailto(href);
        } else if (notifyKind === 'status') {
            const href = buildStatusMailto(saved, { actorName: ui.meName });
            if (href) openMailto(href);
        }

        setMessage(push.ok || push.skipped ? 'ok' : 'warn', msg);
        render();
        return push;
    }

    function state() {
        return loadState();
    }

    function syncAccount() {
        const acc = getAufgabenAccount();
        if (acc.loggedIn) {
            ui.meEmail = acc.email;
            ui.meName = acc.name || acc.email;
        }
    }

    function admin() {
        return isAufgabenAdmin(ui.meEmail);
    }

    function actor() {
        return { actorName: ui.meName, actorEmail: ui.meEmail };
    }

    function listFilters(extra) {
        return {
            meEmail: ui.meEmail,
            admin: admin(),
            showAll: ui.showAll,
            ...(extra || {})
        };
    }

    function navViews() {
        const isAdm = admin();
        if (!isAdm) return MELDER_VIEWS.slice();
        return [
            ...ADMIN_EXTRA.filter((v) => v.id === 'alle'),
            ...MELDER_VIEWS,
            ...ADMIN_EXTRA.filter((v) => v.id === 'system')
        ];
    }

    function render() {
        syncAccount();
        if (!admin()) {
            ui.showAll = false;
            if (ui.view === 'alle' || ui.view === 'system') ui.view = 'uebersicht';
            if (ui.filters.scope === 'offen' || ui.filters.scope === '') ui.filters.scope = 'mir';
        }
        root.innerHTML =
            '<div class="ag-shell">' +
            renderNav() +
            '<div class="ag-main">' +
            renderMessage() +
            renderSyncHint() +
            renderView() +
            '</div></div>';
        bind();
    }

    function renderNav() {
        const isAdm = admin();
        const role = getAufgabenAdminAccessCache();
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
        if (canShowAufgabenRoleSwitcher()) {
            const cur = getAufgabenDemoRoleOverride() || (isAdm ? 'admin' : 'melder');
            roleChips =
                '<div class="af-role-chips" role="group" aria-label="Demo-Rolle">' +
                '<button type="button" class="btn' +
                (cur === 'melder' ? ' btn-success' : '') +
                '" data-af-demo-role="melder">Melder</button>' +
                '<button type="button" class="btn' +
                (cur === 'admin' ? ' btn-success' : '') +
                '" data-af-demo-role="admin">Admin</button>' +
                '</div>';
        }

        return (
            '<nav class="ag-nav" aria-label="Aufgaben">' +
            '<div class="ag-nav__brand">' +
            '<i class="bi bi-clipboard-check" aria-hidden="true"></i>' +
            '<div><strong>Aufgaben</strong><span>Zuweisen · Abhaken</span></div>' +
            '</div>' +
            '<div class="ag-nav__list">' +
            buttons +
            '</div>' +
            '<p class="ag-nav__hint">' +
            (isAdm
                ? 'Admin: schulweite Übersicht; Scope „alle“ nur mit Recht.'
                : 'Du legst Aufgaben an und erledigst die dir zugewiesenen.') +
            (isAdm
                ? ' · Admin' +
                  (role.demo || isAufgabenDemoRoleUiEnabled()
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

    function renderSyncHint() {
        const sync = getAufgabenSyncMeta();
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
                    'SharePoint nicht verbunden – Aufgaben bleiben lokal, bis die Liste greift.'
            });
        }
        return '';
    }

    function renderPipelineLegend() {
        const steps = [
            { id: STATUS.OFFEN, label: 'Offen' },
            { id: STATUS.IN_ARBEIT, label: 'In Arbeit' },
            { id: STATUS.ERLEDIGT, label: 'Erledigt' }
        ];
        return (
            '<ul class="af-pipeline" aria-label="Status-Pipeline">' +
            steps
                .map(
                    (s, i) =>
                        (i ? '<li class="af-pipeline__arrow" aria-hidden="true">→</li>' : '') +
                        '<li class="af-pipeline__step">' +
                        esc(s.label) +
                        '</li>'
                )
                .join('') +
            '</ul>'
        );
    }

    function renderCalmBanner(opts) {
        const eyebrow = opts.eyebrow || '';
        const text = opts.text || '';
        const cta = opts.cta || '';
        const kind = opts.kind === 'ops' ? 'af-calm-banner--ops' : 'af-calm-banner--action';
        return (
            '<section class="af-calm-banner ' +
            kind +
            '" role="status">' +
            '<div class="af-calm-banner__text">' +
            (eyebrow ? '<p class="af-calm-banner__eyebrow">' + esc(eyebrow) + '</p>' : '') +
            '<p>' +
            esc(text) +
            '</p></div>' +
            (cta || '') +
            '</section>'
        );
    }

    function quietEmpty(text) {
        return (
            '<p class="af-quiet-status" role="status"><i class="bi bi-check-circle" aria-hidden="true"></i><span>' +
            esc(text) +
            '</span></p>'
        );
    }

    function renderView() {
        if (ui.view === 'liste') return renderListe();
        if (ui.view === 'neu') return renderNeu();
        if (ui.view === 'erledigt') return renderErledigt();
        if (ui.view === 'alle') return renderAlle();
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

    function renderUebersicht() {
        const s = state();
        const visible = filterAufgaben(s.aufgaben, listFilters({ scope: '' }));
        const st = aufgabenStats(visible, ui.meEmail);
        const meine = sortAufgaben(
            filterAufgaben(s.aufgaben, listFilters({ scope: 'mir' }))
        ).filter((a) => a.status === STATUS.OFFEN || a.status === STATUS.IN_ARBEIT);
        const banner =
            st.meineOffen > 0
                ? renderCalmBanner({
                      eyebrow: 'Offene Aufgaben',
                      text:
                          st.ueberfaellig > 0
                              ? st.meineOffen +
                                ' dir zugewiesen · ' +
                                st.ueberfaellig +
                                ' überfällig – Status setzen oder kommentieren.'
                              : st.meineOffen === 1
                                ? '1 Aufgabe wartet auf dich.'
                                : st.meineOffen + ' Aufgaben warten auf dich.',
                      cta:
                          '<button type="button" class="btn btn-success" data-ag-go="neu">Neue Aufgabe</button>'
                  })
                : renderCalmBanner({
                      eyebrow: 'Alles ruhig',
                      text: 'Keine offenen Aufgaben für dich. Bei Bedarf eine neue zuweisen.',
                      cta:
                          '<button type="button" class="btn btn-success" data-ag-go="neu">Neue Aufgabe</button>'
                  });

        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Meine Aufgaben</h2>' +
            banner +
            renderPipelineLegend() +
            '<div class="ag-cards">' +
            card('Mir offen', st.meineOffen) +
            card('Überfällig', st.ueberfaellig) +
            card('Dringend', st.dringend) +
            card('Erledigt', st.erledigt) +
            '</div>' +
            '<p class="ag-item__meta" style="margin-bottom:10px;">Angemeldet: <strong>' +
            esc(ui.meName) +
            '</strong> (' +
            esc(ui.meEmail) +
            ') · Scope: mir / von mir</p>' +
            '<h3 style="font-size:1rem;margin:0 0 8px;">Mir zugewiesen</h3>' +
            '<div class="ag-list">' +
            (meine.map(renderItem).join('') || quietEmpty('Keine offenen Aufgaben für dich.')) +
            '</div>'
        );
    }

    function renderAlle() {
        ui.showAll = true;
        const s = state();
        const visible = sortAufgaben(
            filterAufgaben(s.aufgaben, listFilters({ showAll: true, scope: 'offen' }))
        );
        const st = aufgabenStats(
            filterAufgaben(s.aufgaben, { admin: true, showAll: true, meEmail: ui.meEmail, scope: '' }),
            ui.meEmail
        );
        const banner =
            st.offen > 0
                ? renderCalmBanner({
                      kind: 'ops',
                      eyebrow: 'Schulweite Übersicht',
                      text:
                          st.dringend > 0
                              ? st.dringend +
                                ' dringend · ' +
                                st.offen +
                                ' offen – Pipeline und Kommentar nutzen.'
                              : st.offen + ' offene Aufgaben in der Schule.'
                  })
                : renderCalmBanner({
                      eyebrow: 'Übersicht leer',
                      text: 'Aktuell keine offenen Aufgaben.'
                  });

        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Alle offenen Aufgaben</h2>' +
            banner +
            renderPipelineLegend() +
            '<div class="ag-cards">' +
            card('Offen', st.offen) +
            card('In Arbeit', st.inArbeit) +
            card('Überfällig', st.ueberfaellig) +
            card('Dringend', st.dringend) +
            '</div>' +
            '<div class="ag-list">' +
            (visible.map(renderItem).join('') || quietEmpty('Keine offenen Aufgaben.')) +
            '</div>'
        );
    }

    function renderListe() {
        const list = sortAufgaben(
            filterAufgaben(
                state().aufgaben,
                listFilters({
                    scope: ui.filters.scope,
                    status: ui.filters.status,
                    q: ui.filters.q
                })
            )
        );
        const isAdm = admin();
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Aufgabenliste</h2>' +
            '<div class="af-filter-row">' +
            '<select id="afScope">' +
            opt('mir', 'Mir zugewiesen', ui.filters.scope) +
            opt('von_mir', 'Von mir erstellt', ui.filters.scope) +
            (isAdm && ui.showAll
                ? opt('offen', 'Alle offenen', ui.filters.scope) + opt('', 'Alle', ui.filters.scope)
                : '') +
            '</select>' +
            '<select id="afStatus">' +
            opt('', 'Status: alle', ui.filters.status) +
            Object.keys(STATUS_LABEL)
                .map((k) => opt(k, STATUS_LABEL[k], ui.filters.status))
                .join('') +
            '</select>' +
            '<input id="afQ" type="search" placeholder="Suche…" value="' +
            esc(ui.filters.q) +
            '">' +
            '<button type="button" class="btn" id="afFilterApply">Filtern</button>' +
            (isAdm
                ? '<label class="ag-item__meta"><input type="checkbox" id="afShowAll"' +
                  (ui.showAll ? ' checked' : '') +
                  '> alle Vorgänge</label>'
                : '') +
            '</div>' +
            '<div class="ag-list">' +
            (list.map(renderItem).join('') || quietEmpty('Keine Aufgaben für diesen Filter.')) +
            '</div>'
        );
    }

    function renderErledigt() {
        const list = sortAufgaben(
            filterAufgaben(state().aufgaben, listFilters({ status: STATUS.ERLEDIGT, scope: '' }))
        );
        const abgelehnt = sortAufgaben(
            filterAufgaben(state().aufgaben, listFilters({ status: STATUS.ABGELEHNT, scope: '' }))
        );
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Erledigt / Abgelehnt</h2>' +
            '<div class="ag-list">' +
            (list.concat(abgelehnt).map(renderItem).join('') ||
                quietEmpty('Noch keine erledigten Aufgaben.')) +
            '</div>'
        );
    }

    function renderVorlagen() {
        return (
            '<p class="ag-item__meta" style="margin:0 0 8px;">Vorlage wählen (optional):</p>' +
            '<div class="af-vorlagen">' +
            AUFGABEN_VORLAGEN.map(
                (v) =>
                    '<button type="button" class="af-vorlage" data-af-vorlage="' +
                    esc(v.id) +
                    '"><strong>' +
                    esc(v.titel) +
                    '</strong><span>' +
                    esc(v.hinweis || v.beschreibung) +
                    '</span></button>'
            ).join('') +
            '</div>'
        );
    }

    function renderTeacherPicker() {
        const teachers = getAufgabenTeachers();
        const opts =
            '<option value="">– Lehrkraft aus Stammdaten –</option>' +
            teachers
                .map((t) => {
                    const label =
                        (t.name || t.email) +
                        (t.code ? ' (' + t.code + ')' : '') +
                        (t.email ? ' – ' + t.email : '');
                    const selected =
                        t.email && t.email === ui.draft.zuweisungEmail ? ' selected' : '';
                    return (
                        '<option value="' +
                        esc(t.email || t.name) +
                        '" data-name="' +
                        esc(t.name) +
                        '" data-email="' +
                        esc(t.email) +
                        '"' +
                        selected +
                        '>' +
                        esc(label) +
                        '</option>'
                    );
                })
                .join('');
        return (
            '<div class="af-zuweisung-picker">' +
            '<label style="flex:1 1 240px;">Zuweisung aus Stammdaten<select id="afTeacherPick">' +
            opts +
            '</select></label>' +
            '<button type="button" class="btn" id="afEntraPick" title="In Entra suchen"><i class="bi bi-people"></i> Entra</button>' +
            '</div>' +
            (teachers.length
                ? ''
                : '<p class="ag-item__meta">Keine Lehrkräfte in Stammdaten – Name/E-Mail manuell oder Entra.</p>')
        );
    }

    function renderNeu() {
        const d = ui.draft;
        const rolleOpts = Object.keys(ROLLE_LABEL)
            .map(
                (k) =>
                    '<option value="' +
                    esc(k) +
                    '"' +
                    (k === d.rolle ? ' selected' : '') +
                    '>' +
                    esc(ROLLE_LABEL[k]) +
                    '</option>'
            )
            .join('');
        const prioOpts = Object.keys(PRIORITAET_LABEL)
            .map(
                (k) =>
                    '<option value="' +
                    esc(k) +
                    '"' +
                    (k === (d.prioritaet || PRIORITAET.MITTEL) ? ' selected' : '') +
                    '>' +
                    esc(PRIORITAET_LABEL[k]) +
                    '</option>'
            )
            .join('');
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Neue Aufgabe</h2>' +
            renderCalmBanner({
                eyebrow: 'Kurz zuweisen',
                text: 'Vorlage wählen oder selbst ausfüllen – Status läuft offen → in Arbeit → erledigt.'
            }) +
            renderVorlagen() +
            '<form class="ag-form" id="afNeuForm">' +
            '<label>Titel<input name="title" required maxlength="200" placeholder="Bitte Unterschrift" value="' +
            esc(d.title) +
            '"></label>' +
            '<label>Beschreibung<textarea name="beschreibung" rows="3" placeholder="Kurz erklären…">' +
            esc(d.beschreibung) +
            '</textarea></label>' +
            renderTeacherPicker() +
            '<div class="ag-form__row">' +
            '<label>Zuweisung Name<input name="zuweisungName" required placeholder="Name" value="' +
            esc(d.zuweisungName) +
            '"></label>' +
            '<label>Zuweisung E-Mail<input name="zuweisungEmail" type="email" placeholder="name@schule.at" value="' +
            esc(d.zuweisungEmail) +
            '"></label>' +
            '</div>' +
            '<div class="ag-form__row">' +
            '<label>Rolle (optional)<select name="rolle"><option value="">–</option>' +
            rolleOpts +
            '</select></label>' +
            '<label>Priorität<select name="prioritaet">' +
            prioOpts +
            '</select></label>' +
            '</div>' +
            '<div class="ag-form__row">' +
            '<label>Frist<input name="frist" type="date" value="' +
            esc(d.frist) +
            '"></label>' +
            '<label>Kommentar<textarea name="kommentar" rows="2">' +
            esc(d.kommentar) +
            '</textarea></label>' +
            '</div>' +
            '<label class="ag-item__meta"><input type="checkbox" name="notify" value="1"' +
            (d.notify ? ' checked' : '') +
            '> Zugewiesene Person per E-Mail benachrichtigen (mailto)</label>' +
            '<div class="ag-form__actions">' +
            '<button type="submit" class="btn btn-success">Speichern</button>' +
            '<button type="button" class="btn" data-ag-go="uebersicht">Abbrechen</button>' +
            '</div>' +
            '</form>'
        );
    }

    function renderSystem() {
        const meta = getAufgabenStammdatenMeta();
        const sync = getAufgabenSyncMeta();
        const site = resolveAufgabenSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
        const s = state();
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">System</h2>' +
            '<div class="af-system-panel">' +
            '<h3>SharePoint (Live)</h3>' +
            '<p><code>' +
            esc(site) +
            '</code></p>' +
            '<p>Sync: <strong>' +
            esc(sync.source) +
            '</strong>' +
            (sync.listTitle ? ' · Liste „' + esc(sync.listTitle) + '“' : '') +
            (sync.lastPullAt ? ' · zuletzt ' + esc(formatDeDate(sync.lastPullAt)) : '') +
            (sync.connected ? ' · verbunden' : ' · nicht verbunden') +
            (sync.error ? ' · ' + esc(sync.error) : '') +
            '</p>' +
            '<p>Stammdaten: ' +
            esc(meta.source) +
            (meta.classCount || meta.teacherCount
                ? ' · ' + meta.classCount + ' Klassen · ' + meta.teacherCount + ' Lehrkräfte'
                : '') +
            (meta.error ? ' · ' + esc(meta.error) : '') +
            '</p>' +
            '<p>Lokal: ' +
            (s.aufgaben || []).length +
            ' Aufgaben' +
            (s.localDemoOnly ? ' · Demo local-first' : '') +
            ' · mit spoItemId: ' +
            (s.aufgaben || []).filter((a) => a.spoItemId).length +
            '</p>' +
            '<p class="ag-item__meta">Automatische Mails: Power-Automate-Rezept „Aufgaben Status-Mail“ (oder mailto im Formular).</p>' +
            '<div class="ag-toolbar">' +
            '<button type="button" class="btn btn-success" data-ag-pull>Von SharePoint laden</button>' +
            '<button type="button" class="btn" data-ag-demo>Demo laden</button>' +
            '<button type="button" class="btn" data-ag-clear-seed>Demo-Seed entfernen</button>' +
            '<button type="button" class="btn" data-ag-clear>Lokal leeren</button>' +
            '<a class="btn" href="sharepoint-liste-aufgaben.html"><i class="bi bi-list-ul"></i> Listen einrichten</a>' +
            '<a class="btn" href="power-automate-rezepte.html"><i class="bi bi-lightning"></i> PA-Rezepte</a>' +
            '</div></div>'
        );
    }

    function opt(value, label, selected) {
        return (
            '<option value="' +
            esc(value) +
            '"' +
            (value === selected ? ' selected' : '') +
            '>' +
            esc(label) +
            '</option>'
        );
    }

    function renderItem(a) {
        const can = canManageAufgabe(a, { email: ui.meEmail, admin: admin() });
        const actions = !can
            ? ''
            : statusPipelineActions(a.status)
                  .map(
                      (act) =>
                          '<button type="button" class="btn' +
                          (act.kind === 'success' ? ' btn-success' : '') +
                          '" data-af-status="' +
                          esc(a.id) +
                          '" data-af-to="' +
                          esc(act.to) +
                          '">' +
                          esc(act.label) +
                          '</button>'
                  )
                  .join('') +
              (a.status === STATUS.ERLEDIGT || a.status === STATUS.ABGELEHNT
                  ? admin() ||
                    String(a.erstelltVonEmail || '').toLowerCase() === ui.meEmail
                      ? '<button type="button" class="btn" data-af-del="' +
                        esc(a.id) +
                        '">Löschen</button>'
                      : ''
                  : '');

        const mailBtn =
            can && a.zuweisungEmail
                ? '<button type="button" class="btn" data-af-mail-zuweisung="' +
                  esc(a.id) +
                  '" title="Zuweisung per Mail">Mail</button>'
                : '';

        return (
            '<article class="ag-item">' +
            '<h3 class="ag-item__title">' +
            esc(a.title) +
            ' <span class="ag-pill ' +
            statusPill(a.status) +
            '">' +
            esc(STATUS_LABEL[a.status] || a.status) +
            '</span> <span class="ag-pill ' +
            prioPill(a.prioritaet) +
            '">' +
            esc(PRIORITAET_LABEL[a.prioritaet] || a.prioritaet) +
            '</span>' +
            (a.spoItemId ? ' <span class="ag-pill" title="SharePoint">SPO</span>' : '') +
            '</h3>' +
            '<p class="ag-item__meta">' +
            'An: ' +
            esc(a.zuweisungName || a.zuweisungEmail || '–') +
            (a.rolle ? ' · ' + esc(ROLLE_LABEL[a.rolle] || a.rolle) : '') +
            ' · Frist: ' +
            esc(formatDeDate(a.frist)) +
            ' · von ' +
            esc(a.erstelltVon || '–') +
            (a.bearbeitetVonName || a.bearbeitetVonEmail
                ? ' · zuletzt ' + esc(a.bearbeitetVonName || a.bearbeitetVonEmail)
                : '') +
            '</p>' +
            (a.beschreibung
                ? '<p class="ag-item__meta">' + esc(a.beschreibung) + '</p>'
                : '') +
            (a.kommentar
                ? '<p class="ag-item__meta" style="white-space:pre-wrap;">' +
                  esc(a.kommentar) +
                  '</p>'
                : '') +
            '<div class="ag-item__actions">' +
            actions +
            (can
                ? '<button type="button" class="btn" data-af-kommentar="' +
                  esc(a.id) +
                  '">Kommentar</button>'
                : '') +
            mailBtn +
            '</div></article>'
        );
    }

    function captureDraftFromForm(form) {
        if (!form) return;
        const fd = new FormData(form);
        ui.draft = {
            title: String(fd.get('title') || ''),
            beschreibung: String(fd.get('beschreibung') || ''),
            rolle: String(fd.get('rolle') || ''),
            prioritaet: String(fd.get('prioritaet') || PRIORITAET.MITTEL),
            zuweisungName: String(fd.get('zuweisungName') || ''),
            zuweisungEmail: String(fd.get('zuweisungEmail') || ''),
            frist: String(fd.get('frist') || ''),
            kommentar: String(fd.get('kommentar') || ''),
            notify: form.querySelector('[name="notify"]')
                ? !!form.querySelector('[name="notify"]').checked
                : true
        };
    }

    function bind() {
        root.querySelectorAll('[data-ag-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ag-view') || 'uebersicht';
                ui.message = null;
                if (ui.view === 'alle') ui.showAll = true;
                if (ui.view === 'neu' && !ui.draft.title) ui.draft = emptyDraft();
                render();
            });
        });
        root.querySelectorAll('[data-ag-go]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ag-go') || 'uebersicht';
                render();
            });
        });

        root.querySelector('#afShowAll')?.addEventListener('change', (ev) => {
            ui.showAll = !!ev.target.checked && admin();
            render();
        });

        root.querySelector('#afFilterApply')?.addEventListener('click', () => {
            ui.filters.scope = root.querySelector('#afScope')?.value || '';
            ui.filters.status = root.querySelector('#afStatus')?.value || '';
            ui.filters.q = root.querySelector('#afQ')?.value || '';
            render();
        });

        root.querySelectorAll('[data-af-demo-role]').forEach((btn) => {
            btn.addEventListener('click', () => {
                setAufgabenDemoRoleOverride(btn.getAttribute('data-af-demo-role'));
                render();
            });
        });

        root.querySelectorAll('[data-af-vorlage]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const draft = draftFromAufgabeVorlage(btn.getAttribute('data-af-vorlage') || '');
                if (!draft) return;
                ui.draft = {
                    ...ui.draft,
                    title: draft.title,
                    beschreibung: draft.beschreibung,
                    rolle: draft.rolle || '',
                    prioritaet: draft.prioritaet || PRIORITAET.MITTEL
                };
                render();
            });
        });

        root.querySelector('#afTeacherPick')?.addEventListener('change', (ev) => {
            const sel = ev.target;
            const optEl = sel.options[sel.selectedIndex];
            if (!optEl || !optEl.value) return;
            ui.draft.zuweisungName = optEl.getAttribute('data-name') || optEl.textContent || '';
            ui.draft.zuweisungEmail = optEl.getAttribute('data-email') || '';
            const nameInput = root.querySelector('[name="zuweisungName"]');
            const mailInput = root.querySelector('[name="zuweisungEmail"]');
            if (nameInput) nameInput.value = ui.draft.zuweisungName;
            if (mailInput) mailInput.value = ui.draft.zuweisungEmail;
        });

        root.querySelector('#afEntraPick')?.addEventListener('click', () => {
            void pickEntraUser({
                title: 'Lehrkraft zuweisen',
                hint: 'Person in Entra suchen und als Zuweisung übernehmen.'
            }).then((user) => {
                if (!user) return;
                ui.draft.zuweisungName = String(user.displayName || '').trim();
                ui.draft.zuweisungEmail = String(user.mail || user.userPrincipalName || '')
                    .trim()
                    .toLowerCase();
                const form = root.querySelector('#afNeuForm');
                if (form) captureDraftFromForm(form);
                ui.draft.zuweisungName =
                    String(user.displayName || '').trim() || ui.draft.zuweisungName;
                ui.draft.zuweisungEmail =
                    String(user.mail || user.userPrincipalName || '')
                        .trim()
                        .toLowerCase() || ui.draft.zuweisungEmail;
                render();
            });
        });

        root.querySelector('#afNeuForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const form = ev.target;
            captureDraftFromForm(form);
            const d = ui.draft;
            const res = upsertAufgabe(
                {
                    title: d.title,
                    beschreibung: d.beschreibung,
                    zuweisungName: d.zuweisungName,
                    zuweisungEmail: d.zuweisungEmail,
                    rolle: d.rolle,
                    prioritaet: d.prioritaet || PRIORITAET.MITTEL,
                    frist: d.frist || toIsoDateOnly(new Date()),
                    status: STATUS.OFFEN,
                    erstelltVon: ui.meName,
                    erstelltVonEmail: ui.meEmail,
                    kommentar: d.kommentar
                },
                actor()
            );
            if (!res.ok) {
                setMessage('err', (res.errors || []).join(' '));
                render();
                return;
            }
            ui.view = 'liste';
            ui.filters.scope = 'von_mir';
            ui.draft = emptyDraft();
            void afterLocalWrite(res.item, 'Aufgabe angelegt.', 'zuweisung');
        });

        root.querySelectorAll('[data-af-status]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-af-status') || '';
                const item = state().aufgaben.find((a) => a.id === id);
                if (!canManageAufgabe(item, { email: ui.meEmail, admin: admin() })) {
                    setMessage('err', 'Keine Berechtigung für diese Aufgabe.');
                    render();
                    return;
                }
                const to = btn.getAttribute('data-af-to') || STATUS.OFFEN;
                const res = setAufgabeStatus(id, to, actor());
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                const notifyKind =
                    to === STATUS.ERLEDIGT || to === STATUS.ABGELEHNT ? 'status' : '';
                void afterLocalWrite(res.item, 'Status: ' + (STATUS_LABEL[to] || to), notifyKind);
            });
        });

        root.querySelectorAll('[data-af-kommentar]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-af-kommentar') || '';
                const item = state().aufgaben.find((a) => a.id === id);
                if (!canManageAufgabe(item, { email: ui.meEmail, admin: admin() })) return;
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

        root.querySelectorAll('[data-af-mail-zuweisung]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-af-mail-zuweisung') || '';
                const item = state().aufgaben.find((a) => a.id === id);
                const href = buildZuweisungMailto(item, { fromName: ui.meName });
                if (!href) {
                    setMessage('warn', 'Keine Zuweisungs-E-Mail hinterlegt.');
                    render();
                    return;
                }
                openMailto(href);
            });
        });

        root.querySelectorAll('[data-af-del]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!window.confirm('Aufgabe löschen?')) return;
                const id = btn.getAttribute('data-af-del') || '';
                const item = state().aufgaben.find((a) => a.id === id);
                removeAufgabe(id);
                void pushAufgabeDelete(item || {}).then((del) => {
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
            void pullAufgabenFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                setMessage(
                    r.ok ? 'ok' : 'warn',
                    r.ok
                        ? 'SharePoint geladen (' + (r.state.aufgaben || []).length + ' Aufgaben).'
                        : r.softError || r.meta.error || 'Pull fehlgeschlagen.'
                );
                render();
            });
        });

        root.querySelector('[data-ag-demo]')?.addEventListener('click', () => {
            if (!admin()) return;
            loadDemoData();
            setMessage('ok', 'Demo geladen (local-first, ohne SPO-Push der Seeds).');
            ui.view = 'alle';
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
            if (!window.confirm('Lokale Aufgaben leeren?')) return;
            clearAll();
            setMessage('ok', 'Geleert.');
            render();
        });
    }

    render();
    void resolveAufgabenAdmin(ui.meEmail).then(() => {
        if (!admin()) ui.showAll = false;
        render();
    });
    void pullAufgabenFromSharePoint().then((r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (r.ok) setMessage('ok', 'Mit SharePoint verbunden.');
        render();
    });
    void hydrateAufgabenStammdaten().then(() => {
        if (ui.view === 'system' || ui.view === 'neu') render();
    });
}
