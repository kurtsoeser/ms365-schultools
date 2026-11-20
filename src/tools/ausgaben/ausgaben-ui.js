/**
 * Ausgabenverwaltung Admin-UI – Inventar, Ausgeben, Zurücknehmen, Typen.
 */

import {
    buildUebergabeText,
    filterGegenstaende,
    filterVorgaenge,
    gegenstaendeToCsv,
    inventarStats
} from './ausgaben-logic.js';
import {
    EMPFAENGER_TYP,
    EMPFAENGER_TYP_LABEL,
    VORGANG_STATUS,
    VORGANG_STATUS_LABEL,
    toIsoDateOnly
} from './ausgaben-model.js';
import { loadAusgabenSettings, resolveAusgabenSiteUrl, saveAusgabenSettings } from './ausgaben-settings.js';
import {
    applyTypVorlage,
    ausgeben,
    clearAll,
    clearDemoSeed,
    importGegenstaendeCsv,
    loadDemoData,
    loadState,
    removeGegenstand,
    removeTyp,
    seedDemoIfEmpty,
    upsertGegenstand,
    upsertTyp,
    zuruecknehmen
} from './ausgaben-store.js';
import {
    getAusgabenSyncMeta,
    pullAusgabenFromSharePoint,
    pushAusgabenGegenstand,
    pushAusgabenGegenstandDelete,
    pushAusgabenVorgang
} from './ausgaben-sync.js';
import { TYP_VORLAGEN } from './ausgaben-templates.js';
import { loadEinsammlerStammdaten } from '../einsammler/einsammler-stammdaten.js';
import {
    getAusgabenAccount,
    getAusgabenAdminAccessCache,
    isAusgabenAdmin,
    resolveAusgabenAdmin
} from './ausgaben-entra-role.js';
import { bindSystemTabClicks, renderSystemAdminPage } from '../../shared/system-admin-shell.js';

const VIEWS = [
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2' },
    { id: 'inventar', label: 'Inventar', icon: 'bi-box-seam' },
    { id: 'ausgeben', label: 'Ausgeben', icon: 'bi-box-arrow-right' },
    { id: 'offen', label: 'Offen / Retour', icon: 'bi-arrow-left-right' },
    { id: 'typen', label: 'Typen & Vorlagen', icon: 'bi-tags' },
    { id: 'einstellungen', label: 'Systemadministration', icon: 'bi-gear' }
];

const ADMIN_ONLY_VIEWS = new Set(['inventar', 'ausgeben', 'typen', 'einstellungen']);

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

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    /** @type {{ view: string, message: {type:string,text:string}|null, filters: object, editId: string, showAll: boolean }} */
    const ui = {
        view: 'uebersicht',
        systemTab: 'site',
        message: null,
        filters: { typ: '', q: '', status: 'ausgegeben', ueberfaellig: false },
        editId: '',
        showAll: true
    };

    seedDemoIfEmpty();

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    function state() {
        return loadState();
    }

    function account() {
        return getAusgabenAccount();
    }

    function admin() {
        return isAusgabenAdmin(account().email);
    }

    function typLabel(typId) {
        const t = state().typen.find((x) => x.id === typId);
        return (t && t.label) || typId || '–';
    }

    function render() {
        if (!admin()) {
            ui.showAll = false;
            if (ADMIN_ONLY_VIEWS.has(ui.view)) ui.view = 'offen';
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
        const role = getAusgabenAdminAccessCache();
        const buttons = VIEWS.filter((v) => isAdm || !ADMIN_ONLY_VIEWS.has(v.id))
            .map((v) => {
                const label =
                    !isAdm && v.id === 'offen' ? 'Meine Ausgaben' : v.id === 'uebersicht' && !isAdm ? 'Hinweis' : v.label;
                return (
                    '<button type="button" class="ag-nav__btn' +
                    (ui.view === v.id ? ' is-active' : '') +
                    '" data-ag-view="' +
                    esc(v.id) +
                    '"><i class="bi ' +
                    esc(v.icon) +
                    '" aria-hidden="true"></i>' +
                    esc(label) +
                    '</button>'
                );
            })
            .join('');
        return (
            '<nav class="ag-nav" aria-label="Ausgaben">' +
            '<div class="ag-nav__brand">' +
            '<i class="bi bi-key" aria-hidden="true"></i>' +
            '<div><strong>Ausgaben</strong><span>Zugang &amp; Inventar · P5</span></div>' +
            '</div>' +
            '<div class="ag-nav__list">' +
            buttons +
            '</div>' +
            '<p class="ag-nav__hint">Admin: Inventar/Ausgeben. Lehrkraft: eigene Empfänger-Vorgänge.' +
            (isAdm
                ? ' · Admin' +
                  (role.demo ? ' (Demo)' : role.sources.length ? ' (' + esc(role.sources.join(', ')) + ')' : '')
                : '') +
            '</p>' +
            (isAdm
                ? '<div class="ag-nav__links"><a href="sharepoint-liste-ausgaben.html"><i class="bi bi-list-ul"></i> SharePoint-Paket</a></div>'
                : '') +
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

    function renderCalmBanner(opts) {
        const eyebrow = opts.eyebrow || '';
        const text = opts.text || '';
        const kind = opts.kind === 'ops' ? 'tm-calm-banner--ops' : 'tm-calm-banner--action';
        return (
            '<section class="tm-calm-banner ' +
            kind +
            '" role="status">' +
            '<div class="tm-calm-banner__text">' +
            (eyebrow ? '<p class="tm-calm-banner__eyebrow">' + esc(eyebrow) + '</p>' : '') +
            '<p>' +
            esc(text) +
            '</p></div></section>'
        );
    }

    function renderSyncHint() {
        const sync = getAusgabenSyncMeta();
        if (sync.source === 'live' && sync.ok) return '';
        if (!sync.siteUrl && sync.source !== 'demo') {
            return renderCalmBanner({
                kind: 'ops',
                eyebrow: 'SharePoint',
                text: 'Noch keine Site – IT richtet die Listen ein. Bis dahin lokal/Demo.'
            });
        }
        if (sync.source === 'demo' || sync.source === 'local') {
            return renderCalmBanner({
                kind: 'ops',
                eyebrow: 'Lokal / Demo',
                text:
                    sync.error ||
                    'SharePoint nicht verbunden – Inventar bleibt lokal, bis die Listen greifen.'
            });
        }
        return '';
    }

    function renderView() {
        if (ui.view === 'inventar') return renderInventar();
        if (ui.view === 'ausgeben') return renderAusgeben();
        if (ui.view === 'offen') return renderOffen();
        if (ui.view === 'typen') return renderTypen();
        if (ui.view === 'einstellungen') return renderEinstellungen();
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
        if (!admin()) {
            const acc = account();
            return (
                '<h2 style="margin:0 0 12px;font-size:1.15rem;">Ausgaben</h2>' +
                '<p class="ag-item__meta">Inventar und Ausgabe nur für Verwaltung/Schulleitung/IT. Sie sehen unter „Meine Ausgaben“ Vorgänge mit Ihrer E-Mail' +
                (acc.email ? ' (<strong>' + esc(acc.email) + '</strong>)' : '') +
                '.</p>' +
                '<div class="ag-toolbar">' +
                '<button type="button" class="btn btn-success" data-ag-go="offen">Meine Ausgaben</button>' +
                '</div>'
            );
        }
        const s = state();
        const st = inventarStats(s.gegenstaende, s.vorgaenge);
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Übersicht</h2>' +
            '<div class="ag-cards">' +
            card('Inventar', st.total) +
            card('Verfügbar', st.verfuegbar) +
            card('Ausgegeben', st.ausgegeben) +
            card('Überfällig', st.ueberfaellig) +
            '</div>' +
            renderSyncHint() +
            '<div class="ag-toolbar">' +
            '<button type="button" class="btn btn-success" data-ag-go="ausgeben">Ausgeben</button>' +
            '<button type="button" class="btn" data-ag-go="offen">Offene Vorgänge</button>' +
            '<button type="button" class="btn" data-ag-go="typen">Typen / Vorlagen</button>' +
            '</div>' +
            '<p class="ag-item__meta">Keine starren Kategorien: Vorlagen für Schlüssel &amp; Karten, eigene Typen jederzeit. SharePoint ist System of Record (Inventar + Vorgänge).</p>'
        );
    }

    function typOptions(selected) {
        return state()
            .typen.map(
                (t) =>
                    '<option value="' +
                    esc(t.id) +
                    '"' +
                    (t.id === selected ? ' selected' : '') +
                    '>' +
                    esc(t.label) +
                    '</option>'
            )
            .join('');
    }

    function renderInventar() {
        const s = state();
        const list = filterGegenstaende(s.gegenstaende, s.vorgaenge, {
            typ: ui.filters.typ,
            q: ui.filters.q,
            nurAktiv: false
        });
        const rows = list
            .map((g) => {
                const open = s.vorgaenge.find(
                    (v) => v.gegenstandId === g.id && v.status === VORGANG_STATUS.AUSGEGEBEN
                );
                return (
                    '<article class="ag-item">' +
                    '<div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;">' +
                    '<h3 class="ag-item__title">' +
                    esc(g.kennung) +
                    ' · ' +
                    esc(g.bezeichnung || '–') +
                    '</h3>' +
                    '<span class="ag-pill ' +
                    (open ? 'ag-pill--offen' : g.aktiv ? 'ag-pill--ok' : '') +
                    '">' +
                    (open ? 'ausgegeben' : g.aktiv ? 'verfügbar' : 'inaktiv') +
                    '</span></div>' +
                    '<p class="ag-item__meta">' +
                    esc(g.typLabel || typLabel(g.typ)) +
                    ' · ' +
                    esc(g.ortZone || 'ohne Ort') +
                    (open ? ' · bei ' + esc(open.empfaengerName) : '') +
                    '</p>' +
                    '<div class="ag-item__actions">' +
                    (!open && g.aktiv
                        ? '<button type="button" class="btn btn-success" data-ag-aus="' +
                          esc(g.id) +
                          '">Ausgeben</button>'
                        : '') +
                    '<button type="button" class="btn" data-ag-edit="' +
                    esc(g.id) +
                    '">Bearbeiten</button>' +
                    '<button type="button" class="btn" data-ag-delg="' +
                    esc(g.id) +
                    '">Löschen</button>' +
                    '</div></article>'
                );
            })
            .join('');

        const edit = ui.editId ? s.gegenstaende.find((g) => g.id === ui.editId) : null;

        return (
            '<h2 style="margin:0 0 10px;font-size:1.15rem;">Inventar</h2>' +
            '<div class="ag-filters">' +
            '<select id="agFilterTyp"><option value="">Alle Typen</option>' +
            typOptions(ui.filters.typ) +
            '</select>' +
            '<input id="agFilterQ" placeholder="Suche Kennung / Name" value="' +
            esc(ui.filters.q || '') +
            '">' +
            '<button type="button" class="btn" id="agFilterApply">Filtern</button>' +
            '</div>' +
            '<div class="ag-toolbar">' +
            '<button type="button" class="btn btn-success" data-ag-newg>Neuer Gegenstand</button>' +
            '<button type="button" class="btn" data-ag-export>CSV export</button>' +
            '</div>' +
            (edit || ui.editId === 'new' ? renderGegenstandForm(edit) : '') +
            '<div class="ag-list">' +
            (rows || '<p class="ag-item__meta">Kein Inventar.</p>') +
            '</div>'
        );
    }

    function renderGegenstandForm(g) {
        const item = g || {
            id: '',
            typ: state().typen[0]?.id || 'sonstiges',
            kennung: '',
            bezeichnung: '',
            ortZone: '',
            aktiv: true,
            notiz: ''
        };
        return (
            '<form class="ag-form" id="agGegenstandForm" style="margin-bottom:14px;">' +
            '<input type="hidden" name="id" value="' +
            esc(item.id) +
            '">' +
            '<div class="ag-form__row">' +
            '<label>Typ<select name="typ">' +
            typOptions(item.typ) +
            '</select></label>' +
            '<label>Kennung<input name="kennung" required value="' +
            esc(item.kennung) +
            '"></label></div>' +
            '<label>Bezeichnung<input name="bezeichnung" value="' +
            esc(item.bezeichnung) +
            '"></label>' +
            '<label>Ort / Zone<input name="ortZone" value="' +
            esc(item.ortZone) +
            '"></label>' +
            '<label>Notiz<textarea name="notiz" rows="2">' +
            esc(item.notiz) +
            '</textarea></label>' +
            '<label class="ag-item__meta"><input type="checkbox" name="aktiv"' +
            (item.aktiv !== false ? ' checked' : '') +
            '> Aktiv</label>' +
            '<div class="ag-form__actions">' +
            '<button type="submit" class="btn btn-success">Speichern</button>' +
            '<button type="button" class="btn" data-ag-cancel-edit>Abbrechen</button>' +
            '</div></form>'
        );
    }

    function renderAusgeben() {
        const s = state();
        const verfuegbar = filterGegenstaende(s.gegenstaende, s.vorgaenge, {
            nurAktiv: true,
            nurVerfuegbar: true,
            typ: ui.filters.typ
        });
        const settings = loadAusgabenSettings();
        const opts = verfuegbar
            .map(
                (g) =>
                    '<option value="' +
                    esc(g.id) +
                    '">' +
                    esc(g.kennung + ' – ' + (g.bezeichnung || typLabel(g.typ))) +
                    '</option>'
            )
            .join('');
        return (
            '<h2 style="margin:0 0 10px;font-size:1.15rem;">Ausgeben</h2>' +
            '<form class="ag-form" id="agAusgebenForm">' +
            '<label>Gegenstand<select name="gegenstandId" required><option value="">– wählen –</option>' +
            opts +
            '</select></label>' +
            '<div class="ag-form__row">' +
            '<label>Empfänger-Typ<select name="empfaengerTyp">' +
            Object.keys(EMPFAENGER_TYP_LABEL)
                .map(
                    (k) =>
                        '<option value="' +
                        esc(k) +
                        '">' +
                        esc(EMPFAENGER_TYP_LABEL[k]) +
                        '</option>'
                )
                .join('') +
            '</select></label>' +
            '<label>Name<input name="empfaengerName" required></label></div>' +
            '<div class="ag-form__row">' +
            '<label>E-Mail<input type="email" name="empfaengerEmail"></label>' +
            (() => {
                const classes = loadEinsammlerStammdaten().classes || [];
                if (!classes.length) {
                    return '<label>Klasse<input name="klasseCode" placeholder="optional"></label>';
                }
                return (
                    '<label>Klasse<select name="klasseCode"><option value="">– optional –</option>' +
                    classes
                        .map(
                            (c) =>
                                '<option value="' +
                                esc(c.code) +
                                '">' +
                                esc(c.code + (c.name && c.name !== c.code ? ' · ' + c.name : '')) +
                                '</option>'
                        )
                        .join('') +
                    '</select></label>'
                );
            })() +
            '</div>' +
            '<div class="ag-form__row">' +
            '<label>Ausgegeben am<input type="date" name="ausgegebenAm" value="' +
            esc(toIsoDateOnly(new Date())) +
            '"></label>' +
            '<label>Fällig am<input type="date" name="faelligAm"></label></div>' +
            '<label>Kaution (nur Hinweis, kein Zahlungsfluss)<input name="kautionHinweis" placeholder="z. B. Kaution hinterlegt"></label>' +
            '<label class="ag-item__meta"><input type="checkbox" name="unterschriftEmpfang"> Unterschrift Empfang erhalten</label>' +
            '<label>Notiz<textarea name="notiz" rows="2"></textarea></label>' +
            '<label>Ausgegeben von<input name="ausgegebenVon" value="' +
            esc(settings.defaultAusgegebenVon) +
            '"></label>' +
            '<div class="ag-form__actions"><button type="submit" class="btn btn-success">Ausgeben</button></div>' +
            '</form>' +
            (verfuegbar.length
                ? ''
                : '<p class="ag-msg ag-msg--warn">Kein verfügbarer Gegenstand – Inventar prüfen oder Retour buchen.</p>')
        );
    }

    function renderOffen() {
        const s = state();
        const isAdm = admin();
        const acc = account();
        const list = filterVorgaenge(
            s.vorgaenge,
            {
                status: ui.filters.status || '',
                typ: ui.filters.typ,
                q: ui.filters.q,
                ueberfaellig: ui.filters.ueberfaellig,
                meEmail: acc.email,
                admin: isAdm,
                showAll: isAdm
            },
            s.gegenstaende
        );
        const today = toIsoDateOnly(new Date());
        const rows = list
            .map((v) => {
                const g = s.gegenstaende.find((x) => x.id === v.gegenstandId);
                const late =
                    v.status === VORGANG_STATUS.AUSGEGEBEN && v.faelligAm && v.faelligAm < today;
                return (
                    '<article class="ag-item">' +
                    '<div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;">' +
                    '<h3 class="ag-item__title">' +
                    esc(v.kennung || g?.kennung || '–') +
                    ' → ' +
                    esc(v.empfaengerName) +
                    '</h3>' +
                    '<span class="ag-pill ' +
                    (late
                        ? 'ag-pill--bad'
                        : v.status === 'ausgegeben'
                          ? 'ag-pill--offen'
                          : 'ag-pill--ok') +
                    '">' +
                    esc(VORGANG_STATUS_LABEL[v.status] || v.status) +
                    (late ? ' · überfällig' : '') +
                    '</span></div>' +
                    '<p class="ag-item__meta">' +
                    esc((g && (g.typLabel || typLabel(g.typ))) || '') +
                    ' · seit ' +
                    formatDeDate(v.ausgegebenAm) +
                    (v.faelligAm ? ' · fällig ' + formatDeDate(v.faelligAm) : '') +
                    (v.zurueckAm ? ' · zurück ' + formatDeDate(v.zurueckAm) : '') +
                    '</p>' +
                    '<div class="ag-item__actions">' +
                    (isAdm && v.status === VORGANG_STATUS.AUSGEGEBEN
                        ? '<button type="button" class="btn btn-success" data-ag-retour="' +
                          esc(v.id) +
                          '">Zurücknehmen</button>' +
                          '<button type="button" class="btn" data-ag-lost="' +
                          esc(v.id) +
                          '">Als verloren</button>'
                        : '') +
                    '<button type="button" class="btn" data-ag-protokoll="' +
                    esc(v.id) +
                    '">Protokoll</button>' +
                    '</div></article>'
                );
            })
            .join('');

        return (
            '<h2 style="margin:0 0 10px;font-size:1.15rem;">' +
            (isAdm ? 'Vorgänge' : 'Meine Ausgaben') +
            '</h2>' +
            (!isAdm
                ? '<p class="ag-item__meta" style="margin-bottom:10px;">Gefiltert auf ' +
                  esc(acc.email || '– bitte anmelden –') +
                  '.</p>'
                : '') +
            '<div class="ag-filters">' +
            '<select id="agVStatus">' +
            '<option value="">Alle Status</option>' +
            Object.keys(VORGANG_STATUS_LABEL)
                .map(
                    (k) =>
                        '<option value="' +
                        esc(k) +
                        '"' +
                        (ui.filters.status === k ? ' selected' : '') +
                        '>' +
                        esc(VORGANG_STATUS_LABEL[k]) +
                        '</option>'
                )
                .join('') +
            '</select>' +
            '<select id="agVTyp"><option value="">Alle Typen</option>' +
            typOptions(ui.filters.typ) +
            '</select>' +
            '<label class="ag-item__meta"><input type="checkbox" id="agVLate"' +
            (ui.filters.ueberfaellig ? ' checked' : '') +
            '> nur überfällig</label>' +
            '<button type="button" class="btn" id="agVApply">Filtern</button>' +
            '</div>' +
            '<div class="ag-list">' +
            (rows || '<p class="ag-item__meta">Keine Vorgänge.</p>') +
            '</div>'
        );
    }

    function renderTypen() {
        const s = state();
        const existing = new Set(s.typen.map((t) => t.id));
        const vorlagen = TYP_VORLAGEN.map(
            (v) =>
                '<button type="button" class="ag-vorlage" data-ag-vorlage="' +
                esc(v.id) +
                '"><strong><i class="bi ' +
                esc(v.icon) +
                '"></i> ' +
                esc(v.label) +
                '</strong><span>' +
                esc(v.hinweis) +
                '</span>' +
                (existing.has(v.id)
                    ? '<br><span class="ag-pill ag-pill--ok" style="margin-top:6px;">bereits aktiv</span>'
                    : '') +
                '</button>'
        ).join('');

        const typRows = s.typen
            .map(
                (t) =>
                    '<article class="ag-item"><h3 class="ag-item__title"><i class="bi ' +
                    esc(t.icon || 'bi-tag') +
                    '"></i> ' +
                    esc(t.label) +
                    ' <span class="ag-pill">' +
                    esc(t.id) +
                    '</span></h3>' +
                    '<p class="ag-item__meta">' +
                    esc(t.hinweis || '') +
                    '</p>' +
                    '<div class="ag-item__actions">' +
                    '<button type="button" class="btn" data-ag-delt="' +
                    esc(t.id) +
                    '">Entfernen</button></div></article>'
            )
            .join('');

        return (
            '<h2 style="margin:0 0 8px;font-size:1.15rem;">Typen &amp; Vorlagen</h2>' +
            '<p class="ag-item__meta" style="margin-bottom:12px;">Vorlagen nur als Start – eigene Typen (z. B. „Beamer-Koffer“, „Sportgeräte“) jederzeit anlegen. SharePoint speichert Typ als Text, nicht starr.</p>' +
            '<div class="ag-vorlagen">' +
            vorlagen +
            '</div>' +
            '<h3 style="font-size:1rem;margin:18px 0 8px;">Eigener Typ</h3>' +
            '<form class="ag-form" id="agTypForm">' +
            '<div class="ag-form__row">' +
            '<label>ID (technisch)<input name="id" required placeholder="beamer_koffer" pattern="[a-z0-9_\\-]+"></label>' +
            '<label>Bezeichnung<input name="label" required placeholder="Beamer-Koffer"></label></div>' +
            '<label>Hinweis<input name="hinweis" placeholder="optional"></label>' +
            '<div class="ag-form__actions"><button type="submit" class="btn btn-success">Typ speichern</button></div>' +
            '</form>' +
            '<h3 style="font-size:1rem;margin:18px 0 8px;">Aktive Typen</h3>' +
            '<div class="ag-list">' +
            (typRows || '<p class="ag-item__meta">Noch keine Typen.</p>') +
            '</div>'
        );
    }

    function renderEinstellungen() {
        if (!admin()) {
            return '<div class="ag-head"><h2>Einstellungen</h2><p>Nur Verwaltung / Schulleitung / IT.</p></div>';
        }
        const set = loadAusgabenSettings();
        const role = getAusgabenAdminAccessCache();
        const sync = getAusgabenSyncMeta();
        const site = resolveAusgabenSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
        const s = state();
        const syncLine =
            'Sync: ' +
            sync.source +
            (sync.listTitle ? ' · ' + sync.listTitle : '') +
            (sync.lastPullAt ? ' · zuletzt ' + formatDeDate(sync.lastPullAt) : '') +
            (sync.connected ? ' · verbunden' : ' · nicht verbunden') +
            (sync.error ? ' · ' + sync.error : '');
        const who =
            'Admin aus Entra (Schulleitung/Verwaltung/IT)' +
            (role.demo ? ' · Demo' : role.sources.length ? ' · ' + role.sources.join(', ') : '') +
            '. Legacy-Mails optional.';
        return (
            renderSyncHint() +
            renderSystemAdminPage({
                prefix: 'ag',
                activeTab: ui.systemTab,
                dataAttr: 'data-ag-sys-tab',
                ariaLabel: 'Systemadministration Ausgaben',
                panels: [
                    {
                        id: 'site',
                        step: '1',
                        title: 'Datenquelle',
                        leadHtml: 'Inventar und Vorgänge in SharePoint.',
                        body:
                            '<p><code>' +
                            esc(site) +
                            '</code></p><p class="muted">' +
                            esc(syncLine) +
                            '</p><p class="muted">Lokal: ' +
                            (s.gegenstaende || []).length +
                            ' Gegenstände · ' +
                            (s.vorgaenge || []).length +
                            ' Vorgänge' +
                            (s.localDemoOnly ? ' · Demo local-first' : '') +
                            ' · spoItemId Inventar: ' +
                            (s.gegenstaende || []).filter((g) => g.spoItemId).length +
                            ' · Vorgänge: ' +
                            (s.vorgaenge || []).filter((v) => v.spoItemId).length +
                            '</p><div class="ms-sys-actions">' +
                            '<button type="button" class="btn ms-sys-btn-primary" data-ag-pull>Von SharePoint laden</button>' +
                            '<a class="btn ms-sys-btn-secondary" href="sharepoint-liste-ausgaben.html"><i class="bi bi-list-ul"></i> Listen einrichten</a>' +
                            '</div>'
                    },
                    {
                        id: 'perms',
                        step: '2',
                        title: 'Berechtigungen',
                        leadHtml: esc(who),
                        body:
                            '<p class="muted">Zusätzliche Admin-Mails speicherst du unter Wartung. Listenrechte setzt der Assistent.</p>' +
                            '<div class="ms-sys-links"><a class="btn ms-sys-btn-secondary" href="sharepoint-liste-ausgaben.html"><i class="bi bi-shield-check"></i>Listen-Assistent</a></div>'
                    },
                    {
                        id: 'kalender',
                        step: '3',
                        title: 'Kalender',
                        leadHtml: 'Rückgaben siehst du unter Offen / Retour. Es gibt keinen Outlook-Feed.',
                        body:
                            '<div class="ms-sys-actions"><button type="button" class="btn ms-sys-btn-secondary" data-ag-go="offen"><i class="bi bi-arrow-left-right"></i>Offen / Retour</button></div>'
                    },
                    {
                        id: 'wartung',
                        step: '4',
                        title: 'Wartung',
                        leadHtml: 'Schulname, Standard-Ausgeber und CSV-Import.',
                        body:
                            '<form class="ag-form" id="agSettingsForm">' +
                            '<label>Schulname<input name="schoolName" value="' +
                            esc(set.schoolName) +
                            '"></label>' +
                            '<label>Standard „Ausgegeben von“<input name="defaultAusgegebenVon" value="' +
                            esc(set.defaultAusgegebenVon) +
                            '"></label>' +
                            '<label>Zusätzliche Admin-E-Mails (Legacy)<textarea name="adminEmails" rows="2" placeholder="verwaltung@schule.at">' +
                            esc(set.adminEmails) +
                            '</textarea></label>' +
                            '<div class="ag-form__actions"><button type="submit" class="btn btn-success">Speichern</button>' +
                            '<button type="button" class="btn" data-ag-demo>Demo laden</button></div></form>' +
                            '<form class="ag-form" id="agImportForm" style="margin-top:14px">' +
                            '<label>Zeilen: Typ;Kennung;Bezeichnung;OrtZone<textarea name="csv" rows="5" placeholder="schluessel;S-01;Raum 01;Trakt A"></textarea></label>' +
                            '<div class="ag-form__actions"><button type="submit" class="btn">Importieren</button></div></form>'
                    },
                    {
                        id: 'danger',
                        title: 'Danger Zone',
                        danger: true,
                        leadHtml: 'Lokal leeren lässt SharePoint unangetastet.',
                        body:
                            '<div class="ms-sys-actions">' +
                            '<button type="button" class="btn ms-sys-btn-danger" data-ag-clear-seed>Demo-Seed entfernen</button>' +
                            '<button type="button" class="btn ms-sys-btn-danger" data-ag-clear>Lokal leeren</button>' +
                            '</div>'
                    }
                ]
            })
        );
    }

    function downloadText(filename, text, mime) {
        const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function bind() {
        bindSystemTabClicks(root, {
            dataAttr: 'data-ag-sys-tab',
            onSelect: (tab) => {
                ui.systemTab = tab;
                render();
            }
        });
        root.querySelectorAll('[data-ag-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ag-view') || 'uebersicht';
                ui.message = null;
                ui.editId = '';
                render();
            });
        });
        root.querySelectorAll('[data-ag-go]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ag-go') || 'uebersicht';
                render();
            });
        });

        root.querySelector('#agFilterApply')?.addEventListener('click', () => {
            ui.filters.typ = root.querySelector('#agFilterTyp')?.value || '';
            ui.filters.q = root.querySelector('#agFilterQ')?.value || '';
            render();
        });
        root.querySelector('#agVApply')?.addEventListener('click', () => {
            ui.filters.status = root.querySelector('#agVStatus')?.value || '';
            ui.filters.typ = root.querySelector('#agVTyp')?.value || '';
            ui.filters.ueberfaellig = !!root.querySelector('#agVLate')?.checked;
            render();
        });

        root.querySelector('[data-ag-newg]')?.addEventListener('click', () => {
            ui.editId = 'new';
            render();
        });
        root.querySelector('[data-ag-cancel-edit]')?.addEventListener('click', () => {
            ui.editId = '';
            render();
        });
        root.querySelectorAll('[data-ag-edit]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.editId = btn.getAttribute('data-ag-edit') || '';
                render();
            });
        });
        root.querySelectorAll('[data-ag-delg]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!window.confirm('Gegenstand löschen?')) return;
                const id = btn.getAttribute('data-ag-delg') || '';
                const item = state().gegenstaende.find((g) => g.id === id);
                const res = removeGegenstand(id);
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                void pushAusgabenGegenstandDelete(item || {}).then((del) => {
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
        root.querySelectorAll('[data-ag-aus]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = 'ausgeben';
                render();
                const sel = root.querySelector('[name="gegenstandId"]');
                if (sel) sel.value = btn.getAttribute('data-ag-aus') || '';
            });
        });

        root.querySelector('#agGegenstandForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const fd = new FormData(ev.target);
            const typ = String(fd.get('typ') || '');
            const existingId = String(fd.get('id') || '').trim();
            const prev = existingId
                ? state().gegenstaende.find((g) => g.id === existingId)
                : null;
            const res = upsertGegenstand({
                id: existingId || undefined,
                typ,
                typLabel: typLabel(typ),
                kennung: String(fd.get('kennung') || ''),
                bezeichnung: String(fd.get('bezeichnung') || ''),
                ortZone: String(fd.get('ortZone') || ''),
                notiz: String(fd.get('notiz') || ''),
                aktiv: !!ev.target.querySelector('[name="aktiv"]')?.checked,
                spoItemId: prev && prev.spoItemId ? prev.spoItemId : '',
                seedTag: prev && prev.seedTag ? prev.seedTag : ''
            });
            if (!res.ok) {
                setMessage('err', (res.errors || []).join(' '));
                render();
                return;
            }
            ui.editId = '';
            void pushAusgabenGegenstand(res.item).then((push) => {
                setMessage(
                    push.ok ? 'ok' : 'warn',
                    push.ok
                        ? push.skipped
                            ? 'Gespeichert (Demo, ohne SPO-Push).'
                            : 'Gespeichert.'
                        : 'Lokal gespeichert. ' + (push.error || 'SharePoint-Push fehlgeschlagen.')
                );
                render();
            });
        });

        root.querySelector('#agAusgebenForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const fd = new FormData(ev.target);
            const res = ausgeben({
                gegenstandId: String(fd.get('gegenstandId') || ''),
                empfaengerTyp: String(fd.get('empfaengerTyp') || EMPFAENGER_TYP.LEHRER),
                empfaengerName: String(fd.get('empfaengerName') || ''),
                empfaengerEmail: String(fd.get('empfaengerEmail') || ''),
                klasseCode: String(fd.get('klasseCode') || ''),
                ausgegebenAm: String(fd.get('ausgegebenAm') || ''),
                faelligAm: String(fd.get('faelligAm') || ''),
                kautionHinweis: String(fd.get('kautionHinweis') || ''),
                unterschriftEmpfang: !!ev.target.querySelector('[name="unterschriftEmpfang"]')?.checked,
                notiz: String(fd.get('notiz') || ''),
                ausgegebenVon: String(fd.get('ausgegebenVon') || '')
            });
            if (!res.ok) {
                setMessage('err', (res.errors || []).join(' '));
                render();
                return;
            }
            ui.view = 'offen';
            ui.filters.status = 'ausgegeben';
            void pushAusgabenVorgang(res.item).then((push) => {
                setMessage(
                    push.ok ? 'ok' : 'warn',
                    push.ok
                        ? push.skipped
                            ? 'Ausgegeben (Demo, ohne SPO-Push).'
                            : 'Ausgegeben.'
                        : 'Lokal ausgegeben. ' + (push.error || 'SharePoint-Push fehlgeschlagen.')
                );
                render();
            });
        });

        root.querySelectorAll('[data-ag-retour]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!admin()) return;
                const res = zuruecknehmen(btn.getAttribute('data-ag-retour') || '');
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                void pushAusgabenVorgang(res.item).then((push) => {
                    setMessage(
                        push.ok ? 'ok' : 'warn',
                        push.ok
                            ? 'Zurückgenommen.'
                            : 'Lokal zurückgenommen. ' + (push.error || 'SharePoint-Push fehlgeschlagen.')
                    );
                    render();
                });
            });
        });
        root.querySelectorAll('[data-ag-lost]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!admin()) return;
                const res = zuruecknehmen(btn.getAttribute('data-ag-lost') || '', {
                    status: VORGANG_STATUS.VERLOREN
                });
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                void pushAusgabenVorgang(res.item).then((push) => {
                    setMessage(
                        push.ok ? 'ok' : 'warn',
                        push.ok
                            ? 'Als verloren markiert.'
                            : 'Lokal markiert. ' + (push.error || 'SharePoint-Push fehlgeschlagen.')
                    );
                    render();
                });
            });
        });
        root.querySelectorAll('[data-ag-protokoll]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const s = state();
                const v = s.vorgaenge.find((x) => x.id === btn.getAttribute('data-ag-protokoll'));
                if (!v) return;
                const g = s.gegenstaende.find((x) => x.id === v.gegenstandId) || {};
                downloadText('uebergabe-' + (v.kennung || v.id) + '.txt', buildUebergabeText(g, v));
            });
        });

        root.querySelectorAll('[data-ag-vorlage]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-ag-vorlage') || '';
                const withExamples = window.confirm(
                    'Vorlage „' + id + '“ aktivieren. Auch Beispiel-Gegenstände anlegen?'
                );
                const before = new Set((state().gegenstaende || []).map((g) => g.id));
                const res = applyTypVorlage(id, { withExamples });
                if (!res.ok) {
                    setMessage('err', (res.errors || []).join(' '));
                    render();
                    return;
                }
                const added = (res.state.gegenstaende || []).filter((g) => !before.has(g.id));
                setMessage('ok', 'Vorlage übernommen.');
                render();
                void Promise.all(added.map((g) => pushAusgabenGegenstand(g))).then((results) => {
                    const failed = results.filter((r) => !r.ok && !r.skipped);
                    if (!failed.length) return;
                    setMessage(
                        'warn',
                        'Vorlage lokal ok. ' + failed.length + ' SharePoint-Push(es) fehlgeschlagen.'
                    );
                    render();
                });
            });
        });
        root.querySelector('#agTypForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const fd = new FormData(ev.target);
            const res = upsertTyp({
                id: String(fd.get('id') || ''),
                label: String(fd.get('label') || ''),
                hinweis: String(fd.get('hinweis') || ''),
                icon: 'bi-tag'
            });
            setMessage(res.ok ? 'ok' : 'err', res.ok ? 'Typ gespeichert.' : (res.errors || []).join(' '));
            render();
        });
        root.querySelectorAll('[data-ag-delt]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const res = removeTyp(btn.getAttribute('data-ag-delt') || '');
                setMessage(res.ok ? 'ok' : 'err', res.ok ? 'Typ entfernt.' : (res.errors || []).join(' '));
                render();
            });
        });

        root.querySelector('[data-ag-export]')?.addEventListener('click', () => {
            downloadText('inventar-ausgaben.csv', gegenstaendeToCsv(state().gegenstaende), 'text/csv;charset=utf-8');
        });

        root.querySelector('#agSettingsForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            if (!admin()) return;
            const fd = new FormData(ev.target);
            saveAusgabenSettings({
                schoolName: String(fd.get('schoolName') || ''),
                defaultAusgegebenVon: String(fd.get('defaultAusgegebenVon') || ''),
                adminEmails: String(fd.get('adminEmails') || '')
            });
            setMessage('ok', 'Einstellungen gespeichert.');
            void resolveAusgabenAdmin(account().email).then(() => render());
        });
        root.querySelector('#agImportForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            if (!admin()) return;
            const before = new Set((state().gegenstaende || []).map((g) => g.id));
            const fd = new FormData(ev.target);
            const res = importGegenstaendeCsv(String(fd.get('csv') || ''));
            const added = (res.state.gegenstaende || []).filter((g) => !before.has(g.id));
            setMessage(
                res.ok || res.added ? 'ok' : 'err',
                res.added + ' importiert.' + (res.errors.length ? ' ' + res.errors.slice(0, 3).join(' ') : '')
            );
            render();
            void Promise.all(added.map((g) => pushAusgabenGegenstand(g))).then((results) => {
                const failed = results.filter((r) => !r.ok && !r.skipped);
                if (!failed.length) return;
                setMessage(
                    'warn',
                    'Import lokal ok. ' + failed.length + ' SharePoint-Push(es) fehlgeschlagen.'
                );
                render();
            });
        });
        root.querySelector('[data-ag-pull]')?.addEventListener('click', () => {
            if (!admin()) return;
            setMessage('ok', 'Lade von SharePoint …');
            render();
            void pullAusgabenFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                setMessage(
                    r.ok ? 'ok' : 'warn',
                    r.ok
                        ? 'SharePoint geladen (' +
                              (r.state.gegenstaende || []).length +
                              ' Gegenstände · ' +
                              (r.state.vorgaenge || []).length +
                              ' Vorgänge).'
                        : r.softError || r.meta.error || 'Pull fehlgeschlagen.'
                );
                render();
            });
        });
        root.querySelector('[data-ag-demo]')?.addEventListener('click', () => {
            if (!admin()) return;
            loadDemoData();
            setMessage('ok', 'Demo geladen (local-first, ohne SPO-Push der Seeds).');
            ui.view = 'inventar';
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
            if (!window.confirm('Lokale Daten leeren?')) return;
            clearAll();
            setMessage('ok', 'Geleert.');
            render();
        });
    }

    render();
    void resolveAusgabenAdmin(account().email).then(() => {
        if (!admin()) ui.showAll = false;
        render();
    });
    void pullAusgabenFromSharePoint().then((r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (r.ok) setMessage('ok', 'Mit SharePoint verbunden.');
        render();
    });
}
