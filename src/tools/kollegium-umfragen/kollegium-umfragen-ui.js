/**
 * Kollegiumsumfragen Admin-UI (MSAL-Kollegium, SharePoint SoR).
 */

import { loadAttachment } from './kollegium-umfragen-attachments.js';
import {
    auswertungToCsv,
    buildAuswertung,
    parseLehrerListe
} from './kollegium-umfragen-logic.js';
import {
    FRAGE_TYP,
    FRAGE_TYP_LABEL,
    STATUS,
    STATUS_LABEL,
    ZIELGRUPPE,
    ZIELGRUPPE_LABEL,
    newId,
    toIsoDateOnly
} from './kollegium-umfragen-model.js';
import {
    loadKollegiumSettings,
    resolveKollegiumSiteUrl,
    saveKollegiumSettings
} from './kollegium-umfragen-settings.js';
import {
    clearAll,
    clearDemoSeed,
    getUmfrage,
    listAntworten,
    loadDemoData,
    loadState,
    removeUmfrage,
    seedDemoIfEmpty,
    seedEmpfaenger,
    setUmfrageStatus,
    statsFor,
    upsertUmfrage
} from './kollegium-umfragen-store.js';
import {
    getKollegiumSyncMeta,
    pullKollegiumFromSharePoint,
    pushKollegiumAntwort,
    pushKollegiumDeleteUmfrage,
    pushKollegiumUmfrage
} from './kollegium-umfragen-sync.js';
import { VORLAGEN, draftFromVorlage } from './kollegium-umfragen-templates.js';
import {
    loadEinsammlerStammdaten,
    parseKlasseCodesField
} from '../einsammler/einsammler-stammdaten.js';
import {
    getKollegiumRoleCache,
    isKollegiumAdminCached,
    resolveKollegiumRoles
} from './kollegium-umfragen-entra-role.js';
import { bindSystemTabClicks, renderSystemAdminPage } from '../../shared/system-admin-shell.js';

const VIEWS = [
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2' },
    { id: 'vorlagen', label: 'Vorlagen', icon: 'bi-layout-text-window' },
    { id: 'liste', label: 'Umfragen', icon: 'bi-collection' },
    { id: 'detail', label: 'Auswertung', icon: 'bi-bar-chart' },
    { id: 'einstellungen', label: 'Systemadministration', icon: 'bi-gear' }
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

function pillClass(st) {
    if (st === 'aktiv') return 'ku-pill--aktiv';
    if (st === 'geschlossen') return 'ku-pill--ok';
    if (st === 'entwurf') return 'ku-pill--offen';
    return 'ku-pill--arch';
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    /** @type {{ view: string, selectedId: string, draft: object|null, message: {type:string,text:string}|null }} */
    const ui = {
        view: 'uebersicht',
        systemTab: 'site',
        selectedId: '',
        draft: null,
        message: null
    };

    seedDemoIfEmpty();

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    function state() {
        return loadState();
    }

    function accountEmail() {
        try {
            if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
                const info = window.ms365AuthGetAccountInfo();
                return String((info && (info.username || info.upn)) || '')
                    .trim()
                    .toLowerCase();
            }
        } catch {
            /* ignore */
        }
        return '';
    }

    function admin() {
        const em = accountEmail();
        return isKollegiumAdminCached(em) || getKollegiumRoleCache().isAdmin;
    }

    function render() {
        root.innerHTML =
            '<div class="ku-shell">' +
            renderNav() +
            '<div class="ku-main">' +
            renderMessage() +
            (admin()
                ? renderView()
                : '<div class="ku-head"><h2>Kollegiumsumfragen</h2><p>Umfragen anlegen und auswerten dürfen Schulleitung / Verwaltung / Schul-IT (Entra). Lehrkräfte antworten unter <a href="kollegium-umfragen-antwort.html">Antworten</a>.</p></div>') +
            '</div></div>';
        bind();
    }

    function renderNav() {
        const isAdm = admin();
        const buttons = (isAdm ? VIEWS : [])
            .map(
                (v) =>
                    '<button type="button" class="ku-nav__btn' +
                    (ui.view === v.id ? ' is-active' : '') +
                    '" data-ku-view="' +
                    esc(v.id) +
                    '"><i class="bi ' +
                    esc(v.icon) +
                    '" aria-hidden="true"></i>' +
                    esc(v.label) +
                    '</button>'
            )
            .join('');
        const role = getKollegiumRoleCache();
        return (
            '<nav class="ku-nav" aria-label="Kollegiumsumfragen">' +
            '<div class="ku-nav__brand">' +
            '<i class="bi bi-people" aria-hidden="true"></i>' +
            '<div><strong>Kollegium</strong><span>Umfragen · SharePoint · P4</span></div>' +
            '</div>' +
            '<div class="ku-nav__list">' +
            buttons +
            '</div>' +
            '<p class="ku-nav__hint">Admin = Schulleitung/Verwaltung (Entra). Lehrkräfte nur Antwort-Seite. Datenbasis SharePoint – nicht Forms.' +
            (isAdm && role.sources.length
                ? ' · ' + esc(role.sources.join(', '))
                : role.demo
                  ? ' · Demo'
                  : '') +
            '</p>' +
            '<div class="ku-nav__links">' +
            '<a href="kollegium-umfragen-antwort.html"><i class="bi bi-pencil-square"></i> Antworten (Lehrkräfte)</a>' +
            '<a href="sharepoint-liste-kollegium-umfragen.html"><i class="bi bi-list-ul"></i> SharePoint-Paket</a>' +
            '</div></nav>'
        );
    }

    function renderMessage() {
        if (!ui.message) return '';
        const cls =
            ui.message.type === 'ok'
                ? 'ku-msg--ok'
                : ui.message.type === 'warn'
                  ? 'ku-msg--warn'
                  : 'ku-msg--err';
        return '<p class="ku-msg ' + cls + '" role="status">' + esc(ui.message.text) + '</p>';
    }

    function renderView() {
        if (ui.view === 'vorlagen') return renderVorlagen();
        if (ui.view === 'liste') return renderListe();
        if (ui.view === 'detail') return renderDetail();
        if (ui.view === 'einstellungen') return renderEinstellungen();
        if (ui.view === 'neu') return renderNeu();
        return renderUebersicht();
    }

    function card(label, value) {
        return (
            '<div class="ku-card"><p class="ku-card__label">' +
            esc(label) +
            '</p><p class="ku-card__value">' +
            esc(value) +
            '</p></div>'
        );
    }

    function renderUebersicht() {
        const s = state();
        const aktiv = s.umfragen.filter((u) => u.status === STATUS.AKTIV).length;
        const offen = s.antworten.filter((a) => a.status === 'offen').length;
        const ab = s.antworten.filter((a) => a.status === 'abgegeben').length;
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Übersicht</h2>' +
            '<div class="ku-cards">' +
            card('Umfragen', s.umfragen.length) +
            card('Aktiv', aktiv) +
            card('Antworten offen', offen) +
            card('Abgegeben', ab) +
            '</div>' +
            '<div class="ku-toolbar">' +
            '<button type="button" class="btn btn-success" data-ku-go="vorlagen"><i class="bi bi-plus-circle"></i> Aus Vorlage</button>' +
            '<button type="button" class="btn" data-ku-go="liste">Alle Umfragen</button>' +
            '<a class="btn" href="kollegium-umfragen-antwort.html">Lehrkräfte-Ansicht</a>' +
            '</div>' +
            renderSyncHint() +
            '<p style="color:#64748b;font-size:0.9rem;margin:0;">SharePoint ist System of Record. Listen über „SharePoint-Paket“ – Auswertung bleibt im Schultool, ohne Forms.</p>'
        );
    }

    function renderSyncHint() {
        const sync = getKollegiumSyncMeta();
        if (sync.source === 'live' && sync.ok) return '';
        if (!sync.siteUrl && sync.source !== 'demo') {
            return (
                '<p class="ku-msg ku-msg--warn" role="status">Noch keine Site – IT richtet die Listen ein. Bis dahin lokal/Demo.</p>'
            );
        }
        if (sync.source === 'demo' || sync.source === 'local') {
            return (
                '<p class="ku-msg ku-msg--warn" role="status">' +
                esc(
                    sync.error ||
                        'SharePoint nicht verbunden – Daten bleiben lokal, bis die Listen greifen.'
                ) +
                '</p>'
            );
        }
        return '';
    }

    function renderVorlagen() {
        const cards = VORLAGEN.map(
            (v) =>
                '<button type="button" class="ku-vorlage" data-ku-vorlage="' +
                esc(v.id) +
                '"><strong>' +
                esc(v.titel) +
                '</strong><span>' +
                esc(v.beschreibung) +
                '</span><br><span class="ku-pill" style="margin-top:8px;">' +
                esc(v.komplexitaet) +
                ' · ' +
                esc(ZIELGRUPPE_LABEL[v.zielgruppe] || v.zielgruppe) +
                '</span></button>'
        ).join('');
        return (
            '<h2 style="margin:0 0 8px;font-size:1.15rem;">Vorlagen</h2>' +
            '<p style="color:#64748b;font-size:0.9rem;margin:0 0 14px;">Einfache Abstimmungen bis komplexere Randbedingungen – danach im Entwurf anpassbar.</p>' +
            '<div class="ku-vorlagen">' +
            cards +
            '</div>' +
            '<button type="button" class="btn" data-ku-blank>Leere Umfrage (komplex)</button>' +
            (ui.draft ? renderNeu() : '')
        );
    }

    function renderNeu() {
        const d = ui.draft;
        if (!d) return '';
        const fragenHtml = (d.fragen || [])
            .map((f, i) => {
                return (
                    '<div class="ku-frage" data-qi="' +
                    i +
                    '">' +
                    '<label>Frage ' +
                    (i + 1) +
                    '<input name="frageText" value="' +
                    esc(f.text) +
                    '"></label>' +
                    '<div class="ku-form__row">' +
                    '<label>Typ<select name="frageTyp">' +
                    Object.keys(FRAGE_TYP_LABEL)
                        .map(
                            (k) =>
                                '<option value="' +
                                esc(k) +
                                '"' +
                                (f.typ === k ? ' selected' : '') +
                                '>' +
                                esc(FRAGE_TYP_LABEL[k]) +
                                '</option>'
                        )
                        .join('') +
                    '</select></label>' +
                    '<label>Optionen (eine pro Zeile)<textarea name="frageOpt" rows="3">' +
                    esc((f.optionen || []).join('\n')) +
                    '</textarea></label>' +
                    '</div></div>'
                );
            })
            .join('');

        return (
            '<h2 style="margin:16px 0 10px;font-size:1.15rem;">Umfrage bearbeiten</h2>' +
            '<form class="ku-form" id="kuDraftForm">' +
            '<label>Titel<input name="titel" required value="' +
            esc(d.titel) +
            '"></label>' +
            '<label>Beschreibung<textarea name="beschreibung" rows="2">' +
            esc(d.beschreibung) +
            '</textarea></label>' +
            '<div class="ku-form__row">' +
            '<label>Zielgruppe<select name="zielgruppe">' +
            Object.keys(ZIELGRUPPE_LABEL)
                .map(
                    (k) =>
                        '<option value="' +
                        esc(k) +
                        '"' +
                        (d.zielgruppe === k ? ' selected' : '') +
                        '>' +
                        esc(ZIELGRUPPE_LABEL[k]) +
                        '</option>'
                )
                .join('') +
            '</select></label>' +
            '<label>Fällig am<input type="date" name="faelligAm" value="' +
            esc(d.faelligAm || '') +
            '"></label></div>' +
            '<label>Abteilung (bei Zielgruppe Abteilung)<input name="abteilung" value="' +
            esc(d.abteilung || '') +
            '"></label>' +
            (() => {
                const sd = loadEinsammlerStammdaten();
                const selected = new Set(
                    parseKlasseCodesField(d.klasseCodes || '').map((c) => c.toLowerCase())
                );
                if (!sd.classes.length) {
                    return (
                        '<label>Klassen-Bezug (optional, Codes)<input name="klasseCodes" id="kuKlasseCodes" value="' +
                        esc(d.klasseCodes || '') +
                        '" placeholder="3A;2B"></label>'
                    );
                }
                return (
                    '<div><span style="font-size:0.85rem;font-weight:600;color:#334155;">Klassen-Bezug (optional)</span>' +
                    '<div class="ku-class-pick" id="kuClassPick" style="display:flex;flex-wrap:wrap;gap:8px;margin:8px 0;">' +
                    sd.classes
                        .map((c) => {
                            const on = selected.has(c.code.toLowerCase());
                            return (
                                '<label style="display:inline-flex;align-items:center;gap:6px;padding:6px 10px;border:1px solid #cbd5e1;border-radius:999px;background:' +
                                (on ? '#e0e7ff' : '#fff') +
                                ';font-size:0.85rem;cursor:pointer;"><input type="checkbox" data-ku-klasse="' +
                                esc(c.code) +
                                '"' +
                                (on ? ' checked' : '') +
                                '> ' +
                                esc(c.code) +
                                '</label>'
                            );
                        })
                        .join('') +
                    '</div>' +
                    '<label>Klassen-Codes<input name="klasseCodes" id="kuKlasseCodes" value="' +
                    esc(d.klasseCodes || '') +
                    '" placeholder="3A;2B"></label></div>'
                );
            })() +
            '<label>Empfänger-Mails (Semikolon; aus Stammdaten einfügen)<textarea name="empfaengerMails" rows="2">' +
            esc(d.empfaengerMails || '') +
            '</textarea></label>' +
            '<label class="ku-check"><input type="checkbox" name="anonym"' +
            (d.anonym ? ' checked' : '') +
            '> Anonym auswerten</label>' +
            '<label class="ku-check"><input type="checkbox" name="anhangErlaubt"' +
            (d.anhangErlaubt ? ' checked' : '') +
            '> Zusätzlichen Datei-Anhang bei Antwort erlauben</label>' +
            '<p class="ku-item__meta">Für Belege/Dokumente: Frage-Typ „Datei-Upload“ oder optionalen Anhang aktivieren.</p>' +
            '<div><strong style="font-size:0.9rem;">Fragen</strong>' +
            fragenHtml +
            '</div>' +
            '<div class="ku-form__actions">' +
            '<button type="submit" class="btn btn-success">Als Entwurf speichern</button>' +
            '<button type="button" class="btn" data-ku-save-aktiv>Speichern &amp; aktivieren</button>' +
            '<button type="button" class="btn" data-ku-cancel-draft>Abbrechen</button>' +
            '</div></form>'
        );
    }

    function renderListe() {
        const s = state();
        const rows = s.umfragen
            .map((u) => {
                const st = statsFor(u);
                return (
                    '<article class="ku-item">' +
                    '<div class="ku-item__top">' +
                    '<h3 class="ku-item__title">' +
                    esc(u.titel) +
                    '</h3>' +
                    '<span class="ku-pill ' +
                    pillClass(u.status) +
                    '">' +
                    esc(STATUS_LABEL[u.status] || u.status) +
                    '</span></div>' +
                    '<p class="ku-item__meta">' +
                    esc(ZIELGRUPPE_LABEL[u.zielgruppe] || u.zielgruppe) +
                    ' · ' +
                    esc(u.komplexitaet) +
                    ' · ' +
                    st.abgegeben +
                    '/' +
                    st.total +
                    ' Antworten · Frist ' +
                    formatDeDate(u.faelligAm) +
                    '</p>' +
                    '<div class="ku-item__actions">' +
                    '<button type="button" class="btn" data-ku-open="' +
                    esc(u.id) +
                    '">Öffnen</button>' +
                    (u.status === STATUS.ENTWURF
                        ? '<button type="button" class="btn btn-success" data-ku-status="' +
                          esc(u.id) +
                          '" data-status="aktiv">Aktivieren</button>'
                        : '') +
                    (u.status === STATUS.AKTIV
                        ? '<button type="button" class="btn" data-ku-status="' +
                          esc(u.id) +
                          '" data-status="geschlossen">Schließen</button>'
                        : '') +
                    '<button type="button" class="btn" data-ku-del="' +
                    esc(u.id) +
                    '">Löschen</button>' +
                    '</div></article>'
                );
            })
            .join('');
        return (
            '<h2 style="margin:0 0 12px;font-size:1.15rem;">Umfragen</h2>' +
            '<div class="ku-toolbar">' +
            '<button type="button" class="btn btn-success" data-ku-go="vorlagen">Neu aus Vorlage</button>' +
            '</div>' +
            '<div class="ku-list">' +
            (rows || '<p class="ku-item__meta">Noch keine Umfragen.</p>') +
            '</div>'
        );
    }

    function renderDetail() {
        const u = getUmfrage(ui.selectedId) || state().umfragen[0];
        if (!u) {
            return '<p>Keine Umfrage gewählt. <button type="button" class="btn" data-ku-go="liste">Zur Liste</button></p>';
        }
        ui.selectedId = u.id;
        const antworten = listAntworten(u.id);
        const st = statsFor(u);
        const aus = buildAuswertung(u, antworten);
        const select =
            '<label style="display:block;margin-bottom:12px;font-size:0.85rem;font-weight:600;">Umfrage ' +
            '<select id="kuSelectUmfrage">' +
            state()
                .umfragen.map(
                    (x) =>
                        '<option value="' +
                        esc(x.id) +
                        '"' +
                        (x.id === u.id ? ' selected' : '') +
                        '>' +
                        esc(x.titel) +
                        '</option>'
                )
                .join('') +
            '</select></label>';

        const bars = aus.fragen
            .map((f) => {
                if (f.typ === FRAGE_TYP.TEXT) {
                    const lis = (f.texte || [])
                        .map((t) => '<li>' + esc(t) + '</li>')
                        .join('');
                    return (
                        '<div class="ku-item"><strong>' +
                        esc(f.text) +
                        '</strong><ul style="margin:8px 0 0;padding-left:1.2em;">' +
                        (lis || '<li class="ku-item__meta">Keine Freitexte</li>') +
                        '</ul></div>'
                    );
                }
                const max = Math.max(1, ...f.buckets.map((b) => b.count));
                const rows = f.buckets
                    .map(
                        (b) =>
                            '<div class="ku-bar"><span style="min-width:110px;">' +
                            esc(b.label) +
                            '</span><div class="ku-bar__track"><div class="ku-bar__fill" style="width:' +
                            Math.round((b.count / max) * 100) +
                            '%"></div></div><span>' +
                            b.count +
                            '</span></div>'
                    )
                    .join('');
                return (
                    '<div class="ku-item"><strong>' +
                    esc(f.text) +
                    '</strong>' +
                    (f.avg != null
                        ? '<p class="ku-item__meta">Durchschnitt: ' + f.avg + '</p>'
                        : '') +
                    rows +
                    '</div>'
                );
            })
            .join('');

        const empfaengerRows = antworten
            .map((a) => {
                const fileBtn =
                    a.anhangName || a.anhangKey
                        ? ' <button type="button" data-ku-anhang="' +
                          esc(a.id) +
                          '" style="font-size:0.78rem;">📎 ' +
                          esc(a.anhangName || 'Anhang') +
                          '</button>'
                        : '';
                return (
                    '<tr><td>' +
                    esc(a.lehrerName || '–') +
                    fileBtn +
                    '</td><td>' +
                    esc(u.anonym && a.status === 'abgegeben' ? '(anonym)' : a.lehrerEmail) +
                    '</td><td>' +
                    esc(a.status) +
                    '</td></tr>'
                );
            })
            .join('');

        return (
            select +
            '<div class="ku-cards">' +
            card('Status', STATUS_LABEL[u.status] || u.status) +
            card('Fortschritt', st.pct + '%') +
            card('Abgegeben', st.abgegeben + '/' + st.total) +
            card('Zielgruppe', ZIELGRUPPE_LABEL[u.zielgruppe] || u.zielgruppe) +
            '</div>' +
            '<div class="ku-toolbar">' +
            '<button type="button" class="btn" data-ku-export>CSV exportieren</button>' +
            '<a class="btn" href="kollegium-umfragen-antwort.html?id=' +
            encodeURIComponent(u.id) +
            '">Antwort-Link</a>' +
            '</div>' +
            '<h3 style="font-size:1rem;margin:16px 0 8px;">Auswertung</h3>' +
            '<div class="ku-list">' +
            bars +
            '</div>' +
            '<h3 style="font-size:1rem;margin:18px 0 8px;">Empfänger / Status</h3>' +
            '<form id="kuSeedForm" class="ku-form" style="margin-bottom:12px;">' +
            '<label>Lehrkräfte hinzufügen (Name;E-Mail[;KV] pro Zeile)<textarea name="liste" rows="3" placeholder="Müller;mueller@schule.at&#10;Schmidt;schmidt@schule.at;KV"></textarea></label>' +
            '<div class="ku-form__actions"><button type="submit" class="btn">Übernehmen</button></div>' +
            '</form>' +
            '<div style="overflow:auto;"><table class="tm-table" style="width:100%;border-collapse:collapse;font-size:0.9rem;">' +
            '<thead><tr><th align="left">Name</th><th align="left">E-Mail</th><th align="left">Status</th></tr></thead>' +
            '<tbody>' +
            (empfaengerRows || '<tr><td colspan="3">Noch keine Empfängerzeilen.</td></tr>') +
            '</tbody></table></div>'
        );
    }

    function renderEinstellungen() {
        const set = loadKollegiumSettings();
        const role = getKollegiumRoleCache();
        const sync = getKollegiumSyncMeta();
        const site = resolveKollegiumSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
        const s = state();
        const who =
            'Admin aus Schulleitung/Verwaltung/Schul-IT (Entra)' +
            (role.demo ? ' · Demo' : role.sources.length ? ' · ' + role.sources.join(', ') : '') +
            '. Zielgruppe „alle Lehrkräfte“ prüft Lehrer-Gruppe/Stammdaten – nicht jedes MSAL-Konto.';
        const syncLine =
            'Sync: ' +
            sync.source +
            (sync.listTitle ? ' · „' + sync.listTitle + '“' : '') +
            (sync.lastPullAt ? ' · zuletzt ' + formatDeDate(sync.lastPullAt) : '') +
            (sync.connected ? ' · verbunden' : ' · nicht verbunden') +
            (sync.error ? ' · ' + sync.error : '');
        return renderSystemAdminPage({
            prefix: 'ku',
            activeTab: ui.systemTab,
            dataAttr: 'data-ku-sys-tab',
            ariaLabel: 'Systemadministration Kollegiumsumfragen',
            panels: [
                {
                    id: 'site',
                    step: '1',
                    title: 'Datenquelle',
                    leadHtml: 'Umfragen und Antworten in SharePoint. Microsoft Forms ist nicht die Quelle.',
                    body:
                        '<p><code>' +
                        esc(site) +
                        '</code></p><p class="muted">' +
                        esc(syncLine) +
                        '</p><p class="muted">Lokal: ' +
                        (s.umfragen || []).length +
                        ' Umfragen · ' +
                        (s.antworten || []).length +
                        ' Antworten' +
                        (s.localDemoOnly ? ' · Demo local-first' : '') +
                        ' · mit spoItemId: ' +
                        (s.umfragen || []).filter((u) => u.spoItemId).length +
                        ' / ' +
                        (s.antworten || []).filter((a) => a.spoItemId).length +
                        '</p><div class="ms-sys-actions">' +
                        '<button type="button" class="btn ms-sys-btn-primary" data-ku-pull>Von SharePoint laden</button>' +
                        '<a class="btn ms-sys-btn-secondary" href="sharepoint-liste-kollegium-umfragen.html">SharePoint-Paket anlegen</a>' +
                        '</div>'
                },
                {
                    id: 'perms',
                    step: '2',
                    title: 'Berechtigungen',
                    leadHtml: esc(who),
                    body:
                        '<div class="ms-sys-links"><a class="btn ms-sys-btn-secondary" href="sharepoint-liste-kollegium-umfragen.html"><i class="bi bi-shield-check"></i>Listen-Paket</a></div>'
                },
                {
                    id: 'kalender',
                    step: '3',
                    title: 'Kalender',
                    leadHtml: 'Kein Terminkalender. Laufzeit und Status stehen an der Umfrage.',
                    body: '<p class="muted">Antworten laufen über die Antwortseite, nicht über Outlook.</p>'
                },
                {
                    id: 'wartung',
                    step: '4',
                    title: 'Wartung',
                    leadHtml: 'Anzeige-Name und Standard für neue Umfragen.',
                    body:
                        '<form class="ku-form" id="kuSettingsForm">' +
                        '<label>Schulname (Anzeige)<input name="schoolName" value="' +
                        esc(set.schoolName) +
                        '"></label>' +
                        '<label class="ku-check"><input type="checkbox" name="defaultAnonym"' +
                        (set.defaultAnonym ? ' checked' : '') +
                        '> Neue Umfragen standardmäßig anonym</label>' +
                        '<div class="ku-form__actions"><button type="submit" class="btn btn-success">Speichern</button>' +
                        '<button type="button" class="btn" data-ku-demo>Demo laden</button></div></form>'
                },
                {
                    id: 'danger',
                    title: 'Danger Zone',
                    danger: true,
                    leadHtml: 'Lokal leeren lässt die SharePoint-Liste stehen.',
                    body:
                        '<div class="ms-sys-actions">' +
                        '<button type="button" class="btn ms-sys-btn-danger" data-ku-clear-seed>Demo-Seed entfernen</button>' +
                        '<button type="button" class="btn ms-sys-btn-danger" data-ku-clear>Lokal leeren</button>' +
                        '</div>'
                }
            ]
        });
    }

    function readDraftFromForm(form, activate) {
        const fd = new FormData(form);
        const fragenBlocks = [...form.querySelectorAll('.ku-frage')];
        const fragen = fragenBlocks.map((block, i) => {
            const prev = (ui.draft && ui.draft.fragen && ui.draft.fragen[i]) || {};
            const typ = String(block.querySelector('[name="frageTyp"]')?.value || 'single');
            const optText = String(block.querySelector('[name="frageOpt"]')?.value || '');
            return {
                id: prev.id || newId('q'),
                typ,
                text: String(block.querySelector('[name="frageText"]')?.value || ''),
                pflicht: prev.pflicht !== false,
                optionen: optText
                    .split(/\r?\n/)
                    .map((x) => x.trim())
                    .filter(Boolean),
                skalaMin: prev.skalaMin || 1,
                skalaMax: prev.skalaMax || 5,
                skalaLabelMin: prev.skalaLabelMin || '',
                skalaLabelMax: prev.skalaLabelMax || ''
            };
        });
        return {
            ...(ui.draft || {}),
            titel: String(fd.get('titel') || ''),
            beschreibung: String(fd.get('beschreibung') || ''),
            zielgruppe: String(fd.get('zielgruppe') || ZIELGRUPPE.ALLE),
            abteilung: String(fd.get('abteilung') || ''),
            klasseCodes: String(fd.get('klasseCodes') || ''),
            empfaengerMails: String(fd.get('empfaengerMails') || ''),
            faelligAm: String(fd.get('faelligAm') || ''),
            anonym: !!form.querySelector('[name="anonym"]')?.checked,
            anhangErlaubt: !!form.querySelector('[name="anhangErlaubt"]')?.checked,
            fragen,
            status: activate ? STATUS.AKTIV : STATUS.ENTWURF,
            komplexitaet: fragen.length > 3 ? 'komplex' : 'einfach'
        };
    }

    function bind() {
        bindSystemTabClicks(root, {
            dataAttr: 'data-ku-sys-tab',
            onSelect: (tab) => {
                ui.systemTab = tab;
                render();
            }
        });
        root.querySelectorAll('[data-ku-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ku-view') || 'uebersicht';
                ui.message = null;
                render();
            });
        });
        root.querySelectorAll('[data-ku-go]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-ku-go') || 'uebersicht';
                render();
            });
        });
        root.querySelectorAll('[data-ku-vorlage]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const d = draftFromVorlage(btn.getAttribute('data-ku-vorlage') || '');
                if (!d) return;
                const set = loadKollegiumSettings();
                if (set.defaultAnonym) d.anonym = true;
                ui.draft = d;
                ui.view = 'vorlagen';
                setMessage('ok', 'Vorlage geladen – bitte prüfen und speichern.');
                render();
            });
        });
        root.querySelector('[data-ku-blank]')?.addEventListener('click', () => {
            ui.draft = {
                id: newId('ku'),
                titel: '',
                beschreibung: '',
                vorlageId: '',
                komplexitaet: 'komplex',
                zielgruppe: ZIELGRUPPE.ALLE,
                abteilung: '',
                klasseCodes: '',
                empfaengerMails: '',
                fragen: [
                    {
                        id: newId('q'),
                        typ: FRAGE_TYP.SINGLE,
                        text: '',
                        pflicht: true,
                        optionen: ['Option A', 'Option B']
                    }
                ],
                faelligAm: toIsoDateOnly(new Date()),
                anonym: loadKollegiumSettings().defaultAnonym,
                status: STATUS.ENTWURF,
                quelle: 'ui'
            };
            render();
        });
        root.querySelector('[data-ku-cancel-draft]')?.addEventListener('click', () => {
            ui.draft = null;
            render();
        });

        root.querySelectorAll('[data-ku-klasse]').forEach((cb) => {
            cb.addEventListener('change', () => {
                const codes = Array.from(root.querySelectorAll('[data-ku-klasse]:checked')).map(
                    (el) => String(el.getAttribute('data-ku-klasse') || '')
                );
                const field = root.querySelector('#kuKlasseCodes');
                if (field) field.value = codes.join(';');
                if (ui.draft) ui.draft.klasseCodes = codes.join(';');
                const label = cb.closest('label');
                if (label) {
                    label.style.background = cb.checked ? '#e0e7ff' : '#fff';
                }
            });
        });

        const form = root.querySelector('#kuDraftForm');
        if (form) {
            form.addEventListener('submit', (ev) => {
                ev.preventDefault();
                saveDraft(false);
            });
            root.querySelector('[data-ku-save-aktiv]')?.addEventListener('click', () => saveDraft(true));
        }

        function saveDraft(activate) {
            if (!form) return;
            const draft = readDraftFromForm(form, activate);
            const res = upsertUmfrage(draft);
            if (!res.ok) {
                setMessage('err', (res.errors || []).join(' '));
                render();
                return;
            }
            // Empfängerzeilen aus Mails anlegen
            const people = parseLehrerListe(
                String(draft.empfaengerMails || '')
                    .split(/[;,\s]+/)
                    .filter((m) => m.includes('@'))
                    .map((m) => m + ';' + m)
                    .join('\n')
            );
            /** @type {object[]} */
            let seeded = [];
            if (people.length) {
                const seed = seedEmpfaenger(
                    res.item.id,
                    people.map((p) => p.name + ';' + p.email).join('\n')
                );
                seeded = seed.addedItems || [];
            }
            const saved = getUmfrage(res.item.id) || res.item;
            ui.draft = null;
            ui.selectedId = saved.id;
            ui.view = 'detail';
            setMessage('ok', activate ? 'Umfrage aktiv – speichere nach SharePoint …' : 'Entwurf gespeichert …');
            render();
            void pushKollegiumUmfrage(saved).then(async (push) => {
                for (let i = 0; i < seeded.length; i++) {
                    await pushKollegiumAntwort(seeded[i]);
                }
                setMessage(
                    push.ok ? 'ok' : 'warn',
                    push.ok
                        ? activate
                            ? 'Umfrage aktiv und in SharePoint.'
                            : 'Entwurf in SharePoint gespeichert.'
                        : (activate ? 'Umfrage lokal aktiv. ' : 'Lokal gespeichert. ') +
                          (push.error || 'SharePoint-Push fehlgeschlagen.')
                );
                render();
            });
        }

        root.querySelectorAll('[data-ku-open]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.selectedId = btn.getAttribute('data-ku-open') || '';
                ui.view = 'detail';
                render();
            });
        });
        root.querySelectorAll('[data-ku-status]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const res = setUmfrageStatus(
                    btn.getAttribute('data-ku-status') || '',
                    btn.getAttribute('data-status') || ''
                );
                setMessage('ok', 'Status aktualisiert …');
                render();
                if (res.ok && res.item) {
                    void pushKollegiumUmfrage(res.item).then((push) => {
                        setMessage(
                            push.ok ? 'ok' : 'warn',
                            push.ok
                                ? 'Status in SharePoint aktualisiert.'
                                : 'Lokal aktualisiert. ' + (push.error || '')
                        );
                        render();
                    });
                }
            });
        });
        root.querySelectorAll('[data-ku-del]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!window.confirm('Umfrage wirklich löschen?')) return;
                const del = removeUmfrage(btn.getAttribute('data-ku-del') || '');
                setMessage('ok', 'Gelöscht …');
                render();
                void pushKollegiumDeleteUmfrage(del.removed || {}, del.antworten || []).then(
                    (push) => {
                        setMessage(
                            push.ok ? 'ok' : 'warn',
                            push.ok
                                ? 'Gelöscht.'
                                : 'Lokal gelöscht. ' + (push.error || 'SharePoint-Löschen fehlgeschlagen.')
                        );
                        render();
                    }
                );
            });
        });

        root.querySelector('#kuSelectUmfrage')?.addEventListener('change', (ev) => {
            ui.selectedId = ev.target.value;
            render();
        });

        root.querySelector('#kuSeedForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const fd = new FormData(ev.target);
            const res = seedEmpfaenger(ui.selectedId, String(fd.get('liste') || ''));
            setMessage(
                res.ok ? 'ok' : 'err',
                res.ok ? res.added + ' Empfänger ergänzt …' : (res.errors || []).join(' ')
            );
            render();
            if (res.ok) {
                void (async () => {
                    if (res.umfrage) await pushKollegiumUmfrage(res.umfrage);
                    for (let i = 0; i < (res.addedItems || []).length; i++) {
                        await pushKollegiumAntwort(res.addedItems[i]);
                    }
                    setMessage('ok', res.added + ' Empfänger ergänzt.');
                    render();
                })();
            }
        });

        root.querySelector('[data-ku-export]')?.addEventListener('click', () => {
            const u = getUmfrage(ui.selectedId);
            if (!u) return;
            const csv = auswertungToCsv(buildAuswertung(u, listAntworten(u.id)));
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'umfrage-' + u.id + '.csv';
            a.click();
            URL.revokeObjectURL(a.href);
        });

        root.querySelectorAll('[data-ku-anhang]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const id = String(btn.getAttribute('data-ku-anhang') || '');
                const a = listAntworten(ui.selectedId).find((x) => x.id === id);
                if (!a) return;
                const file = await loadAttachment(a.anhangKey || a.id);
                if (!file || !file.dataUrl) {
                    setMessage(
                        'warn',
                        'Anhang hier nicht gefunden (anderer Browser) – SharePoint-URL folgt später.'
                    );
                    render();
                    return;
                }
                const link = document.createElement('a');
                link.href = file.dataUrl;
                link.download = file.name || a.anhangName || 'anhang';
                document.body.appendChild(link);
                link.click();
                link.remove();
            });
        });

        root.querySelector('#kuSettingsForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const fd = new FormData(ev.target);
            saveKollegiumSettings({
                schoolName: String(fd.get('schoolName') || ''),
                defaultAnonym: !!ev.target.querySelector('[name="defaultAnonym"]')?.checked
            });
            setMessage('ok', 'Einstellungen gespeichert.');
            render();
        });
        root.querySelector('[data-ku-pull]')?.addEventListener('click', () => {
            setMessage('ok', 'Lade von SharePoint …');
            render();
            void pullKollegiumFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                setMessage(
                    r.ok ? 'ok' : 'warn',
                    r.ok
                        ? 'SharePoint geladen (' +
                          (r.state.umfragen || []).length +
                          ' Umfragen · ' +
                          (r.state.antworten || []).length +
                          ' Antworten).'
                        : r.softError || r.meta.error || 'Pull fehlgeschlagen.'
                );
                render();
            });
        });
        root.querySelector('[data-ku-demo]')?.addEventListener('click', () => {
            loadDemoData();
            setMessage('ok', 'Demo geladen (local-first, ohne SPO-Push der Seeds).');
            ui.view = 'liste';
            render();
        });
        root.querySelector('[data-ku-clear-seed]')?.addEventListener('click', () => {
            clearDemoSeed();
            setMessage('ok', 'Demo-Seed entfernt.');
            render();
        });
        root.querySelector('[data-ku-clear]')?.addEventListener('click', () => {
            if (!window.confirm('Lokale Daten leeren?')) return;
            clearAll();
            setMessage('ok', 'Geleert.');
            render();
        });
    }

    render();
    void resolveKollegiumRoles(accountEmail()).then(() => render());
    void pullKollegiumFromSharePoint().then((r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (r.ok) setMessage('ok', 'Mit SharePoint verbunden.');
        render();
    });
}
