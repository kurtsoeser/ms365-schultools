/**
 * Einsammler Admin-UI (P1 intern + P2 Magic-Links).
 */

import {
    importResponseCode,
    listInbox,
    mergeInboxIntoStore
} from './einsammler-inbox.js';
import {
    allowedAntwortStatusesForModus,
    progressStats,
    resolveAntwortStatus
} from './einsammler-logic.js';
import {
    ANTWORT_STATUS_LABEL,
    EINSAMMLER_STATUS,
    EINSAMMLER_STATUS_LABEL,
    MODUS,
    MODUS_LABEL,
    TYP_LABEL,
    newId
} from './einsammler-model.js';
import { einsammlerRemoteConfigured, remotePublishPacks } from './einsammler-remote.js';
import { loadAttachment } from './einsammler-attachments.js';
import {
    filterAntwortenForHeads,
    filterEinsammlerForViewer,
    getEinsammlerAccount,
    getEinsammlerAdminAccessCache,
    isEinsammlerAdmin,
    isOwnedBy
} from './einsammler-scope.js';
import { resolveEinsammlerAdmin } from './einsammler-entra-role.js';
import {
    loadEinsammlerSettings,
    resolveEinsammlerSiteUrl,
    saveEinsammlerSettings
} from './einsammler-settings.js';
import { readDemoRoleOverrideFromUrl } from '../../shared/app-role-switcher-policy.js';
import { renderToolOpsStatusCard } from '../../shared/tool-ops-status.js';
import {
    addEmpfaengerFromText,
    clearAll,
    clearDemoSeeds,
    loadDemoData,
    loadState,
    removeAntwort,
    removeEinsammler,
    seedDemoIfEmpty,
    setAntwortStatus,
    setEinsammlerStatus,
    upsertEinsammler
} from './einsammler-store.js';
import {
    buildEmpfaengerLinesForKlassen,
    getEinsammlerStammdatenMeta,
    hydrateEinsammlerStammdaten,
    loadEinsammlerStammdaten,
    parseKlasseCodesField
} from './einsammler-stammdaten.js';
import {
    getEinsammlerSyncMeta,
    pullEinsammlerFromSharePoint,
    pushEinsammlerAntwort,
    pushEinsammlerAntwortDelete,
    pushEinsammlerAntworten,
    pushEinsammlerDelete,
    pushEinsammlerKopf
} from './einsammler-sync.js';
import {
    buildPublicPayload,
    buildPublicUrl,
    encodePublicToken
} from './einsammler-token.js';
import { bindSystemTabClicks, renderSystemAdminPage } from '../../shared/system-admin-shell.js';

const VIEWS = [
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2' },
    { id: 'liste', label: 'Einsammler', icon: 'bi-collection' },
    { id: 'neu', label: 'Neu anlegen', icon: 'bi-plus-circle' },
    { id: 'detail', label: 'Antworten', icon: 'bi-people' },
    { id: 'links', label: 'Eltern-Links', icon: 'bi-link-45deg' },
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

function pillClassForStatus(st) {
    const s = String(st || '').toLowerCase();
    if (s === 'zugestimmt' || s === 'gelesen') return 'ein-pill--ok';
    if (s === 'aktiv') return 'ein-pill--aktiv';
    if (s === 'offen' || s === 'entwurf') return 'ein-pill--offen';
    if (s === 'abgelehnt' || s === 'abgelaufen') return 'ein-pill--bad';
    return '';
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    /** @type {{ view: string, selectedId: string, message: { type: string, text: string }|null, form: object, showAll: boolean, stammdatenHint: string }} */
    const ui = {
        view: 'uebersicht',
        systemTab: 'site',
        selectedId: '',
        message: null,
        form: emptyForm(),
        showAll: false,
        stammdatenHint: ''
    };

    const demoUiEnabled = readDemoRoleOverrideFromUrl(false);

    seedDemoIfEmpty();
    try {
        const shot = new URLSearchParams(window.location.search || '');
        if (shot.get('landingShot') === '1') {
            loadDemoData({ email: 'mueller@schule.demo', name: 'Maria Müller' });
            ui.showAll = true;
        }
    } catch {
        /* ignore */
    }
    try {
        mergeInboxIntoStore();
    } catch {
        /* ignore */
    }

    function syncHeaderChrome() {
        const sp = document.getElementById('einHeaderSpSetup');
        if (sp) sp.hidden = !admin();
    }

    function canUseDemoActions() {
        return admin() || demoUiEnabled;
    }

    function emptyForm() {
        const in7 = new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10);
        return {
            id: '',
            titel: '',
            modus: MODUS.EINWILLIGUNG,
            typ: 'exkursion',
            textInhalt: '',
            klasseCodes: '',
            faelligAm: in7,
            empfaengerText: '',
            uploadErlaubt: false,
            uploadPflicht: false,
            status: EINSAMMLER_STATUS.ENTWURF
        };
    }

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    /**
     * @param {object|null} item
     * @param {'kopf'|'antwort'} kind
     * @param {string} [okMsg]
     */
    async function afterLocalWrite(item, kind, okMsg) {
        if (!item) {
            setMessage('ok', okMsg || 'Gespeichert.');
            render();
            return { ok: true, skipped: true };
        }
        const push =
            kind === 'antwort'
                ? await pushEinsammlerAntwort(item)
                : await pushEinsammlerKopf(item);
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

    /**
     * @param {object[]} items
     * @param {string} [okMsg]
     */
    async function afterLocalAntworten(items, okMsg) {
        const list = Array.isArray(items) ? items : [];
        if (!list.length) {
            setMessage('ok', okMsg || 'Gespeichert.');
            render();
            return;
        }
        const push = await pushEinsammlerAntworten(list);
        if (push.ok) {
            setMessage('ok', okMsg || 'In SharePoint gespeichert.');
        } else {
            setMessage(
                'warn',
                (okMsg ? okMsg + ' ' : '') +
                    (push.error || 'SharePoint-Sync teilweise fehlgeschlagen – lokal gespeichert.')
            );
        }
        render();
    }

    function state() {
        return loadState();
    }

    function account() {
        return getEinsammlerAccount();
    }

    function admin() {
        return isEinsammlerAdmin(account().email);
    }

    function visibleHeads(s) {
        const acc = account();
        return filterEinsammlerForViewer(s.einsammler || [], {
            email: acc.email,
            admin: admin(),
            showAll: ui.showAll
        });
    }

    function canAccessHead(head) {
        if (!head) return false;
        if (admin() && ui.showAll) return true;
        const acc = account();
        if (!acc.loggedIn) return false;
        return isOwnedBy(head, acc.email) || (admin() && ui.showAll);
    }

    function selectedHead(s) {
        const id = ui.selectedId;
        const head = (s.einsammler || []).find((e) => e.id === id) || null;
        if (!head || !canAccessHead(head)) return null;
        return head;
    }

    function statusCard() {
        const acc = account();
        const adm = admin();
        return renderToolOpsStatusCard({
            nounPlural: 'Einsammler',
            listWord: 'Listen',
            requireLogin: true,
            loggedIn: !!acc.loggedIn,
            accountLabel: acc.name || acc.email || '',
            isAdmin: adm,
            showAll: ui.showAll,
            scopeOwnLabel: 'nur eigene',
            scopeAllLabel: 'alle Vorgänge',
            scopeToggleId: 'einShowAll',
            scopeToggleLabel: 'Alle Einsammler (Admin)',
            sync: getEinsammlerSyncMeta(),
            setupHref: 'sharepoint-liste-einsammler.html',
            setupLabel: 'Listen einrichten',
            esc
        });
    }

    function render() {
        syncHeaderChrome();
        const s = state();
        if ((ui.view === 'detail' || ui.view === 'links') && !selectedHead(s)) {
            ui.view = 'liste';
            ui.selectedId = '';
        }
        if (ui.view === 'einstellungen' && !admin()) {
            /* Lehrkräfte: persönliche Kurzinfo statt Admin-API */
        }
        root.innerHTML =
            '<div class="ein-shell">' +
            renderNav() +
            '<div class="ein-main">' +
            statusCard() +
            renderMessage() +
            renderView(s) +
            '</div></div>';
        bind(s);
    }

    function renderNav() {
        const isAdm = admin();
        const buttons = VIEWS.filter((v) => {
            if (v.id === 'detail' || v.id === 'links') return !!ui.selectedId;
            if (v.id === 'einstellungen') return true;
            return true;
        })
            .map((v) => {
                const label =
                    v.id === 'einstellungen' && !isAdm ? 'Mein Bereich' : v.label;
                return (
                    '<button type="button" class="ein-nav__btn' +
                    (ui.view === v.id ? ' is-active' : '') +
                    '" data-ein-view="' +
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
            '<nav class="ein-nav" aria-label="Einsammler">' +
            '<div class="ein-nav__brand">' +
            '<i class="bi bi-inbox" aria-hidden="true"></i>' +
            '<div><strong>Einsammler</strong><span>Persönlich · Eltern &amp; Ablauf</span></div>' +
            '</div>' +
            '<div class="ein-nav__list">' +
            buttons +
            '</div>' +
            '<p class="ein-nav__hint">Jeder Lehrer sieht nur die eigenen Einsammler. Admin kann „Alle“ aktivieren und Einstellungen pflegen.</p>' +
            '<div class="ein-nav__links">' +
            '<a href="einsammler-antwort.html" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i> Öffentliche Seite</a>' +
            (isAdm
                ? '<a href="sharepoint-liste-einsammler.html"><i class="bi bi-list-ul"></i> SharePoint-Listen</a>'
                : '') +
            '</div></nav>'
        );
    }

    function quietEmpty(title, text, ctaView, ctaLabel) {
        return (
            '<div class="ein-empty">' +
            '<p class="ein-empty__title">' +
            esc(title) +
            '</p>' +
            '<p class="ein-empty__text">' +
            esc(text) +
            '</p>' +
            (ctaView
                ? '<button type="button" class="btn btn-success" data-ein-view="' +
                  esc(ctaView) +
                  '"><i class="bi bi-plus-lg" aria-hidden="true"></i> ' +
                  esc(ctaLabel || 'Neu anlegen') +
                  '</button>'
                : '') +
            '</div>'
        );
    }

    /**
     * Kontextuelles Action-Banner (kein Dauer-How-to).
     * @param {object[]} heads
     * @param {object[]} antworten
     */
    function actionBanner(heads, antworten) {
        const drafts = (heads || []).filter((h) => h.status === EINSAMMLER_STATUS.ENTWURF);
        if (drafts.length) {
            const first = drafts[0];
            return (
                '<section class="ein-calm-banner ein-calm-banner--action" role="status">' +
                '<div class="ein-calm-banner__text">' +
                '<p class="ein-calm-banner__eyebrow">Nächster Schritt</p>' +
                '<p><strong>' +
                drafts.length +
                '</strong> Entwurf' +
                (drafts.length === 1 ? '' : 'e') +
                ' – aktivieren, damit Eltern-Links entstehen.</p>' +
                '</div>' +
                '<button type="button" class="btn btn-success" data-ein-activate="' +
                esc(first.id) +
                '">„' +
                esc(first.titel || 'Entwurf') +
                '“ aktivieren</button>' +
                '</section>'
            );
        }
        const active = (heads || []).filter((h) => h.status === EINSAMMLER_STATUS.AKTIV);
        let openActive = 0;
        for (const h of active) {
            openActive += progressStats(h, antworten || []).offen;
        }
        if (active.length && openActive > 0) {
            const first = active[0];
            return (
                '<section class="ein-calm-banner ein-calm-banner--action" role="status">' +
                '<div class="ein-calm-banner__text">' +
                '<p class="ein-calm-banner__eyebrow">Offene Rückläufe</p>' +
                '<p><strong>' +
                openActive +
                '</strong> Antwort' +
                (openActive === 1 ? '' : 'en') +
                ' offen – Links teilen oder Status manuell setzen.</p>' +
                '</div>' +
                '<button type="button" class="btn" data-ein-links="' +
                esc(first.id) +
                '">Eltern-Links öffnen</button>' +
                '</section>'
            );
        }
        if (!(heads || []).length && account().loggedIn) {
            return (
                '<section class="ein-calm-banner ein-calm-banner--action" role="status">' +
                '<div class="ein-calm-banner__text">' +
                '<p class="ein-calm-banner__eyebrow">Loslegen</p>' +
                '<p>Neuen Einsammler anlegen, Klassen wählen, aktivieren und Links teilen.</p>' +
                '</div>' +
                '<button type="button" class="btn btn-success" data-ein-view="neu"><i class="bi bi-plus-lg" aria-hidden="true"></i> Neu anlegen</button>' +
                '</section>'
            );
        }
        return '';
    }

    function stammdatenSourceHint() {
        const meta = getEinsammlerStammdatenMeta();
        const src =
            meta.source === 'sharepoint'
                ? 'SharePoint live'
                : meta.source === 'local-fallback'
                  ? 'lokal (Fallback)'
                  : 'lokal';
        let hint =
            'Stammdaten: ' + meta.classCount + ' Klassen · ' + meta.studentCount + ' Schüler · ' + src;
        if (meta.error) hint += ' · ' + meta.error;
        return hint;
    }

    function renderMessage() {
        if (!ui.message) return '';
        const cls =
            ui.message.type === 'ok'
                ? 'ein-msg--ok'
                : ui.message.type === 'warn'
                  ? 'ein-msg--warn'
                  : 'ein-msg--err';
        return '<p class="ein-msg ' + cls + '" role="status">' + esc(ui.message.text) + '</p>';
    }

    function renderView(s) {
        switch (ui.view) {
            case 'liste':
                return renderListe(s);
            case 'neu':
                return renderNeu();
            case 'detail':
                return renderDetail(s);
            case 'links':
                return renderLinks(s);
            case 'einstellungen':
                return renderEinstellungen(s);
            default:
                return renderUebersicht(s);
        }
    }

    function publicHrefFor(head, ant) {
        const settings = loadEinsammlerSettings();
        const payload = buildPublicPayload({
            einsammler: head,
            antwort: ant,
            schoolName: settings.schoolName,
            expDays: settings.expDays
        });
        const token = encodePublicToken(payload);
        const abs = new URL('einsammler-antwort.html', window.location.href).href;
        return buildPublicUrl(abs, token);
    }

    function renderUebersicht(s) {
        const heads = visibleHeads(s);
        const antworten = filterAntwortenForHeads(s.antworten || [], heads);
        const aktiv = heads.filter((h) => h.status === EINSAMMLER_STATUS.AKTIV).length;
        let offen = 0;
        let total = 0;
        for (const h of heads) {
            const st = progressStats(h, antworten);
            offen += st.offen;
            total += st.total;
        }
        const rows = heads
            .slice(0, 8)
            .map((h) => {
                const st = progressStats(h, antworten);
                return (
                    '<tr>' +
                    '<td><button type="button" class="btn" data-ein-open="' +
                    esc(h.id) +
                    '" style="padding:4px 8px;font-size:0.85rem;">' +
                    esc(h.titel) +
                    '</button></td>' +
                    '<td><span class="ein-pill ' +
                    pillClassForStatus(h.status) +
                    '">' +
                    esc(EINSAMMLER_STATUS_LABEL[h.status] || h.status) +
                    '</span></td>' +
                    '<td>' +
                    esc(MODUS_LABEL[h.modus] || h.modus) +
                    '</td>' +
                    '<td>' +
                    st.erledigt +
                    '/' +
                    st.total +
                    ' (' +
                    st.pct +
                    '%)</td>' +
                    '<td>' +
                    esc(formatDeDate(h.faelligAm)) +
                    '</td>' +
                    '</tr>'
                );
            })
            .join('');

        const greetNum = offen;
        const greetText = greetNum === 1 ? 'offene Antwort' : 'offene Antworten';
        return (
            actionBanner(heads, antworten) +
            '<div class="ein-head"><div>' +
            '<h1 class="fb-greet"><span class="fb-num">' +
            greetNum +
            '</span> ' +
            greetText +
            '</h1>' +
            '<p class="fb-sub">Ihre Einsammler' +
            (admin() && ui.showAll ? ' (Admin: alle)' : '') +
            ' – ' +
            heads.length +
            (heads.length === 1 ? ' Vorgang' : ' Vorgänge') +
            (aktiv ? ', ' + aktiv + ' aktiv' : '') +
            '.</p></div>' +
            '<div class="ein-actions">' +
            '<button type="button" class="btn btn-success fb-btn" data-ein-view="neu"><i class="bi bi-plus-lg"></i> Neu</button>' +
            '</div></div>' +
            '<div class="ein-cards ein-cards--quiet" aria-label="Kennzahlen">' +
            card('Meine Einsammler', String(heads.length)) +
            card('Aktiv', String(aktiv)) +
            card('Offene Antworten', String(offen)) +
            card('Empfänger gesamt', String(total)) +
            '</div>' +
            (heads.length
                ? '<div class="ein-table-wrap"><table class="ein-table"><thead><tr>' +
                  '<th>Titel</th><th>Status</th><th>Modus</th><th>Fortschritt</th><th>Fällig</th>' +
                  '</tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : account().loggedIn
                  ? quietEmpty(
                        'Noch keine Einsammler',
                        'Legen Sie den ersten Vorgang an – Klassen aus Stammdaten, dann aktivieren und Links teilen.',
                        'neu',
                        'Neu anlegen'
                    )
                  : quietEmpty(
                        'Noch nichts sichtbar',
                        'Nach der Anmeldung erscheinen hier Ihre Einsammler.',
                        '',
                        ''
                    ))
        );
    }

    function card(label, value) {
        return (
            '<div class="ein-card"><p class="ein-card__label">' +
            esc(label) +
            '</p><p class="ein-card__value">' +
            esc(value) +
            '</p></div>'
        );
    }

    function renderListe(s) {
        const heads = visibleHeads(s);
        const antworten = filterAntwortenForHeads(s.antworten || [], heads);
        const rows = heads
            .map((h) => {
                const st = progressStats(h, antworten);
                return (
                    '<tr>' +
                    '<td><strong>' +
                    esc(h.titel) +
                    '</strong><div class="ein-muted">' +
                    esc(TYP_LABEL[h.typ] || h.typ) +
                    ' · ' +
                    esc(h.klasseCodes || '–') +
                    (admin() && ui.showAll && h.erstellerEmail
                        ? ' · ' + esc(h.erstellerName || h.erstellerEmail)
                        : '') +
                    '</div>' +
                    '<div class="ein-progress" title="' +
                    st.pct +
                    '%"><span style="width:' +
                    st.pct +
                    '%"></span></div></td>' +
                    '<td><span class="ein-pill ' +
                    pillClassForStatus(h.status) +
                    '">' +
                    esc(EINSAMMLER_STATUS_LABEL[h.status] || h.status) +
                    '</span></td>' +
                    '<td>' +
                    st.offen +
                    ' offen</td>' +
                    '<td class="ein-inline-actions">' +
                    '<button type="button" data-ein-open="' +
                    esc(h.id) +
                    '">Antworten</button>' +
                    '<button type="button" data-ein-links="' +
                    esc(h.id) +
                    '">Links</button>' +
                    (h.status === EINSAMMLER_STATUS.ENTWURF
                        ? '<button type="button" data-ein-activate="' + esc(h.id) + '">Aktivieren</button>'
                        : '') +
                    (h.status === EINSAMMLER_STATUS.AKTIV
                        ? '<button type="button" data-ein-close="' + esc(h.id) + '">Schließen</button>'
                        : '') +
                    '<button type="button" data-ein-delete="' +
                    esc(h.id) +
                    '">Löschen</button>' +
                    '</td></tr>'
                );
            })
            .join('');

        return (
            actionBanner(heads, antworten) +
            '<div class="ein-head"><div><h2>Einsammler</h2><p>' +
            (admin() && ui.showAll
                ? 'Admin-Sicht: alle Vorgänge der Schule (dieser Browser).'
                : 'Nur Ihre angelegten Vorgänge.') +
            '</p></div>' +
            '<div class="ein-actions"><button type="button" class="btn btn-success" data-ein-view="neu">Neu</button></div></div>' +
            (rows
                ? '<div class="ein-table-wrap"><table class="ein-table"><thead><tr><th>Vorgang</th><th>Status</th><th>Offen</th><th></th></tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty(
                      'Keine Einträge',
                      'In Ihrer Sicht sind noch keine Vorgänge. Legen Sie einen neuen Einsammler an.',
                      'neu',
                      'Neu anlegen'
                  ))
        );
    }

    function renderNeu() {
        const f = ui.form;
        const sd = loadEinsammlerStammdaten();
        const selected = new Set(parseKlasseCodesField(f.klasseCodes).map((c) => c.toLowerCase()));
        const classChips = sd.classes.length
            ? '<div class="ein-class-pick" id="einClassPick">' +
              sd.classes
                  .map((c) => {
                      const on = selected.has(c.code.toLowerCase());
                      return (
                          '<label class="ein-class-chip' +
                          (on ? ' is-on' : '') +
                          '"><input type="checkbox" data-ein-klasse="' +
                          esc(c.code) +
                          '"' +
                          (on ? ' checked' : '') +
                          '> ' +
                          esc(c.code) +
                          (c.name && c.name !== c.code ? ' · ' + esc(c.name) : '') +
                          '</label>'
                      );
                  })
                  .join('') +
              '</div>' +
              '<div class="ein-form__actions" style="margin:0 0 10px;">' +
              '<button type="button" class="btn" id="einFillFromKlassen">Empfänger aus gewählten Klassen übernehmen</button>' +
              '</div>' +
              '<p class="ein-hint">' +
              esc(stammdatenSourceHint()) +
              ' · ' +
              sd.classes.length +
              ' sichtbar · ' +
              sd.students.length +
              ' Schüler:innen. Eltern-Mail kommt aus hinterlegten Erziehungsberechtigten (sonst leer).</p>'
            : '<p class="ein-hint">Keine Klassen in den Stammdaten gefunden – unter <a href="../tenant.html">Stammdaten</a> pflegen oder Codes manuell eintragen.</p>';

        return (
            '<div class="ein-head"><div><h2>' +
            (f.id ? 'Einsammler bearbeiten' : 'Neuer Einsammler') +
            '</h2><p>Klassen wählen → Empfänger übernehmen → speichern → später aktivieren und Eltern-Links teilen.</p></div></div>' +
            '<form class="ein-form" id="einForm">' +
            '<label>Titel<input name="titel" required value="' +
            esc(f.titel) +
            '"></label>' +
            '<div class="ein-form__row">' +
            '<label>Modus<select name="modus" id="einModus">' +
            option(MODUS.EINWILLIGUNG, MODUS_LABEL.einwilligung, f.modus) +
            option(MODUS.LESEBESTAETIGUNG, MODUS_LABEL.lesebestaetigung, f.modus) +
            option(MODUS.ABGABE, MODUS_LABEL.abgabe, f.modus) +
            '</select></label>' +
            '<label>Typ<select name="typ">' +
            Object.keys(TYP_LABEL)
                .map((k) => option(k, TYP_LABEL[k], f.typ))
                .join('') +
            '</select></label></div>' +
            '<label class="ein-hint" style="display:flex;gap:8px;align-items:center;font-weight:500;">' +
            '<input type="checkbox" name="uploadErlaubt" id="einUploadErlaubt"' +
            (f.uploadErlaubt || f.modus === MODUS.ABGABE ? ' checked' : '') +
            '> Eltern dürfen Datei/Bild hochladen</label>' +
            '<label class="ein-hint" style="display:flex;gap:8px;align-items:center;font-weight:500;">' +
            '<input type="checkbox" name="uploadPflicht" id="einUploadPflicht"' +
            (f.uploadPflicht || f.modus === MODUS.ABGABE ? ' checked' : '') +
            '> Upload ist Pflicht</label>' +
            '<p class="ein-hint">Modus „Datei-/Bildabgabe“ = nur Abgabe (kein Zustimmen). Anhänge landen lokal (IndexedDB) bzw. optional in der API; SharePoint-URL-Spalte ist vorbereitet.</p>' +
            '<div><span class="ein-field-label">Klassen aus Stammdaten</span>' +
            classChips +
            '</div>' +
            '<label>Klassen-Codes (auto aus Auswahl)<input name="klasseCodes" id="einKlasseCodes" placeholder="3AHW;3BHW" value="' +
            esc(f.klasseCodes) +
            '"></label>' +
            '<label>Fällig am<input type="date" name="faelligAm" value="' +
            esc(f.faelligAm) +
            '"></label>' +
            '<label>Text / Inhalt<textarea name="textInhalt" placeholder="Text für Eltern (Einwilligung / Brief)">' +
            esc(f.textInhalt) +
            '</textarea></label>' +
            '<label>Empfänger (eine Zeile: Name;Klasse;Eltern-E-Mail)<textarea name="empfaengerText" id="einEmpfaengerText" placeholder="Anna Berger;3AHW;eltern@example.com">' +
            esc(f.empfaengerText) +
            '</textarea></label>' +
            '<div class="ein-form__actions">' +
            '<button type="submit" class="btn btn-success">Speichern</button>' +
            '<button type="button" class="btn" data-ein-view="liste">Abbrechen</button>' +
            '</div></form>'
        );
    }

    function option(value, label, current) {
        return (
            '<option value="' +
            esc(value) +
            '"' +
            (value === current ? ' selected' : '') +
            '>' +
            esc(label) +
            '</option>'
        );
    }

    function renderDetail(s) {
        const head = selectedHead(s);
        if (!head) return '<div class="ein-empty">Kein Einsammler gewählt.</div>';
        const st = progressStats(head, s.antworten);
        const allowed = allowedAntwortStatusesForModus(head.modus);
        const rows = (s.antworten || [])
            .filter((a) => a.einsammlerId === head.id)
            .map((a) => {
                const eff = resolveAntwortStatus(head, a);
                const opts = allowed
                    .map((x) => option(x, ANTWORT_STATUS_LABEL[x] || x, eff))
                    .join('');
                const linkBtn =
                    head.status === EINSAMMLER_STATUS.AKTIV
                        ? '<button type="button" data-ein-copy-one="' +
                          esc(a.id) +
                          '">Link</button>' +
                          (a.elternEmail
                              ? '<a class="btn" style="padding:4px 8px;font-size:0.78rem;text-decoration:none;" data-ein-mailto="' +
                                esc(a.id) +
                                '" href="#">Mail</a>'
                              : '')
                        : '';
                return (
                    '<tr>' +
                    '<td>' +
                    esc(a.schuelerName) +
                    '<div class="ein-muted">' +
                    esc(a.klasseCode || '–') +
                    (a.elternEmail ? ' · ' + esc(a.elternEmail) : '') +
                    (a.anhangName
                        ? ' · 📎 ' +
                          esc(a.anhangName) +
                          ' <button type="button" data-ein-anhang="' +
                          esc(a.id) +
                          '" style="margin-left:4px;">Öffnen</button>'
                        : '') +
                    '</div></td>' +
                    '<td><span class="ein-pill ' +
                    pillClassForStatus(eff) +
                    '">' +
                    esc(ANTWORT_STATUS_LABEL[eff] || eff) +
                    '</span></td>' +
                    '<td class="ein-inline-actions">' +
                    linkBtn +
                    '<select data-ein-status-for="' +
                    esc(a.id) +
                    '" title="Nur Fallback, wenn Eltern nicht per Link antworten">' +
                    opts +
                    '</select>' +
                    '<button type="button" data-ein-apply-status="' +
                    esc(a.id) +
                    '">Manuell</button>' +
                    '<button type="button" data-ein-remove-ant="' +
                    esc(a.id) +
                    '">Entfernen</button>' +
                    '</td></tr>'
                );
            })
            .join('');

        const activateHint =
            head.status !== EINSAMMLER_STATUS.AKTIV
                ? '<p class="ein-msg ein-msg--warn">Noch Entwurf – bitte <strong>aktivieren</strong>, dann erscheinen Eltern-Links. Manuelles Setzen ist nur Fallback (Telefon o. Ä.).</p>' +
                  '<div class="ein-actions" style="margin-bottom:12px;"><button type="button" class="btn btn-success" data-ein-activate="' +
                  esc(head.id) +
                  '">Jetzt aktivieren &amp; Links erzeugen</button></div>'
                : '<p class="ein-hint">Eltern antworten über ihren persönlichen Link (Tab „Eltern-Links“). Status „Manuell“ nur wenn nötig.</p>';

        return (
            '<div class="ein-head"><div><h2>' +
            esc(head.titel) +
            '</h2><p>' +
            esc(MODUS_LABEL[head.modus] || head.modus) +
            ' · ' +
            esc(TYP_LABEL[head.typ] || head.typ) +
            ' · Fällig ' +
            esc(formatDeDate(head.faelligAm)) +
            '</p></div>' +
            '<div class="ein-actions">' +
            '<button type="button" class="btn" data-ein-view="links">Eltern-Links</button>' +
            '<button type="button" class="btn" data-ein-view="liste">Zurück</button>' +
            '</div></div>' +
            activateHint +
            '<div class="ein-cards">' +
            card('Empfänger', String(st.total)) +
            card('Offen', String(st.offen)) +
            card('Erledigt', String(st.erledigt)) +
            card('Fortschritt', st.pct + '%') +
            '</div>' +
            '<div class="ein-progress" style="margin-bottom:16px;"><span style="width:' +
            st.pct +
            '%"></span></div>' +
            '<form class="ein-form" id="einAddEmp" style="margin-bottom:16px;">' +
            '<label>Weitere Empfänger<textarea name="empfaengerText" placeholder="Name;Klasse;E-Mail"></textarea></label>' +
            '<div class="ein-form__actions">' +
            '<button type="submit" class="btn">Hinzufügen</button>' +
            '<button type="button" class="btn" id="einFillDetailKlassen">Aus Stammdaten-Klassen nachladen</button>' +
            '</div></form>' +
            (rows
                ? '<div class="ein-table-wrap"><table class="ein-table"><thead><tr><th>Schüler:in</th><th>Status</th><th>Aktion</th></tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : '<div class="ein-empty">Noch keine Empfänger – Klassen beim Anlegen wählen oder hier einfügen.</div>')
        );
    }

    function renderLinks(s) {
        const head = selectedHead(s);
        if (!head) return '<div class="ein-empty">Kein Einsammler gewählt.</div>';
        const ants = (s.antworten || []).filter((a) => a.einsammlerId === head.id);
        if (head.status !== EINSAMMLER_STATUS.AKTIV) {
            return (
                '<div class="ein-head"><div><h2>Eltern-Links</h2><p>' +
                esc(head.titel) +
                '</p></div></div>' +
                '<div class="ein-empty">Bitte zuerst <strong>aktivieren</strong>, damit Eltern antworten können.</div>' +
                '<div class="ein-actions" style="margin-top:12px;">' +
                '<button type="button" class="btn btn-success" data-ein-activate="' +
                esc(head.id) +
                '">Jetzt aktivieren</button></div>'
            );
        }
        const rows = ants
            .map((a) => {
                const href = publicHrefFor(head, a);
                const mailHref = a.elternEmail
                    ? 'mailto:' +
                      encodeURIComponent(a.elternEmail).replace(/%40/g, '@') +
                      '?subject=' +
                      encodeURIComponent(head.titel || 'Schul-Rückmeldung') +
                      '&body=' +
                      encodeURIComponent(
                          'Guten Tag,\n\nbitte öffnen Sie diesen persönlichen Link zur Rückmeldung („' +
                              (head.titel || 'Einsammler') +
                              '“):\n\n' +
                              href +
                              '\n\nVielen Dank.\n'
                      )
                    : '';
                return (
                    '<div class="ein-link-row">' +
                    '<strong>' +
                    esc(a.schuelerName) +
                    '</strong>' +
                    '<span class="ein-muted">' +
                    esc(a.elternEmail || 'keine Eltern-Mail') +
                    '</span>' +
                    '<input type="text" readonly value="' +
                    esc(href) +
                    '" data-ein-link-input>' +
                    '<div class="ein-inline-actions">' +
                    '<button type="button" data-ein-copy-link>Kopieren</button>' +
                    (mailHref
                        ? '<a class="btn ein-btn-compact" href="' +
                          esc(mailHref) +
                          '">Mail vorbereiten</a>'
                        : '') +
                    '<a class="btn ein-btn-compact" href="' +
                    esc(href) +
                    '" target="_blank" rel="noopener">Öffnen</a>' +
                    '</div></div>'
                );
            })
            .join('');
        return (
            '<section class="ein-calm-banner ein-calm-banner--action" role="status">' +
            '<div class="ein-calm-banner__text">' +
            '<p class="ein-calm-banner__eyebrow">Teilen</p>' +
            '<p>Links kopieren oder „Mail vorbereiten“ – noch kein automatischer Versand.</p>' +
            '</div></section>' +
            '<div class="ein-head"><div><h2>Eltern-Links</h2><p>' +
            esc(head.titel) +
            ' – je Empfänger ein Unique-Link.</p></div>' +
            '<div class="ein-actions">' +
            '<button type="button" class="btn" data-ein-publish-remote>An API veröffentlichen</button>' +
            '<button type="button" class="btn" data-ein-view="detail">Zu Antworten</button>' +
            '</div></div>' +
            (rows
                ? '<div class="ein-link-list">' + rows + '</div>'
                : quietEmpty(
                      'Keine Empfänger',
                      'Legen Sie zuerst Empfänger unter Antworten an.',
                      'detail',
                      'Zu Antworten'
                  ))
        );
    }

    function renderEinstellungen(s) {
        const settings = loadEinsammlerSettings();
        const inbox = listInbox();
        const acc = account();
        const mine = visibleHeads(s).length;
        const sync = getEinsammlerSyncMeta();
        const site = resolveEinsammlerSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
        const spoHeads = (s.einsammler || []).filter((e) => e.spoItemId).length;
        const spoAnt = (s.antworten || []).filter((a) => a.spoItemId).length;
        const syncPanel =
            '<div class="ein-form ein-system-panel" style="margin-top:14px;">' +
            '<h3 style="margin:0 0 8px;font-size:1rem;">SharePoint (Live)</h3>' +
            '<p class="ein-hint"><code>' +
            esc(site) +
            '</code></p>' +
            '<p class="ein-hint">Sync: <strong>' +
            esc(sync.source) +
            '</strong>' +
            (sync.listTitle ? ' · „' + esc(sync.listTitle) + '“' : '') +
            (sync.listAntwortenTitle ? ' / „' + esc(sync.listAntwortenTitle) + '“' : '') +
            (sync.lastPullAt ? ' · zuletzt ' + esc(formatDeDate(sync.lastPullAt)) : '') +
            (sync.connected ? ' · verbunden' : ' · nicht verbunden') +
            (sync.error ? ' · ' + esc(sync.error) : '') +
            '</p>' +
            '<p class="ein-hint">Lokal: ' +
            (s.einsammler || []).length +
            ' Einsammler · ' +
            (s.antworten || []).length +
            ' Antworten' +
            (s.localDemoOnly ? ' · Demo local-first' : '') +
            ' · spoItemId: ' +
            spoHeads +
            '/' +
            spoAnt +
            '</p>' +
            '<p class="ein-hint">Eltern-Antworten ohne Login: Posteingang/Code oder Remote-API – danach Push durch Absender/Admin.</p>' +
            (admin()
                ? '<div class="ein-form__actions">' +
                  '<button type="button" class="btn btn-success" data-ein-pull>Von SharePoint laden</button>' +
                  '<a class="btn" href="sharepoint-liste-einsammler.html">Listen einrichten</a>' +
                  '</div>'
                : '') +
            '</div>';
        if (!admin()) {
            return (
                '<div class="ein-head"><div><h2>Mein Bereich</h2>' +
                '<p>Persönliche Sicht – Admin-Einstellungen sind der Schulleitung/IT vorbehalten.</p></div></div>' +
                '<div class="ein-form">' +
                '<p class="ein-settings-summary">Konto: <strong>' +
                esc(acc.name || acc.email || 'nicht angemeldet') +
                '</strong><br>Eigene Einsammler: <strong>' +
                mine +
                '</strong> · Posteingang: <strong>' +
                inbox.length +
                '</strong></p>' +
                '<p class="ein-hint">' +
                esc(stammdatenSourceHint()) +
                '</p>' +
                '<div class="ein-form__actions" style="margin-top:12px;">' +
                '<button type="button" class="btn" data-ein-merge-inbox>Posteingang übernehmen</button>' +
                (canUseDemoActions()
                    ? '<button type="button" class="btn" data-ein-demo>Demo als meine Daten laden</button>' +
                      '<button type="button" class="btn" data-ein-clear-demo>Demo entfernen</button>'
                    : '') +
                '</div></div>' +
                syncPanel
            );
        }
        const sources = getEinsammlerAdminAccessCache();
        return (
            '<form class="ein-form" id="einSettingsForm">' +
            '<input type="hidden" name="apiBaseUrl" value="">' +
            '<input type="hidden" name="adminKey" value="">' +
            renderSystemAdminPage({
                prefix: 'ein',
                activeTab: ui.systemTab,
                dataAttr: 'data-ein-sys-tab',
                ariaLabel: 'Systemadministration Einsammler',
                panels: [
                    {
                        id: 'site',
                        step: '1',
                        title: 'Datenquelle',
                        leadHtml: 'Antworten landen in der Schul-SharePoint-Site.',
                        body: syncPanel
                    },
                    {
                        id: 'perms',
                        step: '2',
                        title: 'Berechtigungen',
                        leadHtml:
                            'Admin „Alle sehen“ kommt aus Schulleitung/Verwaltung (Entra) bzw. Schul-IT. Legacy-Mails nur Fallback.',
                        body:
                            '<label>Zusätzliche Admin-E-Mails (Legacy, optional)<textarea name="adminEmails" rows="2" placeholder="direktion@schule.at; it@schule.at">' +
                            esc(settings.adminEmails) +
                            '</textarea></label>' +
                            (sources.sources.length
                                ? '<p class="ein-hint">Ihre Admin-Quellen: ' +
                                  esc(sources.sources.join(', ')) +
                                  (sources.demo ? ' (IT-Vorschau)' : '') +
                                  '</p>'
                                : '') +
                            '<div class="ein-form__actions"><button type="submit" class="btn btn-success">Einstellungen speichern</button></div>'
                    },
                    {
                        id: 'kalender',
                        step: '3',
                        title: 'Kalender',
                        leadHtml: 'Kein Terminkalender. Die Link-Frist steuert, wie lange Eltern antworten können.',
                        body:
                            '<label>Link-Gültigkeit (Tage)<input type="number" min="1" name="expDays" value="' +
                            esc(String(settings.expDays)) +
                            '"></label>' +
                            '<p class="ein-hint">Elternlinks sind HMAC-signiert und stehen im URL-Fragment (<code>#p=</code>).</p>' +
                            '<div class="ein-form__actions"><button type="submit" class="btn btn-success">Einstellungen speichern</button></div>'
                    },
                    {
                        id: 'wartung',
                        step: '4',
                        title: 'Wartung',
                        leadHtml: 'Schulname auf der Elternseite, Stammdaten und Antwortcodes.',
                        body:
                            '<label>Schulname (auf Elternseite)<input name="schoolName" value="' +
                            esc(settings.schoolName) +
                            '"></label>' +
                            '<p class="ein-settings-summary">Gesamt-Store: <strong>' +
                            (s.einsammler || []).length +
                            '</strong> · in Ihrer Filter-Sicht: <strong>' +
                            mine +
                            '</strong> · Posteingang: <strong>' +
                            inbox.length +
                            '</strong>' +
                            (einsammlerRemoteConfigured() ? ' · Remote konfiguriert' : '') +
                            '</p>' +
                            '<p class="ein-hint">' +
                            esc(stammdatenSourceHint()) +
                            '</p>' +
                            '<label>Antwortcode einfügen<textarea id="einImportCode" rows="3" placeholder="Code von der Elternseite"></textarea></label>' +
                            '<div class="ein-form__actions">' +
                            '<button type="submit" class="btn btn-success">Einstellungen speichern</button>' +
                            '<button type="button" class="btn" data-ein-import-code>Code übernehmen</button>' +
                            '<button type="button" class="btn" data-ein-merge-inbox>Posteingang übernehmen</button>' +
                            '<button type="button" class="btn" data-ein-demo>Demo laden</button>' +
                            '</div>'
                    },
                    {
                        id: 'danger',
                        title: 'Danger Zone',
                        danger: true,
                        leadHtml: 'Demo entfernen oder den lokalen Store leeren. SharePoint löscht das nicht automatisch.',
                        body:
                            '<div class="ms-sys-actions">' +
                            '<button type="button" class="btn ms-sys-btn-danger" data-ein-clear-demo>Demo entfernen</button>' +
                            '<button type="button" class="btn ms-sys-btn-danger" data-ein-clear>Alles löschen</button>' +
                            '</div>'
                    }
                ]
            }) +
            '</form>'
        );
    }

    function bind(s) {
        bindSystemTabClicks(root, {
            dataAttr: 'data-ein-sys-tab',
            onSelect: (tab) => {
                ui.systemTab = tab;
                render();
            }
        });
        root.querySelector('#einShowAll')?.addEventListener('change', (ev) => {
            ui.showAll = !!ev.target.checked && admin();
            setMessage(
                'ok',
                ui.showAll ? 'Admin-Sicht: alle Einsammler.' : 'Nur eigene Einsammler.'
            );
            render();
        });

        root.querySelectorAll('[data-ein-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const v = btn.getAttribute('data-ein-view');
                if (v === 'neu') {
                    if (!account().loggedIn) {
                        setMessage('warn', 'Bitte zuerst anmelden, damit der Einsammler Ihnen zugeordnet wird.');
                        render();
                        return;
                    }
                    ui.form = emptyForm();
                }
                ui.view = v || 'uebersicht';
                setMessage('', '');
                render();
            });
        });

        root.querySelectorAll('[data-ein-open]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-ein-open') || '';
                const head = (state().einsammler || []).find((e) => e.id === id);
                if (!canAccessHead(head)) {
                    setMessage('err', 'Kein Zugriff auf diesen Einsammler.');
                    render();
                    return;
                }
                ui.selectedId = id;
                ui.view = 'detail';
                setMessage('', '');
                render();
            });
        });

        root.querySelectorAll('[data-ein-links]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.selectedId = btn.getAttribute('data-ein-links') || '';
                ui.view = 'links';
                setMessage('', '');
                render();
            });
        });

        function syncKlasseCodesFromChips() {
            const codes = [...root.querySelectorAll('[data-ein-klasse]:checked')].map((el) =>
                el.getAttribute('data-ein-klasse')
            );
            const input = root.querySelector('#einKlasseCodes');
            if (input) input.value = codes.join(';');
            root.querySelectorAll('.ein-class-chip').forEach((lab) => {
                const cb = lab.querySelector('input');
                lab.classList.toggle('is-on', !!(cb && cb.checked));
            });
            return codes;
        }

        root.querySelectorAll('[data-ein-klasse]').forEach((cb) => {
            cb.addEventListener('change', () => syncKlasseCodesFromChips());
        });

        root.querySelector('#einFillFromKlassen')?.addEventListener('click', () => {
            const codes = syncKlasseCodesFromChips();
            if (!codes.length) {
                setMessage('warn', 'Bitte zuerst Klassen anhaken.');
                render();
                return;
            }
            const lines = buildEmpfaengerLinesForKlassen(codes);
            const ta = root.querySelector('#einEmpfaengerText');
            if (ta) {
                const prev = String(ta.value || '').trim();
                ta.value = prev ? prev + '\n' + lines.join('\n') : lines.join('\n');
            }
            setMessage(
                'ok',
                lines.length
                    ? lines.length + ' Empfängerzeilen aus Stammdaten übernommen.'
                    : 'Keine Schüler:innen für diese Klassen in den Stammdaten.'
            );
            // keep form values – re-read into ui.form lightly
            const form = root.querySelector('#einForm');
            if (form) {
                const fd = new FormData(form);
                ui.form = {
                    ...ui.form,
                    titel: String(fd.get('titel') || ''),
                    modus: String(fd.get('modus') || ui.form.modus),
                    typ: String(fd.get('typ') || ui.form.typ),
                    textInhalt: String(fd.get('textInhalt') || ''),
                    klasseCodes: String(fd.get('klasseCodes') || codes.join(';')),
                    faelligAm: String(fd.get('faelligAm') || ''),
                    empfaengerText: ta ? String(ta.value || '') : ''
                };
            }
            render();
        });

        root.querySelector('#einFillDetailKlassen')?.addEventListener('click', () => {
            const head = selectedHead(state());
            if (!head) return;
            const codes = parseKlasseCodesField(head.klasseCodes);
            const lines = buildEmpfaengerLinesForKlassen(codes);
            if (!lines.length) {
                setMessage('warn', 'Keine Schüler für die Klassencodes dieses Einsammlers.');
                render();
                return;
            }
            const r = addEmpfaengerFromText(head.id, lines.join('\n'));
            if (!r.ok) {
                setMessage('err', (r.errors && r.errors[0]) || 'Fehler');
                render();
                return;
            }
            void afterLocalAntworten(r.addedItems || [], aAddedMsg(r.added));
        });

        root.querySelectorAll('[data-ein-copy-link]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const row = btn.closest('.ein-link-row');
                const input = row && row.querySelector('[data-ein-link-input]');
                const val = input ? String(input.value || '') : '';
                try {
                    await navigator.clipboard.writeText(val);
                    setMessage('ok', 'Link kopiert.');
                } catch {
                    if (input) {
                        input.focus();
                        input.select();
                    }
                    setMessage('warn', 'Bitte Link manuell kopieren.');
                }
                render();
            });
        });

        root.querySelectorAll('[data-ein-copy-one]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const head = selectedHead(state());
                const ant = (state().antworten || []).find(
                    (a) => a.id === btn.getAttribute('data-ein-copy-one')
                );
                if (!head || !ant) return;
                const href = publicHrefFor(head, ant);
                try {
                    await navigator.clipboard.writeText(href);
                    setMessage('ok', 'Eltern-Link kopiert.');
                } catch {
                    setMessage('warn', href);
                }
                render();
            });
        });

        root.querySelectorAll('[data-ein-mailto]').forEach((btn) => {
            btn.addEventListener('click', (ev) => {
                ev.preventDefault();
                const head = selectedHead(state());
                const ant = (state().antworten || []).find(
                    (a) => a.id === btn.getAttribute('data-ein-mailto')
                );
                if (!head || !ant || !ant.elternEmail) return;
                const href = publicHrefFor(head, ant);
                const mail =
                    'mailto:' +
                    ant.elternEmail +
                    '?subject=' +
                    encodeURIComponent(head.titel || 'Schul-Rückmeldung') +
                    '&body=' +
                    encodeURIComponent(
                        'Guten Tag,\n\nbitte öffnen Sie diesen persönlichen Link („' +
                            (head.titel || '') +
                            '“):\n\n' +
                            href +
                            '\n\nVielen Dank.\n'
                    );
                window.location.href = mail;
            });
        });

        root.querySelectorAll('[data-ein-publish-remote]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                if (!einsammlerRemoteConfigured()) {
                    setMessage('warn', 'Keine API-URL in den Einstellungen.');
                    render();
                    return;
                }
                const head = selectedHead(state());
                if (!head) return;
                const settings = loadEinsammlerSettings();
                const packs = (state().antworten || [])
                    .filter((a) => a.einsammlerId === head.id)
                    .map((a) =>
                        buildPublicPayload({
                            einsammler: head,
                            antwort: a,
                            schoolName: settings.schoolName,
                            expDays: settings.expDays
                        })
                    );
                btn.disabled = true;
                try {
                    await remotePublishPacks(packs);
                    setMessage('ok', packs.length + ' Paket(e) an API veröffentlicht.');
                } catch (e) {
                    setMessage('err', e && e.message ? e.message : String(e));
                }
                render();
            });
        });

        root.querySelectorAll('[data-ein-activate]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const r = setEinsammlerStatus(btn.getAttribute('data-ein-activate'), EINSAMMLER_STATUS.AKTIV);
                if (!r.ok) {
                    setMessage('err', (r.errors && r.errors[0]) || 'Fehler');
                    render();
                    return;
                }
                ui.view = 'links';
                void afterLocalWrite(r.item, 'kopf', 'Aktiviert – jetzt Eltern-Links teilen.');
            });
        });

        root.querySelectorAll('[data-ein-close]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const r = setEinsammlerStatus(
                    btn.getAttribute('data-ein-close'),
                    EINSAMMLER_STATUS.GESCHLOSSEN
                );
                if (!r.ok) {
                    setMessage('err', (r.errors && r.errors[0]) || 'Fehler');
                    render();
                    return;
                }
                void afterLocalWrite(r.item, 'kopf', 'Geschlossen.');
            });
        });

        root.querySelectorAll('[data-ein-delete]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!window.confirm('Einsammler inkl. Antworten löschen?')) return;
                const id = btn.getAttribute('data-ein-delete') || '';
                const del = removeEinsammler(id);
                if (ui.selectedId === id) ui.selectedId = '';
                void pushEinsammlerDelete(del.item || {}, del.antworten || []).then((r) => {
                    setMessage(
                        r.ok || r.skipped ? 'ok' : 'warn',
                        r.ok || r.skipped
                            ? 'Gelöscht.'
                            : 'Lokal gelöscht. ' + (r.error || 'SharePoint-Löschen fehlgeschlagen.')
                    );
                    render();
                });
            });
        });

        root.querySelectorAll('[data-ein-apply-status]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-ein-apply-status');
                const sel = root.querySelector('select[data-ein-status-for="' + id + '"]');
                const status = sel ? sel.value : '';
                const r = setAntwortStatus(id, status);
                if (!r.ok) {
                    setMessage('err', (r.errors && r.errors[0]) || 'Fehler');
                    render();
                    return;
                }
                void afterLocalWrite(r.item, 'antwort', 'Status gesetzt.');
            });
        });

        root.querySelectorAll('[data-ein-remove-ant]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const del = removeAntwort(btn.getAttribute('data-ein-remove-ant'));
                void pushEinsammlerAntwortDelete(del.item || {}).then((r) => {
                    setMessage(
                        r.ok || r.skipped ? 'ok' : 'warn',
                        r.ok || r.skipped
                            ? 'Empfänger entfernt.'
                            : 'Lokal entfernt. ' + (r.error || 'SharePoint-Löschen fehlgeschlagen.')
                    );
                    render();
                });
            });
        });

        root.querySelector('#einModus')?.addEventListener('change', (ev) => {
            const m = ev.target.value;
            const up = root.querySelector('#einUploadErlaubt');
            const pf = root.querySelector('#einUploadPflicht');
            if (m === MODUS.ABGABE) {
                if (up) up.checked = true;
                if (pf) pf.checked = true;
            }
        });

        root.querySelectorAll('[data-ein-anhang]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const id = btn.getAttribute('data-ein-anhang') || '';
                const ant = (state().antworten || []).find((a) => a.id === id);
                if (!ant) return;
                if (ant.anhangUrl) {
                    window.open(ant.anhangUrl, '_blank', 'noopener');
                    return;
                }
                const key = ant.anhangKey || ant.id;
                const file = await loadAttachment(key);
                if (!file || !file.dataUrl) {
                    setMessage(
                        'warn',
                        'Anhang lokal nicht gefunden (anderer Browser/Gerät). Mit Remote-API oder SharePoint-URL verfügbar.'
                    );
                    render();
                    return;
                }
                const a = document.createElement('a');
                a.href = file.dataUrl;
                a.download = file.name || ant.anhangName || 'anhang';
                a.click();
                setMessage('ok', 'Anhang geöffnet / Download gestartet.');
                render();
            });
        });

        const form = root.querySelector('#einForm');
        if (form) {
            form.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(form);
                const id = ui.form.id || newId('ein');
                const acc = account();
                if (!acc.loggedIn) {
                    setMessage('err', 'Anmeldung nötig – sonst fehlt die Zuordnung zu Ihrer Person.');
                    render();
                    return;
                }
                const existing = (state().einsammler || []).find((e) => e.id === id);
                if (existing && !canAccessHead(existing)) {
                    setMessage('err', 'Kein Zugriff zum Bearbeiten.');
                    render();
                    return;
                }
                const modus = String(fd.get('modus') || MODUS.EINWILLIGUNG);
                const draft = {
                    id,
                    titel: String(fd.get('titel') || ''),
                    modus,
                    typ: String(fd.get('typ') || 'sonstiges'),
                    textInhalt: String(fd.get('textInhalt') || ''),
                    klasseCodes: String(fd.get('klasseCodes') || ''),
                    faelligAm: String(fd.get('faelligAm') || ''),
                    uploadErlaubt:
                        modus === MODUS.ABGABE || !!form.querySelector('[name="uploadErlaubt"]')?.checked,
                    uploadPflicht:
                        modus === MODUS.ABGABE || !!form.querySelector('[name="uploadPflicht"]')?.checked,
                    status: existing ? existing.status : EINSAMMLER_STATUS.ENTWURF,
                    erstellerEmail: (existing && existing.erstellerEmail) || acc.email,
                    erstellerName: (existing && existing.erstellerName) || acc.name || acc.email,
                    quelle: 'ui'
                };
                const r = upsertEinsammler(draft);
                if (!r.ok) {
                    setMessage('err', (r.errors && r.errors.join(' ')) || 'Speichern fehlgeschlagen.');
                    render();
                    return;
                }
                ui.selectedId = id;
                ui.view = 'detail';
                ui.form = emptyForm();
                const emp = String(fd.get('empfaengerText') || '');
                void (async () => {
                    await afterLocalWrite(r.item, 'kopf', 'Einsammler gespeichert (Entwurf).');
                    if (!emp.trim()) return;
                    const a = addEmpfaengerFromText(id, emp);
                    if (!a.ok) {
                        setMessage('warn', 'Gespeichert, aber Empfänger: ' + (a.errors || []).join(' '));
                        render();
                        return;
                    }
                    await afterLocalAntworten(a.addedItems || [], 'Gespeichert, ' + a.added + ' Empfänger hinzugefügt.');
                })();
            });
        }

        const addEmp = root.querySelector('#einAddEmp');
        if (addEmp) {
            addEmp.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(addEmp);
                const r = addEmpfaengerFromText(ui.selectedId, String(fd.get('empfaengerText') || ''));
                if (!r.ok) {
                    setMessage('err', (r.errors && r.errors[0]) || 'Fehler');
                    render();
                    return;
                }
                void afterLocalAntworten(r.addedItems || [], aAddedMsg(r.added));
            });
        }

        const settingsForm = root.querySelector('#einSettingsForm');
        if (settingsForm) {
            settingsForm.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(settingsForm);
                if (!admin()) {
                    setMessage('err', 'Nur Admins dürfen diese Einstellungen speichern.');
                    render();
                    return;
                }
                saveEinsammlerSettings({
                    schoolName: String(fd.get('schoolName') || ''),
                    // w2: Remote-API in der UI weder zeigen noch aktivieren
                    apiBaseUrl: '',
                    adminKey: '',
                    expDays: Number(fd.get('expDays') || 45),
                    adminEmails: String(fd.get('adminEmails') || '')
                });
                setMessage('ok', 'Einstellungen gespeichert.');
                render();
            });
        }

        const importBtn = root.querySelector('[data-ein-import-code]');
        if (importBtn) {
            importBtn.addEventListener('click', () => {
                const raw = String((root.querySelector('#einImportCode') || {}).value || '');
                const r = importResponseCode(raw);
                if (!r.ok) {
                    setMessage('err', r.error || 'Fehler');
                    render();
                    return;
                }
                const aid = r.slip && r.slip.aid;
                const ant = aid
                    ? (state().antworten || []).find((a) => a.id === aid)
                    : null;
                void afterLocalWrite(ant, 'antwort', 'Antwortcode übernommen.');
            });
        }

        const mergeBtn = root.querySelector('[data-ein-merge-inbox]');
        if (mergeBtn) {
            mergeBtn.addEventListener('click', () => {
                const before = new Set(
                    (state().antworten || [])
                        .filter((a) => a && a.status !== 'offen')
                        .map((a) => a.id + ':' + a.status + ':' + (a.antwortAm || ''))
                );
                const r = mergeInboxIntoStore();
                const changed = (state().antworten || []).filter((a) => {
                    if (!a || a.status === 'offen') return false;
                    return !before.has(a.id + ':' + a.status + ':' + (a.antwortAm || ''));
                });
                void afterLocalAntworten(
                    changed,
                    'Posteingang: ' +
                        r.applied +
                        ' übernommen' +
                        (r.skipped ? ', ' + r.skipped + ' übersprungen' : '') +
                        '.'
                );
            });
        }

        root.querySelector('[data-ein-pull]')?.addEventListener('click', () => {
            if (!admin()) return;
            setMessage('ok', 'Lade von SharePoint …');
            render();
            void pullEinsammlerFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                setMessage(
                    r.ok ? 'ok' : 'warn',
                    r.ok
                        ? 'SharePoint geladen (' +
                              (r.state.einsammler || []).length +
                              ' Einsammler · ' +
                              (r.state.antworten || []).length +
                              ' Antworten).'
                        : r.softError || r.meta.error || 'Pull fehlgeschlagen.'
                );
                render();
            });
        });

        const demo = root.querySelector('[data-ein-demo]');
        if (demo) {
            demo.addEventListener('click', () => {
                if (!canUseDemoActions()) {
                    setMessage('warn', 'Demo nur für Admin oder mit ?demoRole=1.');
                    render();
                    return;
                }
                const acc = account();
                if (!acc.loggedIn) {
                    setMessage('warn', 'Bitte anmelden – Demo wird Ihnen zugeordnet.');
                    render();
                    return;
                }
                loadDemoData({ email: acc.email, name: acc.name });
                saveEinsammlerSettings({ apiBaseUrl: '', adminKey: '' });
                setMessage(
                    'ok',
                    'Demo geladen (local-first, ohne SPO-Push der Seeds, Remote-API aus). Echte Vorgänge bleiben.'
                );
                ui.view = 'uebersicht';
                render();
            });
        }

        const clearDemo = root.querySelector('[data-ein-clear-demo]');
        if (clearDemo) {
            clearDemo.addEventListener('click', () => {
                if (!canUseDemoActions()) {
                    setMessage('warn', 'Demo nur für Admin oder mit ?demoRole=1.');
                    render();
                    return;
                }
                clearDemoSeeds();
                ui.selectedId = '';
                setMessage('ok', 'Demo-Daten entfernt.');
                render();
            });
        }

        const clear = root.querySelector('[data-ein-clear]');
        if (clear) {
            clear.addEventListener('click', () => {
                if (!admin()) {
                    setMessage('err', 'Nur Admins dürfen den gesamten Store leeren.');
                    render();
                    return;
                }
                if (!window.confirm('Wirklich alle lokalen Einsammler-Daten löschen?')) return;
                clearAll();
                ui.selectedId = '';
                setMessage('ok', 'Geleert.');
                render();
            });
        }

        void s;
    }

    function aAddedMsg(n) {
        return n === 1 ? '1 Empfänger hinzugefügt.' : n + ' Empfänger hinzugefügt.';
    }

    render();
    void resolveEinsammlerAdmin(account().email).then(() => {
        if (!admin()) ui.showAll = false;
        render();
    });
    void pullEinsammlerFromSharePoint().then((r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (r.ok) setMessage('ok', 'Mit SharePoint verbunden.');
        render();
    });
    void hydrateEinsammlerStammdaten()
        .then(() => {
            render();
        })
        .catch(() => {
            /* local-first */
        });
}
