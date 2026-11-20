/**
 * Elternbrief Admin-UI: Vorlagen, Editor, Empfänger, Links, Lesebestätigungen.
 */

import {
    fileToDataUrl,
    saveAttachment,
    validateUploadFile
} from './elternbrief-attachments.js';
import { importLeseCode, mergeInboxIntoStore } from './elternbrief-inbox.js';
import { buildMailto, parseKlasseCodes, progressStats, resolveLeseStatus } from './elternbrief-logic.js';
import {
    ABSENDER_MODUS,
    ABSENDER_MODUS_LABEL,
    LESE_STATUS,
    LESE_STATUS_LABEL,
    STATUS,
    STATUS_LABEL,
    newId
} from './elternbrief-model.js';
import {
    filterBriefeForViewer,
    filterEmpfaengerForBriefe,
    getElternbriefAccount,
    getElternbriefAdminAccessCache,
    isElternbriefAdmin,
    isOwnedBy
} from './elternbrief-scope.js';
import { resolveElternbriefAdmin } from './elternbrief-entra-role.js';
import { readDemoRoleOverrideFromUrl } from '../../shared/app-role-switcher-policy.js';
import {
    loadElternbriefSettings,
    resolveElternbriefSiteUrl,
    saveElternbriefSettings
} from './elternbrief-settings.js';
import {
    addEmpfaenger,
    clearAll,
    clearDemoSeeds,
    getBrief,
    listEmpfaenger,
    loadDemoData,
    loadState,
    removeBrief,
    seedDemoIfEmpty,
    setBriefStatus,
    setLeseStatus,
    upsertBrief
} from './elternbrief-store.js';
import {
    getElternbriefSyncMeta,
    pullElternbriefFromSharePoint,
    pushElternbriefBrief,
    pushElternbriefDelete,
    pushElternbriefEmpfaenger,
    pushElternbriefEmpfaengerForBrief
} from './elternbrief-sync.js';
import { VORLAGEN, draftFromVorlage } from './elternbrief-templates.js';
import {
    buildPublicPayload,
    buildPublicUrl,
    encodePublicToken
} from './elternbrief-token.js';
import { bindSystemTabClicks, renderSystemAdminPage } from '../../shared/system-admin-shell.js';
import {
    buildEmpfaengerLinesForKlassen,
    getEinsammlerStammdatenMeta,
    hydrateEinsammlerStammdaten,
    loadEinsammlerStammdaten,
    parseKlasseCodesField
} from '../einsammler/einsammler-stammdaten.js';

const VIEWS = [
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2' },
    { id: 'liste', label: 'Briefe', icon: 'bi-envelope-paper' },
    { id: 'neu', label: 'Neu / Editor', icon: 'bi-pencil-square' },
    { id: 'detail', label: 'Empfänger', icon: 'bi-people' },
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
    if (s === 'gelesen' || s === 'gesendet' || s === 'bereit') return 'eb-pill--ok';
    if (s === 'offen' || s === 'entwurf') return 'eb-pill--offen';
    if (s === 'abgelaufen' || s === 'geschlossen') return 'eb-pill--bad';
    if (s === 'gesendet') return 'eb-pill--aktiv';
    return '';
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    /** @type {{ view: string, selectedId: string, message: { type: string, text: string }|null, form: object, showAll: boolean, pendingFile: File|null }} */
    const ui = {
        view: 'uebersicht',
        systemTab: 'site',
        selectedId: '',
        message: null,
        form: emptyForm(),
        showAll: false,
        pendingFile: null
    };

    const demoUiEnabled = readDemoRoleOverrideFromUrl(false);

    seedDemoIfEmpty();
    try {
        mergeInboxIntoStore();
    } catch {
        /* ignore */
    }

    function syncHeaderChrome() {
        const sp = document.getElementById('ebHeaderSpSetup');
        if (sp) sp.hidden = !admin();
    }

    function canUseDemoActions() {
        return admin() || demoUiEnabled;
    }

    function emptyForm() {
        const settings = loadElternbriefSettings();
        const acc = getElternbriefAccount();
        const in14 = new Date(Date.now() + 14 * 86400000).toISOString().slice(0, 10);
        return {
            id: '',
            betreff: '',
            inhalt: '',
            vorlageId: 'frei',
            klasseCodes: '',
            absenderModus: settings.absenderDefault || ABSENDER_MODUS.EIGEN,
            absenderEmail:
                settings.absenderDefault === ABSENDER_MODUS.ZENTRAL
                    ? settings.zentralPostfach
                    : acc.email || '',
            absenderName:
                settings.absenderDefault === ABSENDER_MODUS.ZENTRAL
                    ? settings.zentralName
                    : acc.name || '',
            faelligAm: in14,
            lesebestaetigung: true,
            empfaengerText: '',
            anhangName: '',
            status: STATUS.ENTWURF
        };
    }

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    /**
     * Nach lokalem Write: Brief (+ optional Empfänger) nach SPO pushen.
     * @param {object|null} brief
     * @param {string} [okMsg]
     * @param {{ pushEmpfaenger?: boolean }} [opts]
     */
    async function afterLocalBriefWrite(brief, okMsg, opts = {}) {
        if (!brief) {
            setMessage('ok', okMsg || 'Gespeichert.');
            render();
            return;
        }
        const push = await pushElternbriefBrief(brief);
        let empPush = { ok: true };
        if (opts.pushEmpfaenger !== false && push.ok && !push.skipped) {
            empPush = await pushElternbriefEmpfaengerForBrief(brief.id);
        }
        if (push.ok && empPush.ok) {
            setMessage(
                'ok',
                push.skipped
                    ? okMsg || 'Gespeichert (Demo lokal).'
                    : okMsg || 'In SharePoint gespeichert.'
            );
        } else if (push.skipped) {
            setMessage('ok', okMsg || 'Gespeichert (Demo lokal).');
        } else {
            setMessage(
                'warn',
                (okMsg ? okMsg + ' ' : '') +
                    (push.error || empPush.error || 'SharePoint-Sync fehlgeschlagen – lokal gespeichert.')
            );
        }
        render();
        return push;
    }

    function state() {
        return loadState();
    }

    function renderSyncHint() {
        const sync = getElternbriefSyncMeta();
        if (sync.source === 'live' && sync.ok) return '';
        if (!sync.siteUrl && sync.source !== 'demo') {
            return (
                '<section class="eb-calm-banner eb-calm-banner--ops" role="status">' +
                '<div class="eb-calm-banner__text">' +
                '<p class="eb-calm-banner__eyebrow">SharePoint</p>' +
                '<p>Noch keine Site – IT richtet die Listen ein. Bis dahin lokal/Demo.</p>' +
                '</div></section>'
            );
        }
        if (sync.source === 'demo' || sync.source === 'local') {
            return (
                '<section class="eb-calm-banner eb-calm-banner--ops" role="status">' +
                '<div class="eb-calm-banner__text">' +
                '<p class="eb-calm-banner__eyebrow">Lokal / Demo</p>' +
                '<p>' +
                esc(
                    sync.error ||
                        'SharePoint nicht verbunden – Briefe bleiben lokal, bis die Listen greifen.'
                ) +
                '</p></div></section>'
            );
        }
        return '';
    }

    function account() {
        return getElternbriefAccount();
    }

    function admin() {
        return isElternbriefAdmin(account().email);
    }

    function visibleBriefe(s) {
        return filterBriefeForViewer(s.briefe || [], {
            email: account().email,
            admin: admin(),
            showAll: ui.showAll
        });
    }

    function canAccess(brief) {
        if (!brief) return false;
        if (admin() && ui.showAll) return true;
        const acc = account();
        if (!acc.loggedIn) return false;
        return isOwnedBy(brief, acc.email);
    }

    function selectedBrief(s) {
        const brief = (s.briefe || []).find((b) => b.id === ui.selectedId) || null;
        if (!brief || !canAccess(brief)) return null;
        return brief;
    }

    function scopeBanner() {
        const acc = account();
        if (!acc.loggedIn) {
            return (
                '<p class="eb-msg eb-msg--warn">Bitte mit Microsoft-Konto anmelden – dann sehen Sie nur <strong>Ihre</strong> Elternbriefe.</p>'
            );
        }
        const who = esc(acc.name || acc.email);
        const adm = admin();
        return (
            '<div class="eb-scope-bar">' +
            '<span>Angemeldet: <strong>' +
            who +
            '</strong>' +
            (adm ? ' · Admin' : ' · Lehrkraft') +
            (ui.showAll && adm ? ' · <em>alle Briefe</em>' : ' · nur eigene') +
            '</span>' +
            (adm
                ? '<label class="eb-scope-toggle"><input type="checkbox" id="ebShowAll"' +
                  (ui.showAll ? ' checked' : '') +
                  '> Alle Elternbriefe (Admin)</label>'
                : '') +
            '</div>'
        );
    }

    function render() {
        syncHeaderChrome();
        const s = state();
        if ((ui.view === 'detail' || ui.view === 'links') && !selectedBrief(s)) {
            ui.view = 'liste';
            ui.selectedId = '';
        }
        root.innerHTML =
            '<div class="eb-shell">' +
            renderNav() +
            '<div class="eb-main">' +
            scopeBanner() +
            renderSyncHint() +
            renderMessage() +
            renderView(s) +
            '</div></div>';
        bind(s);
    }

    function renderNav() {
        const isAdm = admin();
        const buttons = VIEWS.filter((v) => {
            if (v.id === 'detail' || v.id === 'links') return !!ui.selectedId;
            return true;
        })
            .map((v) => {
                const label = v.id === 'einstellungen' && !isAdm ? 'Mein Bereich' : v.label;
                return (
                    '<button type="button" class="eb-nav__btn' +
                    (ui.view === v.id ? ' is-active' : '') +
                    '" data-eb-view="' +
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
            '<nav class="eb-nav" aria-label="Elternbrief">' +
            '<div class="eb-nav__brand">' +
            '<i class="bi bi-envelope-paper" aria-hidden="true"></i>' +
            '<div><strong>Elternbrief</strong><span>Vorlagen · Links · Lesebestätigung</span></div>' +
            '</div>' +
            '<div class="eb-nav__list">' +
            buttons +
            '</div>' +
            '<p class="eb-nav__hint">Briefe gehen als persönliche Links an Eltern – Mail nur als Versandkanal. Lesebestätigung wie beim Einsammler.</p>' +
            '<div class="eb-nav__links">' +
            '<a href="elternbrief-lesen.html" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i> Öffentliche Leseseite</a>' +
            (isAdm
                ? '<a href="sharepoint-liste-elternbrief.html"><i class="bi bi-list-ul"></i> SharePoint-Listen</a>'
                : '') +
            '</div></nav>'
        );
    }

    function quietEmpty(title, text, ctaView, ctaLabel) {
        return (
            '<div class="eb-empty">' +
            '<p class="eb-empty__title">' +
            esc(title) +
            '</p>' +
            '<p class="eb-empty__text">' +
            esc(text) +
            '</p>' +
            (ctaView
                ? '<button type="button" class="btn btn-success" data-eb-view="' +
                  esc(ctaView) +
                  '"><i class="bi bi-plus-lg" aria-hidden="true"></i> ' +
                  esc(ctaLabel || 'Neu anlegen') +
                  '</button>'
                : '') +
            '</div>'
        );
    }

    /**
     * @param {object[]} briefe
     * @param {object[]} emp
     */
    function actionBanner(briefe, emp) {
        const drafts = (briefe || []).filter(
            (b) => b.status === STATUS.ENTWURF || b.status === STATUS.BEREIT
        );
        if (drafts.length) {
            const first = drafts[0];
            return (
                '<section class="eb-calm-banner eb-calm-banner--action" role="status">' +
                '<div class="eb-calm-banner__text">' +
                '<p class="eb-calm-banner__eyebrow">Nächster Schritt</p>' +
                '<p><strong>' +
                drafts.length +
                '</strong> Brief' +
                (drafts.length === 1 ? '' : 'e') +
                ' bereit zum Freigeben – dann entstehen Eltern-Links.</p>' +
                '</div>' +
                '<button type="button" class="btn btn-success" data-eb-send="' +
                esc(first.id) +
                '">„' +
                esc(first.betreff || 'Brief') +
                '“ freigeben</button>' +
                '</section>'
            );
        }
        const sent = (briefe || []).filter((b) => b.status === STATUS.GESENDET);
        let openRead = 0;
        for (const b of sent) {
            openRead += progressStats(b, emp || []).offen;
        }
        if (sent.length && openRead > 0) {
            const first = sent[0];
            return (
                '<section class="eb-calm-banner eb-calm-banner--action" role="status">' +
                '<div class="eb-calm-banner__text">' +
                '<p class="eb-calm-banner__eyebrow">Offene Lesebestätigungen</p>' +
                '<p><strong>' +
                openRead +
                '</strong> offen – Links teilen oder Status manuell setzen.</p>' +
                '</div>' +
                '<button type="button" class="btn" data-eb-links="' +
                esc(first.id) +
                '">Eltern-Links öffnen</button>' +
                '</section>'
            );
        }
        if (!(briefe || []).length && account().loggedIn) {
            return (
                '<section class="eb-calm-banner eb-calm-banner--action" role="status">' +
                '<div class="eb-calm-banner__text">' +
                '<p class="eb-calm-banner__eyebrow">Loslegen</p>' +
                '<p>Vorlage wählen, Klassen übernehmen, speichern und senden freigeben.</p>' +
                '</div>' +
                '<button type="button" class="btn btn-success" data-eb-view="neu"><i class="bi bi-plus-lg" aria-hidden="true"></i> Neu anlegen</button>' +
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
                ? 'eb-msg--ok'
                : ui.message.type === 'warn'
                  ? 'eb-msg--warn'
                  : 'eb-msg--err';
        return '<p class="eb-msg ' + cls + '" role="status">' + esc(ui.message.text) + '</p>';
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

    function publicHrefFor(brief, emp) {
        const settings = loadElternbriefSettings();
        const payload = buildPublicPayload({
            brief,
            empfaenger: emp,
            schoolName: settings.schoolName,
            expDays: settings.expDays,
            hasAnhang: !!(brief.anhangName || brief.anhangKey)
        });
        const token = encodePublicToken(payload);
        const abs = new URL('elternbrief-lesen.html', window.location.href).href;
        return buildPublicUrl(abs, token);
    }

    function card(label, value) {
        return (
            '<div class="eb-card"><p class="eb-card__label">' +
            esc(label) +
            '</p><p class="eb-card__value">' +
            esc(value) +
            '</p></div>'
        );
    }

    function renderUebersicht(s) {
        const briefe = visibleBriefe(s);
        const emp = filterEmpfaengerForBriefe(s.empfaenger || [], briefe);
        const aktiv = briefe.filter(
            (b) => b.status === STATUS.GESENDET || b.status === STATUS.BEREIT
        ).length;
        let offen = 0;
        let total = 0;
        for (const b of briefe) {
            const st = progressStats(b, emp);
            offen += st.offen;
            total += st.total;
        }
        const rows = briefe
            .slice(0, 8)
            .map((b) => {
                const st = progressStats(b, emp);
                return (
                    '<tr>' +
                    '<td><button type="button" class="btn" data-eb-open="' +
                    esc(b.id) +
                    '" style="padding:4px 8px;font-size:0.85rem;">' +
                    esc(b.betreff) +
                    '</button></td>' +
                    '<td><span class="eb-pill ' +
                    pillClassForStatus(b.status) +
                    '">' +
                    esc(STATUS_LABEL[b.status] || b.status) +
                    '</span></td>' +
                    '<td>' +
                    st.gelesen +
                    '/' +
                    st.total +
                    ' (' +
                    st.pct +
                    '%)</td>' +
                    '<td>' +
                    esc(formatDeDate(b.faelligAm)) +
                    '</td></tr>'
                );
            })
            .join('');

        return (
            actionBanner(briefe, emp) +
            '<div class="eb-head"><div><h2>Übersicht</h2>' +
            '<p>Elternbriefe mit Vorlagen, Anhang und Lesebestätigung – Versand als persönlicher Link.</p></div>' +
            '<div class="eb-actions">' +
            '<button type="button" class="btn btn-success" data-eb-view="neu"><i class="bi bi-plus-lg"></i> Neu</button>' +
            '</div></div>' +
            '<div class="eb-cards">' +
            card('Meine Briefe', String(briefe.length)) +
            card('Aktiv / gesendet', String(aktiv)) +
            card('Offene Lesebestätigungen', String(offen)) +
            card('Empfänger gesamt', String(total)) +
            '</div>' +
            (briefe.length
                ? '<div class="eb-table-wrap"><table class="eb-table"><thead><tr>' +
                  '<th>Betreff</th><th>Status</th><th>Gelesen</th><th>Frist</th>' +
                  '</tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : account().loggedIn
                  ? quietEmpty(
                        'Noch keine Elternbriefe',
                        'Vorlage wählen, Klassen übernehmen, speichern und senden freigeben.',
                        'neu',
                        'Neu anlegen'
                    )
                  : quietEmpty(
                        'Anmeldung nötig',
                        'Bitte mit Microsoft-Konto anmelden, um Ihre Elternbriefe zu sehen.',
                        '',
                        ''
                    ))
        );
    }

    function renderListe(s) {
        const briefe = visibleBriefe(s);
        const emp = filterEmpfaengerForBriefe(s.empfaenger || [], briefe);
        const rows = briefe
            .map((b) => {
                const st = progressStats(b, emp);
                return (
                    '<tr>' +
                    '<td><strong>' +
                    esc(b.betreff) +
                    '</strong><div class="eb-muted">' +
                    esc(b.klasseCodes || '–') +
                    (b.vorlageId ? ' · Vorlage ' + esc(b.vorlageId) : '') +
                    (admin() && ui.showAll && b.erstellerEmail
                        ? ' · ' + esc(b.erstellerName || b.erstellerEmail)
                        : '') +
                    '</div>' +
                    '<div class="eb-progress" title="' +
                    st.pct +
                    '%"><span style="width:' +
                    st.pct +
                    '%"></span></div></td>' +
                    '<td><span class="eb-pill ' +
                    pillClassForStatus(b.status) +
                    '">' +
                    esc(STATUS_LABEL[b.status] || b.status) +
                    '</span></td>' +
                    '<td>' +
                    st.offen +
                    ' offen</td>' +
                    '<td class="eb-inline-actions">' +
                    '<button type="button" data-eb-open="' +
                    esc(b.id) +
                    '">Empfänger</button>' +
                    '<button type="button" data-eb-links="' +
                    esc(b.id) +
                    '">Links</button>' +
                    '<button type="button" data-eb-edit="' +
                    esc(b.id) +
                    '">Bearbeiten</button>' +
                    (b.status === STATUS.ENTWURF || b.status === STATUS.BEREIT
                        ? '<button type="button" data-eb-send="' + esc(b.id) + '">Senden freigeben</button>'
                        : '') +
                    (b.status === STATUS.GESENDET
                        ? '<button type="button" data-eb-close="' + esc(b.id) + '">Schließen</button>'
                        : '') +
                    '<button type="button" data-eb-delete="' +
                    esc(b.id) +
                    '">Löschen</button>' +
                    '</td></tr>'
                );
            })
            .join('');

        return (
            actionBanner(briefe, emp) +
            '<div class="eb-head"><div><h2>Elternbriefe</h2><p>' +
            (admin() && ui.showAll
                ? 'Admin-Sicht: alle Briefe (dieser Browser).'
                : 'Nur Ihre angelegten Briefe.') +
            '</p></div>' +
            '<div class="eb-actions"><button type="button" class="btn btn-success" data-eb-view="neu">Neu</button></div></div>' +
            (rows
                ? '<div class="eb-table-wrap"><table class="eb-table"><thead><tr><th>Brief</th><th>Status</th><th>Offen</th><th></th></tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty(
                      'Keine Einträge',
                      'In Ihrer Sicht sind noch keine Briefe. Legen Sie einen neuen Elternbrief an.',
                      'neu',
                      'Neu anlegen'
                  ))
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

    function renderNeu() {
        const f = ui.form;
        const sd = loadEinsammlerStammdaten();
        const selected = new Set(parseKlasseCodesField(f.klasseCodes).map((c) => c.toLowerCase()));
        const classChips = sd.classes.length
            ? '<div class="eb-class-pick" id="ebClassPick">' +
              sd.classes
                  .map((c) => {
                      const on = selected.has(c.code.toLowerCase());
                      return (
                          '<label class="eb-class-chip' +
                          (on ? ' is-on' : '') +
                          '"><input type="checkbox" data-eb-klasse="' +
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
              '<div class="eb-form__actions" style="margin:0 0 10px;">' +
              '<button type="button" class="btn" id="ebFillFromKlassen">Empfänger neu aus Klassen laden</button>' +
              '</div>' +
              '<p class="eb-hint">Klassen anhaken → Empfänger aus Stammdaten. ' +
              esc(stammdatenSourceHint()) +
              ' · ' +
              sd.classes.length +
              ' sichtbar · ' +
              sd.students.length +
              ' Schüler:innen.</p>' +
              (sd.students.length
                  ? ''
                  : '<p class="eb-hint">Keine Schüler:innen geladen. Unter <a href="../tenant.html">Stammdaten → Schüler</a> prüfen und die Seite neu laden.</p>')
            : '<p class="eb-hint">Keine Klassen in den Stammdaten – unter <a href="../tenant.html">Stammdaten</a> pflegen oder Codes manuell.</p>';

        const vorlagen = VORLAGEN.map(
            (v) =>
                '<button type="button" class="eb-vorlage' +
                (f.vorlageId === v.id ? ' is-on' : '') +
                '" data-eb-vorlage="' +
                esc(v.id) +
                '"><strong>' +
                esc(v.titel) +
                '</strong><span>' +
                esc(v.hinweis) +
                '</span></button>'
        ).join('');

        return (
            '<div class="eb-head"><div><h2>' +
            (f.id ? 'Elternbrief bearbeiten' : 'Neuer Elternbrief') +
            '</h2><p>Vorlage → Text anpassen → Klassen → optional Anhang → speichern → senden freigeben.</p></div></div>' +
            '<div class="eb-vorlagen">' +
            vorlagen +
            '</div>' +
            '<form class="eb-form" id="ebForm">' +
            '<label>Betreff<input name="betreff" required value="' +
            esc(f.betreff) +
            '"></label>' +
            '<label>Inhalt<textarea name="inhalt" required placeholder="Text für Erziehungsberechtigte">' +
            esc(f.inhalt) +
            '</textarea></label>' +
            '<div><span class="eb-field-label">Klassen aus Stammdaten</span>' +
            classChips +
            '</div>' +
            '<label>Klassen-Codes<input name="klasseCodes" id="ebKlasseCodes" placeholder="3A;2B" value="' +
            esc(f.klasseCodes) +
            '"></label>' +
            '<div class="eb-form__row">' +
            '<label>Absender-Modus<select name="absenderModus" id="ebAbsenderModus">' +
            option(ABSENDER_MODUS.EIGEN, ABSENDER_MODUS_LABEL.eigenes_konto, f.absenderModus) +
            option(
                ABSENDER_MODUS.ZENTRAL,
                ABSENDER_MODUS_LABEL.zentrales_postfach,
                f.absenderModus
            ) +
            '</select></label>' +
            '<label>Lesebestätigung bis<input type="date" name="faelligAm" value="' +
            esc(f.faelligAm) +
            '"></label></div>' +
            '<div class="eb-form__row">' +
            '<label>Absender-Name<input name="absenderName" value="' +
            esc(f.absenderName) +
            '"></label>' +
            '<label>Absender-E-Mail<input name="absenderEmail" type="email" value="' +
            esc(f.absenderEmail) +
            '"></label></div>' +
            '<label class="eb-hint" style="display:flex;gap:8px;align-items:center;font-weight:500;">' +
            '<input type="checkbox" name="lesebestaetigung"' +
            (f.lesebestaetigung ? ' checked' : '') +
            '> Lesebestätigung verlangen</label>' +
            '<label>Anhang (PDF/Bild, max. 6 MB)<input type="file" name="anhang" id="ebAnhang" accept=".pdf,.png,.jpg,.jpeg,.doc,.docx,application/pdf,image/*">' +
            (f.anhangName
                ? '<span class="eb-hint" style="margin:4px 0 0;">Aktuell: ' + esc(f.anhangName) + '</span>'
                : '') +
            '</label>' +
            '<label>Empfänger (Name;Klasse;Eltern-E-Mail)<textarea name="empfaengerText" id="ebEmpfaengerText" placeholder="Anna Berger;3A;eltern@example.com">' +
            esc(f.empfaengerText) +
            '</textarea></label>' +
            '<div class="eb-form__actions">' +
            '<button type="submit" class="btn btn-success">Speichern</button>' +
            '<button type="button" class="btn" data-eb-view="liste">Abbrechen</button>' +
            '</div></form>' +
            (f.inhalt
                ? '<div class="eb-preview" aria-label="Vorschau"><strong>Vorschau</strong>\n\n' +
                  esc(f.inhalt) +
                  '</div>'
                : '')
        );
    }

    function renderDetail(s) {
        const brief = selectedBrief(s);
        if (!brief) return '<div class="eb-empty">Kein Brief gewählt.</div>';
        const st = progressStats(brief, s.empfaenger);
        const rows = (s.empfaenger || [])
            .filter((e) => e.briefId === brief.id)
            .map((e) => {
                const eff = resolveLeseStatus(brief, e);
                const opts = [LESE_STATUS.OFFEN, LESE_STATUS.GELESEN, LESE_STATUS.ABGELAUFEN]
                    .map((x) => option(x, LESE_STATUS_LABEL[x] || x, eff))
                    .join('');
                const linkBtn =
                    brief.status === STATUS.GESENDET || brief.status === STATUS.BEREIT
                        ? '<button type="button" data-eb-copy-one="' +
                          esc(e.id) +
                          '">Link</button>' +
                          (e.elternEmail
                              ? '<a class="btn" style="padding:4px 8px;font-size:0.78rem;text-decoration:none;" data-eb-mailto="' +
                                esc(e.id) +
                                '" href="#">Mail</a>'
                              : '')
                        : '';
                return (
                    '<tr>' +
                    '<td>' +
                    esc(e.schuelerName) +
                    '<div class="eb-muted">' +
                    esc(e.klasseCode || '–') +
                    (e.elternEmail ? ' · ' + esc(e.elternEmail) : '') +
                    '</div></td>' +
                    '<td><span class="eb-pill ' +
                    pillClassForStatus(eff) +
                    '">' +
                    esc(LESE_STATUS_LABEL[eff] || eff) +
                    '</span></td>' +
                    '<td class="eb-inline-actions">' +
                    linkBtn +
                    '<select data-eb-status-for="' +
                    esc(e.id) +
                    '">' +
                    opts +
                    '</select>' +
                    '<button type="button" data-eb-apply-status="' +
                    esc(e.id) +
                    '">Manuell</button>' +
                    '</td></tr>'
                );
            })
            .join('');

        const sendHint =
            brief.status === STATUS.ENTWURF || brief.status === STATUS.BEREIT
                ? '<p class="eb-msg eb-msg--warn">Noch nicht freigegeben – bitte <strong>Senden freigeben</strong>, dann erscheinen Eltern-Links.</p>' +
                  '<div class="eb-actions" style="margin-bottom:12px;"><button type="button" class="btn btn-success" data-eb-send="' +
                  esc(brief.id) +
                  '">Jetzt senden freigeben</button></div>'
                : '<p class="eb-hint">Eltern öffnen den persönlichen Link. Status „Manuell“ nur als Fallback.</p>';

        return (
            '<div class="eb-head"><div><h2>' +
            esc(brief.betreff) +
            '</h2><p>' +
            esc(brief.klasseCodes || '–') +
            ' · Frist ' +
            esc(formatDeDate(brief.faelligAm)) +
            (brief.anhangName ? ' · Anhang: ' + esc(brief.anhangName) : '') +
            '</p></div>' +
            '<div class="eb-actions">' +
            '<button type="button" class="btn" data-eb-view="links">Eltern-Links</button>' +
            '<button type="button" class="btn" data-eb-edit="' +
            esc(brief.id) +
            '">Bearbeiten</button>' +
            '<button type="button" class="btn" data-eb-view="liste">Zurück</button>' +
            '</div></div>' +
            sendHint +
            '<div class="eb-cards">' +
            card('Empfänger', String(st.total)) +
            card('Gelesen', String(st.gelesen)) +
            card('Offen', String(st.offen)) +
            card('Fortschritt', st.pct + '%') +
            '</div>' +
            (rows
                ? '<div class="eb-table-wrap"><table class="eb-table"><thead><tr><th>Schüler:in</th><th>Status</th><th></th></tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : '<div class="eb-empty">Noch keine Empfänger – Brief bearbeiten und Klassen übernehmen.</div>')
        );
    }

    function renderLinks(s) {
        const brief = selectedBrief(s);
        if (!brief) return '<div class="eb-empty">Kein Brief gewählt.</div>';
        if (brief.status !== STATUS.GESENDET && brief.status !== STATUS.BEREIT) {
            return (
                '<div class="eb-empty">Links erscheinen nach „Senden freigeben“.</div>' +
                '<div class="eb-actions" style="margin-top:12px;"><button type="button" class="btn btn-success" data-eb-send="' +
                esc(brief.id) +
                '">Senden freigeben</button></div>'
            );
        }
        const emps = (s.empfaenger || []).filter((e) => e.briefId === brief.id);
        const list = emps
            .map((e) => {
                const href = publicHrefFor(brief, e);
                return (
                    '<div class="eb-link-row" data-eb-link-id="' +
                    esc(e.id) +
                    '">' +
                    '<strong>' +
                    esc(e.schuelerName) +
                    '</strong>' +
                    '<input readonly value="' +
                    esc(href) +
                    '">' +
                    '<div class="eb-inline-actions">' +
                    '<button type="button" data-eb-copy-one="' +
                    esc(e.id) +
                    '">Kopieren</button>' +
                    (e.elternEmail
                        ? '<a class="btn" style="padding:4px 8px;font-size:0.78rem;text-decoration:none;" href="' +
                          esc(buildMailto(brief, e, href)) +
                          '">Mail vorbereiten</a>'
                        : '<span class="eb-hint">Keine Eltern-Mail</span>') +
                    '</div></div>'
                );
            })
            .join('');

        return (
            '<section class="eb-calm-banner eb-calm-banner--action" role="status">' +
            '<div class="eb-calm-banner__text">' +
            '<p class="eb-calm-banner__eyebrow">Teilen</p>' +
            '<p>Links kopieren oder „Mail vorbereiten“ – noch kein automatischer Versand.</p>' +
            '</div></section>' +
            '<div class="eb-head"><div><h2>Eltern-Links</h2><p>' +
            esc(brief.betreff) +
            ' – persönlicher Leselink je Empfänger.</p></div>' +
            '<div class="eb-actions"><button type="button" class="btn" data-eb-view="detail">Empfänger</button></div></div>' +
            '<div class="eb-link-list">' +
            (list ||
                quietEmpty(
                    'Keine Empfänger',
                    'Legen Sie zuerst Empfänger im Editor an.',
                    'neu',
                    'Zum Editor'
                )) +
            '</div>'
        );
    }

    function renderEinstellungen(s) {
        const settings = loadElternbriefSettings();
        const isAdm = admin();
        const demoBtns = canUseDemoActions()
            ? '<button type="button" class="btn" id="ebLoadDemo">Demo laden</button>' +
              '<button type="button" class="btn" id="ebClearDemo">Demo entfernen</button>'
            : '';
        const sync = getElternbriefSyncMeta();
        const site = resolveElternbriefSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
        const syncPanel =
            '<div class="eb-system-panel">' +
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
            '<p>Lokal: ' +
            (s.briefe || []).length +
            ' Briefe · ' +
            (s.empfaenger || []).length +
            ' Empfänger' +
            (s.localDemoOnly ? ' · Demo local-first' : '') +
            ' · mit spoItemId: ' +
            (s.briefe || []).filter((b) => b.spoItemId).length +
            '</p>' +
            '<div class="eb-actions">' +
            '<button type="button" class="btn btn-success" id="ebPullSp">Von SharePoint laden</button>' +
            (isAdm
                ? '<a class="btn" href="sharepoint-liste-elternbrief.html">Listen einrichten</a>'
                : '') +
            '</div></div>';

        if (!isAdm) {
            return (
                '<div class="eb-head"><div><h2>Mein Bereich</h2><p>Persönliche Elternbriefe – SharePoint ist System of Record.</p></div></div>' +
                syncPanel +
                '<p class="eb-hint">Briefe: ' +
                (s.briefe || []).filter((b) => isOwnedBy(b, account().email)).length +
                ' · Absender-Standard: ' +
                esc(ABSENDER_MODUS_LABEL[settings.absenderDefault] || settings.absenderDefault) +
                '</p>' +
                '<p class="eb-hint">' +
                esc(stammdatenSourceHint()) +
                '</p>' +
                '<div class="eb-actions">' +
                demoBtns +
                '</div>'
            );
        }
        const ebRole = getElternbriefAdminAccessCache();
        return (
            '<form class="eb-form" id="ebSettingsForm">' +
            renderSystemAdminPage({
                prefix: 'eb',
                activeTab: ui.systemTab,
                dataAttr: 'data-eb-sys-tab',
                ariaLabel: 'Systemadministration Elternbrief',
                panels: [
                    {
                        id: 'site',
                        step: '1',
                        title: 'Datenquelle',
                        leadHtml: 'Lesebestätigungen landen in der Schul-SharePoint-Site.',
                        body: syncPanel
                    },
                    {
                        id: 'perms',
                        step: '2',
                        title: 'Berechtigungen',
                        leadHtml:
                            'Admin „Alle sehen“ aus Schulleitung/Verwaltung (Entra) bzw. Schul-IT. Legacy-Mails optional.',
                        body:
                            '<label>Zusätzliche Admin-E-Mails (Legacy)<textarea name="adminEmails">' +
                            esc(settings.adminEmails) +
                            '</textarea></label>' +
                            (ebRole.sources.length
                                ? '<p class="eb-hint">Ihre Admin-Quellen: ' +
                                  esc(ebRole.sources.join(', ')) +
                                  (ebRole.demo ? ' (IT-Vorschau)' : '') +
                                  '</p>'
                                : '') +
                            '<div class="eb-form__actions"><button type="submit" class="btn btn-success">Speichern</button></div>'
                    },
                    {
                        id: 'kalender',
                        step: '3',
                        title: 'Kalender',
                        leadHtml: 'Kein Terminkalender. Die Link-Frist gilt für den Eltern-Leselink.',
                        body:
                            '<label>Link gültig (Tage)<input name="expDays" type="number" min="1" value="' +
                            esc(String(settings.expDays)) +
                            '"></label>' +
                            '<p class="eb-hint">Elternlinks sind HMAC-signiert und stehen im Fragment (<code>#p=</code>).</p>' +
                            '<div class="eb-form__actions"><button type="submit" class="btn btn-success">Speichern</button></div>'
                    },
                    {
                        id: 'wartung',
                        step: '4',
                        title: 'Wartung',
                        leadHtml: 'Absender und Schulname für alle Elternbriefe.',
                        body:
                            '<label>Schulname<input name="schoolName" value="' +
                            esc(settings.schoolName) +
                            '"></label>' +
                            '<label>Standard-Absender<select name="absenderDefault">' +
                            option(
                                ABSENDER_MODUS.EIGEN,
                                ABSENDER_MODUS_LABEL.eigenes_konto,
                                settings.absenderDefault
                            ) +
                            option(
                                ABSENDER_MODUS.ZENTRAL,
                                ABSENDER_MODUS_LABEL.zentrales_postfach,
                                settings.absenderDefault
                            ) +
                            '</select></label>' +
                            '<div class="eb-form__row">' +
                            '<label>Zentrales Postfach<input name="zentralPostfach" type="email" value="' +
                            esc(settings.zentralPostfach) +
                            '" placeholder="elternbrief@schule.at"></label>' +
                            '<label>Name zentrales Postfach<input name="zentralName" value="' +
                            esc(settings.zentralName) +
                            '"></label></div>' +
                            '<p class="eb-hint">' +
                            esc(stammdatenSourceHint()) +
                            '</p>' +
                            '<div class="eb-form__actions"><button type="submit" class="btn btn-success">Speichern</button>' +
                            demoBtns +
                            '<button type="button" class="btn" id="ebImportCode">Bestätigungscode einlesen</button></div>'
                    },
                    {
                        id: 'danger',
                        title: 'Danger Zone',
                        danger: true,
                        leadHtml: 'Leert den lokalen Store. SharePoint-Briefe bleiben, bis sie dort gelöscht werden.',
                        body: '<button type="button" class="btn ms-sys-btn-danger" id="ebClearAll">Store leeren</button>'
                    }
                ]
            }) +
            '</form>'
        );
    }

    function bind(s) {
        bindSystemTabClicks(root, {
            dataAttr: 'data-eb-sys-tab',
            onSelect: (tab) => {
                ui.systemTab = tab;
                render();
            }
        });
        root.querySelectorAll('[data-eb-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = String(btn.getAttribute('data-eb-view') || 'uebersicht');
                setMessage('', '');
                render();
            });
        });

        const showAll = root.querySelector('#ebShowAll');
        if (showAll) {
            showAll.addEventListener('change', () => {
                ui.showAll = !!showAll.checked;
                render();
            });
        }

        root.querySelectorAll('[data-eb-open]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.selectedId = String(btn.getAttribute('data-eb-open') || '');
                ui.view = 'detail';
                render();
            });
        });

        root.querySelectorAll('[data-eb-links]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.selectedId = String(btn.getAttribute('data-eb-links') || '');
                ui.view = 'links';
                render();
            });
        });

        root.querySelectorAll('[data-eb-edit]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = String(btn.getAttribute('data-eb-edit') || '');
                const brief = (state().briefe || []).find((b) => b.id === id);
                if (!brief || !canAccess(brief)) return;
                ui.form = {
                    ...emptyForm(),
                    ...brief,
                    empfaengerText: ''
                };
                ui.pendingFile = null;
                ui.selectedId = id;
                ui.view = 'neu';
                render();
            });
        });

        root.querySelectorAll('[data-eb-vorlage]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = String(btn.getAttribute('data-eb-vorlage') || '');
                const draft = draftFromVorlage(id);
                if (!draft) return;
                ui.form = {
                    ...emptyForm(),
                    ...draft,
                    id: ui.form.id || '',
                    klasseCodes: ui.form.klasseCodes,
                    empfaengerText: ui.form.empfaengerText
                };
                setMessage('ok', 'Vorlage „' + (draft.vorlageId || id) + '“ übernommen – Text anpassen.');
                render();
            });
        });

        function snapshotFormIntoUi() {
            const formEl = root.querySelector('#ebForm');
            if (!formEl) return;
            const fd = new FormData(formEl);
            ui.form = {
                ...ui.form,
                betreff: String(fd.get('betreff') || ui.form.betreff || ''),
                inhalt: String(fd.get('inhalt') || ui.form.inhalt || ''),
                klasseCodes: String(fd.get('klasseCodes') || ui.form.klasseCodes || ''),
                absenderModus: String(fd.get('absenderModus') || ui.form.absenderModus || ''),
                absenderName: String(fd.get('absenderName') || ui.form.absenderName || ''),
                absenderEmail: String(fd.get('absenderEmail') || ui.form.absenderEmail || ''),
                faelligAm: String(fd.get('faelligAm') || ui.form.faelligAm || ''),
                lesebestaetigung: !!fd.get('lesebestaetigung'),
                empfaengerText: String(fd.get('empfaengerText') || ui.form.empfaengerText || '')
            };
        }

        /**
         * Klassen-Auswahl → Empfängerzeilen aus Stammdaten (ersetzt die Liste).
         * @param {string[]} [codesIn]
         */
        function applyEmpfaengerFromKlassen(codesIn) {
            snapshotFormIntoUi();
            const codes = (Array.isArray(codesIn) ? codesIn : parseKlasseCodes(ui.form.klasseCodes))
                .map((c) => String(c || '').trim())
                .filter(Boolean);
            ui.form.klasseCodes = codes.join(';');
            const lines = buildEmpfaengerLinesForKlassen(
                codes,
                loadEinsammlerStammdaten().students
            );
            ui.form.empfaengerText = lines.join('\n');
            if (!codes.length) {
                setMessage('', '');
            } else if (lines.length) {
                setMessage(
                    'ok',
                    lines.length +
                        ' Empfänger aus Stammdaten übernommen (Name;Klasse;Eltern-E-Mail).'
                );
            } else {
                const total = loadEinsammlerStammdaten().students.length;
                setMessage(
                    'warn',
                    total
                        ? 'Für die gewählten Klassen gibt es keine Schüler:innen in den Stammdaten (Klasse stimmt ggf. nicht mit dem Feld „Klasse“ der Schüler überein).'
                        : 'Keine Schüler:innen in den Stammdaten geladen. Bitte unter Stammdaten → Schüler prüfen und die Seite neu laden.'
                );
            }
            render();
            return lines;
        }

        root.querySelectorAll('[data-eb-klasse]').forEach((cb) => {
            cb.addEventListener('change', () => {
                const codes = Array.from(root.querySelectorAll('[data-eb-klasse]:checked')).map(
                    (el) => String(el.getAttribute('data-eb-klasse') || '')
                );
                applyEmpfaengerFromKlassen(codes);
            });
        });

        const klasseCodesInput = root.querySelector('#ebKlasseCodes');
        if (klasseCodesInput) {
            klasseCodesInput.addEventListener('change', () => {
                applyEmpfaengerFromKlassen(parseKlasseCodes(klasseCodesInput.value));
            });
        }

        const fillKlassen = root.querySelector('#ebFillFromKlassen');
        if (fillKlassen) {
            fillKlassen.addEventListener('click', () => {
                applyEmpfaengerFromKlassen(parseKlasseCodes(ui.form.klasseCodes || ''));
            });
        }

        const absModus = root.querySelector('#ebAbsenderModus');
        if (absModus) {
            absModus.addEventListener('change', () => {
                const settings = loadElternbriefSettings();
                const acc = account();
                const v = String(absModus.value || '');
                const nameEl = root.querySelector('[name="absenderName"]');
                const mailEl = root.querySelector('[name="absenderEmail"]');
                if (v === ABSENDER_MODUS.ZENTRAL) {
                    if (nameEl) nameEl.value = settings.zentralName || '';
                    if (mailEl) mailEl.value = settings.zentralPostfach || '';
                } else {
                    if (nameEl) nameEl.value = acc.name || '';
                    if (mailEl) mailEl.value = acc.email || '';
                }
            });
        }

        const form = root.querySelector('#ebForm');
        if (form) {
            const fileInput = root.querySelector('#ebAnhang');
            if (fileInput) {
                fileInput.addEventListener('change', () => {
                    ui.pendingFile = fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
                });
            }
            form.addEventListener('submit', async (ev) => {
                ev.preventDefault();
                const fd = new FormData(form);
                const acc = account();
                if (!acc.loggedIn) {
                    setMessage('err', 'Bitte anmelden, um einen Brief zu speichern.');
                    render();
                    return;
                }
                const id = ui.form.id || newId('eb');
                let anhangName = ui.form.anhangName || '';
                let anhangMime = ui.form.anhangMime || '';
                let anhangSize = ui.form.anhangSize || 0;
                let anhangKey = ui.form.anhangKey || '';
                if (ui.pendingFile) {
                    const v = validateUploadFile(ui.pendingFile);
                    if (!v.ok) {
                        setMessage('err', v.error || 'Anhang ungültig.');
                        render();
                        return;
                    }
                    try {
                        const dataUrl = await fileToDataUrl(ui.pendingFile);
                        anhangKey = id;
                        await saveAttachment(anhangKey, {
                            name: ui.pendingFile.name,
                            mime: ui.pendingFile.type || 'application/octet-stream',
                            size: ui.pendingFile.size,
                            dataUrl
                        });
                        anhangName = ui.pendingFile.name;
                        anhangMime = ui.pendingFile.type || '';
                        anhangSize = ui.pendingFile.size;
                    } catch (e) {
                        setMessage('err', 'Anhang speichern fehlgeschlagen.');
                        render();
                        return;
                    }
                }
                const existing = getBrief(id);
                const draft = {
                    id,
                    betreff: String(fd.get('betreff') || ''),
                    inhalt: String(fd.get('inhalt') || ''),
                    vorlageId: ui.form.vorlageId || '',
                    klasseCodes: String(fd.get('klasseCodes') || ''),
                    absenderModus: String(fd.get('absenderModus') || ABSENDER_MODUS.EIGEN),
                    absenderEmail: String(fd.get('absenderEmail') || ''),
                    absenderName: String(fd.get('absenderName') || ''),
                    faelligAm: String(fd.get('faelligAm') || ''),
                    lesebestaetigung: !!fd.get('lesebestaetigung'),
                    anhangName,
                    anhangMime,
                    anhangSize,
                    anhangKey,
                    status: ui.form.status || STATUS.ENTWURF,
                    erstellerEmail: acc.email,
                    erstellerName: acc.name || acc.email,
                    spoItemId: (existing && existing.spoItemId) || ui.form.spoItemId || '',
                    seedTag: (existing && existing.seedTag) || '',
                    quelle: (existing && existing.quelle) || 'ui'
                };
                const r = upsertBrief(draft);
                if (!r.ok) {
                    setMessage('err', (r.errors || []).join(' '));
                    render();
                    return;
                }
                let empText = String(fd.get('empfaengerText') || '').trim();
                if (!empText) {
                    const codes = parseKlasseCodes(String(fd.get('klasseCodes') || ''));
                    if (codes.length) {
                        empText = buildEmpfaengerLinesForKlassen(
                            codes,
                            loadEinsammlerStammdaten().students
                        ).join('\n');
                    }
                }
                if (empText) {
                    addEmpfaenger(id, { text: empText });
                }
                ui.selectedId = id;
                ui.form = emptyForm();
                ui.pendingFile = null;
                ui.view = 'detail';
                void afterLocalBriefWrite(r.item, 'Elternbrief gespeichert.', {
                    pushEmpfaenger: true
                });
            });
        }

        root.querySelectorAll('[data-eb-send]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = String(btn.getAttribute('data-eb-send') || '');
                const r = setBriefStatus(id, STATUS.GESENDET);
                if (!r.ok) {
                    setMessage('err', (r.errors || []).join(' '));
                    render();
                    return;
                }
                ui.selectedId = id;
                ui.view = 'links';
                void afterLocalBriefWrite(
                    r.item,
                    'Brief freigegeben – persönliche Links sind bereit. Mail manuell vorbereiten oder Link teilen.',
                    { pushEmpfaenger: true }
                );
            });
        });

        root.querySelectorAll('[data-eb-close]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const r = setBriefStatus(
                    String(btn.getAttribute('data-eb-close') || ''),
                    STATUS.GESCHLOSSEN
                );
                void afterLocalBriefWrite(r.item, 'Brief geschlossen.', {
                    pushEmpfaenger: false
                });
            });
        });

        root.querySelectorAll('[data-eb-delete]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!window.confirm('Elternbrief wirklich löschen?')) return;
                const id = String(btn.getAttribute('data-eb-delete') || '');
                const before = getBrief(id);
                const beforeEmp = listEmpfaenger(id);
                removeBrief(id);
                ui.selectedId = '';
                void pushElternbriefDelete(before || {}, beforeEmp).then((del) => {
                    setMessage(
                        del.ok || del.skipped ? 'ok' : 'warn',
                        del.ok || del.skipped
                            ? 'Gelöscht.'
                            : 'Lokal gelöscht. ' + (del.error || 'SharePoint-Löschen fehlgeschlagen.')
                    );
                    render();
                });
            });
        });

        root.querySelectorAll('[data-eb-apply-status]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = String(btn.getAttribute('data-eb-apply-status') || '');
                const sel = root.querySelector('[data-eb-status-for="' + id + '"]');
                const st = sel ? String(sel.value || '') : LESE_STATUS.OFFEN;
                const r = setLeseStatus(id, st);
                if (!r.ok || !r.item) {
                    setMessage('ok', 'Status manuell gesetzt.');
                    render();
                    return;
                }
                void pushElternbriefEmpfaenger(r.item).then((push) => {
                    setMessage(
                        push.ok || push.skipped ? 'ok' : 'warn',
                        push.ok || push.skipped
                            ? 'Status manuell gesetzt.'
                            : 'Lokal gesetzt. ' + (push.error || 'SharePoint-Sync fehlgeschlagen.')
                    );
                    render();
                });
            });
        });

        root.querySelectorAll('[data-eb-copy-one]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const eid = String(btn.getAttribute('data-eb-copy-one') || '');
                const brief = selectedBrief(state());
                const emp = (state().empfaenger || []).find((e) => e.id === eid);
                if (!brief || !emp) return;
                const href = publicHrefFor(brief, emp);
                try {
                    await navigator.clipboard.writeText(href);
                    setMessage('ok', 'Link kopiert.');
                } catch {
                    setMessage('warn', 'Kopieren nicht möglich – Link manuell markieren.');
                }
                render();
            });
        });

        root.querySelectorAll('[data-eb-mailto]').forEach((a) => {
            a.addEventListener('click', (ev) => {
                ev.preventDefault();
                const eid = String(a.getAttribute('data-eb-mailto') || '');
                const brief = selectedBrief(state());
                const emp = (state().empfaenger || []).find((e) => e.id === eid);
                if (!brief || !emp) return;
                const href = publicHrefFor(brief, emp);
                const mail = buildMailto(brief, emp, href);
                if (mail) window.location.href = mail;
            });
        });

        const settingsForm = root.querySelector('#ebSettingsForm');
        if (settingsForm) {
            settingsForm.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(settingsForm);
                saveElternbriefSettings({
                    schoolName: String(fd.get('schoolName') || ''),
                    absenderDefault: String(fd.get('absenderDefault') || ABSENDER_MODUS.EIGEN),
                    zentralPostfach: String(fd.get('zentralPostfach') || ''),
                    zentralName: String(fd.get('zentralName') || ''),
                    adminEmails: String(fd.get('adminEmails') || ''),
                    expDays: Number(fd.get('expDays')) || 45
                });
                setMessage('ok', 'Einstellungen gespeichert.');
                render();
            });
        }

        const demo = root.querySelector('#ebLoadDemo');
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
                setMessage('ok', 'Demo geladen und Ihnen zugeordnet (echte Briefe bleiben).');
                ui.view = 'liste';
                render();
            });
        }

        const clearDemo = root.querySelector('#ebClearDemo');
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

        const clearBtn = root.querySelector('#ebClearAll');
        if (clearBtn) {
            clearBtn.addEventListener('click', () => {
                if (!admin()) {
                    setMessage('err', 'Nur Admins dürfen den gesamten Store leeren.');
                    render();
                    return;
                }
                if (!window.confirm('Alle lokalen Elternbrief-Daten löschen?')) return;
                clearAll();
                setMessage('ok', 'Store geleert.');
                render();
            });
        }

        const importBtn = root.querySelector('#ebImportCode');
        if (importBtn) {
            importBtn.addEventListener('click', () => {
                const raw = window.prompt('Bestätigungscode der Eltern einfügen:');
                if (!raw) return;
                const r = importLeseCode(raw);
                if (!r.ok) {
                    setMessage('err', r.error || 'Fehler');
                    render();
                    return;
                }
                if (r.item) {
                    void pushElternbriefEmpfaenger(r.item).then((push) => {
                        setMessage(
                            push.ok || push.skipped ? 'ok' : 'warn',
                            push.ok || push.skipped
                                ? 'Lesebestätigung übernommen.'
                                : 'Lokal übernommen. ' +
                                      (push.error || 'SharePoint-Sync fehlgeschlagen.')
                        );
                        render();
                    });
                } else {
                    setMessage('ok', 'Lesebestätigung übernommen.');
                    render();
                }
            });
        }

        const pullBtn = root.querySelector('#ebPullSp');
        if (pullBtn) {
            pullBtn.addEventListener('click', () => {
                setMessage('ok', 'Lade von SharePoint …');
                render();
                void pullElternbriefFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                    setMessage(
                        r.ok ? 'ok' : 'warn',
                        r.ok
                            ? 'SharePoint geladen (' +
                                  (r.state.briefe || []).length +
                                  ' Briefe · ' +
                                  (r.state.empfaenger || []).length +
                                  ' Empfänger).'
                            : r.softError || r.meta.error || 'Pull fehlgeschlagen.'
                    );
                    render();
                });
            });
        }

        void s;
    }

    render();
    void resolveElternbriefAdmin(getElternbriefAccount().email).then(() => {
        if (!isElternbriefAdmin(getElternbriefAccount().email)) ui.showAll = false;
        render();
    });
    void hydrateEinsammlerStammdaten({ siteKind: 'elternbrief' })
        .then(() => {
            render();
        })
        .catch(() => {
            /* local-first */
        });
    void pullElternbriefFromSharePoint().then(async (r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (r.ok) setMessage('ok', 'Mit SharePoint verbunden.');
        try {
            const merged = mergeInboxIntoStore();
            for (const emp of (merged && merged.items) || []) {
                await pushElternbriefEmpfaenger(emp);
            }
        } catch {
            /* ignore */
        }
        render();
    });
}
