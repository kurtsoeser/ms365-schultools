/**
 * Krankmeldung Admin-Inbox.
 */

import {
    fileToDataUrl,
    loadAttachment,
    saveAttachment,
    validateUploadFile
} from './krankmeldung-attachments.js';
import {
    buildPublicFormUrl,
    getKrankmeldungAccount,
    loadStammdatenKlassen,
    resolveAdminKlassen,
    resolveKvKlassenForEmail,
    stammdatenKlassenAsText
} from './krankmeldung-klassen.js';
import { filterMeldungen, inboxStats, meldungenToCsv } from './krankmeldung-logic.js';
import {
    KATEGORIE_LABEL,
    STATUS,
    STATUS_LABEL,
    UPLOAD_MAX_BYTES,
    newId,
    toIsoDateOnly
} from './krankmeldung-model.js';
import {
    krankmeldungRemoteConfigured,
    remoteListMeldungen
} from './krankmeldung-remote.js';
import {
    loadKrankmeldungSettings,
    resolveKrankmeldungSiteUrl,
    saveKrankmeldungSettings
} from './krankmeldung-settings.js';
import { readDemoRoleOverrideFromUrl } from '../../shared/app-role-switcher-policy.js';
import {
    getEinsammlerStammdatenMeta,
    hydrateEinsammlerStammdaten
} from '../einsammler/einsammler-stammdaten.js';
import {
    clearAll,
    clearDemoSeeds,
    loadDemoData,
    loadState,
    markKenntnis,
    removeMeldung,
    seedDemoIfEmpty,
    setMeldungStatus,
    upsertMeldung
} from './krankmeldung-store.js';
import {
    getKrankmeldungSyncMeta,
    pullKrankmeldungFromSharePoint,
    pushKrankmeldungDelete,
    pushKrankmeldungMeldung
} from './krankmeldung-sync.js';
import {
    getKrankmeldungRoleCache,
    isKrankmeldungAdminCached,
    resolveKrankmeldungRoles
} from './krankmeldung-entra-role.js';

const VIEWS = [
    { id: 'inbox', label: 'Inbox', icon: 'bi-inbox' },
    { id: 'neu', label: 'Manuell melden', icon: 'bi-plus-circle' },
    { id: 'einstellungen', label: 'Einstellungen', icon: 'bi-gear' }
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

function formatDeDateTime(iso) {
    const s = String(iso || '');
    if (!s) return '–';
    try {
        const d = new Date(s);
        if (Number.isNaN(d.getTime())) return s;
        return d.toLocaleString('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return s;
    }
}

function pillClass(st) {
    if (st === 'neu') return 'km-pill--neu';
    if (st === 'zur_kenntnis') return 'km-pill--ok';
    if (st === 'archiviert') return 'km-pill--arch';
    return '';
}

/** Auto-Seed nur mit ?demo=1 – nie still in Produktivdaten. */
function wantsKrankmeldungDemo() {
    try {
        return new URLSearchParams(window.location.search).get('demo') === '1';
    } catch {
        return false;
    }
}

/**
 * @param {string} title
 * @param {string} text
 * @param {string} [ctaHtml]
 */
function quietEmpty(title, text, ctaHtml) {
    return (
        '<div class="km-quiet-empty" role="status">' +
        '<p class="km-quiet-empty__title">' +
        title +
        '</p>' +
        '<p class="km-quiet-empty__text">' +
        text +
        '</p>' +
        (ctaHtml || '') +
        '</div>'
    );
}

function softRemoteMessage(err) {
    const msg = err && err.message ? String(err.message) : String(err || '');
    if (/403|401|Forbidden|Unauthorized|Zugriff/i.test(msg)) {
        return 'API-Zugriff gerade nicht möglich – lokale Inbox bleibt unverändert.';
    }
    if (/Failed to fetch|NetworkError|Load failed|offline/i.test(msg)) {
        return 'API nicht erreichbar – lokale Inbox bleibt.';
    }
    return 'API-Antwort unbrauchbar – lokale Inbox bleibt. (' + msg.slice(0, 120) + ')';
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    const today = toIsoDateOnly(new Date());
    /** @type {{ view: string, message: {type:string,text:string}|null, filters: object, nurMeineKlassen: boolean }} */
    const ui = {
        view: 'inbox',
        message: null,
        filters: { day: today, klasse: '', status: '', kategorie: '', q: '', mode: 'tag' },
        nurMeineKlassen: false
    };

    const demoUiEnabled = readDemoRoleOverrideFromUrl(false) || wantsKrankmeldungDemo();

    if (wantsKrankmeldungDemo()) seedDemoIfEmpty();

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    /**
     * @param {object|null} item
     * @param {string} [okMsg]
     */
    async function afterLocalWrite(item, okMsg) {
        if (!item) {
            setMessage('ok', okMsg || 'Gespeichert.');
            render();
            return { ok: true, skipped: true };
        }
        const push = await pushKrankmeldungMeldung(item);
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

    function renderSyncHint() {
        const sync = getKrankmeldungSyncMeta();
        if (sync.source === 'live' && sync.ok) return '';
        if (!sync.siteUrl && sync.source !== 'demo') {
            return (
                '<section class="km-calm-banner km-calm-banner--ops" role="status">' +
                '<div class="km-calm-banner__text">' +
                '<p class="km-calm-banner__eyebrow">SharePoint</p>' +
                '<p>Noch keine Site – IT richtet die Liste ein. Bis dahin lokal/Demo.</p>' +
                '</div></section>'
            );
        }
        if (sync.source === 'demo' || sync.source === 'local') {
            return (
                '<section class="km-calm-banner km-calm-banner--ops" role="status">' +
                '<div class="km-calm-banner__text">' +
                '<p class="km-calm-banner__eyebrow">Lokal / Demo</p>' +
                '<p>' +
                esc(
                    sync.error ||
                        'SharePoint nicht verbunden – Meldungen bleiben lokal, bis die Liste greift.'
                ) +
                '</p></div></section>'
            );
        }
        return '';
    }

    function admin() {
        const acc = getKrankmeldungAccount();
        return isKrankmeldungAdminCached(acc.email) || getKrankmeldungRoleCache().isAdmin;
    }

    function canUseDemoActions() {
        return admin() || demoUiEnabled;
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

    function kvScopedKlassen() {
        const acc = getKrankmeldungAccount();
        return resolveKvKlassenForEmail(acc.email);
    }

    function render() {
        if (ui.view === 'einstellungen' && !admin()) ui.view = 'inbox';
        root.innerHTML =
            '<div class="km-shell">' +
            renderNav() +
            '<div class="km-main">' +
            renderMessage() +
            (ui.view === 'inbox' ? renderSyncHint() : '') +
            renderView() +
            '</div></div>';
        bind();
    }

    function renderNav() {
        const isAdm = admin();
        const buttons = VIEWS.filter((v) => isAdm || v.id !== 'einstellungen')
            .map(
                (v) =>
                    '<button type="button" class="km-nav__btn' +
                    (ui.view === v.id ? ' is-active' : '') +
                    '" data-km-view="' +
                    esc(v.id) +
                    '"><i class="bi ' +
                    esc(v.icon) +
                    '" aria-hidden="true"></i>' +
                    esc(v.label) +
                    '</button>'
            )
            .join('');
        return (
            '<nav class="km-nav" aria-label="Krankmeldung">' +
            '<div class="km-nav__brand">' +
            '<i class="bi bi-thermometer-half" aria-hidden="true"></i>' +
            '<div><strong>Krankmeldung</strong><span>Eltern &amp; Ablauf · P3</span></div>' +
            '</div>' +
            '<div class="km-nav__list">' +
            buttons +
            '</div>' +
            '<p class="km-nav__hint">Eltern wählen die Klasse im öffentlichen Formular. Inbox nach Tag, Woche und Klasse. Admin = Sekretariat/Schulleitung (Entra).</p>' +
            '<div class="km-nav__links">' +
            '<a href="' +
            esc(buildPublicFormUrl()) +
            '" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i> Öffentliches Formular</a>' +
            (isAdm
                ? '<a href="sharepoint-liste-krankmeldung.html" data-ms365-frontend-it-only><i class="bi bi-list-ul"></i> SharePoint-Liste</a>'
                : '') +
            '</div></nav>'
        );
    }

    /**
     * Kontextuelles Action-Banner (kein Dauer-How-to).
     * @param {{ offenHeute?: number, neu?: number, total?: number }} stats
     */
    function actionBanner(stats) {
        const settings = loadKrankmeldungSettings();
        const hasPublicKlassen = !!String(settings.publicKlassen || '').trim();
        if (admin() && !hasPublicKlassen) {
            return (
                '<section class="km-calm-banner km-calm-banner--action" role="status">' +
                '<div class="km-calm-banner__text">' +
                '<p class="km-calm-banner__eyebrow">Nächster Schritt</p>' +
                '<p>Klassen für das Elternformular veröffentlichen und Link teilen.</p>' +
                '</div>' +
                '<button type="button" class="btn btn-success" data-km-view="einstellungen">Einstellungen öffnen</button>' +
                '</section>'
            );
        }
        if (stats && stats.offenHeute) {
            return (
                '<section class="km-calm-banner km-calm-banner--action" role="status">' +
                '<div class="km-calm-banner__text">' +
                '<p class="km-calm-banner__eyebrow">Was braucht deine Aufmerksamkeit?</p>' +
                '<p><strong>' +
                esc(String(stats.offenHeute)) +
                '</strong> neue Meldung' +
                (stats.offenHeute === 1 ? '' : 'en') +
                ' für heute – bitte zur Kenntnis nehmen.</p>' +
                '</div></section>'
            );
        }
        if (stats && stats.neu > 0 && !stats.offenHeute) {
            return (
                '<section class="km-calm-banner km-calm-banner--action" role="status">' +
                '<div class="km-calm-banner__text">' +
                '<p class="km-calm-banner__eyebrow">Offene Meldungen</p>' +
                '<p><strong>' +
                esc(String(stats.neu)) +
                '</strong> noch nicht zur Kenntnis genommen – Filter auf Status „Neu“ setzen.</p>' +
                '</div></section>'
            );
        }
        if (stats && !stats.total && hasPublicKlassen) {
            return (
                '<section class="km-calm-banner km-calm-banner--action" role="status">' +
                '<div class="km-calm-banner__text">' +
                '<p class="km-calm-banner__eyebrow">Bereit</p>' +
                '<p>Formular-Link an Eltern teilen – Meldungen erscheinen hier in der Inbox.</p>' +
                '</div>' +
                '<button type="button" class="btn" id="kmCopyFormLink">Link kopieren</button>' +
                '</section>'
            );
        }
        return '';
    }

    function renderMessage() {
        if (!ui.message) return '';
        const cls =
            ui.message.type === 'ok'
                ? 'km-msg--ok'
                : ui.message.type === 'warn'
                  ? 'km-msg--warn'
                  : 'km-msg--err';
        return '<p class="km-msg ' + cls + '" role="status">' + esc(ui.message.text) + '</p>';
    }

    function renderView() {
        if (ui.view === 'neu') return renderNeu();
        if (ui.view === 'einstellungen') return renderEinstellungen();
        return renderInbox();
    }

    function card(label, value) {
        return (
            '<div class="km-card"><p class="km-card__label">' +
            esc(label) +
            '</p><p class="km-card__value">' +
            esc(value) +
            '</p></div>'
        );
    }

    function filteredList(s) {
        const f = ui.filters;
        const acc = getKrankmeldungAccount();
        const kvKlassen = kvScopedKlassen();
        /** @type {string[]|undefined} */
        let klassenAllow;
        if (admin()) {
            if (ui.nurMeineKlassen && kvKlassen.length) klassenAllow = kvKlassen;
        } else if (kvKlassen.length) {
            klassenAllow = kvKlassen;
        } else {
            klassenAllow = ['__keine_berechtigung__'];
        }
        const base = {
            klasse: f.klasse,
            status: f.status,
            kategorie: f.kategorie,
            q: f.q,
            klassenAllow
        };
        if (f.mode === 'alle') {
            return filterMeldungen(s.meldungen, base);
        }
        if (f.mode === 'woche') {
            return filterMeldungen(s.meldungen, {
                ...base,
                weekOf: f.day || today
            });
        }
        return filterMeldungen(s.meldungen, {
            ...base,
            day: f.day || today
        });
    }

    function renderInbox() {
        const s = state();
        const stats = inboxStats(s.meldungen, today);
        const list = filteredList(s);
        const klassen = resolveAdminKlassen(s.meldungen);

        const rows = list
            .map((m) => {
                return (
                    '<tr>' +
                    '<td><strong>' +
                    esc(m.schuelerName) +
                    '</strong><div class="km-muted">' +
                    esc(m.klasseCode) +
                    ' · ' +
                    esc(KATEGORIE_LABEL[m.kategorie] || m.kategorie) +
                    '</div>' +
                    (m.notiz
                        ? '<div class="km-row-note">' + esc(m.notiz) + '</div>'
                        : '') +
                    (m.anhangName
                        ? '<div class="km-row-note">📎 ' +
                          esc(m.anhangName) +
                          ' <button type="button" data-km-anhang="' +
                          esc(m.id) +
                          '">Öffnen</button></div>'
                        : '') +
                    '</td>' +
                    '<td>' +
                    esc(formatDeDate(m.von)) +
                    (m.bis !== m.von ? '–' + esc(formatDeDate(m.bis)) : '') +
                    '</td>' +
                    '<td>' +
                    esc(formatDeDateTime(m.gemeldetAm)) +
                    '<div class="km-muted">' +
                    esc(m.gemeldetVonName || m.gemeldetVonEmail || m.quelle) +
                    '</div></td>' +
                    '<td><span class="km-pill ' +
                    pillClass(m.status) +
                    '">' +
                    esc(STATUS_LABEL[m.status] || m.status) +
                    '</span></td>' +
                    '<td class="km-inline-actions">' +
                    (m.status === STATUS.NEU
                        ? '<button type="button" data-km-kenntnis="' + esc(m.id) + '">Zur Kenntnis</button>'
                        : '') +
                    (m.status !== STATUS.ARCHIVIERT
                        ? '<button type="button" data-km-archiv="' + esc(m.id) + '">Archiv</button>'
                        : '') +
                    '<button type="button" data-km-delete="' +
                    esc(m.id) +
                    '">Löschen</button>' +
                    '</td></tr>'
                );
            })
            .join('');

        const acc = getKrankmeldungAccount();
        const kvKlassen = kvScopedKlassen();
        const isAdm = admin();
        const role = getKrankmeldungRoleCache();
        const dateLabel =
            ui.filters.mode === 'woche'
                ? 'Woche von'
                : ui.filters.mode === 'alle'
                  ? 'Datum (unbenutzt)'
                  : 'Tag';

        return (
            actionBanner(stats) +
            '<div class="km-head"><div><h2>Inbox</h2>' +
            '<p>Tag / Woche / Alle – Suche, Klasse, Status. Gefiltert: <strong>' +
            list.length +
            '</strong> von ' +
            stats.total +
            '.</p></div>' +
            '<div class="km-form__actions">' +
            (isAdm
                ? '<button type="button" class="btn btn-success" data-km-pull-spo>Von SharePoint</button>' +
                  '<button type="button" class="btn" data-km-pull-remote>Von API laden</button>'
                : '') +
            '<button type="button" class="btn" id="kmExportCsv">CSV</button>' +
            '<a class="btn" href="' +
            esc(buildPublicFormUrl()) +
            '" target="_blank" rel="noopener">Formular</a>' +
            (isAdm
                ? '<button type="button" class="btn" id="kmCopyFormLink">Link kopieren</button>'
                : '') +
            '</div></div>' +
            '<div class="km-cards">' +
            card('Neu gesamt', String(stats.neu)) +
            card('Heute', String(stats.heute)) +
            card('Heute neu', String(stats.offenHeute)) +
            card('Gefiltert', String(list.length)) +
            '</div>' +
            '<div class="km-scope-bar"><span>Angemeldet' +
            (acc.email ? ': <strong>' + esc(acc.name || acc.email) + '</strong>' : '') +
            (isAdm
                ? ' · Sekretariat/Admin' +
                  (role.demo ? ' (Demo)' : role.sources.length ? ' (' + esc(role.sources.join(', ')) + ')' : '')
                : kvKlassen.length
                  ? ' · KV ' + esc(kvKlassen.join(', '))
                  : ' · keine Inbox-Berechtigung') +
            '</span>' +
            (isAdm && kvKlassen.length
                ? '<label class="km-scope-toggle"><input type="checkbox" id="kmNurMeine"' +
                  (ui.nurMeineKlassen ? ' checked' : '') +
                  '> Nur meine Klassen</label>'
                : '') +
            '</div>' +
            '<div class="km-filters">' +
            '<label>Ansicht<select id="kmMode">' +
            '<option value="tag"' +
            (ui.filters.mode === 'tag' ? ' selected' : '') +
            '>Tag</option>' +
            '<option value="woche"' +
            (ui.filters.mode === 'woche' ? ' selected' : '') +
            '>Woche</option>' +
            '<option value="alle"' +
            (ui.filters.mode === 'alle' ? ' selected' : '') +
            '>Alle</option></select></label>' +
            '<label>' +
            esc(dateLabel) +
            '<input type="date" id="kmDay" value="' +
            esc(ui.filters.day || today) +
            '"' +
            (ui.filters.mode === 'alle' ? ' disabled' : '') +
            '></label>' +
            '<label>Suche<input type="search" id="kmQ" placeholder="Name / Notiz" value="' +
            esc(ui.filters.q || '') +
            '"></label>' +
            '<label>Klasse<select id="kmKlasse"><option value="">Alle</option>' +
            klassen
                .map(
                    (k) =>
                        '<option value="' +
                        esc(k) +
                        '"' +
                        (ui.filters.klasse === k ? ' selected' : '') +
                        '>' +
                        esc(k) +
                        '</option>'
                )
                .join('') +
            '</select></label>' +
            '<label>Status<select id="kmStatus"><option value="">Alle</option>' +
            Object.keys(STATUS_LABEL)
                .map(
                    (k) =>
                        '<option value="' +
                        esc(k) +
                        '"' +
                        (ui.filters.status === k ? ' selected' : '') +
                        '>' +
                        esc(STATUS_LABEL[k]) +
                        '</option>'
                )
                .join('') +
            '</select></label>' +
            '<label>Kategorie<select id="kmKategorie"><option value="">Alle</option>' +
            Object.keys(KATEGORIE_LABEL)
                .map(
                    (k) =>
                        '<option value="' +
                        esc(k) +
                        '"' +
                        (ui.filters.kategorie === k ? ' selected' : '') +
                        '>' +
                        esc(KATEGORIE_LABEL[k]) +
                        '</option>'
                )
                .join('') +
            '</select></label>' +
            '<button type="button" class="btn" id="kmFilterApply">Filtern</button>' +
            '</div>' +
            (rows
                ? '<div class="km-table-wrap"><table class="km-table"><thead><tr><th>Schüler:in</th><th>Zeitraum</th><th>Gemeldet</th><th>Status</th><th></th></tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty(
                      'Keine Meldungen für diesen Filter',
                      stats.total
                          ? 'Andere Filter wählen oder alle anzeigen.'
                          : 'Noch nichts eingegangen – Formular-Link teilen' +
                                (canUseDemoActions() ? ' oder Demo laden.' : '.'),
                      ui.filters.mode !== 'alle'
                          ? '<button type="button" class="btn" id="kmShowAll">Alle anzeigen</button>'
                          : ''
                  ))
        );
    }

    function renderNeu() {
        const t = today;
        const klassen = resolveAdminKlassen(state().meldungen);
        const classField = klassen.length
            ? '<div><span class="km-field-label">Klasse aus Stammdaten</span>' +
              '<div class="km-class-pick" id="kmManualClassPick">' +
              klassen
                  .map(
                      (code) =>
                          '<button type="button" class="km-class-chip" data-km-manual-klasse="' +
                          esc(code) +
                          '">' +
                          esc(code) +
                          '</button>'
                  )
                  .join('') +
              '</div>' +
              '<label>Klasse<input name="klasseCode" id="kmManualKlasse" required placeholder="3AHW"></label></div>'
            : '<label>Klasse<input name="klasseCode" id="kmManualKlasse" required placeholder="3AHW"></label>' +
              '<p class="km-hint">Keine Klassen in Stammdaten – Code manuell eintragen oder unter Einstellungen veröffentlichen.</p>';

        return (
            '<div class="km-head"><div><h2>Manuell melden</h2><p>Für Sekretariat / Telefonauskunft – Klasse wie im Elternformular wählen.</p></div></div>' +
            '<form class="km-form" id="kmManualForm">' +
            '<label>Schüler:in<input name="schuelerName" required></label>' +
            classField +
            '<div class="km-form__row">' +
            '<label>Von<input type="date" name="von" required value="' +
            esc(t) +
            '"></label>' +
            '<label>Bis<input type="date" name="bis" required value="' +
            esc(t) +
            '"></label></div>' +
            '<label>Kategorie<select name="kategorie">' +
            Object.keys(KATEGORIE_LABEL)
                .map((k) => '<option value="' + esc(k) + '">' + esc(KATEGORIE_LABEL[k]) + '</option>')
                .join('') +
            '</select></label>' +
            '<label>Kontakt / Name<input name="gemeldetVonName"></label>' +
            '<label>E-Mail<input type="email" name="gemeldetVonEmail"></label>' +
            '<label>Anhang (optional)<input type="file" name="anhang" accept=".pdf,.png,.jpg,.jpeg,.gif,.webp,.doc,.docx,application/pdf,image/*"><span class="km-hint">Max. ' +
            Math.round(UPLOAD_MAX_BYTES / (1024 * 1024)) +
            ' MB</span></label>' +
            '<label>Notiz<textarea name="notiz" rows="3"></textarea></label>' +
            '<div class="km-form__actions"><button type="submit" class="btn btn-success">Speichern</button></div>' +
            '</form>'
        );
    }

    function renderEinstellungen() {
        if (!admin()) {
            return (
                '<div class="km-head"><div><h2>Einstellungen</h2><p>Nur Sekretariat / Schulleitung / Schul-IT (Entra).</p></div></div>'
            );
        }
        const settings = loadKrankmeldungSettings();
        const s = state();
        const sdCount = loadStammdatenKlassen().length;
        const role = getKrankmeldungRoleCache();
        const publicUrl = (() => {
            try {
                return new URL(buildPublicFormUrl(), window.location.href).href;
            } catch {
                return buildPublicFormUrl();
            }
        })();
        return (
            '<div class="km-head"><div><h2>Einstellungen</h2><p>Klassen für Elternformular, Schulname, optional API. Rolle: Sekretariat/Admin' +
            (role.demo ? ' (Demo)' : '') +
            '.</p></div></div>' +
            '<form class="km-form" id="kmSettingsForm">' +
            '<label>Schulname (auf dem Formular)<input name="schoolName" value="' +
            esc(settings.schoolName) +
            '"></label>' +
            '<label>Klassen für öffentliches Formular (Semikolon)<textarea name="publicKlassen" id="kmPublicKlassen" rows="3" placeholder="1A;1B;2A;3AHW">' +
            esc(settings.publicKlassen) +
            '</textarea></label>' +
            '<p class="km-hint">' +
            esc(stammdatenSourceHint()) +
            ' · ' +
            sdCount +
            ' sichtbar. Eltern sehen Klassen nur, wenn sie hier gespeichert und der Link geteilt wird (oder über <code>?klassen=</code>).</p>' +
            '<div class="km-form__actions" style="margin-top:-4px;">' +
            '<button type="button" class="btn" id="kmSyncKlassen">Aus Stammdaten übernehmen</button>' +
            '<button type="button" class="btn" id="kmCopyFormLink">Formular-Link kopieren</button>' +
            '</div>' +
            '<label class="km-hint" style="display:flex;gap:8px;align-items:center;font-weight:500;">' +
            '<input type="checkbox" name="allowFreeKlasse"' +
            (settings.allowFreeKlasse !== false ? ' checked' : '') +
            '> Eltern dürfen abweichende Klasse eintippen</label>' +
            '<label class="km-hint" style="display:flex;gap:8px;align-items:center;font-weight:500;">' +
            '<input type="checkbox" name="allowAnhang"' +
            (settings.allowAnhang !== false ? ' checked' : '') +
            '> Anhang-Upload auf dem Formular erlauben</label>' +
            '<label class="km-hint" style="display:flex;gap:8px;align-items:center;font-weight:500;">' +
            '<input type="checkbox" name="anhangPflicht"' +
            (settings.anhangPflicht ? ' checked' : '') +
            '> Anhang ist Pflicht (z. B. ärztliche Bestätigung)</label>' +
            '<label>Öffentlicher Link (Vorschau)<input readonly value="' +
            esc(publicUrl) +
            '"></label>' +
            '<label>KV-Hinweis-E-Mail (optional)<input name="kvHintEmail" value="' +
            esc(settings.kvHintEmail) +
            '" placeholder="kv@schule.at"></label>' +
            '<p class="km-hint">Remote-API ist abgeschaltet (Sicherheit / Workshop). Öffentliche Meldungen bleiben lokal bzw. landen über Sekretariat in der Schul-SharePoint-Site.</p>' +
            '<input type="hidden" name="apiBaseUrl" value="">' +
            '<input type="hidden" name="adminKey" value="">' +
            '<div class="km-form__actions"><button type="submit" class="btn btn-success">Speichern</button></div>' +
            '</form>' +
            (() => {
                const sync = getKrankmeldungSyncMeta();
                const site = resolveKrankmeldungSiteUrl() || sync.siteUrl || '– nicht gesetzt –';
                return (
                    '<div class="km-form km-system-panel" style="margin-top:14px;">' +
                    '<h3 style="margin:0 0 8px;font-size:1rem;">SharePoint (Live)</h3>' +
                    '<p class="km-hint"><code>' +
                    esc(site) +
                    '</code></p>' +
                    '<p class="km-settings-summary">Sync: <strong>' +
                    esc(sync.source) +
                    '</strong>' +
                    (sync.listTitle ? ' · Liste „' + esc(sync.listTitle) + '“' : '') +
                    (sync.lastPullAt ? ' · zuletzt ' + esc(formatDeDateTime(sync.lastPullAt)) : '') +
                    (sync.connected ? ' · verbunden' : ' · nicht verbunden') +
                    (sync.error ? ' · ' + esc(sync.error) : '') +
                    '</p>' +
                    '<p class="km-settings-summary">Lokal: <strong>' +
                    (s.meldungen || []).length +
                    '</strong> Meldungen' +
                    (s.localDemoOnly ? ' · Demo local-first' : '') +
                    ' · mit spoItemId: ' +
                    (s.meldungen || []).filter((m) => m.spoItemId).length +
                    (krankmeldungRemoteConfigured() ? ' · Remote konfiguriert' : '') +
                    '</p>' +
                    '<p class="km-hint" style="margin-top:10px;">SoR = SharePoint. Demo nur Admin oder <code>?demo=1</code> / <code>?demoRole=1</code> – Seeds werden nicht gepusht. Public ohne Login: lokal/Remote; Sekretariat pusht.</p>' +
                    '<div class="km-form__actions">' +
                    '<button type="button" class="btn btn-success" data-km-pull-spo>Von SharePoint laden</button>' +
                    (canUseDemoActions()
                        ? '<button type="button" class="btn" data-km-demo>Demo laden</button>' +
                          '<button type="button" class="btn" data-km-clear-demo>Demo entfernen</button>'
                        : '') +
                    '<button type="button" class="btn" data-km-clear>Alles löschen</button>' +
                    '<a class="btn" href="sharepoint-liste-krankmeldung.html" data-ms365-frontend-it-only>Liste einrichten</a>' +
                    '</div></div>'
                );
            })()
        );
    }

    function bind() {
        root.querySelectorAll('[data-km-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = btn.getAttribute('data-km-view') || 'inbox';
                setMessage('', '');
                render();
            });
        });

        function readFiltersFromDom() {
            ui.filters.mode = String((root.querySelector('#kmMode') || {}).value || 'tag');
            ui.filters.day = String((root.querySelector('#kmDay') || {}).value || today);
            ui.filters.klasse = String((root.querySelector('#kmKlasse') || {}).value || '');
            ui.filters.status = String((root.querySelector('#kmStatus') || {}).value || '');
            ui.filters.kategorie = String((root.querySelector('#kmKategorie') || {}).value || '');
            ui.filters.q = String((root.querySelector('#kmQ') || {}).value || '');
        }

        const apply = root.querySelector('#kmFilterApply');
        if (apply) {
            apply.addEventListener('click', () => {
                readFiltersFromDom();
                render();
            });
        }
        ;['#kmMode', '#kmKlasse', '#kmStatus', '#kmKategorie', '#kmDay'].forEach((sel) => {
            root.querySelector(sel)?.addEventListener('change', () => {
                readFiltersFromDom();
                render();
            });
        });
        root.querySelector('#kmQ')?.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                readFiltersFromDom();
                render();
            }
        });
        root.querySelector('#kmShowAll')?.addEventListener('click', () => {
            ui.filters.mode = 'alle';
            render();
        });
        root.querySelector('#kmNurMeine')?.addEventListener('change', (ev) => {
            ui.nurMeineKlassen = !!ev.target.checked;
            render();
        });
        root.querySelector('#kmExportCsv')?.addEventListener('click', () => {
            const list = filteredList(state());
            const csv = meldungenToCsv(list);
            const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'krankmeldungen.csv';
            a.click();
            URL.revokeObjectURL(a.href);
            setMessage('ok', list.length + ' Zeile(n) als CSV exportiert.');
            render();
        });

        root.querySelectorAll('[data-km-kenntnis]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const acc = getKrankmeldungAccount();
                const actor =
                    (acc && acc.email) ||
                    loadKrankmeldungSettings().kvHintEmail ||
                    'sekretariat';
                const r = markKenntnis(btn.getAttribute('data-km-kenntnis'), actor);
                if (!r.ok) {
                    setMessage('err', (r.errors && r.errors[0]) || 'Fehler');
                    render();
                    return;
                }
                void afterLocalWrite(r.item, 'Zur Kenntnis – in SharePoint gespeichert.');
            });
        });

        root.querySelectorAll('[data-km-archiv]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const r = setMeldungStatus(btn.getAttribute('data-km-archiv'), STATUS.ARCHIVIERT);
                if (!r.ok) {
                    setMessage('err', (r.errors && r.errors[0]) || 'Fehler');
                    render();
                    return;
                }
                void afterLocalWrite(r.item, 'Archiviert – in SharePoint gespeichert.');
            });
        });

        root.querySelectorAll('[data-km-delete]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!window.confirm('Meldung löschen?')) return;
                const id = btn.getAttribute('data-km-delete') || '';
                const item = (state().meldungen || []).find((m) => m.id === id);
                removeMeldung(id);
                void pushKrankmeldungDelete(item || {}).then((del) => {
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

        root.querySelectorAll('[data-km-manual-klasse]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const code = String(btn.getAttribute('data-km-manual-klasse') || '');
                const input = root.querySelector('#kmManualKlasse');
                if (input) input.value = code;
                root.querySelectorAll('[data-km-manual-klasse]').forEach((b) => {
                    b.classList.toggle('is-on', b === btn);
                });
            });
        });

        root.querySelectorAll('[data-km-anhang]').forEach((btn) => {
            btn.addEventListener('click', async () => {
                const id = String(btn.getAttribute('data-km-anhang') || '');
                const m = (state().meldungen || []).find((x) => x.id === id);
                if (!m) return;
                const file = await loadAttachment(m.anhangKey || m.id);
                if (!file || !file.dataUrl) {
                    setMessage(
                        'warn',
                        'Anhang hier nicht gefunden (anderer Browser/Gerät) – SharePoint-URL folgt später.'
                    );
                    render();
                    return;
                }
                const a = document.createElement('a');
                a.href = file.dataUrl;
                a.download = file.name || m.anhangName || 'anhang';
                document.body.appendChild(a);
                a.click();
                a.remove();
            });
        });

        const form = root.querySelector('#kmManualForm');
        if (form) {
            form.addEventListener('submit', async (ev) => {
                ev.preventDefault();
                const fd = new FormData(form);
                const id = newId('km');
                /** @type {object} */
                const draft = {
                    id,
                    schuelerName: String(fd.get('schuelerName') || ''),
                    klasseCode: String(fd.get('klasseCode') || ''),
                    von: String(fd.get('von') || ''),
                    bis: String(fd.get('bis') || ''),
                    kategorie: String(fd.get('kategorie') || 'krankheit'),
                    gemeldetVonName: String(fd.get('gemeldetVonName') || ''),
                    gemeldetVonEmail: String(fd.get('gemeldetVonEmail') || ''),
                    notiz: String(fd.get('notiz') || ''),
                    status: STATUS.NEU,
                    quelle: 'ui'
                };
                const fileInput = form.querySelector('[name="anhang"]');
                const file =
                    fileInput && fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
                if (file) {
                    const v = validateUploadFile(file);
                    if (!v.ok) {
                        setMessage('err', v.error || 'Anhang ungültig.');
                        render();
                        return;
                    }
                    try {
                        const dataUrl = await fileToDataUrl(file);
                        await saveAttachment(id, {
                            name: file.name,
                            mime: file.type || 'application/octet-stream',
                            size: file.size,
                            dataUrl
                        });
                        draft.anhangKey = id;
                        draft.anhangName = file.name;
                        draft.anhangMime = file.type || '';
                        draft.anhangSize = file.size;
                    } catch {
                        setMessage('err', 'Anhang speichern fehlgeschlagen.');
                        render();
                        return;
                    }
                }
                const r = upsertMeldung(draft);
                if (!r.ok) {
                    setMessage(
                        'err',
                        (r.errors || []).join(' ') +
                            (r.duplicate ? ' (ähnliche Meldung existiert bereits)' : '')
                    );
                    render();
                    return;
                }
                ui.view = 'inbox';
                const baseMsg =
                    (r.warnings && r.warnings[0]) || 'Meldung gespeichert.';
                void afterLocalWrite(r.item, baseMsg);
            });
        }

        const syncKlassen = root.querySelector('#kmSyncKlassen');
        if (syncKlassen) {
            syncKlassen.addEventListener('click', () => {
                const text = stammdatenKlassenAsText();
                if (!text) {
                    setMessage('warn', 'Keine Klassen in den Stammdaten gefunden.');
                    render();
                    return;
                }
                const ta = root.querySelector('#kmPublicKlassen');
                if (ta) ta.value = text;
                saveKrankmeldungSettings({ publicKlassen: text });
                setMessage(
                    'ok',
                    'Klassen aus Stammdaten übernommen und gespeichert – Link enthält jetzt die Auswahl.'
                );
                render();
            });
        }

        async function copyFormLink() {
            const ta = root.querySelector('#kmPublicKlassen');
            const klassenText = ta
                ? String(ta.value || '')
                : loadKrankmeldungSettings().publicKlassen;
            const href = (() => {
                try {
                    return new URL(buildPublicFormUrl(undefined, klassenText), window.location.href)
                        .href;
                } catch {
                    return buildPublicFormUrl(undefined, klassenText);
                }
            })();
            try {
                await navigator.clipboard.writeText(href);
                setMessage('ok', 'Formular-Link kopiert (inkl. Klassen, falls hinterlegt).');
            } catch {
                setMessage('warn', 'Kopieren nicht möglich – Link manuell aus den Einstellungen nehmen.');
            }
            render();
        }

        root.querySelectorAll('#kmCopyFormLink').forEach((btn) => {
            btn.addEventListener('click', () => {
                void copyFormLink();
            });
        });

        const settingsForm = root.querySelector('#kmSettingsForm');
        if (settingsForm) {
            settingsForm.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(settingsForm);
                saveKrankmeldungSettings({
                    schoolName: String(fd.get('schoolName') || ''),
                    publicKlassen: String(fd.get('publicKlassen') || ''),
                    allowFreeKlasse: !!fd.get('allowFreeKlasse'),
                    allowAnhang: !!fd.get('allowAnhang'),
                    anhangPflicht: !!fd.get('anhangPflicht'),
                    kvHintEmail: String(fd.get('kvHintEmail') || ''),
                    // w2: Remote-API weder zeigen noch aktivieren
                    apiBaseUrl: '',
                    adminKey: ''
                });
                setMessage(
                    'ok',
                    'Einstellungen gespeichert. Eltern-Link enthält die veröffentlichten Klassen.'
                );
                render();
            });
        }

        root.querySelectorAll('[data-km-pull-spo]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (!admin()) return;
                setMessage('ok', 'Lade von SharePoint …');
                render();
                void pullKrankmeldungFromSharePoint({ forceDemoIfEmpty: false }).then((r) => {
                    setMessage(
                        r.ok ? 'ok' : 'warn',
                        r.ok
                            ? 'SharePoint geladen (' +
                                  ((r.state && r.state.meldungen) || []).length +
                                  ' Meldungen).'
                            : r.softError || (r.meta && r.meta.error) || 'Pull fehlgeschlagen.'
                    );
                    render();
                });
            });
        });

        const pull = root.querySelector('[data-km-pull-remote]');
        if (pull) {
            pull.addEventListener('click', async () => {
                if (!krankmeldungRemoteConfigured()) {
                    setMessage('warn', 'Keine API-URL gesetzt – lokale Inbox wird verwendet.');
                    render();
                    return;
                }
                try {
                    const data = await remoteListMeldungen();
                    const list = data.meldungen || [];
                    let n = 0;
                    /** @type {object[]} */
                    const imported = [];
                    for (const m of list) {
                        const r = upsertMeldung(
                            { ...m, quelle: m.quelle || 'public' },
                            { allowDuplicate: true }
                        );
                        if (r.ok && r.item) {
                            n++;
                            imported.push(r.item);
                        }
                    }
                    let pushed = 0;
                    for (const item of imported) {
                        const p = await pushKrankmeldungMeldung(item);
                        if (p.ok && !p.skipped) pushed++;
                    }
                    setMessage(
                        'ok',
                        n +
                            ' Meldung(en) von API geladen' +
                            (pushed ? ', ' + pushed + ' nach SharePoint.' : '.')
                    );
                } catch (e) {
                    setMessage('warn', softRemoteMessage(e));
                }
                render();
            });
        }

        const demo = root.querySelector('[data-km-demo]');
        if (demo) {
            demo.addEventListener('click', () => {
                if (!canUseDemoActions()) {
                    setMessage('warn', 'Demo nur für Admin oder mit ?demo=1 / ?demoRole=1.');
                    render();
                    return;
                }
                loadDemoData();
                saveKrankmeldungSettings({ apiBaseUrl: '', adminKey: '' });
                setMessage('ok', 'Demo geladen (local-first, ohne SPO-Push der Seeds, Remote-API aus).');
                ui.view = 'inbox';
                render();
            });
        }

        const clearDemo = root.querySelector('[data-km-clear-demo]');
        if (clearDemo) {
            clearDemo.addEventListener('click', () => {
                if (!canUseDemoActions()) {
                    setMessage('warn', 'Demo nur für Admin oder mit ?demo=1 / ?demoRole=1.');
                    render();
                    return;
                }
                clearDemoSeeds();
                setMessage('ok', 'Demo-Daten entfernt.');
                render();
            });
        }

        const clear = root.querySelector('[data-km-clear]');
        if (clear) {
            clear.addEventListener('click', () => {
                if (!admin()) {
                    setMessage('err', 'Nur Admins dürfen den gesamten Store leeren.');
                    render();
                    return;
                }
                if (!window.confirm('Alle lokalen Krankmeldungen löschen?')) return;
                clearAll();
                setMessage('ok', 'Geleert.');
                render();
            });
        }
    }

    render();
    void resolveKrankmeldungRoles(getKrankmeldungAccount().email).then(() => {
        if (!admin()) {
            ui.nurMeineKlassen = true;
            if (ui.view === 'einstellungen') ui.view = 'inbox';
        }
        render();
    });
    void pullKrankmeldungFromSharePoint().then((r) => {
        if (r.softError) setMessage('warn', r.softError);
        else if (r.ok) setMessage('ok', 'Mit SharePoint verbunden.');
        render();
    });
    void hydrateEinsammlerStammdaten({ siteKind: 'krankmeldung' })
        .then(() => {
            render();
        })
        .catch(() => {
            /* local-first */
        });
}
