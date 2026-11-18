/**
 * Admin/SLG-UI Wahlfach-Zuteilung (wahl-portal) – Calm, Phasen, Soft-403.
 */

import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import {
    countBelegung,
    countUeberbuchung,
    ergebnisListenProAngebot,
    isAnmeldefensterOffen,
    kapazitaetStatus,
    resolveWzActionHint,
    resolveWzAdminActor,
    resolveWzPhase
} from './wahlfach-zuteilung-logic.js';
import { downloadErgebnisCsv } from './wahlfach-zuteilung-export.js';
import { isWzDemoSeedState } from './wahlfach-zuteilung-demo.js';
import {
    getWzSyncMeta,
    softPullWz,
    softPushWzAngebot,
    softPushWzAngebotDelete,
    softPushWzAnmeldung,
    softPushWzRunde,
    softPushWzZuteilung
} from './wahlfach-zuteilung-sync.js';
import {
    ANMELDUNG_STATUS_LABEL,
    RUNDE_STATUS,
    RUNDE_STATUS_LABEL
} from './wahlfach-zuteilung-model.js';
import { loadSettings, saveSettings } from './wahlfach-zuteilung-settings.js';
import {
    applyZuteilung,
    hasLocalWzData,
    loadDemoData,
    loadState,
    removeAngebot,
    seedDemoIfEmpty,
    upsertAnmeldung,
    upsertAngebot,
    upsertRunde
} from './wahlfach-zuteilung-store.js';
import { resolveWahlfachAdminAccess } from './wahlfach-entra-role.js';
import {
    classCodesFromStammdaten,
    createWzUiState,
    ensureWzView,
    loadWzStammdaten,
    navSectionsForWz,
    readWzAccount,
    resolveWzSiteUrl,
    showWzSystemChrome,
    wantsWzDemoFromUrl
} from './wahlfach-zuteilung-state.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function angebotLabel(state, id) {
    const a = (state.angebote || []).find((x) => x.id === id);
    return a ? a.bezeichnung : id || '–';
}

function formatDeDate(iso) {
    const s = String(iso || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s || '–';
    const [y, m, d] = s.split('-');
    return d + '.' + m + '.' + y;
}

function phaseBadgeHtml(runde) {
    const p = resolveWzPhase(runde);
    const cls =
        p.tone === 'ok'
            ? 'wz-badge--ok'
            : p.tone === 'warn'
              ? 'wz-badge--warn'
              : p.tone === 'muted'
                ? ''
                : 'wz-badge--info';
    return '<span class="wz-badge ' + cls + '">' + esc(p.label) + '</span>';
}

function quietEmpty(text) {
    return (
        '<p class="wz-quiet-empty" role="status"><i class="bi bi-check-circle" aria-hidden="true"></i><span>' +
        esc(text) +
        '</span></p>'
    );
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    const ui = createWzUiState();

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    function state() {
        return loadState();
    }

    function showSystem() {
        return showWzSystemChrome() || (ui.access && ui.access.demo);
    }

    async function hydrateStammdaten(siteId) {
        ui.stammdaten = loadWzStammdaten();
        if (!siteId || ui.localDemoOnly) {
            ui.stammdatenLive = {
                ok: !!ui.localDemoOnly,
                error: '',
                classCount: (ui.stammdaten.classes || []).length,
                teacherCount: (ui.stammdaten.teachers || []).length,
                source: ui.localDemoOnly ? 'demo' : 'cache'
            };
            return;
        }
        const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
        if (!api || typeof api.getGraphToken !== 'function') {
            ui.stammdatenLive = {
                ok: false,
                error: 'Graph nicht geladen',
                classCount: (ui.stammdaten.classes || []).length,
                teacherCount: (ui.stammdaten.teachers || []).length,
                source: 'cache'
            };
            return;
        }
        try {
            const tok = await api.getGraphToken([
                'https://graph.microsoft.com/User.Read',
                'https://graph.microsoft.com/Sites.Read.All'
            ]);
            const result = await hydrateStammdatenFromLiveSite({
                token: tok,
                siteId: String(siteId),
                graph: api,
                local: loadWzStammdaten()
            });
            ui.stammdaten = result.stammdaten;
            ui.stammdatenLive = {
                ok: true,
                error: '',
                classCount: (result.stammdaten.classes || []).length,
                teacherCount: (result.stammdaten.teachers || []).length,
                source: 'live'
            };
        } catch (e) {
            ui.stammdatenLive = {
                ok: false,
                error: (e && e.message) || String(e),
                classCount: (ui.stammdaten.classes || []).length,
                teacherCount: (ui.stammdaten.teachers || []).length,
                source: 'cache'
            };
            /* lokaler Stand bleibt */
        }
    }

    function applySoftPushResult(push) {
        if (!push || push.ok || push.skipped) return;
        if (push.softError) ui.softError = push.softError;
    }

    async function softSyncFromSpo() {
        ui.siteUrl = resolveWzSiteUrl();
        ui.softError = '';
        ui.hardError = '';
        if (ui.localDemoOnly) {
            ui.softError = 'Demo lokal aktiv – SharePoint wird nicht überschrieben.';
            return;
        }
        ui.loading = true;
        render();
        const pull = await softPullWz({
            localDemoOnly: ui.localDemoOnly,
            siteUrl: ui.siteUrl
        });
        if (pull.ok && pull.state) {
            ui.siteId = pull.siteId || '';
            ui.spoSyncedAt = new Date().toISOString();
            ui.localDemoOnly = false;
            await hydrateStammdaten(pull.siteId || '');
        } else {
            ui.softError = pull.softError || '';
            ui.hardError = pull.error || '';
            await hydrateStammdaten('');
        }
        ui.loading = false;
    }

    function enableDemo(confirmOverwrite) {
        const local = state();
        if (hasLocalWzData(local) && !isWzDemoSeedState(local) && confirmOverwrite !== false) {
            if (
                !window.confirm(
                    'Demo-Daten laden? Bestehende lokale Runde/Anmeldungen werden überschrieben.'
                )
            ) {
                return false;
            }
        }
        loadDemoData();
        ui.localDemoOnly = true;
        ui.softError = '';
        ui.hardError = '';
        setMessage('ok', 'Demo lokal geladen (15 Anmeldungen, manuelle Zuteilung).');
        return true;
    }

    function render() {
        if (ui.access && !ui.access.ok) {
            root.innerHTML =
                '<div class="wz-shell"><div class="wz-main">' +
                '<h2>Wahlfach-Zuteilung</h2>' +
                '<p class="wz-msg wz-msg--err" role="alert">' +
                esc(ui.access.reason) +
                '</p>' +
                '<p class="wz-nav__hint">Schüler:innen nutzen das <a href="wahlfach-anmeldung.html">Anmeldeformular</a>. IT-Vorschau: ?demoRole=1</p>' +
                '</div></div>';
            return;
        }
        ui.view = ensureWzView(ui.view, { showSystem: showSystem() });
        const s = state();
        root.innerHTML =
            '<div class="wz-shell">' +
            renderNav(s) +
            '<div class="wz-main">' +
            renderTop(s) +
            renderSoftBars() +
            renderMessage() +
            (ui.loading ? '<p class="wz-nav__hint">Lade …</p>' : '') +
            renderView(s) +
            '</div></div>';
        bind(s);
    }

    function renderNav(s) {
        const sections = navSectionsForWz({ showSystem: showSystem() });
        const buttons = sections
            .map((sec) => {
                const head = sec.label
                    ? '<p class="wz-nav__section">' + esc(sec.label) + '</p>'
                    : '';
                const items = (sec.items || [])
                    .map(
                        (v) =>
                            '<button type="button" class="wz-nav__btn' +
                            (ui.view === v.id ? ' is-active' : '') +
                            '" data-wz-view="' +
                            esc(v.id) +
                            '"><i class="bi ' +
                            esc(v.icon) +
                            '" aria-hidden="true"></i>' +
                            esc(v.label) +
                            '</button>'
                    )
                    .join('');
                return head + items;
            })
            .join('');
        const accessHint =
            ui.access && ui.access.demo
                ? 'IT-Vorschau (?demoRole=1).'
                : ui.access && ui.access.reason
                  ? ui.access.reason
                  : 'Schulleitung / Verwaltung – manuelle Zuteilung.';
        const phase = resolveWzPhase(s.runde);
        return (
            '<nav class="wz-nav" aria-label="Wahlfach-Zuteilung">' +
            '<div class="wz-nav__brand">' +
            '<i class="bi bi-diagram-3" aria-hidden="true"></i>' +
            '<div><strong>Wahlfach-Zuteilung</strong><span>' +
            esc(phase.label) +
            '</span></div>' +
            '</div>' +
            '<div class="wz-nav__list">' +
            buttons +
            '</div>' +
            '<div class="wz-nav__role">' +
            '<span class="wz-nav__role-label">Zugang</span>' +
            '<p class="wz-nav__hint">' +
            esc(accessHint) +
            '</p>' +
            '<div class="wz-nav__links">' +
            '<a href="wahlfach-anmeldung.html"><i class="bi bi-pencil-square"></i> Schüler-Formular</a>' +
            (showSystem()
                ? '<a href="sharepoint-liste-wahlfach.html"><i class="bi bi-list-ul"></i> Listen (IT)</a>'
                : '') +
            '</div></div></nav>'
        );
    }

    function renderTop(s) {
        return (
            '<header class="wz-top">' +
            '<div><h2 class="wz-top__title">Wahlfach-Zuteilung</h2>' +
            '<p class="wz-top__hint">Runde → Angebote → Wünsche → manuelle Zuteilung → CSV</p></div>' +
            '<div class="wz-top__meta">' +
            phaseBadgeHtml(s.runde) +
            (ui.localDemoOnly ? '<span class="wz-badge wz-badge--warn">Demo lokal</span>' : '') +
            (ui.accountEmail
                ? '<small>' + esc(ui.accountName || ui.accountEmail) + '</small>'
                : '<small>Bitte anmelden</small>') +
            '</div></header>'
        );
    }

    function renderSoftBars() {
        let html = '';
        if (ui.softError) {
            html +=
                '<div class="wz-alert wz-alert--info" role="status">' + esc(ui.softError) + '</div>';
        }
        if (ui.hardError && !(ui.localDemoOnly && /403|Forbidden|Zugriff/i.test(ui.hardError))) {
            html +=
                '<div class="wz-alert wz-alert--err" role="alert">' + esc(ui.hardError) + '</div>';
        }
        return html;
    }

    function renderMessage() {
        if (!ui.message) return '';
        const cls =
            ui.message.type === 'ok'
                ? 'wz-msg--ok'
                : ui.message.type === 'warn'
                  ? 'wz-msg--warn'
                  : 'wz-msg--err';
        return '<p class="wz-msg ' + cls + '" role="status">' + esc(ui.message.text) + '</p>';
    }

    function renderActionBanner(s) {
        const hint = resolveWzActionHint(s);
        if (!hint) {
            return quietEmpty('Alles ruhig – keine offenen Zuteilungsaufgaben.');
        }
        return (
            '<section class="wz-calm-banner wz-calm-banner--action" role="status">' +
            '<div class="wz-calm-banner__text">' +
            '<p class="wz-calm-banner__eyebrow">' +
            esc(hint.eyebrow) +
            '</p>' +
            '<p>' +
            esc(hint.text) +
            '</p></div>' +
            '<button type="button" class="btn btn-success" data-wz-view="' +
            esc(hint.ctaView) +
            '">' +
            esc(hint.ctaLabel) +
            '</button></section>'
        );
    }

    function renderView(s) {
        switch (ui.view) {
            case 'angebote':
                return renderAngebote(s);
            case 'anmeldungen':
                return renderAnmeldungen(s);
            case 'zuteilung':
                return renderZuteilung(s);
            case 'ergebnis':
                return renderErgebnis(s);
            case 'system':
                return renderSystem();
            default:
                return renderUebersicht(s);
        }
    }

    function renderUebersicht(s) {
        const r = s.runde;
        const offen = isAnmeldefensterOffen(r);
        const ueber = countUeberbuchung(s.angebote, s.zuteilungen);
        const zugeteilt = (s.zuteilungen || []).filter((z) => z.angebotId).length;
        return (
            '<section class="wz-panel">' +
            '<h2 class="wz-panel__title">Übersicht</h2>' +
            '<p class="wz-panel__lead">Phasen, Kapazität und Zuteilungsstand. SharePoint ist System of Record – lokal bleibt der Arbeitsstand.</p>' +
            renderActionBanner(s) +
            '<div class="wz-kpis">' +
            kpi(s.angebote.length, 'Angebote') +
            kpi(s.anmeldungen.length, 'Anmeldungen') +
            kpi(zugeteilt, 'Zugeteilt') +
            kpi(ueber, 'Überbuchungen') +
            '</div>' +
            (r
                ? '<div class="wz-card"><h3>' +
                  esc(r.titel) +
                  '</h3><p>Schuljahr ' +
                  esc(r.schuljahr) +
                  ' · ' +
                  formatDeDate(r.anmeldeVon) +
                  ' – ' +
                  formatDeDate(r.anmeldeBis) +
                  '</p><p>' +
                  phaseBadgeHtml(r) +
                  ' <span class="wz-badge">' +
                  esc(RUNDE_STATUS_LABEL[r.status] || r.status) +
                  '</span> ' +
                  (offen
                      ? '<span class="wz-badge wz-badge--ok">Fenster offen</span>'
                      : '<span class="wz-badge">Fenster zu</span>') +
                  '</p></div>'
                : quietEmpty('Noch keine Runde – unten anlegen.')) +
            '<h3 class="wz-panel__title" style="margin-top:14px;font-size:1rem">Runde bearbeiten</h3>' +
            rundeForm(r) +
            '</section>'
        );
    }

    function kpi(n, label) {
        return (
            '<div class="wz-kpi"><strong>' +
            esc(String(n)) +
            '</strong><span>' +
            esc(label) +
            '</span></div>'
        );
    }

    function rundeForm(r) {
        const statusOpts = RUNDE_STATUS.map(
            (st) =>
                '<option value="' +
                esc(st) +
                '"' +
                ((r && r.status) === st ? ' selected' : '') +
                '>' +
                esc(RUNDE_STATUS_LABEL[st] || st) +
                '</option>'
        ).join('');
        return (
            '<form class="wz-form" id="wzRundeForm">' +
            '<div class="wz-field"><label for="wzSchuljahr">Schuljahr</label>' +
            '<input id="wzSchuljahr" name="schuljahr" required value="' +
            esc((r && r.schuljahr) || '') +
            '" placeholder="2026/27"></div>' +
            '<div class="wz-field"><label for="wzTitel">Titel</label>' +
            '<input id="wzTitel" name="titel" required value="' +
            esc((r && r.titel) || '') +
            '"></div>' +
            '<div class="wz-form wz-form--inline" style="margin:0">' +
            '<div class="wz-field"><label for="wzVon">Anmelde von</label>' +
            '<input type="date" id="wzVon" name="anmeldeVon" required value="' +
            esc((r && r.anmeldeVon) || '') +
            '"></div>' +
            '<div class="wz-field"><label for="wzBis">Anmelde bis</label>' +
            '<input type="date" id="wzBis" name="anmeldeBis" required value="' +
            esc((r && r.anmeldeBis) || '') +
            '"></div>' +
            '<div class="wz-field"><label for="wzStatus">Status / Phase</label>' +
            '<select id="wzStatus" name="status">' +
            statusOpts +
            '</select></div></div>' +
            '<div class="wz-actions"><button type="submit" class="btn">Runde speichern</button></div>' +
            '</form>'
        );
    }

    function renderAngebote(s) {
        const rows = (s.angebote || [])
            .map((a) => {
                const bel = countBelegung(a.id, s.zuteilungen);
                const st = kapazitaetStatus(a, bel);
                const badge =
                    st.ueberbuchung > 0
                        ? 'wz-badge--bad'
                        : st.voll
                          ? 'wz-badge--warn'
                          : 'wz-badge--ok';
                return (
                    '<tr><td>' +
                    esc(a.bezeichnung) +
                    '</td><td>' +
                    esc(String(a.kapazitaetMin)) +
                    '–' +
                    esc(String(a.kapazitaetMax)) +
                    '</td><td>' +
                    esc(a.zielJahrgaenge || '–') +
                    '</td><td><span class="wz-badge ' +
                    badge +
                    '">' +
                    bel +
                    '/' +
                    a.kapazitaetMax +
                    '</span></td><td>' +
                    (a.aktiv ? 'ja' : 'nein') +
                    '</td><td><button type="button" class="btn" data-wz-del-ang="' +
                    esc(a.id) +
                    '">Löschen</button></td></tr>'
                );
            })
            .join('');
        return (
            '<section class="wz-panel">' +
            '<h2 class="wz-panel__title">Angebote</h2>' +
            '<p class="wz-panel__lead">Wahlfächer mit Kapazität (Min/Max) und Zieljahrgängen – Grundlage für manuelle Zuteilung.</p>' +
            '<form class="wz-form wz-form--inline" id="wzAngebotForm">' +
            '<div class="wz-field"><label for="wzAngBez">Bezeichnung</label>' +
            '<input id="wzAngBez" name="bezeichnung" required placeholder="z. B. Chor"></div>' +
            '<div class="wz-field"><label for="wzAngMin">Min.</label>' +
            '<input type="number" id="wzAngMin" name="kapazitaetMin" min="0" value="8"></div>' +
            '<div class="wz-field"><label for="wzAngMax">Max.</label>' +
            '<input type="number" id="wzAngMax" name="kapazitaetMax" min="1" value="20" required></div>' +
            '<div class="wz-field"><label for="wzAngZiel">Jahrgänge</label>' +
            '<input id="wzAngZiel" name="zielJahrgaenge" placeholder="3–4"></div>' +
            '<div class="wz-actions"><button type="submit" class="btn">Anlegen</button></div>' +
            '</form>' +
            (rows
                ? '<div class="wz-table-wrap"><table class="wz-table"><thead><tr>' +
                  '<th>Bezeichnung</th><th>Kapazität</th><th>Jahrgänge</th><th>Belegt</th><th>Aktiv</th><th></th>' +
                  '</tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty('Noch keine Angebote.')) +
            '</section>'
        );
    }

    function angebotOptions(s, selected, emptyLabel) {
        const opts = ['<option value="">' + esc(emptyLabel || '– wählen –') + '</option>'];
        (s.angebote || [])
            .filter((a) => a.aktiv)
            .forEach((a) => {
                opts.push(
                    '<option value="' +
                        esc(a.id) +
                        '"' +
                        (selected === a.id ? ' selected' : '') +
                        '>' +
                        esc(a.bezeichnung) +
                        '</option>'
                );
            });
        return opts.join('');
    }

    function klasseOptions(selected) {
        const codes = classCodesFromStammdaten(ui.stammdaten);
        if (!codes.length) {
            return (
                '<input id="wzKlasse" name="klasse" required placeholder="3AK" value="' +
                esc(selected || '') +
                '">'
            );
        }
        const opts = ['<option value="">– Klasse –</option>'].concat(
            codes.map(
                (c) =>
                    '<option value="' +
                    esc(c) +
                    '"' +
                    (selected === c ? ' selected' : '') +
                    '>' +
                    esc(c) +
                    '</option>'
            )
        );
        return (
            '<select id="wzKlasse" name="klasse" required>' + opts.join('') + '</select>'
        );
    }

    function renderAnmeldungen(s) {
        const rows = (s.anmeldungen || [])
            .map((a) => {
                const z = (s.zuteilungen || []).find((x) => x.anmeldungId === a.id);
                return (
                    '<tr><td>' +
                    esc(a.schuelerName) +
                    (a.schuelerEmail
                        ? '<br><small>' + esc(a.schuelerEmail) + '</small>'
                        : '') +
                    '</td><td>' +
                    esc(a.klasse) +
                    '</td><td>' +
                    esc(angebotLabel(s, a.wunsch1Id)) +
                    ' / ' +
                    esc(angebotLabel(s, a.wunsch2Id)) +
                    (a.wunsch3Id ? ' / ' + esc(angebotLabel(s, a.wunsch3Id)) : '') +
                    '</td><td>' +
                    esc(ANMELDUNG_STATUS_LABEL[a.status] || a.status) +
                    '</td><td>' +
                    esc(z && z.angebotId ? angebotLabel(s, z.angebotId) : '–') +
                    '</td></tr>'
                );
            })
            .join('');
        return (
            '<section class="wz-panel">' +
            '<h2 class="wz-panel__title">Anmeldungen</h2>' +
            '<p class="wz-panel__lead">Schüler:innen melden sich über das Anmeldeformular (1.–3. Wunsch). Optional Demo-Erfassung hier.</p>' +
            '<p class="wz-actions" style="margin-bottom:12px"><a class="btn btn-success" href="wahlfach-anmeldung.html"><i class="bi bi-box-arrow-up-right"></i>Schüler-Formular</a></p>' +
            '<form class="wz-form" id="wzAnmeldungForm">' +
            '<div class="wz-form wz-form--inline" style="margin:0">' +
            '<div class="wz-field"><label for="wzName">Name</label>' +
            '<input id="wzName" name="schuelerName" required placeholder="Max Mustermann"></div>' +
            '<div class="wz-field"><label for="wzKlasse">Klasse</label>' +
            klasseOptions('') +
            '</div></div>' +
            '<div class="wz-form wz-form--inline" style="margin:0">' +
            '<div class="wz-field"><label for="wzW1">1. Wunsch</label>' +
            '<select id="wzW1" name="wunsch1Id" required>' +
            angebotOptions(s, '', '– 1. Wunsch –') +
            '</select></div>' +
            '<div class="wz-field"><label for="wzW2">2. Wunsch</label>' +
            '<select id="wzW2" name="wunsch2Id">' +
            angebotOptions(s, '', '– optional –') +
            '</select></div>' +
            '<div class="wz-field"><label for="wzW3">3. Wunsch</label>' +
            '<select id="wzW3" name="wunsch3Id">' +
            angebotOptions(s, '', '– optional –') +
            '</select></div></div>' +
            '<div class="wz-actions"><button type="submit" class="btn">Anmeldung speichern</button></div>' +
            '</form>' +
            (rows
                ? '<div class="wz-table-wrap"><table class="wz-table"><thead><tr>' +
                  '<th>Schüler:in</th><th>Klasse</th><th>Wünsche</th><th>Status</th><th>Zuteilung</th>' +
                  '</tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty('Noch keine Anmeldungen.')) +
            '</section>'
        );
    }

    function renderZuteilung(s) {
        const offen = (s.anmeldungen || []).filter((a) => a.status !== 'storniert');
        if (!offen.length) {
            return (
                '<section class="wz-panel"><h2 class="wz-panel__title">Zuteilung</h2>' +
                quietEmpty('Zuerst Anmeldungen erfassen.') +
                '</section>'
            );
        }
        const rows = offen
            .map((a) => {
                const z = (s.zuteilungen || []).find((x) => x.anmeldungId === a.id);
                const current = (z && z.angebotId) || '';
                return (
                    '<tr><td>' +
                    esc(a.schuelerName) +
                    '<br><small>' +
                    esc(a.klasse) +
                    ' · Wünsche: ' +
                    esc(angebotLabel(s, a.wunsch1Id)) +
                    ', ' +
                    esc(angebotLabel(s, a.wunsch2Id)) +
                    (a.wunsch3Id ? ', ' + esc(angebotLabel(s, a.wunsch3Id)) : '') +
                    '</small></td><td><select data-wz-zuteil-select="' +
                    esc(a.id) +
                    '">' +
                    angebotOptions(s, current, '– nicht zugeteilt –') +
                    '</select></td><td><button type="button" class="btn" data-wz-zuteil="' +
                    esc(a.id) +
                    '">Zuweisen</button> ' +
                    '<button type="button" class="btn" data-wz-zuteil-clear="' +
                    esc(a.id) +
                    '">Aufheben</button></td></tr>'
                );
            })
            .join('');
        const ueber = countUeberbuchung(s.angebote, s.zuteilungen);
        return (
            '<section class="wz-panel">' +
            '<h2 class="wz-panel__title">Manuelle Zuteilung</h2>' +
            '<p class="wz-panel__lead">Kapazität wird geprüft. Kein Auto-Matching, keine Genehmigungslogik – nur manuelle Zuweisung.</p>' +
            (ueber
                ? '<p class="wz-msg wz-msg--warn">Aktuell ' +
                  ueber +
                  ' Überbuchung(en) über alle Angebote.</p>'
                : '') +
            '<div class="wz-table-wrap"><table class="wz-table"><thead><tr>' +
            '<th>Anmeldung</th><th>Angebot</th><th></th></tr></thead><tbody>' +
            rows +
            '</tbody></table></div></section>'
        );
    }

    function renderErgebnis(s) {
        const listen = ergebnisListenProAngebot(s.angebote, s.anmeldungen, s.zuteilungen);
        if (!listen.length) {
            return (
                '<section class="wz-panel"><h2 class="wz-panel__title">Ergebnis</h2>' +
                quietEmpty('Keine Angebote.') +
                '</section>'
            );
        }
        const cards = listen
            .map((L) => {
                const st = L.status;
                const badge =
                    st.ueberbuchung > 0
                        ? 'wz-badge--bad'
                        : !st.kommtZustande
                          ? 'wz-badge--warn'
                          : 'wz-badge--ok';
                const people = L.teilnehmende.length
                    ? '<ul>' +
                      L.teilnehmende
                          .map(
                              (t) =>
                                  '<li>' +
                                  esc(t.anmeldung.schuelerName) +
                                  ' (' +
                                  esc(t.anmeldung.klasse) +
                                  ')</li>'
                          )
                          .join('') +
                      '</ul>'
                    : '<p class="wz-empty">Noch niemand zugeteilt.</p>';
                return (
                    '<article class="wz-card"><h3>' +
                    esc(L.angebot.bezeichnung) +
                    ' <span class="wz-badge ' +
                    badge +
                    '">' +
                    L.belegt +
                    '/' +
                    L.angebot.kapazitaetMax +
                    '</span></h3>' +
                    (st.warnung ? '<p>' + esc(st.warnung) + '</p>' : '<p>Min. ' + st.min + '</p>') +
                    people +
                    '</article>'
                );
            })
            .join('');
        return (
            '<section class="wz-panel">' +
            '<h2 class="wz-panel__title">Ergebnis</h2>' +
            '<p class="wz-panel__lead">Teilnehmerlisten pro Angebot. CSV (UTF-8 BOM, Semikolon) für Access-Schülerakte.</p>' +
            '<div class="wz-actions" style="margin-bottom:14px">' +
            '<button type="button" class="btn btn-success" id="wzBtnExportCsv"><i class="bi bi-filetype-csv"></i>Ergebnis als CSV</button>' +
            '</div>' +
            '<div class="wz-card-list">' +
            cards +
            '</div></section>'
        );
    }

    function renderSystem() {
        const cfg = loadSettings();
        const live = ui.stammdatenLive || {};
        const meta = getWzSyncMeta();
        const syncNote =
            'Quelle ' +
            (meta.source || 'none') +
            (meta.lastPullAt ? ' · Pull ' + String(meta.lastPullAt).slice(0, 19).replace('T', ' ') : '') +
            (meta.lastPushAt ? ' · Push ' + String(meta.lastPushAt).slice(0, 19).replace('T', ' ') : '') +
            (ui.spoSyncedAt && !meta.lastPullAt
                ? ' · Sync ' + String(ui.spoSyncedAt).slice(0, 19).replace('T', ' ')
                : '');
        return (
            '<section class="wz-panel wz-panel--system">' +
            '<h2 class="wz-panel__title">System</h2>' +
            '<p class="wz-panel__lead">Site, Listen und Demo – nur IT / bewusste System-Shell. SharePoint ist System of Record (Pull + Push).</p>' +
            '<form class="wz-form" id="wzSettingsForm" style="max-width:640px">' +
            '<div class="wz-field"><label for="wzSiteUrl">SharePoint-Site-URL (kind wahlfach)</label>' +
            '<input type="url" id="wzSiteUrl" name="siteUrl" value="' +
            esc(cfg.siteUrl || ui.siteUrl || '') +
            '" placeholder="https://schule.sharepoint.com/sites/Intranet"></div>' +
            '<div class="wz-form wz-form--inline" style="margin:0">' +
            '<div class="wz-field"><label for="wzListRunden">Liste Runden</label>' +
            '<input id="wzListRunden" name="listRunden" value="' +
            esc(cfg.listRunden) +
            '"></div>' +
            '<div class="wz-field"><label for="wzListAngebote">Liste Angebote</label>' +
            '<input id="wzListAngebote" name="listAngebote" value="' +
            esc(cfg.listAngebote) +
            '"></div></div>' +
            '<div class="wz-form wz-form--inline" style="margin:0">' +
            '<div class="wz-field"><label for="wzListAnmeldungen">Liste Anmeldungen</label>' +
            '<input id="wzListAnmeldungen" name="listAnmeldungen" value="' +
            esc(cfg.listAnmeldungen) +
            '"></div>' +
            '<div class="wz-field"><label for="wzListZuteilungen">Liste Zuteilungen</label>' +
            '<input id="wzListZuteilungen" name="listZuteilungen" value="' +
            esc(cfg.listZuteilungen) +
            '"></div></div>' +
            '<div class="wz-field"><label for="wzAccessHinweis">Hinweis Access / Schülerakte</label>' +
            '<textarea id="wzAccessHinweis" name="accessHinweis" rows="3">' +
            esc(cfg.accessHinweis) +
            '</textarea></div>' +
            '<div class="wz-actions">' +
            '<button type="submit" class="btn btn-success">Einstellungen speichern</button>' +
            '<button type="button" class="btn" id="wzBtnSpoSync"><i class="bi bi-cloud-download"></i>Von SharePoint laden</button>' +
            '<a class="btn" href="sharepoint-liste-wahlfach.html"><i class="bi bi-list-ul"></i>Listen anlegen</a>' +
            '</div></form>' +
            '<div class="wz-settings-note">' +
            '<strong>SPO Live Sync:</strong> ' +
            esc(syncNote) +
            '</div>' +
            '<div class="wz-settings-note">' +
            '<strong>Stammdaten:</strong> ' +
            esc(String(live.classCount || 0)) +
            ' Klassen · ' +
            esc(String(live.teacherCount || 0)) +
            ' Lehrkräfte · Quelle ' +
            esc(live.source || 'cache') +
            (live.error ? ' · ' + esc(live.error) : '') +
            '</div>' +
            '<div class="wz-actions" style="margin-top:14px">' +
            '<button type="button" class="btn" id="wzBtnDemoLoad"><i class="bi bi-database"></i>Demo lokal laden</button>' +
            '<span class="wz-nav__hint">Oder URL: ?demo=1 – Demo wird nicht nach SharePoint gepusht</span>' +
            '</div></section>'
        );
    }

    function adminActor() {
        const acc = readWzAccount();
        ui.accountEmail = acc.email || ui.accountEmail;
        ui.accountName = acc.name || ui.accountName;
        return resolveWzAdminActor({
            accountEmail: ui.accountEmail,
            accountName: ui.accountName
        });
    }

    function bind(s) {
        root.querySelectorAll('[data-wz-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.view = ensureWzView(btn.getAttribute('data-wz-view') || 'uebersicht', {
                    showSystem: showSystem()
                });
                setMessage(null, null);
                render();
            });
        });

        const rundeFormEl = root.querySelector('#wzRundeForm');
        if (rundeFormEl) {
            rundeFormEl.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(rundeFormEl);
                const res = upsertRunde({
                    id: s.runde && s.runde.id,
                    spoItemId: s.runde && s.runde.spoItemId,
                    schuljahr: fd.get('schuljahr'),
                    titel: fd.get('titel'),
                    anmeldeVon: fd.get('anmeldeVon'),
                    anmeldeBis: fd.get('anmeldeBis'),
                    status: fd.get('status')
                });
                if (!res.ok) {
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                setMessage('ok', 'Runde gespeichert.');
                render();
                void softPushWzRunde(res.state.runde, { localDemoOnly: ui.localDemoOnly }).then(
                    (push) => {
                        applySoftPushResult(push);
                        if (push.ok && !push.skipped) setMessage('ok', 'Runde in SharePoint gespeichert.');
                        render();
                    }
                );
            });
        }

        const angForm = root.querySelector('#wzAngebotForm');
        if (angForm) {
            angForm.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(angForm);
                const res = upsertAngebot({
                    bezeichnung: fd.get('bezeichnung'),
                    kapazitaetMin: Number(fd.get('kapazitaetMin')),
                    kapazitaetMax: Number(fd.get('kapazitaetMax')),
                    zielJahrgaenge: fd.get('zielJahrgaenge'),
                    aktiv: true
                });
                if (!res.ok) {
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                setMessage('ok', 'Angebot angelegt.');
                render();
                void softPushWzAngebot(res.angebot, { localDemoOnly: ui.localDemoOnly }).then(
                    (push) => {
                        applySoftPushResult(push);
                        if (push.ok && !push.skipped) setMessage('ok', 'Angebot in SharePoint gespeichert.');
                        render();
                    }
                );
            });
        }

        root.querySelectorAll('[data-wz-del-ang]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-wz-del-ang');
                const prev = (s.angebote || []).find((a) => a && a.id === id) || null;
                removeAngebot(id);
                setMessage('ok', 'Angebot entfernt.');
                render();
                void softPushWzAngebotDelete(prev, { localDemoOnly: ui.localDemoOnly }).then(
                    (push) => {
                        applySoftPushResult(push);
                        render();
                    }
                );
            });
        });

        const anmForm = root.querySelector('#wzAnmeldungForm');
        if (anmForm) {
            anmForm.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(anmForm);
                const res = upsertAnmeldung({
                    schuelerName: fd.get('schuelerName'),
                    klasse: fd.get('klasse'),
                    wunsch1Id: fd.get('wunsch1Id'),
                    wunsch2Id: fd.get('wunsch2Id'),
                    wunsch3Id: fd.get('wunsch3Id'),
                    status: 'eingereicht',
                    schuelerEmail: ui.accountEmail || undefined
                });
                if (!res.ok) {
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                setMessage('ok', 'Anmeldung gespeichert.');
                render();
                void softPushWzAnmeldung(res.anmeldung, { localDemoOnly: ui.localDemoOnly }).then(
                    (push) => {
                        applySoftPushResult(push);
                        if (push.ok && !push.skipped) {
                            setMessage('ok', 'Anmeldung in SharePoint gespeichert.');
                        }
                        render();
                    }
                );
            });
        }

        root.querySelectorAll('[data-wz-zuteil]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-wz-zuteil');
                const sel = root.querySelector('[data-wz-zuteil-select="' + id + '"]');
                const angebotId = sel ? sel.value : '';
                if (!angebotId) {
                    setMessage('err', 'Bitte ein Angebot wählen.');
                    render();
                    return;
                }
                const prevZ =
                    (s.zuteilungen || []).find((z) => z && z.anmeldungId === id) || null;
                const res = applyZuteilung(id, angebotId, {
                    zugeteiltVon: adminActor(),
                    begruendung: 'Manuelle Zuteilung'
                });
                if (!res.ok) {
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                if (res.warnings.length) setMessage('warn', res.warnings.join(' '));
                else setMessage('ok', 'Zuteilung gespeichert.');
                render();
                const z =
                    (res.state.zuteilungen || []).find((row) => row && row.anmeldungId === id) ||
                    null;
                void softPushWzZuteilung(
                    { anmeldungId: id, angebotId, prevZuteilung: prevZ, zuteilung: z },
                    { localDemoOnly: ui.localDemoOnly }
                ).then((push) => {
                    applySoftPushResult(push);
                    if (push.ok && !push.skipped) setMessage('ok', 'Zuteilung in SharePoint gespeichert.');
                    render();
                });
            });
        });

        root.querySelectorAll('[data-wz-zuteil-clear]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-wz-zuteil-clear');
                const prevZ =
                    (s.zuteilungen || []).find((z) => z && z.anmeldungId === id) || null;
                const res = applyZuteilung(id, null, { zugeteiltVon: adminActor() });
                if (!res.ok) {
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                setMessage('ok', 'Zuteilung aufgehoben.');
                render();
                void softPushWzZuteilung(
                    { anmeldungId: id, angebotId: null, prevZuteilung: prevZ },
                    { localDemoOnly: ui.localDemoOnly }
                ).then((push) => {
                    applySoftPushResult(push);
                    render();
                });
            });
        });

        const exportBtn = root.querySelector('#wzBtnExportCsv');
        if (exportBtn) {
            exportBtn.addEventListener('click', () => {
                try {
                    downloadErgebnisCsv(s);
                    setMessage('ok', 'CSV heruntergeladen (Access-Import vorbereiten).');
                    render();
                } catch (e) {
                    setMessage('err', e && e.message ? e.message : String(e));
                    render();
                }
            });
        }

        const settingsForm = root.querySelector('#wzSettingsForm');
        if (settingsForm) {
            settingsForm.addEventListener('submit', (ev) => {
                ev.preventDefault();
                const fd = new FormData(settingsForm);
                saveSettings({
                    siteUrl: fd.get('siteUrl'),
                    listRunden: fd.get('listRunden'),
                    listAngebote: fd.get('listAngebote'),
                    listAnmeldungen: fd.get('listAnmeldungen'),
                    listZuteilungen: fd.get('listZuteilungen'),
                    accessHinweis: fd.get('accessHinweis')
                });
                ui.siteUrl = resolveWzSiteUrl();
                setMessage('ok', 'Einstellungen gespeichert (Site kind wahlfach).');
                render();
            });
        }

        const spoBtn = root.querySelector('#wzBtnSpoSync');
        if (spoBtn) {
            spoBtn.addEventListener('click', () => {
                void softSyncFromSpo().then(() => {
                    if (!ui.softError && !ui.hardError) {
                        setMessage('ok', 'SharePoint geladen.');
                    }
                    render();
                });
            });
        }

        const demoBtn = root.querySelector('#wzBtnDemoLoad');
        if (demoBtn) {
            demoBtn.addEventListener('click', () => {
                if (enableDemo(true)) {
                    ui.view = 'anmeldungen';
                    render();
                }
            });
        }
    }

    root.innerHTML =
        '<div class="wz-shell"><div class="wz-main"><p class="wz-nav__hint">Prüfe Zugang …</p></div></div>';

    void (async () => {
        try {
            ui.access = await resolveWahlfachAdminAccess();
        } catch {
            ui.access = {
                ok: false,
                reason: 'Zugangsprüfung fehlgeschlagen. Bitte neu anmelden.',
                demo: false
            };
        }
        const acc = readWzAccount();
        ui.accountEmail = acc.email;
        ui.accountName = acc.name;

        if (ui.access && ui.access.ok) {
            if (wantsWzDemoFromUrl()) {
                seedDemoIfEmpty();
                ui.localDemoOnly = true;
                setMessage('ok', 'Demo aktiv (?demo=1).');
            } else if (isWzDemoSeedState(state())) {
                ui.localDemoOnly = true;
            } else {
                await softSyncFromSpo();
            }
            if (!ui.stammdatenLive || ui.stammdatenLive.source === 'cache') {
                await hydrateStammdaten(ui.siteId);
            }
        }
        render();
    })();
}

export { mountApp as mountWahlfachZuteilung };
