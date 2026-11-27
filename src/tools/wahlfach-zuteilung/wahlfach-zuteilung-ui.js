/**
 * Admin/SLG-UI Wahlfach-Zuteilung (wahl-portal) – Calm, Phasen, Soft-403.
 */

import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import { bindSystemTabClicks, renderSystemAdminPage } from '../../shared/system-admin-shell.js';
import {
    countBelegung,
    countUeberbuchung,
    ergebnisListenProAngebot,
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
import { resolveWahlfachAppRole } from './wahlfach-entra-role.js';
import {
    classCodesFromStammdaten,
    classResultRows,
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
        p.tone === 'ok' ? 'fb-pill--ok' : p.tone === 'warn' ? 'fb-pill--warn' : '';
    return '<span class="fb-pill ' + cls + '">' + esc(p.label) + '</span>';
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
            ui.softError = 'Änderungen bleiben auf diesem Computer.';
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
        setMessage('ok', 'Beispieldaten geladen. Nichts geht an SharePoint.');
        return true;
    }

    function renderRolePage(inner) {
        root.innerHTML =
            '<div class="wz-shell wz-shell--form"><div class="wz-main">' + inner + '</div></div>';
    }

    function renderLehrerQuiet() {
        renderRolePage(
            '<h1 class="fb-greet">Wahlfach-Zuteilung</h1>' +
                '<p class="fb-sub">Die Zuteilung macht die Direktion. Als Lehrkraft siehst du hier nichts weiter.</p>' +
                renderMessage() +
                '<p class="wz-quiet-empty" role="status"><i class="bi bi-eye-slash" aria-hidden="true"></i><span>Klassenvorstände sehen das Ergebnis ihrer eigenen Klasse. Schülerinnen und Schüler füllen nur ihr Formular aus.</span></p>'
        );
    }

    function renderMeineKlasse(s) {
        const codes = (ui.access && ui.access.kvKlassen) || [];
        const rows = classResultRows(s, codes, (ui.stammdaten && ui.stammdaten.students) || [])
            .map((row) => {
                const z = row.anmeldung
                    ? (s.zuteilungen || []).find((x) => x && x.anmeldungId === row.anmeldung.id)
                    : null;
                return { row, z: z && z.angebotId ? z : null };
            })
            .sort((a, b) => {
                const rank = (item) => (!item.row.anmeldung ? 0 : !item.z ? 1 : 2);
                const byRank = rank(a) - rank(b);
                if (byRank) return byRank;
                return a.row.name.localeCompare(b.row.name, 'de');
            });
        const body = rows.length
            ? rows
                  .map(({ row, z }) => {
                      const a = row.anmeldung;
                      const zugeteilt = z && z.angebotId ? angebotLabel(s, z.angebotId) : '';
                      return (
                          '<tr><td>' +
                          esc(row.name) +
                          (row.email ? '<br><small>' + esc(row.email) + '</small>' : '') +
                          '</td><td>' +
                          esc(row.klasse) +
                          '</td><td>' +
                          (a ? esc(angebotLabel(s, a.wunsch1Id)) : '–') +
                          '</td><td>' +
                          (a && a.wunsch2Id ? esc(angebotLabel(s, a.wunsch2Id)) : '–') +
                          '</td><td>' +
                          (a && a.wunsch3Id ? esc(angebotLabel(s, a.wunsch3Id)) : '–') +
                          '</td><td>' +
                          (a
                              ? esc(zugeteilt || 'noch offen')
                              : 'noch keine Anmeldung') +
                          '</td></tr>'
                      );
                  })
                  .join('')
            : '';
        const assigned = rows.filter((item) => item.z).length;
        const missing = rows.filter((item) => !item.row.anmeldung).length;
        const open = rows.length - assigned - missing;
        /** @type {Map<string, number>} */
        const byOffer = new Map();
        rows.forEach((item) => {
            if (!item.z) return;
            const name = angebotLabel(s, item.z.angebotId);
            byOffer.set(name, (byOffer.get(name) || 0) + 1);
        });
        const offerLine = Array.from(byOffer.entries())
            .map(([name, n]) => esc(name) + ' ' + n)
            .join(' · ');
        renderRolePage(
            pageHead(
                'Meine Klasse',
                (codes.length ? codes.join(', ') : 'Deine Klasse') +
                    ': ' +
                    assigned +
                    ' von ' +
                    rows.length +
                    ' zugeteilt.'
            ) +
            '<p class="wz-kv-facts"><span>' +
            open +
            ' noch offen</span><span>' +
            missing +
            ' ohne Anmeldung</span>' +
            (offerLine ? '<span>' + offerLine + '</span>' : '') +
            '</p>' +
                renderMessage() +
                (ui.loading
                    ? '<p class="fb-sub">Lade …</p>'
                    : !rows.length
                      ? quietEmpty('In dieser Klasse liegt noch keine Anmeldung.')
                      : '<div class="wz-table-wrap"><table class="wz-table"><thead><tr>' +
                        '<th>Schüler/in</th><th>Klasse</th><th>1. Wunsch</th><th>2. Wunsch</th><th>3. Wunsch</th><th>Zuteilung</th>' +
                        '</tr></thead><tbody>' +
                        body +
                        '</tbody></table></div>')
        );
    }

    function render() {
        const role = ui.access && ui.access.role;
        if (!ui.access || role === 'pending') {
            renderRolePage(
                '<h1 class="fb-greet">Wahlfach-Zuteilung</h1><p class="fb-sub">Konto wird gelesen …</p>'
            );
            return;
        }
        if (role === 'lehrer') {
            renderLehrerQuiet();
            return;
        }
        if (role === 'kv') {
            renderMeineKlasse(state());
            return;
        }
        ui.view = ensureWzView(ui.view, { showSystem: showSystem() });
        const s = state();
        root.innerHTML =
            '<div class="wz-shell">' +
            renderNav(s) +
            '<div class="wz-main">' +
            renderSoftBars() +
            renderMessage() +
            (ui.loading ? '<p class="fb-sub">Lade …</p>' : '') +
            renderView(s) +
            '</div></div>';
        bind(s);
    }

    function renderNav(s) {
        const sections = navSectionsForWz({ showSystem: showSystem() });
        const buttons = sections
            .map((sec) => {
                const head = sec.label
                    ? '<span class="wz-nav__section-label">' + esc(sec.label) + '</span>'
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
                            '" aria-hidden="true"></i><span class="wz-nav__btn-label">' +
                            esc(v.label) +
                            '</span></button>'
                    )
                    .join('');
                return '<div class="wz-nav__section">' + head + items + '</div>';
            })
            .join('');
        const accessHint =
            ui.access && ui.access.demo
                ? 'IT-Vorschau (?demoRole=1).'
                : ui.access && ui.access.reason
                  ? ui.access.reason
                  : 'Schulleitung teilt manuell zu.';
        const phase = resolveWzPhase(s.runde);
        return (
            '<aside class="wz-nav" aria-label="Wahlfach-Zuteilung">' +
            '<div class="wz-nav__brand">' +
            '<i class="bi bi-diagram-3" aria-hidden="true"></i>' +
            '<div><strong>Wahlfach-Zuteilung</strong><span>' +
            esc(phase.label) +
            '</span></div>' +
            '</div>' +
            '<nav class="wz-nav__list">' +
            buttons +
            '</nav>' +
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
            '</div></div></aside>'
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

    function renderNextAction(s) {
        const hint = resolveWzActionHint(s);
        if (!hint || hint.ctaView === ui.view) return '';
        return (
            '<div class="fb-actions-row">' +
            '<button type="button" class="fb-btn" data-wz-view="' +
            esc(hint.ctaView) +
            '">' +
            esc(hint.ctaLabel) +
            '</button></div>'
        );
    }

    function pageHead(titleHtml, lead) {
        return (
            '<h1 class="fb-greet">' +
            titleHtml +
            '</h1><p class="fb-sub">' +
            lead +
            '</p>'
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
        const ueber = countUeberbuchung(s.angebote, s.zuteilungen);
        const zugeteilt = (s.zuteilungen || []).filter((z) => z.angebotId).length;
        const hint = resolveWzActionHint(s);
        const headline = !r
            ? 'Noch keine Runde.'
            : '<span class="fb-num">' +
              esc(String(s.anmeldungen.length)) +
              '</span> Anmeldungen.';
        const lead = hint
            ? esc(hint.text)
            : 'Stand der Runde und der Zuteilung.';
        return (
            pageHead(headline, lead) +
            '<div class="wz-meta-row">' +
            (r ? phaseBadgeHtml(r) : '') +
            '</div>' +
            '<div class="wz-kpis">' +
            kpi(s.angebote.length, 'Angebote') +
            kpi(s.anmeldungen.length, 'Anmeldungen') +
            kpi(zugeteilt, 'Zugeteilt') +
            kpi(ueber, 'Überbuchungen') +
            '</div>' +
            renderNextAction(s) +
            (r
                ? '<p class="wz-round-line"><strong>' +
                  esc(r.titel) +
                  '</strong> · Schuljahr ' +
                  esc(r.schuljahr) +
                  ' · ' +
                  formatDeDate(r.anmeldeVon) +
                  ' – ' +
                  formatDeDate(r.anmeldeBis) +
                  '</p>'
                : quietEmpty('Noch keine Runde – unten anlegen.')) +
            '<details class="wz-details"><summary>Runde bearbeiten</summary>' +
            rundeForm(r) +
            '</details>'
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
            '<div class="fb-actions-row"><button type="submit" class="fb-btn">Runde speichern</button></div>' +
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
                        ? 'fb-pill--bad'
                        : st.voll
                          ? 'fb-pill--warn'
                          : 'fb-pill--ok';
                return (
                    '<tr><td>' +
                    esc(a.bezeichnung) +
                    '</td><td>' +
                    esc(String(a.kapazitaetMin)) +
                    '–' +
                    esc(String(a.kapazitaetMax)) +
                    '</td><td>' +
                    esc(a.zielJahrgaenge || '–') +
                    '</td><td><span class="fb-pill ' +
                    badge +
                    '">' +
                    bel +
                    '/' +
                    a.kapazitaetMax +
                    '</span></td><td>' +
                    (a.aktiv ? 'ja' : 'nein') +
                    '</td><td class="wz-row-actions"><button type="button" class="wz-link-btn" data-wz-del-ang="' +
                    esc(a.id) +
                    '">Löschen</button></td></tr>'
                );
            })
            .join('');
        return (
            pageHead(
                'Angebote',
                'Wahlfächer mit Kapazität (Min/Max) und Zieljahrgängen – Grundlage für die manuelle Zuteilung.'
            ) +
            '<form class="wz-form wz-form--inline" id="wzAngebotForm">' +
            '<div class="wz-field"><label for="wzAngBez">Bezeichnung</label>' +
            '<input id="wzAngBez" name="bezeichnung" required placeholder="z. B. Chor"></div>' +
            '<div class="wz-field"><label for="wzAngMin">Min.</label>' +
            '<input type="number" id="wzAngMin" name="kapazitaetMin" min="0" value="8"></div>' +
            '<div class="wz-field"><label for="wzAngMax">Max.</label>' +
            '<input type="number" id="wzAngMax" name="kapazitaetMax" min="1" value="20" required></div>' +
            '<div class="wz-field"><label for="wzAngZiel">Jahrgänge</label>' +
            '<input id="wzAngZiel" name="zielJahrgaenge" placeholder="3–4"></div>' +
            '<div class="wz-field wz-field--action"><button type="submit" class="fb-btn">Anlegen</button></div>' +
            '</form>' +
            (rows
                ? '<div class="wz-table-wrap"><table class="wz-table"><thead><tr>' +
                  '<th>Bezeichnung</th><th>Kapazität</th><th>Jahrgänge</th><th>Belegt</th><th>Aktiv</th><th></th>' +
                  '</tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty('Noch keine Angebote.'))
        );
    }

    function angebotOption(a, selected) {
        return (
            '<option value="' +
            esc(a.id) +
            '"' +
            (selected === a.id ? ' selected' : '') +
            '>' +
            esc(a.bezeichnung) +
            '</option>'
        );
    }

    function angebotOptions(s, selected, emptyLabel, wishIds) {
        const opts = ['<option value="">' + esc(emptyLabel || '– wählen –') + '</option>'];
        const active = (s.angebote || []).filter((a) => a.aktiv);
        const wishSet = new Set((wishIds || []).filter(Boolean));
        const wishes = active.filter((a) => wishSet.has(a.id));
        const rest = active.filter((a) => !wishSet.has(a.id));
        if (wishes.length && rest.length) {
            opts.push('<optgroup label="Wünsche">' + wishes.map((a) => angebotOption(a, selected)).join('') + '</optgroup>');
            opts.push(
                '<optgroup label="Weitere Angebote">' +
                    rest.map((a) => angebotOption(a, selected)).join('') +
                    '</optgroup>'
            );
        } else {
            active.forEach((a) => opts.push(angebotOption(a, selected)));
        }
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
            pageHead(
                'Anmeldungen',
                'Schüler:innen tragen ihre Wünsche im Formular ein. Diese Liste ist die Übersicht für die Zuteilung.'
            ) +
            '<div class="fb-actions-row"><a class="fb-btn fb-btn--ghost" href="wahlfach-anmeldung.html"><i class="bi bi-box-arrow-up-right"></i>Schüler-Formular</a></div>' +
            (rows
                ? '<div class="wz-table-wrap"><table class="wz-table"><thead><tr>' +
                  '<th>Schüler:in</th><th>Klasse</th><th>Wünsche</th><th>Status</th><th>Zuteilung</th>' +
                  '</tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty('Noch keine Anmeldungen.')) +
            '<details class="wz-details"' +
            (ui.captureOpen ? ' open' : '') +
            '><summary>Anmeldung manuell erfassen</summary>' +
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
            '<div class="fb-actions-row"><button type="submit" class="fb-btn">Anmeldung speichern</button></div>' +
            '</form></details>'
        );
    }

    function renderZuteilung(s) {
        const all = (s.anmeldungen || []).filter((a) => a.status !== 'storniert');
        if (!all.length) {
            return (
                pageHead('Zuteilung', 'Zuerst Anmeldungen erfassen.') +
                quietEmpty('Noch keine Anmeldungen.')
            );
        }
        const classSet = [];
        const seenClass = new Set();
        all.forEach((a) => {
            const code = String(a.klasse || '').trim();
            const key = code.toLowerCase();
            if (!code || seenClass.has(key)) return;
            seenClass.add(key);
            classSet.push(code);
        });
        classSet.sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));
        const klasse = String(ui.zuteilKlasse || '');
        const nurOffen = !!ui.zuteilNurOffen;
        const rows = all
            .filter((a) => !klasse || String(a.klasse || '').trim().toLowerCase() === klasse.toLowerCase())
            .filter((a) => {
                if (!nurOffen) return true;
                const z = (s.zuteilungen || []).find((x) => x && x.anmeldungId === a.id);
                return !(z && z.angebotId);
            })
            .map((a) => {
                const z = (s.zuteilungen || []).find((x) => x.anmeldungId === a.id);
                const current = (z && z.angebotId) || '';
                const wishes = [a.wunsch1Id, a.wunsch2Id, a.wunsch3Id].filter(Boolean);
                return (
                    '<tr><td>' +
                    esc(a.schuelerName) +
                    '<br><small>' +
                    esc(a.klasse) +
                    ' · ' +
                    wishes.map((id) => esc(angebotLabel(s, id))).join(' · ') +
                    '</small></td><td><select data-wz-zuteil-select="' +
                    esc(a.id) +
                    '" aria-label="Zuteilung für ' +
                    esc(a.schuelerName) +
                    '">' +
                    angebotOptions(s, current, '– nicht zugeteilt –', wishes) +
                    '</select></td></tr>'
                );
            })
            .join('');
        const classOpts =
            '<option value="">Alle Klassen</option>' +
            classSet
                .map(
                    (c) =>
                        '<option value="' +
                        esc(c) +
                        '"' +
                        (klasse === c ? ' selected' : '') +
                        '>' +
                        esc(c) +
                        '</option>'
                )
                .join('');
        const ueber = countUeberbuchung(s.angebote, s.zuteilungen);
        return (
            pageHead(
                'Manuelle Zuteilung',
                'Angebot wählen. Die Zeile wird sofort gespeichert. „Nicht zugeteilt“ hebt sie auf.'
            ) +
            '<div class="wz-form wz-form--inline">' +
            '<div class="wz-field"><label for="wzZuteilKlasse">Klasse</label>' +
            '<select id="wzZuteilKlasse" data-wz-zuteil-klasse>' +
            classOpts +
            '</select></div>' +
            '<div class="wz-field wz-field--action"><label class="wz-check"><input type="checkbox" data-wz-zuteil-offen' +
            (nurOffen ? ' checked' : '') +
            '> Nur noch offen</label></div></div>' +
            (ueber
                ? '<p class="wz-msg wz-msg--warn">Aktuell ' +
                  ueber +
                  ' Überbuchung(en) über alle Angebote.</p>'
                : '') +
            (rows
                ? '<div class="wz-table-wrap"><table class="wz-table"><thead><tr>' +
                  '<th>Anmeldung</th><th>Angebot</th></tr></thead><tbody>' +
                  rows +
                  '</tbody></table></div>'
                : quietEmpty(nurOffen ? 'In dieser Auswahl ist niemand mehr offen.' : 'Keine Anmeldung in dieser Klasse.'))
        );
    }

    function renderErgebnis(s) {
        const listen = ergebnisListenProAngebot(s.angebote, s.anmeldungen, s.zuteilungen);
        if (!listen.length) {
            return pageHead('Ergebnis', 'Keine Angebote.') + quietEmpty('Zuerst Angebote anlegen.');
        }
        const cards = listen
            .map((L) => {
                const st = L.status;
                const badge =
                    st.ueberbuchung > 0
                        ? 'fb-pill--bad'
                        : !st.kommtZustande
                          ? 'fb-pill--warn'
                          : 'fb-pill--ok';
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
                    ' <span class="fb-pill ' +
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
            pageHead(
                'Ergebnis',
                'Teilnehmerlisten pro Angebot. Die CSV (UTF-8, Semikolon) ist für die Access-Schülerakte gedacht.'
            ) +
            '<div class="fb-actions-row">' +
            '<button type="button" class="fb-btn" id="wzBtnExportCsv"><i class="bi bi-filetype-csv"></i>Ergebnis als CSV</button>' +
            '</div>' +
            '<div class="wz-card-list">' +
            cards +
            '</div>'
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
            '<form class="wz-form" id="wzSettingsForm">' +
            renderSystemAdminPage({
                prefix: 'wz',
                activeTab: ui.systemTab,
                dataAttr: 'data-wz-sys-tab',
                ariaLabel: 'Systemadministration Wahlfach',
                panels: [
                    {
                        id: 'site',
                        step: '1',
                        title: 'Datenquelle',
                        leadHtml: 'SharePoint ist System of Record (Pull + Push). Demo wird nicht gepusht.',
                        body:
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
                            '<div class="wz-settings-note"><strong>SPO Live Sync:</strong> ' +
                            esc(syncNote) +
                            '</div>' +
                            '<div class="fb-actions-row">' +
                            '<button type="submit" class="btn ms-sys-btn-primary">Einstellungen speichern</button>' +
                            '<button type="button" class="btn ms-sys-btn-secondary" id="wzBtnSpoSync"><i class="bi bi-cloud-download"></i>Von SharePoint laden</button>' +
                            '<a class="btn ms-sys-btn-secondary" href="sharepoint-liste-wahlfach.html"><i class="bi bi-list-ul"></i>Listen anlegen</a>' +
                            '</div>'
                    },
                    {
                        id: 'perms',
                        step: '2',
                        title: 'Berechtigungen',
                        leadHtml: 'Listenrechte setzt der Assistent. Diese Ansicht speichert nur die Listennamen.',
                        body:
                            '<div class="ms-sys-links"><a class="btn ms-sys-btn-secondary" href="sharepoint-liste-wahlfach.html"><i class="bi bi-shield-check"></i>Listen-Assistent</a></div>'
                    },
                    {
                        id: 'kalender',
                        step: '3',
                        title: 'Kalender',
                        leadHtml: 'Anmeldefrist und Rundenstatus pflegst du an der Runde, nicht als Outlook-Feed.',
                        body:
                            '<div class="ms-sys-actions"><button type="button" class="btn ms-sys-btn-secondary" data-wz-view="uebersicht"><i class="bi bi-calendar3"></i>Zur Runde</button></div>'
                    },
                    {
                        id: 'wartung',
                        step: '4',
                        title: 'Wartung',
                        leadHtml: 'Stammdaten-Diagnose, Access-Hinweis und lokale Demo.',
                        body:
                            '<div class="wz-settings-note"><strong>Stammdaten:</strong> ' +
                            esc(String(live.classCount || 0)) +
                            ' Klassen · ' +
                            esc(String(live.teacherCount || 0)) +
                            ' Lehrkräfte · Quelle ' +
                            esc(live.source || 'cache') +
                            (live.error ? ' · ' + esc(live.error) : '') +
                            '</div>' +
                            '<div class="wz-field"><label for="wzAccessHinweis">Hinweis Access / Schülerakte</label>' +
                            '<textarea id="wzAccessHinweis" name="accessHinweis" rows="3">' +
                            esc(cfg.accessHinweis) +
                            '</textarea></div>' +
                            '<div class="fb-actions-row">' +
                            '<button type="submit" class="btn ms-sys-btn-primary">Einstellungen speichern</button>' +
                            '<button type="button" class="btn ms-sys-btn-secondary" id="wzBtnDemoLoad"><i class="bi bi-database"></i>Demo lokal laden</button>' +
                            '</div>' +
                            '<p class="wz-nav__hint">Oder URL: ?demo=1 – Demo wird nicht nach SharePoint gepusht</p>'
                    },
                    {
                        id: 'danger',
                        title: 'Danger Zone',
                        danger: true,
                        leadHtml: 'Es gibt hier kein Listen-Leeren. Demo bleibt lokal.',
                        body: '<p class="muted">Produktive Zuteilungen löschst du in der Arbeitsansicht, nicht in diesem Tab.</p>'
                    }
                ]
            }) +
            '</form>'
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
        bindSystemTabClicks(root, {
            dataAttr: 'data-wz-sys-tab',
            onSelect: (tab) => {
                ui.systemTab = tab;
                render();
            }
        });
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
                    ui.captureOpen = true;
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                ui.captureOpen = false;
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

        const klasseSel = root.querySelector('[data-wz-zuteil-klasse]');
        if (klasseSel) {
            klasseSel.addEventListener('change', () => {
                ui.zuteilKlasse = klasseSel.value || '';
                render();
            });
        }
        const offenBox = root.querySelector('[data-wz-zuteil-offen]');
        if (offenBox) {
            offenBox.addEventListener('change', () => {
                ui.zuteilNurOffen = !!offenBox.checked;
                render();
            });
        }

        root.querySelectorAll('[data-wz-zuteil-select]').forEach((sel) => {
            sel.addEventListener('change', () => {
                const id = sel.getAttribute('data-wz-zuteil-select');
                const angebotId = sel.value || '';
                const prevZ =
                    (s.zuteilungen || []).find((z) => z && z.anmeldungId === id) || null;
                const res = applyZuteilung(id, angebotId || null, {
                    zugeteiltVon: adminActor(),
                    begruendung: angebotId ? 'Manuelle Zuteilung' : ''
                });
                if (!res.ok) {
                    setMessage('err', res.errors.join(' '));
                    render();
                    return;
                }
                if (res.warnings.length) setMessage('warn', res.warnings.join(' '));
                else setMessage('ok', angebotId ? 'Zuteilung gespeichert.' : 'Zuteilung aufgehoben.');
                render();
                const z =
                    (res.state.zuteilungen || []).find((row) => row && row.anmeldungId === id) ||
                    null;
                void softPushWzZuteilung(
                    {
                        anmeldungId: id,
                        angebotId: angebotId || null,
                        prevZuteilung: prevZ,
                        zuteilung: z
                    },
                    { localDemoOnly: ui.localDemoOnly }
                ).then((push) => {
                    applySoftPushResult(push);
                    if (push.ok && !push.skipped && angebotId) {
                        setMessage('ok', 'Zuteilung in SharePoint gespeichert.');
                    }
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
        '<div class="wz-shell wz-shell--form"><div class="wz-main"><p class="wz-nav__hint">Prüfe Zugang …</p></div></div>';

    /** @type {{ ok: boolean, reason: string, source: string }|null} */
    let direktionKnown = null;
    let roleRetryWired = false;

    async function settleAccess() {
        ui.stammdaten = loadWzStammdaten();
        try {
            ui.access = await resolveWahlfachAppRole({
                stammdaten: ui.stammdaten,
                direktion: direktionKnown
            });
        } catch {
            ui.access = {
                role: 'lehrer',
                ok: false,
                kvKlassen: [],
                reason: 'Zugangsprüfung fehlgeschlagen. Bitte neu anmelden.',
                source: '',
                demo: false,
                email: ''
            };
            setMessage('err', ui.access.reason);
        }
        if (ui.access && ui.access.role === 'direktion') {
            direktionKnown = {
                ok: true,
                reason: ui.access.reason,
                source: ui.access.source
            };
        } else if (ui.access && !ui.access.demo) {
            direktionKnown = { ok: false, reason: '', source: '' };
        }
        const acc = readWzAccount();
        ui.accountEmail = acc.email;
        ui.accountName = acc.name;

        const role = ui.access && ui.access.role;
        if (role === 'direktion' || role === 'kv') {
            if (wantsWzDemoFromUrl()) {
                seedDemoIfEmpty();
                ui.localDemoOnly = true;
            } else if (isWzDemoSeedState(state())) {
                ui.localDemoOnly = true;
            } else {
                await softSyncFromSpo();
            }
            if (!ui.stammdatenLive || ui.stammdatenLive.source === 'cache') {
                await hydrateStammdaten(ui.siteId);
            }
            if (role === 'kv' && ui.access && !ui.access.demo) {
                ui.access = await resolveWahlfachAppRole({
                    stammdaten: ui.stammdaten,
                    direktion: direktionKnown
                });
            }
        } else if (role === 'pending' && !roleRetryWired && typeof window !== 'undefined') {
            roleRetryWired = true;
            window.addEventListener(
                'ms365-auth-widget-ready',
                () => {
                    void settleAccess();
                },
                { once: true }
            );
        }
        render();
    }

    void settleAccess();
}

export { mountApp as mountWahlfachZuteilung };
