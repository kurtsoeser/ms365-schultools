/**
 * Kalender und Termine – School Calendar Hub (Outlook bleibt führend).
 * Agenda | Woche | Monat · Drawer · IT-Quellen.
 */

import {
    WEEK_GRID_END_HOUR,
    WEEK_GRID_START_HOUR,
    formatTerminRange,
    formatTimeHm,
    fromLocalInputValue,
    groupAgendaByDay,
    layoutTimedEventsForDay,
    listForDay,
    monthGridDays,
    monthLabelDe,
    roundToTimeGrid,
    startOfWeekMonday,
    toLocalInputValue,
    emptyPeriodLabel,
    weekDays,
    weekRangeLabelDe
} from './termine-hub-logic.js';
import {
    SICHTBARKEIT,
    SICHTBARKEIT_LABEL,
    TERMIN_TYP,
    TERMIN_TYP_LABEL,
    toIsoDateOnly
} from './termine-hub-model.js';
import {
    createOutlookEvent,
    deleteOutlookEvent,
    fetchEventsForSources,
    listMyUnifiedGroups,
    probeCalendarSource,
    signedInEmail
} from './termine-hub-outlook.js';
import { getDashboardPersonas } from '../../shared/dashboard-persona-session.js';
import { buildTermineHubDemoState } from './termine-hub-demo.js';
import { announceAppRole } from '../../shared/app-role-chrome.js';
import { isToolVisibleForView } from '../../shared/dashboard-audience-catalog.js';
import {
    getEffectiveToolRule,
    loadDashboardToolAccessConfig
} from '../../shared/dashboard-audience-store.js';
import {
    SOURCE_IDS,
    SOURCE_LABEL,
    audienceKeyFromPersonas,
    ensureSettingsWithSuggestions,
    filterEventsByVisibleSources,
    groupIdFromSourceId,
    isGroupSourceId,
    loadSettings,
    resolveEnabledSources,
    saveSettings,
    setSelectedGroupCalendars,
    sourceWritableForAudience,
    suggestUpnsFromKalenderPostfaecher,
    toggleVisibleSource
} from './termine-hub-settings.js';

const SCHULAKTIVITAETEN_TOOL_ID = 'schulaktivitaeten-planer';
const SCHULAKTIVITAETEN_HREF = 'schulaktivitaeten-planer.html';
const CAL_MODE_KEY = 'ms365-termine-hub-cal-mode-v1';

const SOURCE_HINT = {
    me: 'Persönlicher Outlook-Kalender',
    schul: 'Für schulweit relevante Termine',
    lehrer: 'Für Lehrkräfte relevante Termine',
    schueler: 'Für Schüler:innen relevante Termine'
};

const VIEWS = [
    { id: 'kalender', label: 'Kalender', icon: 'bi-calendar3' },
    { id: 'neu', label: 'Termin anlegen', icon: 'bi-plus-circle' },
    { id: 'abwesenheit', label: 'Abwesenheit', icon: 'bi-person-walking' },
    { id: 'quellen', label: 'Kalenderquellen', icon: 'bi-sliders', itOnly: true }
];

const LEHRER_FREISTELLUNG_HREF = 'lehrer-freistellung-planer.html';
const SCHUELER_FREISTELLUNG_HREF = 'freistellung-planer.html';
const KALENDER_POSTFAECHER_HREF = 'kalender-postfaecher.html';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function loadSavedCalMode() {
    try {
        const raw = String(localStorage.getItem(CAL_MODE_KEY) || '').trim();
        if (raw === 'agenda' || raw === 'week' || raw === 'month') return raw;
    } catch {
        /* ignore */
    }
    if (typeof window !== 'undefined' && window.matchMedia) {
        try {
            if (window.matchMedia('(max-width: 720px)').matches) return 'agenda';
        } catch {
            /* ignore */
        }
    }
    return 'week';
}

function saveCalMode(mode) {
    try {
        localStorage.setItem(CAL_MODE_KEY, mode);
    } catch {
        /* ignore */
    }
}

/**
 * @param {HTMLElement} root
 */
export function mountApp(root) {
    if (!root) return;

    const now = new Date();
    /** @type {{
     *   view: string,
     *   message: {type:string,text:string}|null,
     *   calMode: 'agenda'|'week'|'month',
     *   weekAnchor: Date,
     *   calYear: number,
     *   calMonth: number,
     *   events: object[],
     *   loading: boolean,
     *   lastErrors: string[],
     *   truncated: boolean,
     *   drawerEventId: string,
     *   moreOpen: boolean,
     *           success: null|{ title: string, calendarLabel: string, webLink: string, eventId: string },
     *   probe: Record<string, { status: string, message: string, at?: string }>,
     *   groupPickerOpen: boolean,
     *   groupPickerLoading: boolean,
     *   groupPickerItems: Array<{ id: string, displayName: string, mail: string }>,
     *   groupPickerError: string
     * }} */
    const ui = {
        view: 'kalender',
        message: null,
        calMode: loadSavedCalMode(),
        weekAnchor: startOfWeekMonday(now),
        calYear: now.getFullYear(),
        calMonth: now.getMonth(),
        events: [],
        loading: false,
        lastErrors: [],
        truncated: false,
        drawerEventId: '',
        moreOpen: false,
        success: null,
        probe: {},
        groupPickerOpen: false,
        groupPickerLoading: false,
        groupPickerItems: [],
        groupPickerError: ''
    };

    /** @type {import('../../shared/dashboard-audience-resolve.js').DashboardPersonas} */
    let personas = {
        filtering: false,
        loggedIn: false,
        isIt: false,
        isLehrer: false,
        isSchueler: false,
        schularbeitenRoles: [],
        freistellungRoles: []
    };
    let personasReady = false;

    ensureSettingsWithSuggestions();

    function setMessage(type, text) {
        ui.message = text ? { type, text } : null;
    }

    function canManageQuellen() {
        return !!(personas && personas.isIt);
    }

    function audienceKey() {
        if (!personasReady) {
            // Vor Persona-Resolve: wenn Auth schon eingeloggt ist, Eigenkalender erlauben
            if (isAuthLoggedIn()) return 'self';
            return '';
        }
        const key = audienceKeyFromPersonas(personas);
        if (key) return key;
        if (personas && personas.loggedIn) return 'self';
        if (isAuthLoggedIn()) return 'self';
        return '';
    }

    function isAuthLoggedIn() {
        try {
            if (typeof window.ms365AuthIsLoggedIn === 'function') {
                return !!window.ms365AuthIsLoggedIn();
            }
        } catch {
            /* ignore */
        }
        return !!(personas && personas.loggedIn);
    }

    function settings() {
        return loadSettings();
    }

    function ownCalendarSource() {
        return {
            id: SOURCE_IDS.ME,
            label: SOURCE_LABEL.me,
            kind: 'me',
            upn: '',
            groupId: '',
            visible: true
        };
    }

    function sources() {
        const aud = audienceKey();
        if (!aud) return isAuthLoggedIn() ? [ownCalendarSource()] : [];
        const list = resolveEnabledSources(settings(), { audience: aud });
        if (list.some((s) => s.id === SOURCE_IDS.ME)) return list;
        return [ownCalendarSource(), ...list];
    }

    function writableSources() {
        const aud = audienceKey() || (isAuthLoggedIn() ? 'self' : '');
        const list = sources().filter((s) => sourceWritableForAudience(s.id, aud));
        if (list.length) return list;
        // Harter Fallback: angemeldet → zumindest eigener Outlook-Kalender
        if (aud || isAuthLoggedIn()) return [ownCalendarSource()];
        return [];
    }

    function canWriteSource(sourceId) {
        const aud = audienceKey() || (isAuthLoggedIn() ? 'self' : '');
        if (String(sourceId || '') === SOURCE_IDS.ME && (aud || isAuthLoggedIn())) return true;
        return sourceWritableForAudience(sourceId, aud);
    }

    function visibleEvents() {
        const aud = audienceKey();
        if (!aud) return [];
        return filterEventsByVisibleSources(ui.events, settings(), { audience: aud });
    }

    function writeTargetFromSettings() {
        const s = settings();
        const id = s.defaultWriteTarget || SOURCE_IDS.ME;
        const all = writableSources();
        const hit = all.find((x) => x.id === id);
        if (hit) return hit;
        return all[0] || { id: SOURCE_IDS.ME, label: SOURCE_LABEL.me, kind: 'me', upn: '' };
    }

    function abwesenheitHref() {
        if (personas && personas.isSchueler && !personas.isLehrer && !personas.isIt) {
            return SCHUELER_FREISTELLUNG_HREF;
        }
        return LEHRER_FREISTELLUNG_HREF;
    }

    function catalogViewFromPersonas() {
        if (!personasReady || !personas) return '';
        if (personas.isIt) return 'it';
        if (personas.isLehrer) return 'lehrer';
        if (personas.isSchueler) return 'schueler';
        return '';
    }

    function isToolLicenseAllowed(toolId) {
        const id = String(toolId || '').trim();
        if (!id) return true;
        try {
            if (typeof window.ms365LicenseHasFeature === 'function') {
                return !!window.ms365LicenseHasFeature(id);
            }
            const core = window.ms365LicenseGateCore;
            if (core && typeof core.hasFeature === 'function') {
                return !!core.hasFeature(id);
            }
        } catch {
            /* ignore */
        }
        return true;
    }

    function canAccessSchulaktivitaeten() {
        if (!personasReady || !personas || !personas.loggedIn) return false;
        if (!isToolLicenseAllowed(SCHULAKTIVITAETEN_TOOL_ID)) return false;
        const view = catalogViewFromPersonas();
        if (!view) return false;
        const rule = getEffectiveToolRule(
            SCHULAKTIVITAETEN_TOOL_ID,
            loadDashboardToolAccessConfig()
        );
        return isToolVisibleForView(
            SCHULAKTIVITAETEN_TOOL_ID,
            {
                schularbeitenRoles: personas.schularbeitenRoles || [],
                freistellungRoles: personas.freistellungRoles || []
            },
            view,
            rule
        );
    }

    function rangeForLoad() {
        if (ui.calMode === 'agenda') {
            const from = new Date();
            from.setHours(0, 0, 0, 0);
            const to = new Date(from);
            to.setDate(to.getDate() + 14);
            to.setHours(23, 59, 59, 999);
            return { from, to };
        }
        if (ui.calMode === 'week') {
            const days = weekDays(ui.weekAnchor);
            const from = new Date(days[0]);
            from.setHours(0, 0, 0, 0);
            const to = new Date(days[6]);
            to.setHours(23, 59, 59, 999);
            return { from, to };
        }
        const grid = monthGridDays(ui.calYear, ui.calMonth);
        return {
            from: grid[0],
            to: new Date(grid[41].getFullYear(), grid[41].getMonth(), grid[41].getDate(), 23, 59, 59, 999)
        };
    }

    async function refreshOutlook(opts) {
        const silent = !!(opts && opts.silent);
        const src = sources().filter((s) => s.visible);
        if (!sources().length) {
            ui.events = [];
            ui.lastErrors = [
                canManageQuellen()
                    ? 'Keine Kalender-Quelle aktiv. Unter „Kalenderquellen“ UPNs eintragen oder „Mein Kalender“ aktivieren.'
                    : 'Keine Kalender freigeschaltet. Die IT legt die Kalenderquellen (UPNs) zentral fest.'
            ];
            if (!silent) setMessage('warn', ui.lastErrors[0]);
            render();
            return;
        }
        if (!src.length) {
            ui.events = [];
            ui.lastErrors = [];
            if (!silent) setMessage('warn', 'Mindestens einen Kalender oben einblenden.');
            render();
            return;
        }
        ui.loading = true;
        if (!silent) setMessage(null, '');
        render();
        const { from, to } = rangeForLoad();
        try {
            const res = await fetchEventsForSources(src, from, to, {
                signedInEmail: signedInEmail()
            });
            ui.events = res.events || [];
            ui.lastErrors = res.errors || [];
            ui.truncated = !!res.truncated;
            if (ui.lastErrors.length && !silent) {
                setMessage(
                    'warn',
                    ui.lastErrors.length +
                        ' Kalender teilweise nicht lesbar. Details unter der Ansicht.'
                );
            } else if (!silent) {
                setMessage(
                    'ok',
                    ui.events.length
                        ? ui.events.length + ' Termine aus Outlook geladen.'
                        : 'Keine Termine im gewählten Zeitraum.'
                );
            }
        } catch (e) {
            ui.events = [];
            ui.lastErrors = [e && e.message ? e.message : String(e)];
            if (!silent) setMessage('err', ui.lastErrors[0]);
        } finally {
            ui.loading = false;
            render();
        }
    }

    function render() {
        root.innerHTML =
            '<div class="ag-shell">' +
            renderNav() +
            '<div class="ag-main">' +
            renderMessage() +
            renderView() +
            '</div></div>' +
            renderDrawer();
        bind();
    }

    function renderNav() {
        const userViews = VIEWS.filter((v) => !v.itOnly);
        const itViews = VIEWS.filter((v) => v.itOnly && canManageQuellen());
        const mkBtn = (v) =>
            '<button type="button" class="ag-nav__btn' +
            (ui.view === v.id ? ' is-active' : '') +
            '" data-th-view="' +
            esc(v.id) +
            '"><i class="bi ' +
            esc(v.icon) +
            '" aria-hidden="true"></i>' +
            esc(v.label) +
            '</button>';
        const itExtra = canManageQuellen()
            ? '<div class="ag-nav__links"><a href="' +
              KALENDER_POSTFAECHER_HREF +
              '"><i class="bi bi-mailbox"></i> Kalender-Postfächer</a></div>'
            : '';
        return (
            '<nav class="ag-nav" aria-label="Kalender und Termine">' +
            '<div class="ag-nav__brand">' +
            '<i class="bi bi-calendar-event" aria-hidden="true"></i>' +
            '<div><strong>Kalender und Termine</strong></div>' +
            '</div>' +
            '<div class="ag-nav__list">' +
            userViews.map(mkBtn).join('') +
            '</div>' +
            (itViews.length
                ? '<p class="ag-nav__hint th-nav-section">System</p><div class="ag-nav__list">' +
                  itViews.map(mkBtn).join('') +
                  '</div>' +
                  itExtra
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

    function renderView() {
        if (ui.view === 'neu') return renderNeu();
        if (ui.view === 'success') return renderSuccess();
        if (ui.view === 'quellen') {
            return canManageQuellen()
                ? renderQuellen()
                : '<p class="ag-item__meta">Kalenderquellen sind nur für die IT sichtbar.</p>';
        }
        if (ui.view === 'abwesenheit') return renderAbwesenheit();
        return renderKalender();
    }

    function sourcePill(t) {
        const label = t.sourceLabel || SOURCE_LABEL[t.sourceId] || t.sourceId || 'Outlook';
        return (
            '<span class="ag-pill th-source-pill th-source-pill--' +
            esc(t.sourceId || '') +
            '">' +
            esc(label) +
            '</span>'
        );
    }

    function errorsBlock() {
        if (!canManageQuellen()) return '';
        if (!ui.lastErrors.length && !ui.truncated) return '';
        let html = '<div class="th-errors" role="status"><p class="th-errors__label">Diagnose (IT)</p>';
        if (ui.truncated) {
            html += '<p>Hinweis: Sehr viele Termine – Ansicht ggf. unvollständig.</p>';
        }
        if (ui.lastErrors.length) {
            html +=
                '<ul>' +
                ui.lastErrors.map((e) => '<li>' + esc(e) + '</li>').join('') +
                '</ul>';
        }
        html += '</div>';
        return html;
    }

    function emptyCalendarHint(mode) {
        const src = sources();
        if (!src.length) {
            return canManageQuellen()
                ? '<p class="th-empty">Keine Kalenderquelle aktiv. <button type="button" class="th-textlink" data-th-go="quellen">Kalenderquellen konfigurieren</button></p>'
                : '<p class="th-empty">Kalenderquellen sind noch nicht freigeschaltet. Bitte IT kontaktieren.</p>';
        }
        const hidden = src.filter((s) => !s.visible);
        if (hidden.length === src.length) {
            return '<p class="th-empty">Alle Kalender sind ausgeblendet. Oben mindestens einen Kalender einblenden.</p>';
        }
        return (
            '<p class="th-empty th-empty--period" role="status">' +
            esc(emptyPeriodLabel(mode || ui.calMode)) +
            '</p>'
        );
    }

    function periodHasVisibleEvents(mode) {
        const items = visibleEvents();
        if (!items.length) return false;
        if (mode === 'week') {
            return weekDays(ui.weekAnchor).some(function (d) {
                return listForDay(items, toIsoDateOnly(d)).length > 0;
            });
        }
        if (mode === 'month') {
            return monthGridDays(ui.calYear, ui.calMonth).some(function (d) {
                return d.getMonth() === ui.calMonth && listForDay(items, toIsoDateOnly(d)).length > 0;
            });
        }
        return groupAgendaByDay(items, { days: 14 }).length > 0;
    }

    function emptyOverlayHtml(mode) {
        if (periodHasVisibleEvents(mode)) return '';
        return '<div class="ms365-cal__empty-overlay">' + emptyCalendarHint(mode) + '</div>';
    }

    function layerClassForSource(s) {
        if (s && s.kind === 'group') return 'group';
        return s && s.id ? s.id : 'me';
    }

    function renderGroupPicker() {
        if (!ui.groupPickerOpen) return '';
        const selected = new Set((settings().groupCalendars || []).map((g) => g.id));
        const body = ui.groupPickerLoading
            ? '<p class="ag-item__meta">Gruppen werden geladen…</p>'
            : ui.groupPickerError
              ? '<p class="ag-item__meta">' + esc(ui.groupPickerError) + '</p>'
              : !ui.groupPickerItems.length
                ? '<p class="ag-item__meta">Keine M365-Gruppen gefunden, in denen du Mitglied bist.</p>'
                : '<div class="ms365-group-picker__list">' +
                  ui.groupPickerItems
                      .map(
                          (g) =>
                              '<label class="ms365-group-picker__opt">' +
                              '<input type="checkbox" name="groupCal" value="' +
                              esc(g.id) +
                              '"' +
                              (selected.has(g.id) ? ' checked' : '') +
                              '>' +
                              '<span><strong>' +
                              esc(g.displayName) +
                              '</strong>' +
                              (g.mail
                                  ? '<small>' + esc(g.mail) + '</small>'
                                  : '') +
                              '</span></label>'
                      )
                      .join('') +
                  '</div>';
        return (
            '<div class="ms365-group-picker" role="dialog" aria-label="Gruppenkalender">' +
            '<div class="ms365-group-picker__head">' +
            '<strong>Gruppenkalender</strong>' +
            '<span>M365-Gruppen, in denen du Mitglied bist (ARGE, Lehrergruppe, …)</span>' +
            '</div>' +
            body +
            '<div class="ms365-group-picker__actions">' +
            '<button type="button" class="btn btn-success" data-th-group-save' +
            (ui.groupPickerLoading ? ' disabled' : '') +
            '>Übernehmen</button>' +
            '<button type="button" class="btn" data-th-group-cancel>Schließen</button>' +
            '</div></div>'
        );
    }

    function renderLayerToggles() {
        const src = sources();
        if (!src.length && !personasReady) return '';
        return (
            '<div class="ms365-layers" role="group" aria-label="Kalender ein- und ausblenden">' +
            '<span class="ms365-layers__label">Kalender</span>' +
            src
                .map((s) => {
                    return (
                        '<button type="button" class="ms365-layer ms365-layer--' +
                        esc(layerClassForSource(s)) +
                        (s.visible ? ' is-on' : '') +
                        '" data-ms365-layer="' +
                        esc(s.id) +
                        '" aria-pressed="' +
                        (s.visible ? 'true' : 'false') +
                        '" title="' +
                        esc(s.visible ? 'Ausblenden' : 'Einblenden') +
                        '"><span class="ms365-layer__swatch" aria-hidden="true"></span>' +
                        esc(s.label) +
                        '</button>'
                    );
                })
                .join('') +
            '<button type="button" class="ms365-layer ms365-layer--add" data-th-group-open title="M365-Gruppenkalender hinzufügen">' +
            '<i class="bi bi-people" aria-hidden="true"></i> Gruppen…' +
            '</button>' +
            '</div>' +
            renderGroupPicker()
        );
    }

    function calModeSwitch(mode) {
        const modes = [
            { id: 'agenda', label: 'Agenda' },
            { id: 'week', label: 'Woche' },
            { id: 'month', label: 'Monat' }
        ];
        return (
            '<div class="ms365-cal__seg" role="group" aria-label="Ansicht">' +
            modes
                .map(
                    (m) =>
                        '<button type="button" class="ms365-cal__seg-btn' +
                        (mode === m.id ? ' is-active' : '') +
                        '" data-th-mode="' +
                        m.id +
                        '">' +
                        m.label +
                        '</button>'
                )
                .join('') +
            '</div>'
        );
    }

    function calToolbar(mode, title) {
        const showNav = mode === 'week' || mode === 'month';
        return (
            '<div class="ms365-cal__toolbar">' +
            '<div class="ms365-cal__title-block">' +
            (showNav
                ? '<div class="ms365-cal__nav" role="group" aria-label="Zeitraum">' +
                  '<button type="button" class="ms365-cal__iconbtn" data-th-nav="-1" aria-label="Zurück">‹</button>' +
                  '<button type="button" class="ms365-cal__textbtn ms365-cal__textbtn--today" data-th-nav="0">Heute</button>' +
                  '<button type="button" class="ms365-cal__iconbtn" data-th-nav="1" aria-label="Weiter">›</button>' +
                  '</div>' +
                  '<h2 class="ms365-cal__title">' +
                  esc(title) +
                  '</h2>'
                : '<h2 class="ms365-cal__title">' + esc(title) + '</h2>') +
            '</div>' +
            '<div class="ms365-cal__actions">' +
            calModeSwitch(mode) +
            '<button type="button" class="ms365-cal__iconbtn" data-th-refresh title="Aktualisieren" aria-label="Aktualisieren"' +
            (ui.loading ? ' disabled' : '') +
            '>' +
            (ui.loading
                ? '<span class="ms365-cal__busy" aria-hidden="true">…</span>'
                : '<i class="bi bi-arrow-clockwise" aria-hidden="true"></i>') +
            '</button>' +
            '<button type="button" class="ms365-cal__textbtn ms365-cal__textbtn--primary" data-th-go="neu">Neu</button>' +
            '</div></div>'
        );
    }

    function agendaTimeLabel(t) {
        if (t.allDay) return 'Ganztägig';
        const a = formatTimeHm(t.start);
        const b = formatTimeHm(t.end || t.start);
        return a && b ? a + '–' + b : a || '–';
    }

    function renderAgenda() {
        const groups = groupAgendaByDay(visibleEvents(), { days: 14 });
        if (!groups.length) return emptyCalendarHint();
        return (
            '<div class="ms365-agenda">' +
            groups
                .map((g) => {
                    return (
                        '<section class="ms365-agenda__day">' +
                        '<h3 class="ms365-agenda__heading">' +
                        esc(g.label) +
                        '</h3>' +
                        '<ul class="ms365-agenda__list">' +
                        g.items
                            .map(
                                (t) =>
                                    '<li>' +
                                    '<button type="button" class="ms365-agenda__row ms365-ev--' +
                                    esc(isGroupSourceId(t.sourceId) ? 'group' : t.sourceId || 'me') +
                                    '" data-th-open="' +
                                    esc(t.id) +
                                    '">' +
                                    '<span class="ms365-agenda__time">' +
                                    esc(agendaTimeLabel(t)) +
                                    '</span>' +
                                    '<span class="ms365-agenda__body">' +
                                    '<span class="ms365-agenda__title">' +
                                    esc(t.title) +
                                    '</span>' +
                                    '<span class="ms365-agenda__source">' +
                                    esc(
                                        t.sourceLabel ||
                                            SOURCE_LABEL[t.sourceId] ||
                                            t.sourceId ||
                                            'Outlook'
                                    ) +
                                    '</span>' +
                                    '</span>' +
                                    '</button></li>'
                            )
                            .join('') +
                        '</ul></section>'
                    );
                })
                .join('') +
            '</div>'
        );
    }

    function weekdayShort(day) {
        try {
            return new Intl.DateTimeFormat('de-AT', { weekday: 'short' }).format(day);
        } catch {
            return ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'][day.getDay()] || '';
        }
    }

    function renderWeekView() {
        const days = weekDays(ui.weekAnchor);
        const startH = WEEK_GRID_START_HOUR;
        const endH = WEEK_GRID_END_HOUR;
        const hours = [];
        for (let h = startH; h < endH; h++) hours.push(h);
        const todayIso = toIsoDateOnly(new Date());

        const head =
            '<div class="ms365-week__head">' +
            '<div class="ms365-week__gutter"></div>' +
            '<div class="ms365-week__days">' +
            days
                .map((d) => {
                    const iso = toIsoDateOnly(d);
                    return (
                        '<div class="ms365-week__dayhead' +
                        (iso === todayIso ? ' is-today' : '') +
                        '"><span class="ms365-week__wd">' +
                        esc(weekdayShort(d)) +
                        '</span><span class="ms365-week__dn">' +
                        d.getDate() +
                        '</span></div>'
                    );
                })
                .join('') +
            '</div></div>';

        const allDayRow =
            '<div class="ms365-week__allday">' +
            '<div class="ms365-week__gutter"><span>Ganztägig</span></div>' +
            '<div class="ms365-week__days">' +
            days
                .map((d) => {
                    const iso = toIsoDateOnly(d);
                    const allDay = listForDay(visibleEvents(), iso).filter((t) => t.allDay);
                    return (
                        '<div class="ms365-week__allday-col">' +
                        allDay
                            .slice(0, 4)
                            .map(
                                (t) =>
                                    '<button type="button" class="ms365-week__allday-ev ms365-ev--' +
                                    esc(isGroupSourceId(t.sourceId) ? 'group' : t.sourceId || 'me') +
                                    '" data-th-open="' +
                                    esc(t.id) +
                                    '">' +
                                    esc(t.title) +
                                    '</button>'
                            )
                            .join('') +
                        (allDay.length > 4
                            ? '<span class="ms365-week__more">+' + (allDay.length - 4) + '</span>'
                            : '') +
                        '</div>'
                    );
                })
                .join('') +
            '</div></div>';

        const hourCount = Math.max(1, endH - startH);
        const gutter = hours
            .map(
                (h) =>
                    '<div class="ms365-week__hourlabel" style="top:' +
                    ((h - startH) / hourCount) * 100 +
                    '%">' +
                    String(h).padStart(2, '0') +
                    ':00</div>'
            )
            .join('');

        const columns = days
            .map((d) => {
                const iso = toIsoDateOnly(d);
                const laid = layoutTimedEventsForDay(listForDay(visibleEvents(), iso), iso, {
                    startHour: startH,
                    endHour: endH
                });
                const blocks = laid
                    .map((item) => {
                        const t = item.termin;
                        const width = 100 / item.cols;
                        const left = item.col * width;
                        const time = formatTimeHm(t.start);
                        return (
                            '<button type="button" class="ms365-week__event ms365-ev--' +
                            esc(isGroupSourceId(t.sourceId) ? 'group' : t.sourceId || 'me') +
                            '" style="top:' +
                            item.topPct.toFixed(2) +
                            '%;height:' +
                            Math.max(item.heightPct, 3.2).toFixed(2) +
                            '%;left:calc(' +
                            left.toFixed(2) +
                            '% + 2px);width:calc(' +
                            width.toFixed(2) +
                            '% - 4px);" data-th-open="' +
                            esc(t.id) +
                            '" title="' +
                            esc((time ? time + ' · ' : '') + t.title) +
                            '"><span class="ms365-week__event-time">' +
                            esc(time) +
                            '</span><span class="ms365-week__event-title">' +
                            esc(t.title) +
                            '</span></button>'
                        );
                    })
                    .join('');
                return (
                    '<div class="ms365-week__col' +
                    (iso === todayIso ? ' is-today' : '') +
                    '">' +
                    blocks +
                    '</div>'
                );
            })
            .join('');

        return (
            '<div class="ms365-week' +
            (periodHasVisibleEvents('week') ? '' : ' is-empty') +
            '">' +
            head +
            allDayRow +
            '<div class="ms365-week__body">' +
            '<div class="ms365-week__gutter ms365-week__gutter--times">' +
            gutter +
            '</div>' +
            '<div class="ms365-week__cols">' +
            columns +
            '</div>' +
            emptyOverlayHtml('week') +
            '</div></div>'
        );
    }

    function renderMonthView() {
        const days = monthGridDays(ui.calYear, ui.calMonth);
        const todayIso = toIsoDateOnly(new Date());
        const heads = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']
            .map((d) => '<div class="ms365-month__head">' + d + '</div>')
            .join('');
        const cells = days
            .map((d) => {
                const iso = toIsoDateOnly(d);
                const inMonth = d.getMonth() === ui.calMonth;
                const items = listForDay(visibleEvents(), iso);
                const max = 3;
                const list = items
                    .slice(0, max)
                    .map((t) => {
                        return (
                            '<button type="button" class="ms365-month__ev ms365-ev--' +
                            esc(isGroupSourceId(t.sourceId) ? 'group' : t.sourceId || 'me') +
                            '" data-th-open="' +
                            esc(t.id) +
                            '"><span>' +
                            esc(t.title) +
                            '</span></button>'
                        );
                    })
                    .join('');
                const more =
                    items.length > max
                        ? '<div class="ms365-month__more">+' + (items.length - max) + ' weitere</div>'
                        : '';
                return (
                    '<div class="ms365-month__cell' +
                    (inMonth ? '' : ' is-out') +
                    (iso === todayIso ? ' is-today' : '') +
                    '">' +
                    '<div class="ms365-month__day">' +
                    d.getDate() +
                    '</div>' +
                    '<div class="ms365-month__list">' +
                    list +
                    more +
                    '</div></div>'
                );
            })
            .join('');
        return (
            '<div class="ms365-month' +
            (periodHasVisibleEvents('month') ? '' : ' is-empty') +
            '">' +
            heads +
            cells +
            emptyOverlayHtml('month') +
            '</div>'
        );
    }

    function renderKalender() {
        const mode = ui.calMode;
        const title =
            mode === 'agenda'
                ? 'Nächste 14 Tage'
                : mode === 'week'
                  ? weekRangeLabelDe(ui.weekAnchor)
                  : monthLabelDe(ui.calYear, ui.calMonth);
        const body =
            mode === 'agenda'
                ? renderAgenda()
                : mode === 'week'
                  ? renderWeekView()
                  : renderMonthView();

        return (
            '<div class="ms365-cal">' +
            calToolbar(mode, title) +
            renderLayerToggles() +
            errorsBlock() +
            body +
            '<p class="ms365-cal__hint">Kalender ein-/ausblenden · Termin für Details antippen.</p>' +
            '</div>'
        );
    }

    function findEvent(id) {
        return visibleEvents().find((t) => t.id === id) || ui.events.find((t) => t.id === id) || null;
    }

    function renderDrawer() {
        const id = ui.drawerEventId;
        if (!id) return '';
        const t = findEvent(id);
        if (!t) return '';
        const canWrite = canWriteSource(t.sourceId || SOURCE_IDS.ME);
        const sourceLabel =
            t.sourceLabel || SOURCE_LABEL[t.sourceId] || t.sourceId || 'Outlook';
        const more = ui.moreOpen
            ? '<div class="th-drawer__more">' +
              (canWrite && t.graphEventId
                  ? '<button type="button" class="btn th-btn-danger" data-th-del-outlook="' +
                    esc(t.graphEventId) +
                    '" data-th-source="' +
                    esc(t.sourceId || '') +
                    '">Löschen</button>'
                  : '<p class="ag-item__meta">Löschen nur bei Schreibrecht auf diesen Outlook-Kalender.</p>') +
              '</div>'
            : '';
        return (
            '<div class="th-drawer-root" role="dialog" aria-modal="true" aria-label="Termindetails">' +
            '<button type="button" class="th-drawer__backdrop" data-th-close-drawer aria-label="Schließen"></button>' +
            '<aside class="th-drawer ms365-ev--' +
            esc(isGroupSourceId(t.sourceId) ? 'group' : t.sourceId || 'me') +
            '">' +
            '<div class="th-drawer__head">' +
            '<div>' +
            '<h2 class="th-drawer__title">' +
            esc(t.title) +
            '</h2>' +
            '<p class="th-drawer__sub">' +
            sourcePill(t) +
            '</p>' +
            '</div>' +
            '<button type="button" class="btn th-drawer__close" data-th-close-drawer aria-label="Schließen"><i class="bi bi-x-lg"></i></button>' +
            '</div>' +
            '<div class="th-drawer__core">' +
            '<p class="th-drawer__when"><i class="bi bi-clock" aria-hidden="true"></i>' +
            esc(formatTerminRange(t)) +
            '</p>' +
            (t.ort
                ? '<p class="th-drawer__when"><i class="bi bi-geo-alt" aria-hidden="true"></i>' +
                  esc(t.ort) +
                  '</p>'
                : '') +
            '</div>' +
            '<div class="th-drawer__body">' +
            '<section class="th-drawer__section"><h4>Kalender</h4><p>' +
            esc(sourceLabel) +
            '</p></section>' +
            (t.beschreibung
                ? '<section class="th-drawer__section"><h4>Beschreibung</h4><p>' +
                  esc(t.beschreibung) +
                  '</p></section>'
                : '') +
            (t.sichtbarkeit || t.terminTyp
                ? '<section class="th-drawer__section"><h4>Klassifikation</h4><p>' +
                  (t.sichtbarkeit
                      ? 'Sichtbarkeit: ' +
                        esc(SICHTBARKEIT_LABEL[t.sichtbarkeit] || t.sichtbarkeit)
                      : '') +
                  (t.sichtbarkeit && t.terminTyp ? ' · ' : '') +
                  (t.terminTyp
                      ? 'Typ: ' + esc(TERMIN_TYP_LABEL[t.terminTyp] || t.terminTyp)
                      : '') +
                  '</p></section>'
                : '') +
            more +
            '</div>' +
            '<div class="th-drawer__foot">' +
            (t.webLink
                ? '<a class="btn btn-success" href="' +
                  esc(t.webLink) +
                  '" target="_blank" rel="noopener">In Outlook öffnen</a>'
                : '') +
            '<button type="button" class="btn" data-th-toggle-more>' +
            (ui.moreOpen ? 'Weniger' : 'Weitere Aktionen') +
            '</button>' +
            '</div></aside></div>'
        );
    }

    function optionsFrom(map, selected) {
        return Object.keys(map)
            .map(
                (k) =>
                    '<option value="' +
                    esc(k) +
                    '"' +
                    (k === selected ? ' selected' : '') +
                    '>' +
                    esc(map[k]) +
                    '</option>'
            )
            .join('');
    }

    function writeTargetChecks() {
        const s = settings();
        const opts = writableSources();
        const aud = audienceKey();
        const preferred =
            opts.some((o) => o.id === s.defaultWriteTarget) && s.defaultWriteTarget
                ? s.defaultWriteTarget
                : opts[0]
                  ? opts[0].id
                  : SOURCE_IDS.ME;
        if (!opts.length) {
            return (
                '<p class="ag-item__meta">' +
                (aud === 'self'
                    ? 'Dein eigener Outlook-Kalender ist gerade nicht verfügbar – bitte Seite neu laden oder IT kontaktieren.'
                    : 'Kein beschreibbarer Kalender für deine Rolle.') +
                '</p>'
            );
        }
        const selfHint =
            aud === 'self'
                ? '<p class="ag-item__meta" style="margin:0 0 10px;">Portal-Rolle noch nicht zugeordnet – Termine speicherst du in deinen eigenen Outlook-Kalender. Shared-Kalender schaltet die IT frei (Lehrerliste / Stellvertreter).</p>'
                : '';
        return (
            selfHint +
            '<fieldset class="th-target">' +
            '<legend>Wo soll der Termin erscheinen?</legend>' +
            opts
                .map(
                    (o) =>
                        '<label class="th-target__opt">' +
                        '<input type="checkbox" name="writeTarget" value="' +
                        esc(o.id) +
                        '"' +
                        (o.id === preferred ? ' checked' : '') +
                        '>' +
                        '<span>' +
                        esc(o.label) +
                        (SOURCE_HINT[o.id]
                            ? ' <span class="th-target__hint">(' +
                              esc(SOURCE_HINT[o.id]) +
                              ')</span>'
                            : '') +
                        '</span></label>'
                )
                .join('') +
            '</fieldset>'
        );
    }

    function renderNeu() {
        const startDefault = toLocalInputValue(roundToTimeGrid(new Date()).toISOString());
        const endDefault = toLocalInputValue(
            roundToTimeGrid(new Date(), { addHours: 1 }).toISOString()
        );
        return (
            '<h2 style="margin:0 0 8px;font-size:1.15rem;">Termin anlegen</h2>' +
            '<p class="ag-item__meta" style="margin-bottom:12px;">Wird direkt in Outlook gespeichert – das Schultool bleibt der Hub, Outlook die Kalenderplattform.</p>' +
            '<form class="ag-form" id="thNeuForm">' +
            writeTargetChecks() +
            '<label>Titel<input name="title" required maxlength="200" placeholder="z. B. Fachkonferenz"></label>' +
            '<label class="th-allday">' +
            '<input type="checkbox" name="allDay" id="thAllDay"> Ganztägig</label>' +
            '<div class="ag-form__row th-range" id="thRangeRow">' +
            '<label>Beginn<input type="datetime-local" name="start" id="thStart" required step="900" value="' +
            esc(startDefault) +
            '"></label>' +
            '<label>Ende<input type="datetime-local" name="end" id="thEnd" required step="900" value="' +
            esc(endDefault) +
            '"></label>' +
            '</div>' +
            '<div class="ag-form__row">' +
            '<label>Ort<input name="ort" maxlength="160" placeholder="Raum / Ort"></label>' +
            '<label>Sichtbarkeit<select name="sichtbarkeit">' +
            optionsFrom(SICHTBARKEIT_LABEL, SICHTBARKEIT.INTERN) +
            '</select></label>' +
            '</div>' +
            '<label>Typ <span class="ag-item__meta">(Was für ein Termin?)</span><select name="terminTyp">' +
            optionsFrom(TERMIN_TYP_LABEL, TERMIN_TYP.ORGANISATION) +
            '</select></label>' +
            '<label>Beschreibung<textarea name="beschreibung" rows="3" placeholder="Optional"></textarea></label>' +
            '<div class="ag-form__actions">' +
            '<button type="submit" class="btn btn-success">In Outlook speichern</button>' +
            '<button type="button" class="btn" data-th-go="kalender">Abbrechen</button>' +
            '</div>' +
            '</form>'
        );
    }

    function renderSuccess() {
        const s = ui.success;
        if (!s) return '';
        return (
            '<div class="th-success">' +
            '<p class="th-success__check" aria-hidden="true">✓</p>' +
            '<h2>Termin angelegt</h2>' +
            '<p><strong>' +
            esc(s.title) +
            '</strong> wurde im <strong>' +
            esc(s.calendarLabel) +
            '</strong> gespeichert.</p>' +
            '<div class="ag-form__actions">' +
            '<button type="button" class="btn btn-success" data-th-show-created="' +
            esc(s.eventId) +
            '">Termin anzeigen</button>' +
            (s.webLink
                ? '<a class="btn" href="' +
                  esc(s.webLink) +
                  '" target="_blank" rel="noopener">In Outlook öffnen</a>'
                : '') +
            '</div></div>'
        );
    }

    function probeBadge(status) {
        if (status === 'ok') return '<span class="th-status th-status--ok">Erreichbar</span>';
        if (status === 'denied') return '<span class="th-status th-status--bad">Keine Berechtigung</span>';
        if (status === 'unreachable') {
            return '<span class="th-status th-status--bad">Nicht erreichbar</span>';
        }
        if (status === 'missing') {
            return '<span class="th-status th-status--muted">Nicht konfiguriert</span>';
        }
        return '<span class="th-status th-status--muted">Noch nicht geprüft</span>';
    }

    function renderQuellen() {
        const s = settings();
        const sug = suggestUpnsFromKalenderPostfaecher();
        const enabled = new Set(s.enabledSources);
        const rows = [
            {
                id: SOURCE_IDS.ME,
                label: SOURCE_LABEL.me,
                hint: 'Für berechtigte Benutzer verfügbar',
                upn: '',
                upnName: '',
                placeholder: ''
            },
            {
                id: SOURCE_IDS.SCHUL,
                label: SOURCE_LABEL.schul,
                hint: 'Für berechtigte Benutzer verfügbar',
                upn: s.schulkalenderUpn,
                upnName: 'schulkalenderUpn',
                placeholder: sug.schulkalenderUpn || 'schulkalender@schule.at'
            },
            {
                id: SOURCE_IDS.LEHRER,
                label: SOURCE_LABEL.lehrer,
                hint: 'Für Lehrkräfte verfügbar',
                upn: s.lehrerkalenderUpn,
                upnName: 'lehrerkalenderUpn',
                placeholder: sug.lehrerkalenderUpn || 'lehrerkalender@schule.at'
            },
            {
                id: SOURCE_IDS.SCHUELER,
                label: SOURCE_LABEL.schueler,
                hint: 'Für Schüler:innen verfügbar',
                upn: s.schuelerkalenderUpn,
                upnName: 'schuelerkalenderUpn',
                placeholder: sug.schuelerkalenderUpn || 'schuelerkalender@schule.at'
            }
        ];

        const writeOpts = resolveEnabledSources({
            ...s,
            enabledSources: [
                SOURCE_IDS.ME,
                SOURCE_IDS.SCHUL,
                SOURCE_IDS.LEHRER,
                SOURCE_IDS.SCHUELER
            ]
        });
        const writeOptsHtml = writeOpts
            .map(
                (o) =>
                    '<option value="' +
                    esc(o.id) +
                    '"' +
                    (o.id === s.defaultWriteTarget ? ' selected' : '') +
                    '>' +
                    esc(o.label) +
                    '</option>'
            )
            .join('');

        return (
            '<h2 style="margin:0 0 8px;font-size:1.15rem;">Kalenderadministration</h2>' +
            '<p class="ag-item__meta" style="margin-bottom:14px;">UPNs und Verfügbarkeit der Shared-Kalender. Normale Lehrkräfte konfigurieren hier nichts.</p>' +
            '<form class="ag-form" id="thQuellenForm">' +
            rows
                .map((r) => {
                    const probe = ui.probe[r.id];
                    return (
                        '<article class="th-source-card th-source-card--' +
                        esc(r.id) +
                        '">' +
                        '<div class="th-source-card__top">' +
                        '<div><strong>' +
                        esc(r.label) +
                        '</strong><span>' +
                        esc(r.hint) +
                        '</span></div>' +
                        '<label class="th-switch">' +
                        '<span class="th-switch__text">Kalender verfügbar</span>' +
                        '<input type="checkbox" name="src_' +
                        esc(r.id) +
                        '"' +
                        (enabled.has(r.id) ? ' checked' : '') +
                        '>' +
                        '</label></div>' +
                        (r.upnName
                            ? '<label>UPN<input name="' +
                              esc(r.upnName) +
                              '" type="email" value="' +
                              esc(r.upn) +
                              '" placeholder="' +
                              esc(r.placeholder) +
                              '"></label>' +
                              '<div class="th-source-card__status">' +
                              probeBadge(probe && probe.status) +
                              (probe && probe.message
                                  ? '<span class="th-source-card__msg">' +
                                    esc(probe.message) +
                                    '</span>'
                                  : '') +
                              '<button type="button" class="btn" data-th-probe="' +
                              esc(r.id) +
                              '">Verbindung prüfen</button></div>'
                            : '<p class="ag-item__meta">Kein UPN nötig – persönlicher Kalender der angemeldeten Person.</p>') +
                        '</article>'
                    );
                })
                .join('') +
            '<label>Standard beim Anlegen<select name="defaultWriteTarget">' +
            writeOptsHtml +
            '</select></label>' +
            '<div class="ag-form__actions">' +
            '<button type="submit" class="btn btn-success">Speichern</button>' +
            '<button type="button" class="btn" data-th-suggest>Kalender aus eingerichteten Postfächern übernehmen</button>' +
            '<a class="btn" href="' +
            KALENDER_POSTFAECHER_HREF +
            '">Postfächer einrichten</a>' +
            '</div>' +
            '</form>' +
            '<p class="ag-item__meta">Hinweis: Fremde Kalender brauchen Freigabe mit Bearbeiten-Recht für die Nutzerkonten (nicht nur Global Admin).</p>'
        );
    }

    function renderAbwesenheit() {
        const href = abwesenheitHref();
        const showSchulaktivitaeten = canAccessSchulaktivitaeten();
        return (
            '<div class="th-absence">' +
            '<h2 style="margin:0 0 8px;font-size:1.15rem;">Abwesenheit</h2>' +
            '<p class="ag-item__meta" style="margin-bottom:16px;">Abwesenheiten werden im Freistellungsplaner verwaltet – hier nur der Einstieg, keine Doppelpflege.</p>' +
            '<div class="th-absence__list">' +
            '<a class="th-absence__row" href="' +
            href +
            '"><div><strong>Neue Abwesenheit</strong><span>Antrag, Zeitraum und Vertretung erfassen</span></div><span class="th-absence__cta">Freistellung planen →</span></a>' +
            '<a class="th-absence__row" href="' +
            href +
            '#uebersicht"><div><strong>Meine Abwesenheiten</strong><span>Status bestehender Freistellungen ansehen</span></div><span class="th-absence__cta">Meine Freistellungen →</span></a>' +
            (showSchulaktivitaeten
                ? '<a class="th-absence__row" href="' +
                  SCHULAKTIVITAETEN_HREF +
                  '"><div><strong>Schulaktivitäten einreichen</strong><span>Exkursion / Aktivität beantragen</span></div><span class="th-absence__cta">Öffnen →</span></a>'
                : '') +
            '</div></div>'
        );
    }

    function resolveSourceTarget(sourceId) {
        const s = settings();
        const email = signedInEmail();
        const gid = groupIdFromSourceId(sourceId);
        if (gid) {
            const g = (s.groupCalendars || []).find((x) => x.id === gid);
            return {
                kind: 'group',
                groupId: gid,
                upn: (g && g.mail) || '',
                signedInEmail: email,
                sourceId: sourceId,
                sourceLabel: (g && g.displayName) || 'Gruppenkalender'
            };
        }
        if (sourceId === SOURCE_IDS.ME) {
            return {
                kind: 'me',
                upn: '',
                signedInEmail: email,
                sourceId: SOURCE_IDS.ME,
                sourceLabel: SOURCE_LABEL.me
            };
        }
        if (sourceId === SOURCE_IDS.SCHUL) {
            return {
                kind: 'user',
                upn: s.schulkalenderUpn,
                signedInEmail: email,
                sourceId: SOURCE_IDS.SCHUL,
                sourceLabel: SOURCE_LABEL.schul
            };
        }
        if (sourceId === SOURCE_IDS.LEHRER) {
            return {
                kind: 'user',
                upn: s.lehrerkalenderUpn,
                signedInEmail: email,
                sourceId: SOURCE_IDS.LEHRER,
                sourceLabel: SOURCE_LABEL.lehrer
            };
        }
        if (sourceId === SOURCE_IDS.SCHUELER) {
            return {
                kind: 'user',
                upn: s.schuelerkalenderUpn,
                signedInEmail: email,
                sourceId: SOURCE_IDS.SCHUELER,
                sourceLabel: SOURCE_LABEL.schueler
            };
        }
        const fallback = writeTargetFromSettings();
        return {
            ...fallback,
            signedInEmail: email,
            sourceId: fallback.id,
            sourceLabel: fallback.label
        };
    }

    async function confirmDelete() {
        const msg =
            'Termin in Outlook wirklich löschen? Bei Besprechungen mit Teilnehmer:innen kann Outlook Absagen erwarten – prüfe ggf. in Outlook.';
        try {
            if (typeof window.ms365Confirm === 'function') {
                return !!(await window.ms365Confirm(msg));
            }
        } catch {
            /* ignore */
        }
        return window.confirm(msg);
    }

    function bind() {
        root.querySelectorAll('[data-th-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                let next = btn.getAttribute('data-th-view') || 'kalender';
                if (next === 'quellen' && !canManageQuellen()) next = 'kalender';
                ui.view = next;
                ui.message = null;
                ui.success = null;
                render();
                if (ui.view === 'kalender') {
                    void refreshOutlook({ silent: ui.events.length > 0 });
                }
            });
        });
        root.querySelectorAll('[data-th-go]').forEach((btn) => {
            btn.addEventListener('click', () => {
                let next = btn.getAttribute('data-th-go') || 'kalender';
                if (next === 'quellen' && !canManageQuellen()) next = 'kalender';
                ui.view = next;
                ui.success = null;
                render();
            });
        });
        root.querySelectorAll('[data-th-mode]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const mode = btn.getAttribute('data-th-mode') || 'week';
                ui.calMode =
                    mode === 'agenda' || mode === 'month' || mode === 'week' ? mode : 'week';
                saveCalMode(ui.calMode);
                render();
                void refreshOutlook({ silent: true });
            });
        });
        root.querySelectorAll('[data-th-nav]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const delta = Number(btn.getAttribute('data-th-nav') || '0');
                if (ui.calMode === 'week') {
                    if (delta === 0) ui.weekAnchor = startOfWeekMonday(new Date());
                    else {
                        const d = new Date(ui.weekAnchor);
                        d.setDate(d.getDate() + delta * 7);
                        ui.weekAnchor = startOfWeekMonday(d);
                    }
                } else if (ui.calMode === 'month') {
                    if (delta === 0) {
                        const n = new Date();
                        ui.calYear = n.getFullYear();
                        ui.calMonth = n.getMonth();
                    } else {
                        const d = new Date(ui.calYear, ui.calMonth + delta, 1);
                        ui.calYear = d.getFullYear();
                        ui.calMonth = d.getMonth();
                    }
                }
                render();
                void refreshOutlook({ silent: true });
            });
        });
        root.querySelectorAll('[data-th-refresh]').forEach((btn) => {
            btn.addEventListener('click', () => {
                void refreshOutlook();
            });
        });
        root.querySelectorAll('[data-ms365-layer]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const id = btn.getAttribute('data-ms365-layer') || '';
                const before = sources()
                    .filter((s) => s.visible)
                    .map((s) => s.id)
                    .join('|');
                toggleVisibleSource(id, undefined, { audience: audienceKey() });
                const after = sources()
                    .filter((s) => s.visible)
                    .map((s) => s.id)
                    .join('|');
                if (after === before) {
                    setMessage('warn', 'Mindestens ein Kalender muss eingeblendet bleiben.');
                } else {
                    ui.message = null;
                }
                if (ui.drawerEventId) {
                    const still = visibleEvents().some((t) => t.id === ui.drawerEventId);
                    if (!still) {
                        ui.drawerEventId = '';
                        ui.moreOpen = false;
                    }
                }
                render();
                void refreshOutlook({ silent: true });
            });
        });
        root.querySelector('[data-th-group-open]')?.addEventListener('click', () => {
            ui.groupPickerOpen = true;
            ui.groupPickerLoading = true;
            ui.groupPickerError = '';
            ui.groupPickerItems = [];
            render();
            void listMyUnifiedGroups()
                .then((items) => {
                    ui.groupPickerItems = items || [];
                    ui.groupPickerLoading = false;
                    render();
                })
                .catch((e) => {
                    ui.groupPickerLoading = false;
                    ui.groupPickerError =
                        (e && e.message ? e.message : String(e)) ||
                        'Gruppen konnten nicht geladen werden (Graph-Recht Group.Read.All).';
                    render();
                });
        });
        root.querySelector('[data-th-group-cancel]')?.addEventListener('click', () => {
            ui.groupPickerOpen = false;
            ui.groupPickerError = '';
            render();
        });
        root.querySelector('[data-th-group-save]')?.addEventListener('click', () => {
            const checked = Array.from(root.querySelectorAll('input[name="groupCal"]:checked'));
            const byId = new Map(ui.groupPickerItems.map((g) => [g.id, g]));
            const selected = checked
                .map((el) => byId.get(String(el.value || '').trim()))
                .filter(Boolean);
            setSelectedGroupCalendars(selected);
            ui.groupPickerOpen = false;
            setMessage(
                'ok',
                selected.length
                    ? selected.length + ' Gruppenkalender übernommen.'
                    : 'Keine Gruppenkalender ausgewählt.'
            );
            render();
            void refreshOutlook({ silent: true });
        });
        root.querySelectorAll('[data-th-open]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.drawerEventId = btn.getAttribute('data-th-open') || '';
                ui.moreOpen = false;
                render();
            });
        });
        root.querySelectorAll('[data-th-close-drawer]').forEach((btn) => {
            btn.addEventListener('click', () => {
                ui.drawerEventId = '';
                ui.moreOpen = false;
                render();
            });
        });
        root.querySelector('[data-th-toggle-more]')?.addEventListener('click', () => {
            ui.moreOpen = !ui.moreOpen;
            render();
        });
        root.querySelectorAll('[data-th-del-outlook]').forEach((btn) => {
            btn.addEventListener('click', () => {
                void (async () => {
                    if (!(await confirmDelete())) return;
                    const id = btn.getAttribute('data-th-del-outlook') || '';
                    const sourceId = btn.getAttribute('data-th-source') || '';
                    if (!canWriteSource(sourceId || SOURCE_IDS.ME)) {
                        setMessage(
                            'err',
                            'Löschen in diesem Kalender ist für deine Rolle nicht freigeschaltet.'
                        );
                        render();
                        return;
                    }
                    const target = resolveSourceTarget(sourceId);
                    try {
                        await deleteOutlookEvent(id, target);
                        setMessage('ok', 'In Outlook gelöscht.');
                        ui.drawerEventId = '';
                        ui.moreOpen = false;
                        await refreshOutlook({ silent: true });
                    } catch (e) {
                        setMessage('err', e && e.message ? e.message : String(e));
                        render();
                    }
                })();
            });
        });
        root.querySelectorAll('[data-th-probe]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const sourceId = btn.getAttribute('data-th-probe') || '';
                const target = resolveSourceTarget(sourceId);
                btn.setAttribute('disabled', 'disabled');
                void probeCalendarSource(target)
                    .then((res) => {
                        ui.probe[sourceId] = {
                            status: res.status,
                            message: res.message,
                            at: new Date().toISOString()
                        };
                        setMessage(
                            res.ok ? 'ok' : 'warn',
                            SOURCE_LABEL[sourceId] + ': ' + res.message
                        );
                        render();
                    })
                    .catch((e) => {
                        ui.probe[sourceId] = {
                            status: 'unreachable',
                            message: e && e.message ? e.message : String(e)
                        };
                        render();
                    });
            });
        });
        root.querySelector('[data-th-suggest]')?.addEventListener('click', () => {
            if (!canManageQuellen()) return;
            const sug = suggestUpnsFromKalenderPostfaecher();
            if (!sug.schulkalenderUpn && !sug.lehrerkalenderUpn && !sug.schuelerkalenderUpn) {
                setMessage(
                    'warn',
                    'Keine Vorschläge – zuerst unter Kalender-Postfächer Schul-/Lehrer-/Schülerkalender anlegen.'
                );
                render();
                return;
            }
            const cur = settings();
            let update = 0;
            let neu = 0;
            const keys = [
                ['schulkalenderUpn', sug.schulkalenderUpn],
                ['lehrerkalenderUpn', sug.lehrerkalenderUpn],
                ['schuelerkalenderUpn', sug.schuelerkalenderUpn]
            ];
            keys.forEach(([k, v]) => {
                if (!v) return;
                if (cur[k]) {
                    if (cur[k] !== v) update += 1;
                } else neu += 1;
            });
            const recognized = keys.filter(([, v]) => v).length;
            const ok = window.confirm(
                recognized +
                    ' Kalender erkannt.\n' +
                    update +
                    ' bestehende Zuordnungen werden aktualisiert.\n' +
                    neu +
                    ' neue Zuordnung wird angelegt.\n\nÜbernehmen?'
            );
            if (!ok) return;
            saveSettings({
                schulkalenderUpn: sug.schulkalenderUpn || cur.schulkalenderUpn,
                lehrerkalenderUpn: sug.lehrerkalenderUpn || cur.lehrerkalenderUpn,
                schuelerkalenderUpn: sug.schuelerkalenderUpn || cur.schuelerkalenderUpn
            });
            setMessage('ok', 'Kalender aus eingerichteten Postfächern übernommen.');
            render();
        });
        root.querySelector('#thQuellenForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            if (!canManageQuellen()) {
                setMessage('err', 'Nur die IT darf Kalenderquellen speichern.');
                ui.view = 'kalender';
                render();
                return;
            }
            const fd = new FormData(ev.target);
            const enabled = [];
            if (ev.target.querySelector('[name="src_me"]')?.checked) enabled.push(SOURCE_IDS.ME);
            if (ev.target.querySelector('[name="src_schul"]')?.checked) {
                enabled.push(SOURCE_IDS.SCHUL);
            }
            if (ev.target.querySelector('[name="src_lehrer"]')?.checked) {
                enabled.push(SOURCE_IDS.LEHRER);
            }
            if (ev.target.querySelector('[name="src_schueler"]')?.checked) {
                enabled.push(SOURCE_IDS.SCHUELER);
            }
            const enabledFinal = enabled.length ? enabled : [SOURCE_IDS.ME];
            const prevVisible = settings().visibleSources.filter((id) => enabledFinal.includes(id));
            saveSettings({
                schulkalenderUpn: String(fd.get('schulkalenderUpn') || ''),
                lehrerkalenderUpn: String(fd.get('lehrerkalenderUpn') || ''),
                schuelerkalenderUpn: String(fd.get('schuelerkalenderUpn') || ''),
                enabledSources: enabledFinal,
                visibleSources: prevVisible.length ? prevVisible : enabledFinal.slice(),
                defaultWriteTarget: String(fd.get('defaultWriteTarget') || SOURCE_IDS.ME)
            });
            setMessage('ok', 'Quellen gespeichert.');
            ui.view = 'kalender';
            render();
            void refreshOutlook();
        });

        const allDay = root.querySelector('#thAllDay');
        const startInput = root.querySelector('#thStart');
        const endInput = root.querySelector('#thEnd');
        /** Zwischenspeicher, damit Umschalten Ganztägig Uhrzeit nicht verliert. */
        let timedStart = startInput ? String(startInput.value || '') : '';
        let timedEnd = endInput ? String(endInput.value || '') : '';
        const syncAllDayUi = () => {
            if (!startInput || !endInput) return;
            const on = !!(allDay && allDay.checked);
            const wantType = on ? 'date' : 'datetime-local';
            if (startInput.type === wantType && endInput.type === wantType) return;
            if (on) {
                if (startInput.type === 'datetime-local') {
                    timedStart = String(startInput.value || timedStart);
                    timedEnd = String(endInput.value || timedEnd);
                }
                const startDate = String(startInput.value || timedStart || '').slice(0, 10);
                const endDate = String(endInput.value || timedEnd || startDate || '').slice(0, 10);
                startInput.type = 'date';
                endInput.type = 'date';
                startInput.removeAttribute('step');
                endInput.removeAttribute('step');
                startInput.value = startDate;
                endInput.value = endDate || startDate;
            } else {
                const startDate = String(startInput.value || '').slice(0, 10);
                const endDate = String(endInput.value || startDate || '').slice(0, 10);
                const fallbackStart =
                    timedStart ||
                    toLocalInputValue(roundToTimeGrid(new Date()).toISOString());
                const fallbackEnd =
                    timedEnd ||
                    toLocalInputValue(roundToTimeGrid(new Date(), { addHours: 1 }).toISOString());
                const startTime = (timedStart || fallbackStart).slice(11, 16) || '09:00';
                const endTime = (timedEnd || fallbackEnd).slice(11, 16) || '10:00';
                startInput.type = 'datetime-local';
                endInput.type = 'datetime-local';
                startInput.step = '900';
                endInput.step = '900';
                startInput.value = startDate
                    ? startDate + 'T' + startTime
                    : fallbackStart;
                endInput.value = endDate ? endDate + 'T' + endTime : fallbackEnd;
                timedStart = String(startInput.value || '');
                timedEnd = String(endInput.value || '');
            }
            startInput.required = true;
            endInput.required = true;
        };
        if (allDay) {
            allDay.addEventListener('change', syncAllDayUi);
            syncAllDayUi();
        }

        root.querySelector('[data-th-show-created]')?.addEventListener('click', () => {
            const id = root.querySelector('[data-th-show-created]')?.getAttribute('data-th-show-created') || '';
            ui.view = 'kalender';
            ui.calMode = 'agenda';
            saveCalMode('agenda');
            ui.success = null;
            ui.drawerEventId = id;
            render();
            void refreshOutlook({ silent: true });
        });

        root.querySelector('#thNeuForm')?.addEventListener('submit', (ev) => {
            ev.preventDefault();
            const fd = new FormData(ev.target);
            let writeIds = fd
                .getAll('writeTarget')
                .map((v) => String(v || '').trim())
                .filter(Boolean);
            // Kein Häkchen gesetzt (oder UI ohne Checkboxen): Eigenkalender als Default
            if (!writeIds.length) {
                const fallback = writableSources();
                if (fallback.some((s) => s.id === SOURCE_IDS.ME) || canWriteSource(SOURCE_IDS.ME)) {
                    writeIds = [SOURCE_IDS.ME];
                } else if (fallback[0]) {
                    writeIds = [fallback[0].id];
                } else {
                    setMessage(
                        'err',
                        isAuthLoggedIn()
                            ? 'Kein Zielkalender verfügbar – bitte neu anmelden oder IT kontaktieren.'
                            : 'Bitte zuerst anmelden, dann Termin speichern.'
                    );
                    render();
                    return;
                }
            }
            for (const writeId of writeIds) {
                if (!canWriteSource(writeId)) {
                    setMessage('err', 'Ein gewählter Zielkalender ist für deine Rolle nicht freigeschaltet.');
                    render();
                    return;
                }
                const target = resolveSourceTarget(writeId);
                if (target.kind === 'group' && !target.groupId) {
                    setMessage('err', 'Gruppenkalender unvollständig – bitte unter „Gruppen…“ neu wählen.');
                    render();
                    return;
                }
                if (target.kind === 'user' && !target.upn) {
                    setMessage(
                        'err',
                        canManageQuellen()
                            ? 'UPN für ' +
                                  (SOURCE_LABEL[writeId] || writeId) +
                                  ' fehlt – unter Quellen eintragen.'
                            : 'UPN für ' +
                                  (SOURCE_LABEL[writeId] || writeId) +
                                  ' fehlt – bitte IT kontaktieren.'
                    );
                    render();
                    return;
                }
            }
            const allDayOn = !!ev.target.querySelector('[name="allDay"]')?.checked;
            const startRaw = String(fd.get('start') || '');
            const endRaw = String(fd.get('end') || startRaw || '');
            const draft = {
                title: String(fd.get('title') || ''),
                start: allDayOn ? startRaw.slice(0, 10) : fromLocalInputValue(startRaw),
                end: allDayOn
                    ? (endRaw || startRaw).slice(0, 10)
                    : fromLocalInputValue(endRaw),
                allDay: allDayOn,
                ort: String(fd.get('ort') || ''),
                beschreibung: String(fd.get('beschreibung') || ''),
                sichtbarkeit: String(fd.get('sichtbarkeit') || SICHTBARKEIT.INTERN),
                terminTyp: String(fd.get('terminTyp') || TERMIN_TYP.ORGANISATION)
            };
            const submitBtn = ev.target.querySelector('[type="submit"]');
            if (submitBtn) submitBtn.disabled = true;
            void (async () => {
                try {
                    /** @type {object[]} */
                    const createdList = [];
                    for (const writeId of writeIds) {
                        const target = resolveSourceTarget(writeId);
                        const created = await createOutlookEvent(draft, target);
                        createdList.push({ writeId, target, created });
                    }
                    const primary = createdList[0];
                    if (canManageQuellen() && primary) {
                        saveSettings({ defaultWriteTarget: primary.writeId });
                    }
                    const labels = createdList
                        .map(
                            (x) =>
                                x.target.sourceLabel ||
                                SOURCE_LABEL[x.writeId] ||
                                x.writeId
                        )
                        .join(', ');
                    ui.success = {
                        title: (primary && primary.created && primary.created.title) || draft.title,
                        calendarLabel: labels,
                        webLink:
                            (primary && primary.created && primary.created.webLink) || '',
                        eventId: (primary && primary.created && primary.created.id) || ''
                    };
                    ui.view = 'success';
                    ui.message = null;
                    await refreshOutlook({ silent: true });
                } catch (e) {
                    setMessage('err', e && e.message ? e.message : String(e));
                    render();
                } finally {
                    if (submitBtn) submitBtn.disabled = false;
                }
            })();
        });
    }

    function publishTermineRole() {
        let email = '';
        try {
            if (typeof window.ms365AuthGetAccountInfo === 'function') {
                const a = window.ms365AuthGetAccountInfo();
                email = String((a && (a.username || a.upn)) || '')
                    .trim()
                    .toLowerCase();
            }
        } catch {
            /* ignore */
        }
        let role = 'self';
        if (personas && personas.isIt) role = 'it';
        else if (personas && personas.isLehrer) role = 'lehrer';
        else if (personas && personas.isSchueler) role = 'schueler';
        announceAppRole({ appId: 'termine', role, source: 'portal', email });
    }

    render();

    const landingShot = new URLSearchParams(window.location.search || '').get('landingShot') === '1';
    if (landingShot) {
        personas = {
            loggedIn: true,
            isIt: false,
            isLehrer: true,
            isSchueler: false,
            schularbeitenRoles: ['lehrer'],
            freistellungRoles: ['lehrer']
        };
        personasReady = true;
        const demo = buildTermineHubDemoState();
        ui.events = (demo.termine || []).map((t) => ({
            ...t,
            sourceId: SOURCE_IDS.ME,
            sourceLabel: 'Schulkalender'
        }));
        ui.calMode = 'agenda';
        ui.lastErrors = [];
        publishTermineRole();
        render();
        return;
    }

    void getDashboardPersonas()
        .then((p) => {
            if (p) personas = p;
            personasReady = true;
            publishTermineRole();
            if (!canManageQuellen() && ui.view === 'quellen') ui.view = 'kalender';
            render();
            void refreshOutlook({ silent: true });
        })
        .catch(() => {
            personasReady = true;
            publishTermineRole();
            render();
        });
}
