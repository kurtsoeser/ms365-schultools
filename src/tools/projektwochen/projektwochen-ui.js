/**
 * Projektwochen UI – Template-Strings + Full-Rerender.
 */
import {
    VIEWS,
    navSectionsForRole,
    filterAngebote,
    scopeFromState,
    canEditAngebot,
    canAdminDecide,
    labelMaps,
    emptyForm
} from './projektwochen-state.js';
import {
    PW_ROLE_ORDER,
    canUsePwRoleSwitcher,
    isPwDemoRoleUiEnabled,
    sortPwRoles
} from './projektwochen-entra-role.js';
import {
    validateAngebot,
    computeDashboardKpis,
    computePwAttention,
    buildWeekPlan,
    buildWeekPlanDisplayRows,
    buildMonthCells,
    effectiveBuchungAb,
    isBookingOpen,
    toDatetimeLocalValue,
    weekdayLabelDeFromIso,
    STATUS_ANGEBOT,
    SLOTS,
    PLAN_SLOTS,
    TAGE,
    KATEGORIEN
} from './projektwochen-logic.js';
import { countPendingBookingsSync } from './projektwochen-bookings-logic.js';
import { formatPwSpoSyncLine } from './projektwochen-sync.js';
import { renderSystemAdminPage } from '../../shared/system-admin-shell.js';

function occupancyLabel(state, angebot) {
    const sid = String((angebot && angebot.bookingsServiceId) || '').trim();
    if (!sid || !state.occupancy || !state.occupancy[sid]) return '';
    const o = state.occupancy[sid];
    return String(o.filled || 0) + '/' + String(o.max || angebot.kapazitaet || '?');
}

function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function optionList(items, valueKey, labelKey, selected) {
    return (items || [])
        .map((it) => {
            const v = typeof it === 'string' ? it : String(it[valueKey] || '');
            const lab = typeof it === 'string' ? it : String(it[labelKey] || it[valueKey] || v);
            return (
                '<option value="' +
                esc(v) +
                '"' +
                (String(selected) === v ? ' selected' : '') +
                '>' +
                esc(lab) +
                '</option>'
            );
        })
        .join('');
}

function statusBadge(status) {
    const s = String(status || '');
    return '<span class="pw-badge pw-badge--' + esc(s) + '">' + esc(s || '–') + '</span>';
}

function scopedItems(state) {
    return filterAngebote(state.items, state.filters, scopeFromState(state), state.aktion);
}

function filterBar(state) {
    const f = state.filters || {};
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    const classes = (state.stammdaten && state.stammdaten.classes) || [];
    const activeCount = [
        f.status,
        f.kategorie,
        f.tag,
        f.lehrer,
        f.klasse,
        f.buchung,
        f.q
    ].filter(Boolean).length;

    function filterSelect(id, label, icon, optionsHtml) {
        return (
            '<div class="pw-filters__field">' +
            '<label class="pw-filters__label" for="' +
            id +
            '"><i class="bi ' +
            icon +
            '" aria-hidden="true"></i>' +
            esc(label) +
            '</label>' +
            '<div class="pw-select">' +
            '<select id="' +
            id +
            '" class="pw-select__control">' +
            optionsHtml +
            '</select>' +
            '<i class="bi bi-chevron-down pw-select__caret" aria-hidden="true"></i>' +
            '</div></div>'
        );
    }

    return (
        '<div class="pw-filters" role="search" aria-label="Angebote filtern">' +
        '<div class="pw-filters__head">' +
        '<span class="pw-filters__title"><i class="bi bi-funnel" aria-hidden="true"></i>Filter &amp; Suche</span>' +
        (activeCount
            ? '<button type="button" class="pw-filters__reset" id="pwFilterReset" title="Filter zurücksetzen">' +
              '<i class="bi bi-x-circle" aria-hidden="true"></i>' +
              activeCount +
              ' aktiv · zurücksetzen</button>'
            : '<span class="pw-filters__hint">Alle Angebote</span>') +
        '</div>' +
        '<div class="pw-filters__row">' +
        filterSelect(
            'pwFilterStatus',
            'Status',
            'bi-flag',
            '<option value="">alle</option>' + optionList(STATUS_ANGEBOT, null, null, f.status)
        ) +
        filterSelect(
            'pwFilterKat',
            'Kategorie',
            'bi-tags',
            '<option value="">alle</option>' + optionList(KATEGORIEN, null, null, f.kategorie)
        ) +
        filterSelect(
            'pwFilterTag',
            'Tag',
            'bi-calendar-week',
            '<option value="">alle</option>' + optionList(TAGE, null, null, f.tag)
        ) +
        filterSelect(
            'pwFilterLehrer',
            'Lehrer',
            'bi-person-badge',
            '<option value="">alle</option>' + optionList(teachers, 'code', 'name', f.lehrer)
        ) +
        filterSelect(
            'pwFilterKlasse',
            'Klasse',
            'bi-mortarboard',
            '<option value="">alle</option>' + optionList(classes, 'code', 'name', f.klasse)
        ) +
        filterSelect(
            'pwFilterBuchung',
            'Buchung',
            'bi-clock-history',
            '<option value="">alle</option>' +
                '<option value="offen"' +
                (f.buchung === 'offen' ? ' selected' : '') +
                '>bereits möglich</option>' +
                '<option value="gesperrt"' +
                (f.buchung === 'gesperrt' ? ' selected' : '') +
                '>noch gesperrt</option>'
        ) +
        '<div class="pw-filters__field pw-filters__field--search">' +
        '<label class="pw-filters__label" for="pwFilterQ"><i class="bi bi-search" aria-hidden="true"></i>Suche</label>' +
        '<div class="pw-search">' +
        '<i class="bi bi-search pw-search__icon" aria-hidden="true"></i>' +
        '<input type="search" id="pwFilterQ" class="pw-search__input" value="' +
        esc(f.q || '') +
        '" placeholder="Titel, Ort, Beschreibung …" autocomplete="off" spellcheck="false">' +
        '</div></div>' +
        '</div></div>'
    );
}

function quietEmpty(title, hint, ctaHtml) {
    return (
        '<div class="pw-quiet-empty" role="status">' +
        '<strong>' +
        esc(title) +
        '</strong>' +
        (hint ? '<p>' + esc(hint) + '</p>' : '') +
        (ctaHtml || '') +
        '</div>'
    );
}

function calmBanner(eyebrow, bodyHtml, opts) {
    const o = opts || {};
    const tone = o.tone === 'ops' ? 'ops' : 'action';
    const jump = o.jumpView
        ? '<button type="button" class="btn btn-sm" data-pw-view="' +
          esc(o.jumpView) +
          '">' +
          esc(o.jumpLabel || 'Öffnen') +
          '</button>'
        : '';
    return (
        '<section class="pw-calm-banner pw-calm-banner--' +
        tone +
        '" role="status">' +
        (o.icon
            ? '<div class="pw-calm-banner__icon" aria-hidden="true"><i class="bi ' +
              esc(o.icon) +
              '"></i></div>'
            : '') +
        '<div class="pw-calm-banner__text">' +
        '<p class="pw-calm-banner__eyebrow">' +
        esc(eyebrow) +
        '</p>' +
        '<div>' +
        bodyHtml +
        '</div></div>' +
        jump +
        '</section>'
    );
}

function renderNav(state) {
    const sections = navSectionsForRole(state.role);
    const pendingSync = state.role === 'admin' ? countPendingBookingsSync(state.items) : 0;
    const openQueue =
        state.role === 'admin'
            ? (state.items || []).filter((a) => a && a.status === 'beantragt').length
            : 0;
    const navHtml = sections
        .map((sec) => {
            const label = sec.label
                ? '<span class="pw-nav__section-label">' + esc(sec.label) + '</span>'
                : '';
            const btns = (sec.items || [])
                .map((v) => {
                    let badge = '';
                    if (v.id === 'admin' && openQueue) {
                        badge = '<span class="pw-nav__badge">' + openQueue + '</span>';
                    } else if (v.id === 'bookings' && pendingSync) {
                        badge = '<span class="pw-nav__badge">' + pendingSync + '</span>';
                    }
                    return (
                        '<button type="button" class="pw-nav__btn' +
                        (state.view === v.id ? ' is-active' : '') +
                        (v.id === 'neu' ? ' pw-nav__btn--cta' : '') +
                        '" data-pw-view="' +
                        esc(v.id) +
                        '"><i class="bi ' +
                        esc(v.icon) +
                        '" aria-hidden="true"></i><span class="pw-nav__btn-label">' +
                        esc(v.label) +
                        '</span>' +
                        badge +
                        '</button>'
                    );
                })
                .join('');
            return '<div class="pw-nav__section">' + label + btns + '</div>';
        })
        .join('');
    const demoRoleUi = isPwDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    const switchable = demoRoleUi ? PW_ROLE_ORDER : sortPwRoles(state.planerRoles || []);
    const showRoleSwitcher = canUsePwRoleSwitcher(state.planerRoles || [], demoRoleUi);
    const roleLabels = { lehrer: 'Lehrer', admin: 'Admin', schueler: 'Schüler' };
    const brandSub =
        state.role === 'admin'
            ? 'Freigabe &amp; Bookings'
            : state.role === 'schueler'
              ? 'Plan lesen'
              : 'Angebote planen';
    const rolesHtml = showRoleSwitcher
        ? '<span class="pw-nav__roles-label">Rolle' +
          (demoRoleUi ? ' (Vorschau)' : '') +
          '</span><div class="pw-chips">' +
          switchable
              .map((r) => {
                  return (
                      '<button type="button" class="pw-chip' +
                      (state.role === r ? ' is-active' : '') +
                      '" data-pw-role="' +
                      r +
                      '">' +
                      (roleLabels[r] || r) +
                      '</button>'
                  );
              })
              .join('') +
          '</div>'
        : '<span class="pw-nav__roles-label">Rolle</span><p class="pw-nav__hint"><span class="pw-chip is-active">' +
          (roleLabels[state.role] || state.role) +
          '</span></p>' +
          (state.planerAccessDenied
              ? '<p class="pw-nav__hint">' + esc(String(state.roleHint || 'Kein Zugang.')) + '</p>'
              : '');
    return (
        '<aside class="pw-nav" aria-label="Navigation">' +
        '<div class="pw-nav__brand"><i class="bi bi-calendar2-week" aria-hidden="true"></i><div><strong>Projektwochen</strong><span>' +
        brandSub +
        '</span></div></div>' +
        '<nav class="pw-nav__list">' +
        navHtml +
        '</nav>' +
        '<div class="pw-nav__roles">' +
        rolesHtml +
        '</div>' +
        '</aside>'
    );
}

function renderTop(state) {
    const aktion = state.aktion;
    const aktionen = Array.isArray(state.aktionen) ? state.aktionen : [];
    const aktionLabel = aktion
        ? esc(aktion.title) +
          ' · ' +
          esc(aktion.startdatum || '?') +
          ' – ' +
          esc(aktion.enddatum || '?') +
          ' · ' +
          statusBadge(aktion.status)
        : '<span class="muted">Keine Aktion geladen</span>';
    const aktionSwitch =
        state.role === 'admin' && aktionen.length > 1
            ? '<label class="pw-top__aktion-select">Aktion<select id="pwAktionSelect">' +
              aktionen
                  .map((a) => {
                      const id = String((a && a.aktionId) || '');
                      const sel = aktion && aktion.aktionId === id ? ' selected' : '';
                      const label =
                          (a.title || id || 'Aktion') +
                          (a.startdatum ? ' (' + a.startdatum + ')' : '') +
                          (a.status ? ' · ' + a.status : '');
                      return '<option value="' + esc(id) + '"' + sel + '>' + esc(label) + '</option>';
                  })
                  .join('') +
              '</select></label>'
            : '';
    const syncSrc = state.spoSync && state.spoSync.source;
    const conn =
        state.localDemoOnly || syncSrc === 'demo'
            ? '<span class="pw-conn pw-conn--demo">Demo lokal</span>'
            : state.ctx || syncSrc === 'live'
              ? '<span class="pw-conn pw-conn--ok">SPO live</span>'
              : state.siteUrl
                ? '<span class="pw-conn pw-conn--wait">Site hinterlegt</span>'
                : '<span class="pw-conn">Nicht verbunden</span>';
    const adminHint =
        state.role === 'admin' && !state.bootstrapped
            ? '<button type="button" class="btn btn-sm" data-pw-view="einstellungen"><i class="bi bi-gear"></i>Systemadministration</button>'
            : state.role === 'admin'
              ? '<button type="button" class="btn btn-sm" id="pwBtnReload" title="Daten neu laden"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>'
              : '';
    return (
        '<div class="pw-top">' +
        '<div class="pw-top__aktion">' +
        '<span class="pw-top__aktion-label">Aktive Projektwoche</span>' +
        '<div class="pw-top__aktion-row">' +
        '<div>' +
        (aktionSwitch || aktionLabel) +
        '</div>' +
        '<div class="pw-top__actions">' +
        conn +
        adminHint +
        '</div></div>' +
        (state.accountName || state.accountEmail
            ? '<div class="muted" style="margin-top:4px;font-size:0.85em">' +
              esc(state.accountName || '') +
              (state.accountEmail ? ' · ' + esc(state.accountEmail) : '') +
              '</div>'
            : '') +
        (state.roleHint && state.view === 'einstellungen'
            ? '<div class="pw-alert pw-alert--warn">' + esc(state.roleHint) + '</div>'
            : '') +
        (state.localDemoOnly && state.view === 'einstellungen'
            ? '<div class="pw-alert pw-alert--info">Lokaler Demo-Modus – Anzeige ohne SharePoint.</div>'
            : '') +
        (state.error ? '<div class="pw-alert pw-alert--err">' + esc(state.error) + '</div>' : '') +
        '</div></div>'
    );
}

function renderDashboardAttention(state) {
    const att = computePwAttention(state);
    const parts = [];
    if (state.role === 'admin' && att.beantragt > 0) {
        parts.push(
            calmBanner(
                'Was braucht Aufmerksamkeit?',
                '<p><strong>' +
                    att.beantragt +
                    '</strong> Angebot/Angebote warten auf Freigabe.</p>',
                { jumpView: 'admin', jumpLabel: 'Zur Freigabe', icon: 'bi-hourglass-split' }
            )
        );
    }
    if (state.role === 'admin' && att.pendingSync > 0) {
        parts.push(
            calmBanner(
                'Bookings-Sync',
                '<p><strong>' +
                    att.pendingSync +
                    '</strong> freigegebene Angebot/Angebote noch nicht in Microsoft Bookings.</p>',
                { tone: 'ops', jumpView: 'bookings', jumpLabel: 'Sync öffnen', icon: 'bi-cloud-upload' }
            )
        );
    }
    if (state.role === 'lehrer' && !att.myCount && state.bootstrapped) {
        parts.push(
            calmBanner(
                'Loslegen',
                '<p>Noch kein eigenes Angebot – reichen Sie eines für die Projektwoche ein.</p>',
                { jumpView: 'neu', jumpLabel: 'Neues Angebot', icon: 'bi-plus-circle' }
            )
        );
    }
    if (state.role === 'schueler') {
        parts.push(
            calmBanner(
                'Buchung',
                '<p>Angebote hier lesen; buchen Sie in der <strong>Microsoft Bookings</strong>-App (Link nach Freigabe/Sync).</p>',
                { icon: 'bi-info-circle' }
            )
        );
    }
    if (!state.bootstrapped && !state.localDemoOnly) {
        if (state.role === 'admin') {
            parts.push(
                calmBanner(
                    'Setup',
                    '<p>SharePoint verbinden oder Demo laden – unter System.</p>',
                    { jumpView: 'einstellungen', jumpLabel: 'System', icon: 'bi-gear' }
                )
            );
        } else {
            parts.push(
                quietEmpty(
                    'Noch keine Daten',
                    'Bitte warten Sie auf die Freischaltung durch die Administration.'
                )
            );
        }
    }
    return parts.join('');
}

function renderDashboard(state) {
    const items = scopedItems(state);
    const kpis = computeDashboardKpis(items, state.aktion);
    const emptyBody =
        state.bootstrapped && !items.length
            ? quietEmpty(
                  'Keine Angebote in dieser Sicht',
                  state.role === 'lehrer'
                      ? 'Filter zurücksetzen oder neues Angebot anlegen.'
                      : 'Filter zurücksetzen oder warten, bis Angebote freigegeben sind.',
                  state.role === 'lehrer'
                      ? '<button type="button" class="btn btn-sm" data-pw-view="neu"><i class="bi bi-plus-circle"></i>Neues Angebot</button>'
                      : ''
              )
            : '';
    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head">' +
        '<div><h2>Dashboard</h2>' +
        '<p class="pw-view__sub">Überblick zur aktiven Projektwoche – filtern, dann Plan oder Liste öffnen.</p></div>' +
        '</header>' +
        renderDashboardAttention(state) +
        filterBar(state) +
        '<div class="pw-kpi-grid">' +
        kpiCard('Angebote', kpis.total, 'bi-collection', '') +
        kpiCard('Beantragt', kpis.beantragt, 'bi-hourglass-split', 'warn') +
        kpiCard('Freigegeben', kpis.freigegeben, 'bi-check2-circle', 'ok') +
        kpiCard('Abgelehnt', kpis.abgelehnt, 'bi-x-circle', 'bad') +
        kpiCard('Plätze (Summe)', kpis.kapTotal, 'bi-people', '') +
        kpiCard('Buchung offen', kpis.buchungOffen, 'bi-unlock', 'ok') +
        kpiCard('Buchung gesperrt', kpis.buchungGesperrt, 'bi-lock', '') +
        '</div>' +
        emptyBody +
        (state.bootstrapped && items.length
            ? '<p class="pw-view__meta"><i class="bi bi-funnel" aria-hidden="true"></i> Gefilterte Sicht: <strong>' +
              items.length +
              '</strong> Angebot/Angebote</p>'
            : '') +
        '</section>'
    );
}

function kpiCard(label, value, icon, tone) {
    return (
        '<div class="pw-kpi' +
        (tone ? ' pw-kpi--' + tone : '') +
        '">' +
        '<div class="pw-kpi__top">' +
        (icon ? '<i class="bi ' + esc(icon) + '" aria-hidden="true"></i>' : '') +
        '<span class="pw-kpi__value">' +
        esc(String(value)) +
        '</span></div>' +
        '<div class="pw-kpi__label">' +
        esc(label) +
        '</div></div>'
    );
}

function renderPlan(state) {
    const items = scopedItems(state);
    const grid = buildWeekPlan(items);
    const displayRows = buildWeekPlanDisplayRows(grid);
    const slotHeads = PLAN_SLOTS.map((s) => '<th>' + esc(s) + '</th>').join('');

    function planCards(list) {
        return (list || [])
            .map((a) => {
                const occ = occupancyLabel(state, a);
                return (
                    '<button type="button" class="pw-plan-card" data-pw-detail="' +
                    esc(a.itemId || a.angebotId) +
                    '"><strong>' +
                    esc(a.title) +
                    '</strong><span>' +
                    statusBadge(a.status) +
                    ' · ' +
                    (occ ? esc(occ) + ' TN · ' : '') +
                    esc(String(a.kapazitaet || '')) +
                    ' Plätze' +
                    (a.syncStatus ? ' · sync:' + esc(a.syncStatus) : '') +
                    '</span></button>'
                );
            })
            .join('');
    }

    const rows = displayRows
        .map((day) => {
            const bandCount = day.bands.length;
            return day.bands
                .map((band, bi) => {
                    const th =
                        bi === 0
                            ? '<th' +
                              (bandCount > 1 ? ' rowspan="' + bandCount + '"' : '') +
                              '>' +
                              esc(day.tag) +
                              '</th>'
                            : '';
                    if (band.type === 'ganztags') {
                        const cards = planCards(band.items);
                        return (
                            '<tr>' +
                            th +
                            '<td class="pw-plan__ganztags" colspan="' +
                            PLAN_SLOTS.length +
                            '">' +
                            (cards || '<span class="muted">–</span>') +
                            '</td></tr>'
                        );
                    }
                    const cells = PLAN_SLOTS.map((slot) => {
                        const cards = planCards(band.bySlot[slot]);
                        return '<td>' + (cards || '<span class="muted">–</span>') + '</td>';
                    }).join('');
                    return '<tr>' + th + cells + '</tr>';
                })
                .join('');
        })
        .join('');

    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head"><div><h2>Wochenplan</h2>' +
        '<p class="pw-view__sub">Angebote nach Tag und Slot – ganztägige spannen Vormittag bis Abend.</p></div></header>' +
        filterBar(state) +
        '<div class="pw-plan-wrap"><table class="pw-plan"><thead><tr><th>Tag</th>' +
        slotHeads +
        '</tr></thead><tbody>' +
        rows +
        '</tbody></table></div>' +
        '<p class="muted" style="margin-top:8px;font-size:0.85em">Ganztägige Angebote spannen Vormittag–Abend.</p>' +
        '</section>'
    );
}

function renderKalender(state) {
    const items = scopedItems(state);
    const cells = buildMonthCells(state.calYear, state.calMonth, items);
    const monthLabel = new Date(state.calYear, state.calMonth, 1).toLocaleDateString('de-AT', {
        month: 'long',
        year: 'numeric'
    });
    const head = TAGE.map((t) => '<th>' + esc(t) + '</th>').join('');
    let body = '';
    for (let i = 0; i < cells.length; i += 7) {
        body += '<tr>';
        for (let j = 0; j < 7; j++) {
            const c = cells[i + j];
            if (!c.iso) {
                body += '<td class="pw-cal__empty"></td>';
                continue;
            }
            const bits = (c.items || [])
                .slice(0, 4)
                .map(
                    (a) =>
                        '<button type="button" class="pw-cal__item" data-pw-detail="' +
                        esc(a.itemId || a.angebotId) +
                        '">' +
                        esc(a.title) +
                        '</button>'
                )
                .join('');
            body +=
                '<td class="pw-cal__day"><div class="pw-cal__num">' +
                c.day +
                '</div>' +
                bits +
                (c.items.length > 4 ? '<div class="muted">+' + (c.items.length - 4) + '</div>' : '') +
                '</td>';
        }
        body += '</tr>';
    }
    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head"><div><h2>Kalender</h2>' +
        '<p class="pw-view__sub">Monatsübersicht der gefilterten Angebote.</p></div></header>' +
        filterBar(state) +
        '<div class="pw-cal-nav">' +
        '<button type="button" class="btn" id="pwCalPrev"><i class="bi bi-chevron-left"></i></button>' +
        '<strong>' +
        esc(monthLabel) +
        '</strong>' +
        '<button type="button" class="btn" id="pwCalNext"><i class="bi bi-chevron-right"></i></button>' +
        '</div>' +
        '<table class="pw-cal"><thead><tr>' +
        head +
        '</tr></thead><tbody>' +
        body +
        '</tbody></table></section>'
    );
}

function renderListe(state) {
    const items = scopedItems(state);
    const labels = labelMaps(state.stammdaten);
    const scope = scopeFromState(state);
    const rows = items
        .slice()
        .sort((a, b) => String(a.datum).localeCompare(String(b.datum)))
        .map((a) => {
            const buchung = effectiveBuchungAb(a, state.aktion);
            const open = a.status === 'freigegeben' && isBookingOpen(buchung);
            const edit = canEditAngebot(a, scope);
            const occ = occupancyLabel(state, a);
            return (
                '<tr>' +
                '<td><button type="button" class="pw-link" data-pw-detail="' +
                esc(a.itemId || a.angebotId) +
                '">' +
                esc(a.title) +
                '</button></td>' +
                '<td>' +
                esc(a.datum) +
                ' ' +
                esc(a.tag || '') +
                '</td>' +
                '<td>' +
                esc(a.slot) +
                '</td>' +
                '<td>' +
                esc(labels.teachers[a.lehrerCode] || a.lehrerCode || a.lehrerEmail) +
                '</td>' +
                '<td>' +
                statusBadge(a.status) +
                '</td>' +
                '<td>' +
                (occ ? esc(occ) : esc(String(a.kapazitaet))) +
                '</td>' +
                '<td>' +
                (buchung
                    ? esc(toDatetimeLocalValue(buchung).replace('T', ' ')) +
                      (open ? ' <span class="pw-badge pw-badge--freigegeben">offen</span>' : ' <span class="pw-badge">gesperrt</span>')
                    : '–') +
                '</td>' +
                '<td>' +
                esc(a.syncStatus || '–') +
                '</td>' +
                '<td class="pw-table__actions">' +
                (edit
                    ? '<button type="button" class="btn" data-pw-edit="' + esc(a.itemId) + '">Bearbeiten</button>'
                    : '') +
                (a.bookingsBookingUrl
                    ? ' <a class="btn" href="' +
                      esc(a.bookingsBookingUrl) +
                      '" target="_blank" rel="noopener">Bookings</a>'
                    : '') +
                '</td></tr>'
            );
        })
        .join('');
    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head"><div><h2>Liste</h2>' +
        '<p class="pw-view__sub">Tabellarische Übersicht – Details per Klick auf den Titel.</p></div></header>' +
        filterBar(state) +
        '<div class="pw-table-wrap"><table class="pw-table"><thead><tr>' +
        '<th>Titel</th><th>Datum</th><th>Slot</th><th>Lehrer</th><th>Status</th><th>Belegung</th><th>Buchung ab</th><th>Sync</th><th></th>' +
        '</tr></thead><tbody>' +
        (rows || '<tr><td colspan="9" class="muted">Keine Angebote</td></tr>') +
        '</tbody></table></div></section>'
    );
}

function renderNeu(state) {
    const editing = !!state.editingItemId;
    const isAdmin = state.role === 'admin';
    const steps =
        '<div class="pw-form-steps" aria-hidden="true">' +
        formStep('bi-card-heading', 'Inhalt') +
        formStep('bi-people', 'Betreuung') +
        formStep('bi-calendar3', 'Termin & Ort') +
        formStep('bi-cash-coin', 'Kosten') +
        (isAdmin ? formStep('bi-shield-lock', 'Verwaltung') : '') +
        '</div>';

    return (
        '<section class="pw-view pw-neu">' +
        '<header class="pw-detail__hero pw-neu__hero">' +
        '<p class="pw-detail__kicker">' +
        (editing ? 'Bearbeiten' : 'Antrag') +
        '</p>' +
        '<h2>' +
        (editing ? 'Angebot bearbeiten' : 'Neues Angebot') +
        '</h2>' +
        '<p class="pw-neu__lead">' +
        (editing
            ? 'Alle Felder sind nach Themen gruppiert – speichern, wenn alles passt.'
            : 'Fülle die Gruppen aus: Inhalt, Termin, Betreuung und Kosten. Danach kannst du den Antrag absenden.') +
        '</p>' +
        steps +
        '</header>' +
        '<div class="pw-neu__form">' +
        renderAngebotForm(state, {
            editing: editing,
            submitLabel: editing ? 'Speichern' : 'Antrag stellen'
        }) +
        '</div></section>'
    );
}

function formStep(icon, label) {
    return (
        '<span class="pw-form-step"><i class="bi ' +
        esc(icon) +
        '" aria-hidden="true"></i>' +
        esc(label) +
        '</span>'
    );
}

/**
 * Gemeinsames Angebot-Formular (Neu / Detail) – thematische Gruppen.
 * @param {object} state
 * @param {{ editing?: boolean, submitLabel?: string, showCancel?: boolean }} [opts]
 */
function renderAngebotForm(state, opts) {
    const o = opts || {};
    const form = state.form || emptyForm();
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    const classes = (state.stammdaten && state.stammdaten.classes) || [];
    const classCodes = classes.map((c) => c.code).filter(Boolean);
    const draft = { ...form, angebotId: form.angebotId || 'draft' };
    const peers = state.items.filter((i) => i.itemId !== state.editingItemId);
    const v = validateAngebot({
        draft,
        existing: peers,
        aktion: state.aktion,
        classCodes
    });
    const isAdmin = state.role === 'admin';
    const editing = !!o.editing || !!state.editingItemId;
    const showVal = editing || String(form.title || '').trim() || String(form.datum || '').trim();
    const showCancel = o.showCancel !== false && editing && state.view !== 'detail';
    const deleteId = state.editingItemId || state.detailId || '';
    const showDelete = state.view === 'detail' && editing && !!deleteId;

    function selControl(id, optionsHtml) {
        return (
            '<div class="pw-select"><select id="' +
            id +
            '" class="pw-select__control">' +
            optionsHtml +
            '</select><i class="bi bi-chevron-down pw-select__caret" aria-hidden="true"></i></div>'
        );
    }

    return (
        '<form class="pw-form pw-form--grouped" id="pwForm">' +
        '<div class="pw-form__layout">' +
        '<div class="pw-form__main">' +
        formSection(
            'bi-card-heading',
            'Inhalt',
            'Titel, Kategorie und Texte für Lehrkräfte und Eltern',
            field('Titel', 'bi-type', '<input id="pwFTitle" required value="' + esc(form.title) + '">', 'full') +
                field(
                    'Kategorie',
                    'bi-tags',
                    selControl('pwFKat', optionList(KATEGORIEN, null, null, form.kategorie)),
                    'half'
                ) +
                field(
                    'Beschreibung',
                    'bi-text-left',
                    '<textarea id="pwFDesc" rows="3" placeholder="Was erwartet die Schüler:innen?">' +
                        esc(form.beschreibung) +
                        '</textarea>',
                    'full'
                ) +
                field(
                    'Hinweis Eltern',
                    'bi-info-circle',
                    '<textarea id="pwFEltern" rows="2" placeholder="z. B. Einverständnis / Kosten empfohlen">' +
                        esc(form.hinweisEltern) +
                        '</textarea>',
                    'full'
                )
        ) +
        formSection(
            'bi-people',
            'Teilnehmer & Betreuung',
            'Kapazität, Zielgruppen und verantwortliche Lehrkräfte',
            field(
                'Kapazität',
                'bi-person-bounding-box',
                '<input type="number" id="pwFKap" min="1" value="' + esc(String(form.kapazitaet)) + '">'
            ) +
                field(
                    'Zielklassen',
                    'bi-mortarboard',
                    '<input id="pwFKlassen" list="pwKlassenList" value="' +
                        esc(form.zielklassen) +
                        '" placeholder="alle oder 1AK;1BK">' +
                        '<datalist id="pwKlassenList">' +
                        classes.map((c) => '<option value="' + esc(c.code) + '">').join('') +
                        '</datalist>',
                    'half'
                ) +
                field(
                    'Lehrer',
                    'bi-person-badge',
                    selControl(
                        'pwFLehrer',
                        '<option value="">–</option>' + optionList(teachers, 'code', 'name', form.lehrerCode)
                    ),
                    'half'
                ) +
                field(
                    'Lehrer-E-Mail',
                    'bi-envelope',
                    '<input id="pwFLehrerMail" type="email" value="' + esc(form.lehrerEmail) + '">',
                    'half'
                ) +
                field(
                    'Begleitung',
                    'bi-person-plus',
                    '<input id="pwFBegleitung" value="' +
                        esc(form.begleitung) +
                        '" placeholder="Code oder Name">',
                    'half'
                )
        ) +
        '</div>' +
        '<aside class="pw-form__aside">' +
        formSection(
            'bi-calendar3',
            'Termin & Ort',
            'Wann und wo findet das Angebot statt?',
            field('Datum', 'bi-calendar-event', '<input type="date" id="pwFDatum" value="' + esc(form.datum) + '">', 'full') +
                field(
                    'Wochentag',
                    'bi-calendar-week',
                    selControl(
                        'pwFTag',
                        '<option value="">auto</option>' + optionList(TAGE, null, null, form.tag)
                    )
                ) +
                field('Slot', 'bi-grid-3x2', selControl('pwFSlot', optionList(SLOTS, null, null, form.slot))) +
                field('Start', 'bi-clock', '<input type="time" id="pwFStart" value="' + esc(form.startzeit) + '">') +
                field('Ende', 'bi-clock-history', '<input type="time" id="pwFEnd" value="' + esc(form.endzeit) + '">') +
                field(
                    'Ort',
                    'bi-geo-alt',
                    '<input id="pwFOrt" value="' + esc(form.ort) + '" placeholder="Ort / Adresse">',
                    'full'
                ) +
                field(
                    'Treffpunkt',
                    'bi-signpost',
                    '<input id="pwFTreff" value="' + esc(form.treffpunkt) + '" placeholder="z. B. Schulhof">',
                    'full'
                ),
            'pw-form-section__body--aside'
        ) +
        formSection(
            'bi-cash-coin',
            'Kosten',
            'Preis und Hinweis für Buchung / Eltern',
            field(
                'Preis (€)',
                'bi-currency-euro',
                '<input type="number" id="pwFPreis" min="0" step="0.5" value="' + esc(String(form.preisEuro)) + '">',
                'full'
            ) +
                field(
                    'Kostenhinweis',
                    'bi-receipt',
                    '<input id="pwFKosten" value="' +
                        esc(form.kostenHinweis) +
                        '" placeholder="z. B. inkl. Eintritt">',
                    'full'
                ),
            'pw-form-section__body--aside'
        ) +
        (isAdmin
            ? formSection(
                  'bi-shield-lock',
                  'Verwaltung',
                  'Nur für Admin – Buchungsfenster und interne Notizen',
                  field(
                      'Buchung möglich ab',
                      'bi-unlock',
                      '<input type="datetime-local" id="pwFBuchungAb" value="' +
                          esc(
                              toDatetimeLocalValue(
                                  form.buchungAb || (state.aktion && state.aktion.buchungAbDefault) || ''
                              )
                          ) +
                          '">',
                      'full'
                  ) +
                      field(
                          'Interne Notiz',
                          'bi-journal-text',
                          '<textarea id="pwFNotiz" rows="2" placeholder="Nur intern sichtbar">' +
                              esc(form.notizIntern) +
                              '</textarea>',
                          'full'
                      ),
                  'pw-form-section__body--aside'
              )
            : '') +
        '</aside></div>' +
        validationBlock(showVal ? v : { errors: [], warnings: [] }) +
        '<div class="pw-form__actions pw-form__actions--sticky">' +
        '<button type="button" class="btn btn-success" id="pwBtnSave"><i class="bi bi-check2-circle"></i>' +
        esc(o.submitLabel || (editing ? 'Speichern' : 'Beantragen')) +
        '</button>' +
        (showCancel ? '<button type="button" class="btn" id="pwBtnCancelEdit">Abbrechen</button>' : '') +
        (showDelete
            ? '<button type="button" class="btn pw-form__danger" data-pw-del="' +
              esc(deleteId) +
              '"><i class="bi bi-trash"></i>Löschen</button>'
            : '') +
        '</div></form>'
    );
}

function formSection(icon, title, hint, bodyHtml, bodyMod) {
    return (
        '<section class="pw-form-section">' +
        '<header class="pw-form-section__head">' +
        '<span class="pw-form-section__icon" aria-hidden="true"><i class="bi ' +
        esc(icon) +
        '"></i></span>' +
        '<div><h3 class="pw-form-section__title">' +
        esc(title) +
        '</h3>' +
        (hint ? '<p class="pw-form-section__hint">' + esc(hint) + '</p>' : '') +
        '</div></header>' +
        '<div class="pw-form-section__body' +
        (bodyMod ? ' ' + esc(bodyMod) : '') +
        '">' +
        bodyHtml +
        '</div></section>'
    );
}

/**
 * @param {string} label
 * @param {string} icon
 * @param {string} control
 * @param {'full'|'half'|string} [span]
 */
function field(label, icon, control, span) {
    const spanClass =
        span === 'full' ? ' pw-field--full' : span === 'half' ? ' pw-field--half' : '';
    return (
        '<div class="pw-field' +
        spanClass +
        '"><label>' +
        (icon ? '<i class="bi ' + esc(icon) + '" aria-hidden="true"></i>' : '') +
        esc(label) +
        '</label>' +
        control +
        '</div>'
    );
}

function validationBlock(v) {
    let html = '';
    if (v.errors && v.errors.length) {
        html +=
            '<div class="pw-alert pw-alert--err"><ul>' +
            v.errors.map((e) => '<li>' + esc(e) + '</li>').join('') +
            '</ul></div>';
    }
    if (v.warnings && v.warnings.length) {
        html +=
            '<div class="pw-alert pw-alert--warn"><ul>' +
            v.warnings.map((e) => '<li>' + esc(e) + '</li>').join('') +
            '</ul></div>';
    }
    return html;
}

function renderMeine(state) {
    const scope = { ...scopeFromState(state), onlyMine: true };
    const items = filterAngebote(state.items, state.filters, scope, state.aktion);
    state._meineCount = items.length;
    const rows = items
        .map((a) => {
            return (
                '<tr><td>' +
                esc(a.title) +
                '</td><td>' +
                esc(a.datum) +
                '</td><td>' +
                statusBadge(a.status) +
                '</td><td>' +
                '<button type="button" class="btn" data-pw-detail="' +
                esc(a.itemId) +
                '">Details</button> ' +
                (canEditAngebot(a, scope)
                    ? '<button type="button" class="btn" data-pw-edit="' + esc(a.itemId) + '">Bearbeiten</button> '
                    : '') +
                (canEditAngebot(a, scope)
                    ? '<button type="button" class="btn" data-pw-del="' + esc(a.itemId) + '">Löschen</button>'
                    : '') +
                '</td></tr>'
            );
        })
        .join('');
    return (
        '<section class="pw-view"><h2>Meine Angebote</h2>' +
        (rows
            ? '<div class="pw-table-wrap"><table class="pw-table"><thead><tr><th>Titel</th><th>Datum</th><th>Status</th><th></th></tr></thead><tbody>' +
              rows +
              '</tbody></table></div>'
            : quietEmpty(
                  'Noch keine eigenen Angebote',
                  'Legen Sie ein Angebot an – bis zur Freigabe können Sie es bearbeiten.',
                  '<button type="button" class="btn btn-sm" data-pw-view="neu"><i class="bi bi-plus-circle"></i>Neues Angebot</button>'
              )) +
        '</section>'
    );
}

function renderAdmin(state) {
    if (!canAdminDecide(scopeFromState(state))) {
        return '<section class="pw-view"><p class="pw-alert pw-alert--warn">Nur für Admin.</p></section>';
    }
    const queue = state.items.filter((a) => a.status === 'beantragt');
    const freigegeben = state.items.filter((a) => a.status === 'freigegeben');
    const aktion = state.aktion;

    const queueRows = queue
        .map((a) => {
            return (
                '<tr><td>' +
                esc(a.title) +
                '</td><td>' +
                esc(a.datum) +
                ' · ' +
                esc(a.slot) +
                '</td><td>' +
                esc(a.lehrerCode || a.lehrerEmail) +
                '</td><td>' +
                esc(String(a.kapazitaet)) +
                ' / ' +
                esc(String(a.preisEuro)) +
                ' €</td><td class="pw-table__actions">' +
                '<button type="button" class="btn btn-success" data-pw-approve="' +
                esc(a.itemId) +
                '">Freigeben</button> ' +
                '<button type="button" class="btn" data-pw-reject="' +
                esc(a.itemId) +
                '">Ablehnen</button> ' +
                '<button type="button" class="btn" data-pw-detail="' +
                esc(a.itemId) +
                '">Details</button>' +
                '</td></tr>'
            );
        })
        .join('');

    const buchungRows = freigegeben
        .map((a) => {
            const val = toDatetimeLocalValue(a.buchungAb || (aktion && aktion.buchungAbDefault) || '');
            return (
                '<tr><td>' +
                esc(a.title) +
                '</td><td>' +
                esc(a.datum) +
                '</td><td>' +
                '<input type="datetime-local" data-pw-buchung-id="' +
                esc(a.itemId) +
                '" value="' +
                esc(val) +
                '">' +
                '</td><td>' +
                '<button type="button" class="btn" data-pw-buchung-save="' +
                esc(a.itemId) +
                '">Speichern</button>' +
                '</td></tr>'
            );
        })
        .join('');

    return (
        '<section class="pw-view"><h2>Administration</h2>' +
        '<h3>Freigabe-Queue (' +
        queue.length +
        ')</h3>' +
        '<div class="pw-table-wrap"><table class="pw-table"><thead><tr><th>Titel</th><th>Wann</th><th>Lehrer</th><th>Kap./Preis</th><th></th></tr></thead><tbody>' +
        (queueRows || '<tr><td colspan="5" class="muted">Keine offenen Anträge</td></tr>') +
        '</tbody></table></div>' +
        '<h3 style="margin-top:1.5rem">Buchungsstart (freigegebene Angebote)</h3>' +
        '<p class="muted">Ab diesem Zeitpunkt dürfen Schüler in Bookings buchen (Phase 3). Leer = Standard der Aktion.</p>' +
        '<div class="pw-table-wrap"><table class="pw-table"><thead><tr><th>Angebot</th><th>Datum</th><th>Buchung ab</th><th></th></tr></thead><tbody>' +
        (buchungRows || '<tr><td colspan="4" class="muted">Noch keine freigegebenen Angebote</td></tr>') +
        '</tbody></table></div>' +
        (aktion
            ? '<h3 style="margin-top:1.5rem">Aktions-Standard</h3>' +
              '<div class="pw-form__grid">' +
              field(
                  'Standard-Buchungsstart',
                  '<input type="datetime-local" id="pwAktionBuchungDefault" value="' +
                      esc(toDatetimeLocalValue(aktion.buchungAbDefault || '')) +
                      '">'
              ) +
              '</div>' +
              '<button type="button" class="btn" id="pwBtnAktionSave"><i class="bi bi-save"></i>Aktion speichern</button>'
            : '') +
        '<p class="muted" style="margin-top:1rem">Bookings: siehe Ansicht <strong>Bookings</strong> (Sync-Button). Teilnehmer: Ansicht <strong>Teilnehmer</strong>.</p>' +
        '</section>'
    );
}

function renderEinstellungen(state) {
    if (!canAdminDecide(scopeFromState(state))) {
        return '<section class="pw-view"><p class="pw-alert pw-alert--warn">Nur für Projektwochen-Admin.</p></section>';
    }
    const siteHost = (() => {
        try {
            return state.siteUrl ? new URL(state.siteUrl).host : '';
        } catch {
            return '';
        }
    })();
    const stammLive = state.stammdatenLive || {};
    const spoCount = (state.items || []).filter(
        (a) => a && a.itemId && String(a.itemId).indexOf('local-') !== 0
    ).length;
    return renderSystemAdminPage({
        prefix: 'pw',
        activeTab: state.systemTab,
        dataAttr: 'data-pw-sys-tab',
        ariaLabel: 'Systemadministration Projektwochen',
        panels: [
            {
                id: 'site',
                step: '1',
                title: 'Datenquelle',
                leadHtml:
                    'SharePoint und Listen sind Admin-/IT-Sache. Lehrkräfte sehen nur Planung und eigene Angebote.',
                body:
                    '<div class="ms-sys-field"><label for="pwSiteUrl">SharePoint-Website</label>' +
                    '<div class="ms-sys-field__row"><input type="url" id="pwSiteUrl" class="ms-sys-input" value="' +
                    esc(state.siteUrl) +
                    '" placeholder="https://…sharepoint.com/sites/…" spellcheck="false" autocomplete="off">' +
                    '<button type="button" class="btn ms-sys-btn-primary" id="pwBtnLoad"><i class="bi bi-cloud-download"></i>Speichern &amp; laden</button></div></div>' +
                    (siteHost ? '<p class="muted">Aktuell: ' + esc(siteHost) + '</p>' : '') +
                    '<p class="muted">Stammdaten: ' +
                    esc(String(stammLive.source || 'local')) +
                    ' · ' +
                    esc(String(stammLive.teacherCount || 0)) +
                    ' Lehrkräfte / ' +
                    esc(String(stammLive.classCount || 0)) +
                    ' Klassen' +
                    (stammLive.error ? ' · ' + esc(stammLive.error) : '') +
                    '</p>' +
                    '<p class="muted">SPO Live Sync: ' +
                    esc(formatPwSpoSyncLine(state.spoSync)) +
                    ' · Angebote: ' +
                    esc(String((state.items || []).length)) +
                    ' · mit SPO-ID: ' +
                    esc(String(spoCount)) +
                    '</p>' +
                    '<div class="ms-sys-actions">' +
                    '<button type="button" class="btn ms-sys-btn-secondary" id="pwBtnListsHealth"><i class="bi bi-heart-pulse"></i>Listen prüfen</button>' +
                    '<button type="button" class="btn ms-sys-btn-secondary" id="pwBtnListsCreate"><i class="bi bi-plus-square"></i>Listen anlegen / aktualisieren</button>' +
                    '<a class="btn ms-sys-btn-ghost" href="sharepoint-liste-projektwochen.html"><i class="bi bi-box-arrow-up-right"></i>Listen-Assistent</a>' +
                    '</div><pre class="pw-setup-log" id="pwSetupLog" hidden></pre>'
            },
            {
                id: 'perms',
                step: '2',
                title: 'Berechtigungen',
                leadHtml: 'Gruppen und Listenrechte setzt der Listen-Assistent, nicht diese Ansicht.',
                body:
                    '<div class="ms-sys-links"><a class="btn ms-sys-btn-secondary" href="sharepoint-liste-projektwochen.html"><i class="bi bi-shield-check"></i>Listen-Assistent</a></div>'
            },
            {
                id: 'kalender',
                step: '3',
                title: 'Kalender',
                leadHtml:
                    'Wochenplan und Buchungsstart bleiben in den Arbeitsansichten. Der Export nach Microsoft Bookings ist ein eigener Schritt.',
                body:
                    '<div class="ms-sys-actions">' +
                    '<button type="button" class="btn ms-sys-btn-secondary" data-pw-view="kalender"><i class="bi bi-calendar3"></i>Kalender</button>' +
                    '<button type="button" class="btn ms-sys-btn-secondary" data-pw-view="bookings"><i class="bi bi-box-arrow-up-right"></i>Bookings-Sync</button>' +
                    '<button type="button" class="btn ms-sys-btn-ghost" data-pw-view="admin"><i class="bi bi-shield-check"></i>Freigabe &amp; Buchungsstart</button>' +
                    '</div>'
            },
            {
                id: 'wartung',
                step: '4',
                title: 'Wartung',
                leadHtml:
                    'Demo nur mit Button oder ?demo=1. Lokal first – nicht automatisch nach SharePoint. Rollen: ?demoRole=1.',
                body:
                    '<div class="ms-sys-actions">' +
                    '<button type="button" class="btn ms-sys-btn-secondary" id="pwBtnDemo"><i class="bi bi-database"></i>Demo laden</button>' +
                    '<label class="btn ms-sys-btn-ghost" for="pwImportJsonAdmin"><i class="bi bi-upload"></i>JSON-Datei wählen</label>' +
                    '<input type="file" id="pwImportJsonAdmin" accept=".json,application/json" hidden>' +
                    '<a class="btn ms-sys-btn-ghost" href="../tenant.html"><i class="bi bi-gear"></i>Stammdaten</a>' +
                    '<a class="btn ms-sys-btn-ghost" href="../hilfe.html#tool-projektwochen"><i class="bi bi-life-preserver"></i>Hilfe</a>' +
                    '</div>' +
                    '<p class="muted" style="margin:10px 0 0"><code>docs/demo-data/projektwochen-demo.json</code></p>'
            },
            {
                id: 'danger',
                title: 'Danger Zone',
                danger: true,
                leadHtml:
                    'Diese Ansicht löscht keine SharePoint-Listen. Demo ersetzt nur die lokale Anzeige, bis du neu lädst.',
                body: '<p class="muted">Destruktives Leeren der Produktivlisten bleibt dem Listen-Assistenten bzw. SharePoint vorbehalten.</p>'
            }
        ]
    });
}

function renderBookings(state) {
    if (!canAdminDecide(scopeFromState(state))) {
        return '<section class="pw-view"><p class="pw-alert pw-alert--warn">Nur für Admin.</p></section>';
    }
    const aktion = state.aktion;
    const freigegeben = state.items.filter((a) => a.status === 'freigegeben');
    const synced = freigegeben.filter((a) => a.bookingsServiceId);
    const rows = freigegeben
        .map((a) => {
            return (
                '<tr>' +
                '<td><label><input type="checkbox" class="pw-sync-check" data-pw-sync-id="' +
                esc(a.itemId) +
                '" checked> ' +
                esc(a.title) +
                '</label></td>' +
                '<td>' +
                esc(a.datum) +
                '</td>' +
                '<td>' +
                esc(a.syncStatus || '–') +
                '</td>' +
                '<td>' +
                (a.bookingsServiceId ? esc(a.bookingsServiceId.slice(0, 12)) + '…' : '–') +
                '</td>' +
                '<td>' +
                (a.bookingsBookingUrl
                    ? '<a href="' + esc(a.bookingsBookingUrl) + '" target="_blank" rel="noopener">Link</a>'
                    : '–') +
                '</td>' +
                '<td class="pw-table__actions">' +
                '<button type="button" class="btn" data-pw-sync-one="' +
                esc(a.itemId) +
                '">Sync</button>' +
                '</td></tr>'
            );
        })
        .join('');

    const pending = countPendingBookingsSync(state.items);
    const summary = state.bookingsSyncSummary;
    const summaryBanner =
        summary && summary.headline
            ? calmBanner(
                  summary.tone === 'warn' ? 'Sync mit Fehlern' : 'Sync-Ergebnis',
                  '<p><strong>' +
                      esc(summary.headline) +
                      '</strong>' +
                      (summary.detail ? '<br>' + esc(summary.detail) : '') +
                      '</p>',
                  {
                      tone: summary.tone === 'ok' ? 'action' : 'ops',
                      icon: summary.tone === 'warn' ? 'bi-exclamation-triangle' : 'bi-check2-circle'
                  }
              )
            : pending
              ? calmBanner(
                    'Nächster Schritt',
                    '<p><strong>' +
                        pending +
                        '</strong> freigegebene Angebot/Angebote brauchen einen Sync nach Microsoft Bookings.</p>',
                    { tone: 'ops', icon: 'bi-cloud-upload' }
                )
              : '';
    const failHints = freigegeben
        .filter((a) => String(a.syncStatus || '').toLowerCase() === 'fehler' && a.syncFehler)
        .slice(0, 3)
        .map(
            (a) =>
                '<li><strong>' +
                esc(a.title) +
                ':</strong> ' +
                esc(a.syncFehler) +
                '</li>'
        )
        .join('');
    return (
        '<section class="pw-view"><h2>Microsoft Bookings</h2>' +
        '<p class="pw-view__sub">Klarer Export-Schritt: 1 Projektwoche = 1 Business, freigegebene Angebote = Dienste. Sync nur per Button.</p>' +
        summaryBanner +
        (aktion
            ? '<div class="pw-alert pw-alert--info">Aktion: <strong>' +
              esc(aktion.title) +
              '</strong><br>Business: ' +
              esc(aktion.bookingsBusinessName || aktion.bookingsBusinessId || '(noch nicht angelegt)') +
              '</div>'
            : '<div class="pw-alert pw-alert--warn">Keine aktive Aktion.</div>') +
        '<div class="pw-form__actions" style="margin-bottom:12px">' +
        '<button type="button" class="btn btn-success" id="pwBtnEnsureBiz"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-building"></i>Business anlegen/binden</button> ' +
        '<button type="button" class="btn btn-success" id="pwBtnSyncSelected"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-arrow-repeat"></i>Ausgewählte syncen</button> ' +
        '<button type="button" class="btn" id="pwBtnSyncAll"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-cloud-upload"></i>Alle freigegebenen syncen (' +
        freigegeben.length +
        ')</button> ' +
        '<button type="button" class="btn" id="pwBtnLoadTn"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-people"></i>Teilnehmer laden</button>' +
        '</div>' +
        '<p class="muted">Freigegeben: ' +
        freigegeben.length +
        ' · mit Service: ' +
        synced.length +
        ' · Sync ausstehend: ' +
        pending +
        (state.localDemoOnly ? ' · Demo: Sync lokal simuliert' : '') +
        '</p>' +
        (failHints
            ? '<div class="pw-alert pw-alert--warn"><p>Letzte Sync-Fehler:</p><ul>' +
              failHints +
              '</ul></div>'
            : '') +
        (rows
            ? '<div class="pw-table-wrap"><table class="pw-table"><thead><tr>' +
              '<th>Angebot</th><th>Datum</th><th>Sync</th><th>Service</th><th>URL</th><th></th>' +
              '</tr></thead><tbody>' +
              rows +
              '</tbody></table></div>'
            : quietEmpty(
                  'Keine freigegebenen Angebote',
                  'Zuerst unter Freigabe Angebote freigeben, dann hier nach Bookings syncen.',
                  '<button type="button" class="btn btn-sm" data-pw-view="admin">Zur Freigabe</button>'
              )) +
        '<details class="pw-tech" style="margin-top:12px"><summary>Technisches Protokoll</summary>' +
        '<pre class="pw-log" id="pwBookingsLog">' +
        esc(state.bookingsLog || 'Noch keine Bookings-Aktion.') +
        '</pre></details>' +
        '</section>'
    );
}

function renderTeilnehmer(state) {
    const f = state.attendeeFilter || {};
    const angebote = state.items.filter((a) => a.status === 'freigegeben');
    let rows = Array.isArray(state.attendeeRows) ? state.attendeeRows.slice() : [];
    const q = String(f.q || '')
        .trim()
        .toLowerCase();
    if (f.angebotId) rows = rows.filter((r) => r.angebotId === f.angebotId);
    if (f.klasse) rows = rows.filter((r) => String(r.klasse || '') === f.klasse);
    if (q) {
        rows = rows.filter((r) => {
            const hay = [r.name, r.email, r.angebotTitle, r.klasse].join(' ').toLowerCase();
            return hay.indexOf(q) !== -1;
        });
    }
    const classes = Array.from(
        new Set(
            (state.attendeeRows || [])
                .map((r) => r.klasse)
                .filter(Boolean)
        )
    ).sort();

    const body = rows
        .map((r) => {
            return (
                '<tr><td>' +
                esc(r.name) +
                '</td><td>' +
                esc(r.email) +
                '</td><td>' +
                esc(r.klasse || '–') +
                '</td><td>' +
                esc(r.angebotTitle) +
                '</td><td>' +
                esc(r.datum || '') +
                '</td></tr>'
            );
        })
        .join('');

    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head"><div><h2>Teilnehmer (aus Bookings)</h2>' +
        '<p class="pw-view__sub">Live aus Microsoft Bookings. ' +
        (state.attendeeLoadedAt
            ? 'Zuletzt geladen: ' + esc(state.attendeeLoadedAt)
            : 'Noch nicht geladen.') +
        '</p></div></header>' +
        '<div class="pw-form__actions" style="margin-bottom:12px">' +
        '<button type="button" class="btn" id="pwBtnLoadTn2"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-arrow-clockwise"></i>Aktualisieren</button> ' +
        '<button type="button" class="btn" id="pwBtnTnCsv"><i class="bi bi-download"></i>CSV exportieren</button>' +
        '</div>' +
        '<div class="pw-filters" role="search" aria-label="Teilnehmer filtern">' +
        '<div class="pw-filters__head">' +
        '<span class="pw-filters__title"><i class="bi bi-funnel" aria-hidden="true"></i>Filter &amp; Suche</span>' +
        '<span class="pw-filters__hint">' +
        rows.length +
        ' Treffer</span></div>' +
        '<div class="pw-filters__row">' +
        '<div class="pw-filters__field"><label class="pw-filters__label" for="pwTnAngebot"><i class="bi bi-journal-text" aria-hidden="true"></i>Angebot</label>' +
        '<div class="pw-select"><select id="pwTnAngebot" class="pw-select__control"><option value="">alle</option>' +
        angebote
            .map(
                (a) =>
                    '<option value="' +
                    esc(a.angebotId) +
                    '"' +
                    (f.angebotId === a.angebotId ? ' selected' : '') +
                    '>' +
                    esc(a.title) +
                    '</option>'
            )
            .join('') +
        '</select><i class="bi bi-chevron-down pw-select__caret" aria-hidden="true"></i></div></div>' +
        '<div class="pw-filters__field"><label class="pw-filters__label" for="pwTnKlasse"><i class="bi bi-mortarboard" aria-hidden="true"></i>Klasse</label>' +
        '<div class="pw-select"><select id="pwTnKlasse" class="pw-select__control"><option value="">alle</option>' +
        optionList(classes, null, null, f.klasse) +
        '</select><i class="bi bi-chevron-down pw-select__caret" aria-hidden="true"></i></div></div>' +
        '<div class="pw-filters__field pw-filters__field--search">' +
        '<label class="pw-filters__label" for="pwTnQ"><i class="bi bi-search" aria-hidden="true"></i>Suche</label>' +
        '<div class="pw-search"><i class="bi bi-search pw-search__icon" aria-hidden="true"></i>' +
        '<input type="search" id="pwTnQ" class="pw-search__input" value="' +
        esc(f.q || '') +
        '" placeholder="Name, E-Mail …" autocomplete="off"></div></div>' +
        '</div></div>' +
        '<div class="pw-table-wrap"><table class="pw-table"><thead><tr>' +
        '<th>Name</th><th>E-Mail</th><th>Klasse</th><th>Angebot</th><th>Datum</th>' +
        '</tr></thead><tbody>' +
        (body || '<tr><td colspan="5" class="muted">Keine Teilnehmer – zuerst unter Bookings laden.</td></tr>') +
        '</tbody></table></div></section>'
    );
}

function renderExport(state) {
    const items = scopedItems(state);
    const freigegeben = items.filter((a) => a.status === 'freigegeben');
    const tn = Array.isArray(state.attendeeRows) ? state.attendeeRows.length : 0;
    return (
        '<section class="pw-view"><h2>Export &amp; Druck</h2>' +
        '<p class="muted">Exporte nutzen die aktuelle Filter-/Rollensicht (' +
        items.length +
        ' Angebote' +
        (state.aktion ? ' · ' + esc(state.aktion.title) : '') +
        ').</p>' +
        '<div class="pw-export-grid">' +
        '<div class="pw-export-card">' +
        '<h3><i class="bi bi-filetype-csv"></i> Angebote CSV</h3>' +
        '<p>Alle sichtbaren Angebote inkl. Sync- und Buchungsfelder.</p>' +
        '<button type="button" class="btn" id="pwBtnExportAngebote"><i class="bi bi-download"></i>CSV herunterladen</button>' +
        '</div>' +
        '<div class="pw-export-card">' +
        '<h3><i class="bi bi-calendar-event"></i> iCal (.ics)</h3>' +
        '<p>Freigegebene Angebote als Kalenderdatei (' +
        freigegeben.length +
        ').</p>' +
        '<button type="button" class="btn" id="pwBtnExportIcs"><i class="bi bi-download"></i>ICS herunterladen</button>' +
        '</div>' +
        '<div class="pw-export-card">' +
        '<h3><i class="bi bi-printer"></i> Wochenplan drucken</h3>' +
        '<p>Öffnet ein druckfreundliches Raster Mo–Fr × Slots.</p>' +
        '<button type="button" class="btn" id="pwBtnPrintPlan"><i class="bi bi-printer"></i>Druckansicht</button>' +
        '</div>' +
        '<div class="pw-export-card">' +
        '<h3><i class="bi bi-people"></i> Teilnehmer</h3>' +
        '<p>CSV und Druck der geladenen Bookings-Teilnehmer (' +
        tn +
        ' Zeilen). Zuerst unter „Teilnehmer“ aktualisieren.</p>' +
        '<div class="pw-form__actions">' +
        '<button type="button" class="btn" id="pwBtnExportTn"><i class="bi bi-download"></i>CSV</button> ' +
        '<button type="button" class="btn" id="pwBtnPrintTn"><i class="bi bi-printer"></i>Druck</button>' +
        '</div></div>' +
        '</div>' +
        '<p class="muted" style="margin-top:1rem">Status-Mail bei Freigabe/Ablehnung: ' +
        '<a href="pa-projektwochen-mail.html">Power-Automate-Rezept</a>.</p>' +
        '</section>'
    );
}

function renderDetail(state) {
    const id = state.detailId;
    const item = (state.items || []).find((a) => a.itemId === id || a.angebotId === id);
    if (!item) {
        return (
            '<section class="pw-view pw-detail">' +
            '<button type="button" class="btn" id="pwDetailBack"><i class="bi bi-arrow-left"></i>Zurück</button>' +
            '<p class="pw-alert pw-alert--warn" style="margin-top:12px">Angebot nicht gefunden.</p></section>'
        );
    }
    const scope = scopeFromState(state);
    const editable = canEditAngebot(item, scope);
    const isAdmin = canAdminDecide(scope);
    const buchung = effectiveBuchungAb(item, state.aktion);
    const occ = occupancyLabel(state, item);
    const tnRows = (state.attendeeRows || []).filter(
        (r) => r && (r.angebotId === item.angebotId || r.angebotTitle === item.title)
    );
    const showTnNames = state.role !== 'schueler';

    const metaChips =
        '<div class="pw-detail__chips">' +
        statusBadge(item.status) +
        '<span class="pw-tag">' +
        esc(item.kategorie || 'sonstiges') +
        '</span>' +
        (occ ? '<span class="pw-tag pw-tag--accent">' + esc(occ) + ' TN</span>' : '') +
        (item.syncStatus ? '<span class="pw-tag">Sync: ' + esc(item.syncStatus) + '</span>' : '') +
        '</div>';

    const weekday = item.tag || weekdayLabelDeFromIso(item.datum) || '';
    const readOnly =
        '<div class="pw-detail__ro pw-form__layout">' +
        '<div class="pw-form__main">' +
        formSection(
            'bi-card-heading',
            'Inhalt',
            'Titeltexte und Hinweise',
            '<div class="pw-detail__grid">' +
                detailCard('Kategorie', 'bi-tags', esc(item.kategorie || 'sonstiges')) +
                detailCard('Beschreibung', 'bi-text-left', esc(item.beschreibung || '–'), true) +
                detailCard('Hinweis Eltern', 'bi-info-circle', esc(item.hinweisEltern || '–'), true) +
                (item.ablehnungsGrund
                    ? detailCard(
                          'Ablehnungsgrund',
                          'bi-x-octagon',
                          '<span class="pw-alert pw-alert--warn" style="display:block;margin:0">' +
                              esc(item.ablehnungsGrund) +
                              '</span>',
                          true
                      )
                    : '') +
                '</div>',
            'pw-form-section__body--flush'
        ) +
        formSection(
            'bi-people',
            'Teilnehmer & Betreuung',
            'Kapazität und Lehrkräfte',
            '<div class="pw-detail__grid">' +
                detailCard('Kapazität', 'bi-person-bounding-box', esc(String(item.kapazitaet ?? '–'))) +
                detailCard('Zielklassen', 'bi-mortarboard', esc(item.zielklassen || 'alle')) +
                detailCard(
                    'Lehrer',
                    'bi-person-badge',
                    esc(item.lehrerCode || '–') +
                        (item.lehrerEmail ? '<br><span class="muted">' + esc(item.lehrerEmail) + '</span>' : '')
                ) +
                detailCard('Begleitung', 'bi-person-plus', esc(item.begleitung || '–')) +
                '</div>',
            'pw-form-section__body--flush'
        ) +
        '</div>' +
        '<aside class="pw-form__aside">' +
        formSection(
            'bi-calendar3',
            'Termin & Ort',
            'Wann und wo',
            '<div class="pw-detail__grid pw-detail__grid--aside">' +
                detailCard(
                    'Datum',
                    'bi-calendar-event',
                    esc(item.datum) + (weekday ? ' · ' + esc(weekday) : '')
                ) +
                detailCard('Slot', 'bi-grid-3x2', esc(item.slot || '–')) +
                detailCard(
                    'Zeit',
                    'bi-clock',
                    esc(item.startzeit || '') + '–' + esc(item.endzeit || '')
                ) +
                detailCard('Ort', 'bi-geo-alt', esc(item.ort || '–')) +
                detailCard('Treffpunkt', 'bi-signpost', esc(item.treffpunkt || '–')) +
                '</div>',
            'pw-form-section__body--flush'
        ) +
        formSection(
            'bi-cash-coin',
            'Kosten',
            'Preis und Hinweise',
            '<div class="pw-detail__grid pw-detail__grid--aside">' +
                detailCard('Preis', 'bi-currency-euro', esc(String(item.preisEuro ?? 0)) + ' €') +
                detailCard('Kostenhinweis', 'bi-receipt', esc(item.kostenHinweis || '–')) +
                '</div>',
            'pw-form-section__body--flush'
        ) +
        (isAdmin
            ? formSection(
                  'bi-shield-lock',
                  'Verwaltung',
                  'Nur für Admin',
                  '<div class="pw-detail__grid pw-detail__grid--aside">' +
                      detailCard(
                          'Buchung ab',
                          'bi-unlock',
                          buchung ? esc(toDatetimeLocalValue(buchung).replace('T', ' ')) : '–'
                      ) +
                      (item.notizIntern
                          ? detailCard('Interne Notiz', 'bi-journal-text', esc(item.notizIntern))
                          : '') +
                      '</div>',
                  'pw-form-section__body--flush'
              )
            : '') +
        '</aside></div>';

    const adminBar =
        isAdmin && item.status === 'beantragt'
            ? '<div class="pw-detail__actions">' +
              '<button type="button" class="btn btn-success" data-pw-approve="' +
              esc(item.itemId) +
              '"><i class="bi bi-check2"></i>Freigeben</button> ' +
              '<button type="button" class="btn" data-pw-reject="' +
              esc(item.itemId) +
              '"><i class="bi bi-x-lg"></i>Ablehnen</button></div>'
            : '';

    const bookingsLink = item.bookingsBookingUrl
        ? '<div class="pw-detail__actions">' +
          '<a class="btn" href="' +
          esc(item.bookingsBookingUrl) +
          '" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i>In Bookings öffnen</a></div>'
        : '';

    const tnBody = tnRows
        .map((r) => {
            return (
                '<tr><td>' +
                esc(r.name) +
                '</td><td>' +
                esc(r.email) +
                '</td><td>' +
                esc(r.klasse || '–') +
                '</td></tr>'
            );
        })
        .join('');

    const tnSection =
        '<div class="pw-detail__tn tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-people"></i><div><h3>Angemeldete Schüler (Bookings)</h3>' +
        '<p>' +
        (occ ? 'Belegung ' + esc(occ) + '. ' : '') +
        (state.attendeeLoadedAt
            ? 'Zuletzt geladen: ' + esc(state.attendeeLoadedAt)
            : 'Noch nicht aus Bookings geladen.') +
        '</p></div></div>' +
        '<div class="pw-form__actions" style="margin-bottom:10px">' +
        '<button type="button" class="btn" id="pwBtnDetailLoadTn"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-arrow-clockwise"></i>Teilnehmer laden</button></div>' +
        (showTnNames
            ? '<div class="pw-table-wrap"><table class="pw-table"><thead><tr><th>Name</th><th>E-Mail</th><th>Klasse</th></tr></thead><tbody>' +
              (tnBody ||
                  '<tr><td colspan="3" class="muted">Keine Teilnehmer für dieses Angebot – Button oben nutzen (Demo oder Bookings).</td></tr>') +
              '</tbody></table></div>'
            : '<p class="muted">Als Schüler siehst du die Belegung, aber keine Namensliste.</p>') +
        '</div>';

    return (
        '<section class="pw-view pw-detail">' +
        '<div class="pw-detail__toolbar">' +
        '<button type="button" class="btn" id="pwDetailBack"><i class="bi bi-arrow-left"></i>Zurück</button>' +
        '</div>' +
        '<header class="pw-detail__hero">' +
        '<p class="pw-detail__kicker">Angebot</p>' +
        '<h2>' +
        esc(item.title) +
        '</h2>' +
        metaChips +
        '</header>' +
        adminBar +
        (editable
            ? '<div class="pw-detail__edit">' +
              renderAngebotForm(state, {
                  editing: true,
                  submitLabel: 'Änderungen speichern',
                  showCancel: false
              }) +
              '</div>'
            : readOnly) +
        bookingsLink +
        tnSection +
        '</section>'
    );
}

/**
 * @param {string} label
 * @param {string} icon
 * @param {string} valueHtml
 * @param {boolean} [wide]
 */
function detailCard(label, icon, valueHtml, wide) {
    return (
        '<div class="pw-detail__card' +
        (wide ? ' pw-detail__card--wide' : '') +
        '"><span class="pw-detail__card-label">' +
        (icon ? '<i class="bi ' + esc(icon) + '" aria-hidden="true"></i>' : '') +
        esc(label) +
        '</span><div class="pw-detail__card-value">' +
        valueHtml +
        '</div></div>'
    );
}

function renderView(state) {
    switch (state.view) {
        case 'plan':
            return renderPlan(state);
        case 'kalender':
            return renderKalender(state);
        case 'liste':
            return renderListe(state);
        case 'neu':
            return renderNeu(state);
        case 'meine':
            return renderMeine(state);
        case 'admin':
            return renderAdmin(state);
        case 'einstellungen':
            return renderEinstellungen(state);
        case 'bookings':
            return renderBookings(state);
        case 'teilnehmer':
            return renderTeilnehmer(state);
        case 'export':
            return renderExport(state);
        case 'detail':
            return renderDetail(state);
        case 'dashboard':
        default:
            return renderDashboard(state);
    }
}

/**
 * @param {object} state
 * @param {HTMLElement} root
 */
export function renderApp(state, root) {
    if (state.loading) {
        root.innerHTML =
            '<div class="pw-shell"><div class="pw-loading">Lade Daten …</div></div>';
        return;
    }
    const roleMod =
        state.role === 'admin'
            ? ' pw-shell--admin'
            : state.role === 'schueler'
              ? ' pw-shell--schueler'
              : ' pw-shell--lehrer';
    root.innerHTML =
        '<div class="pw-shell' +
        roleMod +
        '">' +
        renderNav(state) +
        '<div class="pw-main">' +
        renderTop(state) +
        '<div id="pwView">' +
        renderView(state) +
        '</div></div></div>';
}

export function readFiltersFromDom() {
    return {
        status: val('pwFilterStatus'),
        kategorie: val('pwFilterKat'),
        tag: val('pwFilterTag'),
        lehrer: val('pwFilterLehrer'),
        klasse: val('pwFilterKlasse'),
        buchung: val('pwFilterBuchung'),
        q: val('pwFilterQ')
    };
}

export function readFormFromDom() {
    const lehrerSel = document.getElementById('pwFLehrer');
    let lehrerCode = val('pwFLehrer');
    let lehrerEmail = val('pwFLehrerMail');
    if (lehrerSel && lehrerSel.selectedOptions && lehrerSel.selectedOptions[0]) {
        /* code already from value */
    }
    return {
        title: val('pwFTitle'),
        beschreibung: val('pwFDesc'),
        hinweisEltern: val('pwFEltern'),
        ort: val('pwFOrt'),
        treffpunkt: val('pwFTreff'),
        tag: val('pwFTag'),
        datum: val('pwFDatum'),
        slot: val('pwFSlot'),
        startzeit: val('pwFStart'),
        endzeit: val('pwFEnd'),
        kapazitaet: Number(val('pwFKap') || 0),
        preisEuro: Number(val('pwFPreis') || 0),
        kostenHinweis: val('pwFKosten'),
        zielklassen: val('pwFKlassen') || 'alle',
        lehrerCode,
        lehrerEmail,
        begleitung: val('pwFBegleitung'),
        kategorie: val('pwFKat'),
        buchungAb: val('pwFBuchungAb'),
        notizIntern: val('pwFNotiz')
    };
}

function val(id) {
    const el = document.getElementById(id);
    return el ? String(el.value || '').trim() : '';
}

export { VIEWS };
