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
    computePwAttention,
    dashboardHomeModel,
    studentBookingsUrl,
    buildWeekPlan,
    buildWeekPlanDisplayRows,
    buildMonthCells,
    effectiveBuchungAb,
    isBookingOpen,
    toDatetimeLocalValue,
    weekdayLabelDeFromIso,
    weekdayNameDe,
    toIsoDateOnly,
    statusLabel,
    STATUS_ANGEBOT,
    SLOTS,
    PLAN_SLOTS,
    TAGE,
    KATEGORIEN
} from './projektwochen-logic.js';
import { attendeesForEmail, countPendingBookingsSync } from './projektwochen-bookings-logic.js';
import { formatPwSpoSyncLine } from './projektwochen-sync.js';
import { renderSystemAdminPage } from '../../shared/system-admin-shell.js';

function occupancyLabel(state, angebot) {
    const sid = String((angebot && angebot.bookingsServiceId) || '').trim();
    if (!sid || !state.occupancy || !state.occupancy[sid]) return '';
    const o = state.occupancy[sid];
    return String(o.filled || 0) + '/' + String(o.max || angebot.kapazitaet || '?');
}

function mineIds(state) {
    if (!state || state.role !== 'schueler') return new Set();
    const ids = new Set();
    attendeesForEmail(state.attendeeRows, state.accountEmail).forEach((r) => {
        if (r.angebotId) ids.add(String(r.angebotId));
    });
    return ids;
}

function isMine(state, angebot) {
    if (!angebot) return false;
    return mineIds(state).has(String(angebot.angebotId || ''));
}

function mineMark() {
    return '<span class="fb-pill pw-mine">Du</span>';
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

function pillTone(kind) {
    const s = String(kind || '').toLowerCase();
    if (s === 'freigegeben' || s === 'offen' || s === 'ok') return 'ok';
    if (s === 'beantragt') return 'warn';
    if (s === 'abgelehnt' || s === 'abgesagt' || s === 'fehler') return 'bad';
    return '';
}

function statusBadge(status) {
    const s = String(status || '');
    const tone = pillTone(s);
    return '<span class="fb-pill' + (tone ? ' fb-pill--' + tone : '') + '">' + esc(statusLabel(s)) + '</span>';
}

function teacherLabel(state, item) {
    const labels = labelMaps(state && state.stammdaten);
    const code = String((item && item.lehrerCode) || '').trim();
    const name = code ? labels.teachers[code] : '';
    if (name && name !== code) return name;
    return code || String((item && item.lehrerEmail) || '').trim() || '–';
}

function syncPill(status) {
    const s = String(status || '').trim();
    if (!s || s === '–') return '<span class="fb-pill">–</span>';
    const tone = pillTone(s);
    return '<span class="fb-pill' + (tone ? ' fb-pill--' + tone : '') + '">' + esc(s) + '</span>';
}

function scopedItems(state) {
    return filterAngebote(state.items, state.filters, scopeFromState(state), state.aktion);
}

function filterBar(state) {
    const f = state.filters || {};
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    const classes = (state.stammdaten && state.stammdaten.classes) || [];
    const moreCount = [f.kategorie, f.tag, f.lehrer, f.klasse, f.buchung].filter(Boolean).length;
    const activeCount = moreCount + [f.status, f.q].filter(Boolean).length;
    const moreOpen = !!state.filtersMore;

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

    const moreFields =
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
                '>buchbar</option>' +
                '<option value="gesperrt"' +
                (f.buchung === 'gesperrt' ? ' selected' : '') +
                '>noch nicht buchbar</option>'
        );
    return (
        '<div class="pw-filters' +
        (moreOpen ? ' is-open' : ' is-quiet') +
        '" role="search" aria-label="Angebote filtern">' +
        '<div class="pw-filters__row pw-filters__row--main">' +
        filterSelect(
            'pwFilterStatus',
            'Status',
            'bi-flag',
            '<option value="">alle</option>' +
            STATUS_ANGEBOT.map((s) => {
                return (
                    '<option value="' +
                    esc(s) +
                    '"' +
                    (f.status === s ? ' selected' : '') +
                    '>' +
                    esc(statusLabel(s)) +
                    '</option>'
                );
            }).join('')
        ) +
        '<div class="pw-filters__field pw-filters__field--search">' +
        '<label class="pw-filters__label" for="pwFilterQ"><i class="bi bi-search" aria-hidden="true"></i>Suche</label>' +
        '<div class="pw-search">' +
        '<i class="bi bi-search pw-search__icon" aria-hidden="true"></i>' +
        '<input type="search" id="pwFilterQ" class="pw-search__input" value="' +
        esc(f.q || '') +
        '" placeholder="Titel, Ort, Beschreibung …" autocomplete="off" spellcheck="false">' +
        '</div></div>' +
        '<div class="pw-filters__tools">' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="pwFilterMore" aria-expanded="' +
        (moreOpen ? 'true' : 'false') +
        '">' +
        (moreOpen ? 'Weniger Filter' : 'Mehr Filter') +
        (moreCount ? ' · ' + moreCount : '') +
        '</button>' +
        (activeCount
            ? '<button type="button" class="pw-filters__reset" id="pwFilterReset">Zurücksetzen</button>'
            : '') +
        '</div></div>' +
        (moreOpen ? '<div class="pw-filters__more"><div class="pw-filters__row">' + moreFields + '</div></div>' : '') +
        '</div>'
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
        ? '<button type="button" class="fb-btn' +
          (o.primary ? '' : ' fb-btn--ghost') +
          '" data-pw-view="' +
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
                    const onCatalog =
                        v.id === 'plan' &&
                        (state.view === 'plan' ||
                            state.view === 'kalender' ||
                            state.view === 'liste' ||
                            state.view === 'neu');
                    return (
                        '<button type="button" class="pw-nav__btn' +
                        (state.view === v.id || onCatalog ? ' is-active' : '') +
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

function deDate(iso) {
    const s = String(iso || '').slice(0, 10);
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (!m) return s;
    return m[3] + '.' + m[2] + '.' + m[1];
}

function aktionRange(aktion) {
    if (!aktion) return '';
    const start = deDate(aktion.startdatum);
    const end = deDate(aktion.enddatum);
    if (start && end) return start + ' – ' + end;
    return start || end || '';
}

function renderTop(state) {
    const aktion = state.aktion;
    const aktionen = Array.isArray(state.aktionen) ? state.aktionen : [];
    const range = aktionRange(aktion);
    const weekCore = aktion
        ? aktionen.length > 1
            ? '<label class="pw-weekline__select"><span class="pw-weekline__sr">Projektwoche</span><select id="pwAktionSelect">' +
              aktionen
                  .map((a) => {
                      const id = String((a && a.aktionId) || '');
                      const sel = aktion && aktion.aktionId === id ? ' selected' : '';
                      const label =
                          (a.title || id || 'Projektwoche') +
                          (aktionRange(a) ? ' · ' + aktionRange(a) : '');
                      return '<option value="' + esc(id) + '"' + sel + '>' + esc(label) + '</option>';
                  })
                  .join('') +
              '</select></label>'
            : '<strong class="pw-weekline__title">' + esc(aktion.title || 'Projektwoche') + '</strong>'
        : '<span class="muted">Keine Projektwoche geladen</span>';
    const aktionLabel =
        '<div class="pw-weekline">' +
        weekCore +
        (range && aktionen.length < 2 ? '<span class="pw-weekline__range">' + esc(range) + '</span>' : '') +
        (aktion ? statusBadge(aktion.status) : '') +
        '</div>';
    const syncSrc = state.spoSync && state.spoSync.source;
    const conn =
        state.localDemoOnly || syncSrc === 'demo'
            ? '<span class="fb-pill">Demo lokal</span>'
            : state.ctx || syncSrc === 'live'
              ? '<span class="fb-pill fb-pill--ok">SPO live</span>'
              : state.siteUrl
                ? '<span class="fb-pill fb-pill--warn">Site hinterlegt</span>'
                : '<span class="fb-pill">Nicht verbunden</span>';
    const adminHint =
        state.role === 'admin' && !state.bootstrapped
            ? '<button type="button" class="fb-btn fb-btn--ghost" data-pw-view="einstellungen"><i class="bi bi-gear" aria-hidden="true"></i>Systemadministration</button>'
            : state.role === 'admin'
              ? '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnReload" title="Daten neu laden"><i class="bi bi-arrow-repeat" aria-hidden="true"></i>Aktualisieren</button>'
              : '';
    return (
        '<div class="pw-top">' +
        '<div class="pw-top__aktion">' +
        '<div class="pw-top__aktion-row">' +
        aktionLabel +
        '<div class="pw-top__actions">' +
        conn +
        adminHint +
        '</div></div>' +
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
    if (state.role === 'admin' && att.pendingSync > 0) {
        const n = att.pendingSync;
        const sentence =
            n === 1
                ? '<p><strong>1</strong> freies Angebot ist noch nicht in Microsoft Bookings.</p>'
                : '<p><strong>' + n + '</strong> freie Angebote sind noch nicht in Microsoft Bookings.</p>';
        parts.push(
            calmBanner('Bookings-Sync', sentence, {
                tone: 'ops',
                jumpView: 'bookings',
                jumpLabel: 'Sync öffnen'
            })
        );
    }
    if (state.role === 'schueler') {
        parts.push(
            calmBanner(
                'Buchung',
                '<p>Angebote hier lesen; buchen Sie in der <strong>Microsoft Bookings</strong>-App (Link nach Freigabe/Sync).</p>',
                {}
            )
        );
    }
    if (!state.bootstrapped && !state.localDemoOnly) {
        if (state.role === 'admin') {
            parts.push(
                calmBanner(
                    'Setup',
                    '<p>SharePoint verbinden oder Demo laden – unter System.</p>',
                    { jumpView: 'einstellungen', jumpLabel: 'System' }
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

function homeAction(model) {
    const btn = model && (model.primary || model.ghost);
    if (!btn) return '';
    return (
        '<div class="pw-view__home-actions"><button type="button" class="fb-btn' +
        (model.primary ? '' : ' fb-btn--ghost') +
        '" data-pw-view="' +
        esc(btn.view) +
        '">' +
        esc(btn.label) +
        '</button></div>'
    );
}

function bookingsAnchor(url, label) {
    return (
        '<a class="fb-btn fb-btn--ghost" href="' +
        esc(url) +
        '" target="_blank" rel="noopener">' +
        esc(label) +
        '</a>'
    );
}

function renderDashboard(state) {
    const home = dashboardHomeModel(state);
    const num = home.num == null ? '' : '<span class="fb-num">' + esc(String(home.num)) + '</span> ';
    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head"><div>' +
        '<h1 class="fb-greet">' +
        num +
        esc(home.title) +
        '</h1>' +
        '<p class="fb-sub pw-view__sub">' +
        esc(home.sub) +
        '</p>' +
        homeAction(home) +
        '</div></header>' +
        renderDashboardAttention(state) +
        '</section>'
    );
}

function catalogHead(state, title, sub) {
    const modes = [
        ['plan', 'Plan'],
        ['liste', 'Liste'],
        ['kalender', 'Kalender']
    ];
    const seg = modes
        .map((pair) => {
            const on = state.view === pair[0];
            return (
                '<button type="button" class="pw-seg__btn' +
                (on ? ' is-active' : '') +
                '" data-pw-view="' +
                pair[0] +
                '" aria-pressed="' +
                (on ? 'true' : 'false') +
                '">' +
                esc(pair[1]) +
                '</button>'
            );
        })
        .join('');
    const staff = state.role !== 'schueler';
    const menuItems = staff
        ? '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnExportAngebote">Angebote als CSV</button>' +
          '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnExportIcs">Kalenderdatei</button>' +
          '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnPrintPlan">Wochenplan drucken</button>' +
          '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnExportTn">Teilnehmer als CSV</button>' +
          '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnPrintTn">Teilnehmer drucken</button>'
        : '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnPrintMine">Meine Termine drucken</button>';
    return (
        '<header class="pw-view__head"><div><h2>' +
        esc(title) +
        '</h2><p class="pw-view__sub">' +
        esc(sub) +
        '</p></div><div class="pw-catalog-tools">' +
        (staff
            ? '<button type="button" class="fb-btn fb-btn--ghost" data-pw-view="neu">Neues Angebot</button>'
            : '') +
        '<div class="pw-seg" role="group" aria-label="Ansicht">' +
        seg +
        '</div><details class="pw-menu"><summary class="fb-btn fb-btn--ghost">Export</summary><div class="pw-menu__list">' +
        menuItems +
        '</div></details></div></header>'
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
                const bookUrl = studentBookingsUrl(state.role, a, isMine(state, a));
                return (
                    '<div class="pw-plan-card">' +
                    '<button type="button" class="pw-plan-card__open" data-pw-detail="' +
                    esc(a.itemId || a.angebotId) +
                    '"><strong>' +
                    esc(a.title) +
                    '</strong><span class="pw-plan-card__meta">' +
                    statusBadge(a.status) +
                    (isMine(state, a) ? mineMark() : '') +
                    (occ ? '<span>' + esc(occ) + ' TN</span>' : '') +
                    '<span>' +
                    esc(String(a.kapazitaet || '')) +
                    ' Plätze</span>' +
                    (a.syncStatus ? syncPill(a.syncStatus) : '') +
                    '</span></button>' +
                    (bookUrl ? bookingsAnchor(bookUrl, 'In Bookings buchen') : '') +
                    '</div>'
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
        catalogHead(state, 'Wochenplan', 'Angebote nach Tag und Tageszeit. Ganztägige spannen Vormittag bis Abend.') +
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
    const now = new Date();
    const todayIso =
        now.getFullYear() +
        '-' +
        String(now.getMonth() + 1).padStart(2, '0') +
        '-' +
        String(now.getDate()).padStart(2, '0');
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
                .map((a) => {
                    const tone = pillTone(a.status);
                    const mine = isMine(state, a);
                    return (
                        '<button type="button" class="pw-cal__item' +
                        (tone ? ' pw-cal__item--' + tone : '') +
                        (mine ? ' pw-cal__item--mine' : '') +
                        '" data-pw-detail="' +
                        esc(a.itemId || a.angebotId) +
                        '" title="' +
                        esc(a.title + (mine ? ' · Du' : '')) +
                        '">' +
                        esc(a.title) +
                        (mine ? ' · Du' : '') +
                        '</button>'
                    );
                })
                .join('');
            body +=
                '<td class="pw-cal__day' +
                (c.iso === todayIso ? ' is-today' : '') +
                '"><div class="pw-cal__num">' +
                c.day +
                '</div>' +
                bits +
                (c.items.length > 4 ? '<div class="pw-cal__more">+' + (c.items.length - 4) + '</div>' : '') +
                '</td>';
        }
        body += '</tr>';
    }
    return (
        '<section class="pw-view">' +
        catalogHead(state, 'Kalender', 'Monatsübersicht der gefilterten Angebote.') +
        filterBar(state) +
        '<div class="pw-cal-nav">' +
        '<button type="button" class="fb-btn fb-btn--ghost pw-cal-nav__btn" id="pwCalPrev" aria-label="Voriger Monat"><i class="bi bi-chevron-left" aria-hidden="true"></i></button>' +
        '<strong>' +
        esc(monthLabel) +
        '</strong>' +
        '<button type="button" class="fb-btn fb-btn--ghost pw-cal-nav__btn" id="pwCalNext" aria-label="Nächster Monat"><i class="bi bi-chevron-right" aria-hidden="true"></i></button>' +
        '</div>' +
        '<p class="pw-cal-legend">' +
        '<span class="pw-cal__item pw-cal__item--ok">' +
        esc(statusLabel('freigegeben')) +
        '</span>' +
        '<span class="pw-cal__item pw-cal__item--warn">' +
        esc(statusLabel('beantragt')) +
        '</span>' +
        '<span class="pw-cal__item pw-cal__item--bad">' +
        esc(statusLabel('abgelehnt')) +
        '</span>' +
        '<span class="pw-cal__item">' +
        esc(statusLabel('entwurf')) +
        '</span>' +
        '</p>' +
        '<table class="pw-cal"><thead><tr>' +
        head +
        '</tr></thead><tbody>' +
        body +
        '</tbody></table></section>'
    );
}

function renderListe(state) {
    const items = scopedItems(state);
    const scope = scopeFromState(state);
    const rows = items
        .slice()
        .sort((a, b) => String(a.datum).localeCompare(String(b.datum)))
        .map((a) => {
            const buchung = effectiveBuchungAb(a, state.aktion);
            const open = a.status === 'freigegeben' && isBookingOpen(buchung);
            const edit = canEditAngebot(a, scope);
            const occ = occupancyLabel(state, a);
            const bookUrl = studentBookingsUrl(state.role, a, isMine(state, a));
            const staffUrl =
                state.role !== 'schueler' ? String(a.bookingsBookingUrl || '').trim() : '';
            return (
                '<tr>' +
                '<td><button type="button" class="pw-link" data-pw-detail="' +
                esc(a.itemId || a.angebotId) +
                '">' +
                esc(a.title) +
                '</button>' +
                (isMine(state, a) ? ' ' + mineMark() : '') +
                '</td>' +
                '<td>' +
                esc(a.datum) +
                ' ' +
                esc(a.tag || '') +
                '</td>' +
                '<td>' +
                esc(a.slot) +
                '</td>' +
                '<td>' +
                esc(teacherLabel(state, a)) +
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
                      ' ' +
                      (open
                          ? '<span class="fb-pill fb-pill--ok">' + esc(statusLabel('offen')) + '</span>'
                          : '<span class="fb-pill">' + esc(statusLabel('gesperrt')) + '</span>')
                    : '–') +
                '</td>' +
                '<td>' +
                syncPill(a.syncStatus) +
                '</td>' +
                '<td class="pw-table__actions">' +
                (edit
                    ? '<button type="button" class="fb-btn fb-btn--ghost" data-pw-edit="' +
                      esc(a.itemId) +
                      '">Bearbeiten</button>'
                    : '') +
                (bookUrl ? bookingsAnchor(bookUrl, 'In Bookings buchen') : '') +
                (staffUrl ? bookingsAnchor(staffUrl, 'Bookings') : '') +
                '</td></tr>'
            );
        })
        .join('');
    return (
        '<section class="pw-view">' +
        catalogHead(state, 'Liste', 'Tabellarische Übersicht. Details per Klick auf den Titel.') +
        filterBar(state) +
        '<div class="pw-table-wrap"><table class="pw-table"><thead><tr>' +
        '<th>Titel</th><th>Datum</th><th>Tageszeit</th><th>Lehrkraft</th><th>Status</th><th>Belegung</th><th>Buchung ab</th><th>Sync</th><th></th>' +
        '</tr></thead><tbody>' +
        (rows || '<tr><td colspan="9" class="muted">Keine Angebote</td></tr>') +
        '</tbody></table></div></section>'
    );
}

function renderNeu(state) {
    const editing = !!state.editingItemId;
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
            ? 'Titel, Termin, Ort, Plätze und Lehrkraft. Weitere Angaben liegen darunter.'
            : 'Titel, Termin, Ort, Plätze und Lehrkraft reichen für den Antrag.') +
        '</p>' +
        '</header>' +
        '<div class="pw-neu__form">' +
        renderAngebotForm(state, {
            editing: editing,
            submitLabel: editing ? 'Speichern' : 'Antrag stellen'
        }) +
        '</div></section>'
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
    const showCancel = o.showCancel !== false && editing && state.view !== 'detail';
    const deleteId = state.editingItemId || state.detailId || '';
    const showDelete = state.view === 'detail' && editing && !!deleteId;

    function selControl(id, optionsHtml, attrs) {
        return (
            '<div class="pw-select"><select id="' +
            id +
            '" class="pw-select__control"' +
            (attrs || '') +
            '>' +
            optionsHtml +
            '</select><i class="bi bi-chevron-down pw-select__caret" aria-hidden="true"></i></div>'
        );
    }
    function teacherOptions(selected) {
        return (
            '<option value="">–</option>' +
            teachers
                .map((t) => {
                    const code = String((t && t.code) || '');
                    const name = String((t && t.name) || code);
                    const mail = String((t && t.email) || '').trim();
                    return (
                        '<option value="' +
                        esc(code) +
                        '" data-email="' +
                        esc(mail) +
                        '"' +
                        (String(selected) === code ? ' selected' : '') +
                        '>' +
                        esc(name) +
                        '</option>'
                    );
                })
                .join('')
        );
    }
    function labeledOptions(values, labels, selected) {
        return (values || [])
            .map((value) => {
                return (
                    '<option value="' +
                    esc(value) +
                    '"' +
                    (String(selected) === String(value) ? ' selected' : '') +
                    '>' +
                    esc(labels[value] || value) +
                    '</option>'
                );
            })
            .join('');
    }
    const katLabel = {
        exkursion: 'Exkursion',
        workshop: 'Workshop',
        kultur: 'Kultur',
        sport: 'Sport',
        sonstiges: 'Sonstiges'
    };
    const slotLabel = {
        ganztags: 'Ganztags',
        vormittag: 'Vormittag',
        nachmittag: 'Nachmittag',
        abend: 'Abend'
    };
    const tag = form.tag || weekdayLabelDeFromIso(form.datum) || '';
    const tagLine = weekdayNameDe(tag) || 'Der Wochentag folgt aus dem Datum.';
    const weekStart = toIsoDateOnly(state.aktion && state.aktion.startdatum);
    const weekEnd = toIsoDateOnly(state.aktion && state.aktion.enddatum);
    const klassen = String(form.zielklassen || '').trim();
    const mailKnown = String(form.lehrerEmail || '').trim();
    const moreOpen =
        [form.hinweisEltern, form.treffpunkt, form.begleitung, form.kostenHinweis].some((s) =>
            String(s || '').trim()
        ) ||
        Number(form.preisEuro) > 0 ||
        (klassen && klassen.toLowerCase() !== 'alle');

    return (
        '<form class="pw-form pw-form--grouped" id="pwForm">' +
        '<div class="pw-form__layout">' +
        formSection(
            '',
            'Angebot',
            'Was die Schülerinnen und Schüler lesen',
            field('Titel', '', '<input id="pwFTitle" required value="' + esc(form.title) + '" placeholder="z. B. Zoo Schönbrunn">', 'full') +
                field('Kategorie', '', selControl('pwFKat', labeledOptions(KATEGORIEN, katLabel, form.kategorie))) +
                field(
                    'Beschreibung',
                    '',
                    '<textarea id="pwFDesc" rows="3" placeholder="Was erwartet die Schüler:innen?">' +
                        esc(form.beschreibung) +
                        '</textarea>',
                    'full'
                )
        ) +
        formSection(
            '',
            'Termin und Ort',
            'Datum legt den Wochentag fest',
            '<div class="pw-field"><label>Datum</label><input type="date" id="pwFDatum" required value="' +
                esc(form.datum) +
                '"' +
                (weekStart ? ' min="' + esc(weekStart) + '"' : '') +
                (weekEnd ? ' max="' + esc(weekEnd) + '"' : '') +
                '><p class="pw-form__when" id="pwFTagLine">' +
                esc(tagLine) +
                '</p><input type="hidden" id="pwFTag" value="' +
                esc(tag) +
                '"></div>' +
                field(
                    'Tageszeit',
                    '',
                    selControl(
                        'pwFSlot',
                        labeledOptions(SLOTS, slotLabel, form.slot),
                        ' data-pw-slot="' + esc(form.slot || 'vormittag') + '"'
                    )
                ) +
                field('Beginn', '', '<input type="time" id="pwFStart" value="' + esc(form.startzeit) + '">') +
                field('Ende', '', '<input type="time" id="pwFEnd" value="' + esc(form.endzeit) + '">') +
                field(
                    'Ort',
                    '',
                    '<input id="pwFOrt" required value="' + esc(form.ort) + '" placeholder="z. B. Tiergarten Schönbrunn">',
                    'full'
                )
        ) +
        formSection(
            '',
            'Plätze und Lehrkraft',
            'Wie viele Plätze, und wer das Angebot verantwortet',
            field(
                'Plätze',
                '',
                '<input type="number" id="pwFKap" min="1" required value="' + esc(String(form.kapazitaet)) + '">',
                'full'
            ) +
                field('Lehrkraft', '', selControl('pwFLehrer', teacherOptions(form.lehrerCode)), 'full') +
                '<div class="pw-field pw-field--full' +
                (mailKnown ? ' is-quiet' : '') +
                '" id="pwFLehrerMailWrap"><label>E-Mail</label><input id="pwFLehrerMail" type="email" value="' +
                esc(form.lehrerEmail) +
                '" placeholder="für die Zuordnung des Angebots"></div>'
        ) +
        '<details class="pw-form-more"' +
        (moreOpen ? ' open' : '') +
        '><summary>Weitere Angaben</summary><div class="pw-form-section__body">' +
        field(
            'Zielklassen',
            '',
            '<input id="pwFKlassen" list="pwKlassenList" value="' +
                esc(form.zielklassen || 'alle') +
                '" placeholder="alle, oder z. B. 1AK;1BK">' +
                '<datalist id="pwKlassenList">' +
                classes.map((c) => '<option value="' + esc(c.code) + '">').join('') +
                '</datalist>',
            'full'
        ) +
        field(
            'Hinweis für Eltern',
            '',
            '<textarea id="pwFEltern" rows="2" placeholder="z. B. Einverständnis nötig">' +
                esc(form.hinweisEltern) +
                '</textarea>',
            'full'
        ) +
        field('Treffpunkt', '', '<input id="pwFTreff" value="' + esc(form.treffpunkt) + '" placeholder="z. B. Schulhof">') +
        field('Begleitung', '', '<input id="pwFBegleitung" value="' + esc(form.begleitung) + '" placeholder="Name">') +
        field(
            'Preis (€)',
            '',
            '<input type="number" id="pwFPreis" min="0" step="0.5" value="' + esc(String(form.preisEuro)) + '">'
        ) +
        field(
            'Kostenhinweis',
            '',
            '<input id="pwFKosten" value="' + esc(form.kostenHinweis) + '" placeholder="z. B. inkl. Eintritt">',
            'full'
        ) +
        '</div></details>' +
        (isAdmin
            ? formSection(
                  '',
                  'Verwaltung',
                  'Buchungsfenster und interne Notiz',
                  field(
                      'Buchung möglich ab',
                      '',
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
                          '',
                          '<textarea id="pwFNotiz" rows="2">' + esc(form.notizIntern) + '</textarea>',
                          'full'
                      )
              )
            : '') +
        '</div>' +
        validationBlock(state.formTried ? v : { errors: [], warnings: [] }) +
        '<div class="pw-form__actions">' +
        '<button type="button" class="fb-btn" id="pwBtnSave"><i class="bi bi-check2" aria-hidden="true"></i>' +
        esc(o.submitLabel || (editing ? 'Speichern' : 'Beantragen')) +
        '</button>' +
        (showCancel ? '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnCancelEdit">Abbrechen</button>' : '') +
        (showDelete
            ? '<button type="button" class="fb-btn fb-btn--ghost pw-form__danger" data-pw-del="' +
              esc(deleteId) +
              '">Löschen</button>'
            : '') +
        '</div></form>'
    );
}

function formSection(_icon, title, hint, bodyHtml, bodyMod) {
    return (
        '<section class="pw-form-section">' +
        '<header class="pw-form-section__head">' +
        '<h3 class="pw-form-section__title">' +
        esc(title) +
        '</h3>' +
        (hint ? '<p class="pw-form-section__hint">' + esc(hint) + '</p>' : '') +
        '</header>' +
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
function field(label, _icon, control, span) {
    const spanClass =
        span === 'full' ? ' pw-field--full' : span === 'half' ? ' pw-field--half' : '';
    return (
        '<div class="pw-field' +
        spanClass +
        '"><label>' +
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
            const id = a.itemId || a.angebotId;
            const editable = canEditAngebot(a, scope);
            return (
                '<tr><td><button type="button" class="pw-link" data-pw-detail="' +
                esc(id) +
                '">' +
                esc(a.title) +
                '</button></td><td>' +
                esc(a.datum) +
                '</td><td>' +
                statusBadge(a.status) +
                '</td><td class="pw-table__actions">' +
                (editable
                    ? '<button type="button" class="fb-btn fb-btn--ghost" data-pw-edit="' +
                      esc(a.itemId) +
                      '">Bearbeiten</button>' +
                      '<button type="button" class="fb-btn fb-btn--ghost" data-pw-del="' +
                      esc(a.itemId) +
                      '">Löschen</button>'
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
                  '<button type="button" class="fb-btn" data-pw-view="neu"><i class="bi bi-plus-circle" aria-hidden="true"></i>Neues Angebot</button>'
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
                esc(teacherLabel(state, a)) +
                '</td><td>' +
                esc(String(a.kapazitaet)) +
                ' / ' +
                esc(String(a.preisEuro)) +
                ' €</td><td class="pw-table__actions">' +
                '<button type="button" class="fb-btn" data-pw-approve="' +
                esc(a.itemId) +
                '">Freigeben</button>' +
                '<button type="button" class="fb-btn fb-btn--ghost" data-pw-reject="' +
                esc(a.itemId) +
                '">Ablehnen</button>' +
                '<button type="button" class="fb-btn fb-btn--ghost" data-pw-detail="' +
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
                '<input type="datetime-local" class="pw-table__input" data-pw-buchung-id="' +
                esc(a.itemId) +
                '" value="' +
                esc(val) +
                '">' +
                '</td><td class="pw-table__actions">' +
                '<button type="button" class="fb-btn fb-btn--ghost" data-pw-buchung-save="' +
                esc(a.itemId) +
                '">Speichern</button>' +
                '</td></tr>'
            );
        })
        .join('');

    const home = dashboardHomeModel(state);
    const greetNum = home.num == null ? '' : '<span class="fb-num">' + esc(String(home.num)) + '</span> ';
    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head"><div>' +
        '<h1 class="fb-greet">' +
        greetNum +
        esc(home.title) +
        '</h1>' +
        '<p class="fb-sub pw-view__sub">' +
        esc(home.sub) +
        '</p></div></header>' +
        renderDashboardAttention(state) +
        '<h2>Freigabe</h2>' +
        '<div class="pw-table-wrap"><table class="pw-table"><thead><tr><th>Titel</th><th>Wann</th><th>Lehrer</th><th>Kap./Preis</th><th></th></tr></thead><tbody>' +
        (queueRows || '<tr><td colspan="5" class="muted">Keine offenen Anträge</td></tr>') +
        '</tbody></table></div>' +
        '<h3 style="margin-top:1.5rem">Buchungsstart (freie Angebote)</h3>' +
        '<p class="muted">Ab diesem Zeitpunkt dürfen Schüler in Bookings buchen (Phase 3). Leer = Standard der Aktion.</p>' +
        '<div class="pw-table-wrap"><table class="pw-table"><thead><tr><th>Angebot</th><th>Datum</th><th>Buchung ab</th><th></th></tr></thead><tbody>' +
        (buchungRows || '<tr><td colspan="4" class="muted">Noch keine freien Angebote</td></tr>') +
        '</tbody></table></div>' +
        (aktion
            ? '<h3 style="margin-top:1.5rem">Aktions-Standard</h3>' +
              '<div class="pw-form__grid">' +
              field(
                  'Standard-Buchungsstart',
                  'bi-unlock',
                  '<input type="datetime-local" id="pwAktionBuchungDefault" value="' +
                      esc(toDatetimeLocalValue(aktion.buchungAbDefault || '')) +
                      '">'
              ) +
              '</div>' +
              '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnAktionSave"><i class="bi bi-save" aria-hidden="true"></i>Aktion speichern</button>'
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
                syncPill(a.syncStatus) +
                '</td>' +
                '<td>' +
                (a.bookingsServiceId ? esc(a.bookingsServiceId.slice(0, 12)) + '…' : '–') +
                '</td>' +
                '<td>' +
                (a.bookingsBookingUrl
                    ? '<a class="pw-link" href="' +
                      esc(a.bookingsBookingUrl) +
                      '" target="_blank" rel="noopener">Link</a>'
                    : '–') +
                '</td>' +
                '<td class="pw-table__actions">' +
                '<button type="button" class="fb-btn fb-btn--ghost" data-pw-sync-one="' +
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
                        '</strong> ' +
                        (pending === 1
                            ? 'freies Angebot braucht einen Sync nach Microsoft Bookings.'
                            : 'freie Angebote brauchen einen Sync nach Microsoft Bookings.') +
                        '</p>',
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
        '<p class="pw-view__sub">Eine Projektwoche ist ein Bookings-Business. Freie Angebote sind die Dienste. Der Sync läuft nur über den Knopf.</p>' +
        summaryBanner +
        (aktion
            ? '<div class="pw-alert pw-alert--info">Aktion: <strong>' +
              esc(aktion.title) +
              '</strong><br>Business: ' +
              esc(aktion.bookingsBusinessName || aktion.bookingsBusinessId || '(noch nicht angelegt)') +
              '</div>'
            : '<div class="pw-alert pw-alert--warn">Keine aktive Aktion.</div>') +
        '<div class="fb-actions-row">' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnEnsureBiz"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-building" aria-hidden="true"></i>Business anlegen/binden</button>' +
        '<button type="button" class="fb-btn" id="pwBtnSyncSelected"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-arrow-repeat" aria-hidden="true"></i>Ausgewählte syncen</button>' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnSyncAll"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-cloud-upload" aria-hidden="true"></i>Alle freien syncen (' +
        freigegeben.length +
        ')</button>' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnLoadTn"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-people" aria-hidden="true"></i>Teilnehmer laden</button>' +
        '</div>' +
        '<p class="muted">Frei: ' +
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
                  'Keine freien Angebote',
                  'Zuerst in der Freigabe entscheiden, dann hier nach Bookings syncen.',
                  '<button type="button" class="fb-btn" data-pw-view="admin">Zur Freigabe</button>'
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
        '<header class="pw-view__head"><div><h2>Teilnehmer</h2>' +
        '<p class="pw-view__sub">' +
        (state.bookingsBusy
            ? 'Buchungen werden geladen.'
            : state.attendeeLoadedAt
              ? 'Zuletzt geladen: ' + esc(state.attendeeLoadedAt)
              : 'Die Buchungen werden beim Öffnen geladen.') +
        '</p></div></header>' +
        '<div class="fb-actions-row">' +
        '<button type="button" class="fb-btn" id="pwBtnLoadTn2"' +
        (state.bookingsBusy ? ' disabled' : '') +
        '><i class="bi bi-arrow-clockwise" aria-hidden="true"></i>Aktualisieren</button>' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnTnCsv"><i class="bi bi-download" aria-hidden="true"></i>CSV exportieren</button>' +
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
        (Array.isArray(state.attendeeRows) && state.attendeeRows.length
            ? '<div class="pw-table-wrap"><table class="pw-table"><thead><tr>' +
              '<th>Name</th><th>E-Mail</th><th>Klasse</th><th>Angebot</th><th>Datum</th>' +
              '</tr></thead><tbody>' +
              (body || '<tr><td colspan="5" class="muted">Keine Treffer für diesen Filter.</td></tr>') +
              '</tbody></table></div>'
            : quietEmpty(
                  state.bookingsBusy ? 'Teilnehmer werden geladen' : 'Keine Buchungen',
                  state.bookingsBusy
                      ? 'Die Buchungen kommen aus Microsoft Bookings.'
                      : 'In Microsoft Bookings steht für diese Projektwoche noch niemand.'
              )) +
        '</section>'
    );
}

function renderDetail(state) {
    const id = state.detailId;
    const item = (state.items || []).find((a) => a.itemId === id || a.angebotId === id);
    if (!item) {
        return (
            '<section class="pw-view pw-detail">' +
            '<button type="button" class="fb-btn fb-btn--ghost" id="pwDetailBack"><i class="bi bi-arrow-left" aria-hidden="true"></i>Zurück</button>' +
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
        '<span class="fb-pill">' +
        esc(item.kategorie || 'sonstiges') +
        '</span>' +
        (occ ? '<span class="fb-pill">' + esc(occ) + ' TN</span>' : '') +
        (isMine(state, item) ? mineMark() : '') +
        (item.syncStatus ? syncPill(item.syncStatus) : '') +
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
                    'Lehrkraft',
                    'bi-person-badge',
                    esc(teacherLabel(state, item)) +
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
                detailCard('Tageszeit', 'bi-grid-3x2', esc(item.slot || '–')) +
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
              '<button type="button" class="fb-btn' +
              (editable ? ' fb-btn--ghost' : '') +
              '" data-pw-approve="' +
              esc(item.itemId) +
              '">Freigeben</button>' +
              '<button type="button" class="fb-btn fb-btn--ghost" data-pw-reject="' +
              esc(item.itemId) +
              '">Ablehnen</button></div>'
            : '';

    const rawBookUrl = String(item.bookingsBookingUrl || '').trim();
    const studentBookUrl = studentBookingsUrl(state.role, item, isMine(state, item));
    const bookingsLink = rawBookUrl
        ? '<div class="pw-detail__actions">' +
          bookingsAnchor(rawBookUrl, studentBookUrl ? 'In Bookings buchen' : 'In Bookings öffnen') +
          '</div>'
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

    const tnHint =
        (occ ? 'Belegung ' + occ + '. ' : '') +
        (state.attendeeLoadedAt
            ? 'Zuletzt geladen: ' + state.attendeeLoadedAt
            : 'Noch nicht aus Bookings geladen.');
    const tnSection = formSection(
        '',
        'Angemeldete Schüler',
        tnHint,
        '<div class="fb-actions-row">' +
            '<button type="button" class="fb-btn fb-btn--ghost" id="pwBtnDetailLoadTn"' +
            (state.bookingsBusy ? ' disabled' : '') +
            '>Teilnehmer laden</button></div>' +
            (showTnNames
                ? '<div class="pw-table-wrap"><table class="pw-table"><thead><tr><th>Name</th><th>E-Mail</th><th>Klasse</th></tr></thead><tbody>' +
                  (tnBody ||
                      '<tr><td colspan="3" class="muted">Keine Teilnehmer für dieses Angebot.</td></tr>') +
                  '</tbody></table></div>'
                : '<p class="muted">Als Schüler siehst du die Belegung, aber keine Namensliste.</p>')
    );

    return (
        '<section class="pw-view pw-detail">' +
        '<div class="pw-detail__toolbar">' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="pwDetailBack"><i class="bi bi-arrow-left" aria-hidden="true"></i>Zurück</button>' +
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
function detailCard(label, _icon, valueHtml, wide) {
    return (
        '<div class="pw-detail__card' +
        (wide ? ' pw-detail__card--wide' : '') +
        '"><span class="pw-detail__card-label">' +
        esc(label) +
        '</span><div class="pw-detail__card-value">' +
        valueHtml +
        '</div></div>'
    );
}

function angebotById(state, id) {
    return (state.items || []).find((a) => a && (a.angebotId === id || a.itemId === id)) || null;
}

function renderBuchung(state) {
    const mine = attendeesForEmail(state.attendeeRows, state.accountEmail).slice().sort((a, b) => {
        return String(a.datum || '').localeCompare(String(b.datum || '')) || String(a.angebotTitle || '').localeCompare(String(b.angebotTitle || ''));
    });
    const n = mine.length;
    const greet = n === 1 ? 'Termin gebucht' : 'Termine gebucht';
    const week = state.aktion ? esc(state.aktion.title || '') + (aktionRange(state.aktion) ? ' · ' + esc(aktionRange(state.aktion)) : '') : '';
    const cards = mine
        .map((r) => {
            const item = angebotById(state, r.angebotId);
            const when = [r.datum ? deDate(r.datum) : '', item && item.tag, item && item.slot, item && item.startzeit && item.endzeit ? item.startzeit + '–' + item.endzeit : '']
                .filter(Boolean)
                .join(' · ');
            const place = item && (item.ort || item.treffpunkt) ? item.ort || item.treffpunkt : '';
            const url = item && item.bookingsBookingUrl ? String(item.bookingsBookingUrl) : '';
            return (
                '<article class="pw-buchung">' +
                '<div class="pw-buchung__main">' +
                '<button type="button" class="pw-link" data-pw-detail="' +
                esc((item && (item.itemId || item.angebotId)) || r.angebotId) +
                '">' +
                esc(r.angebotTitle || (item && item.title) || 'Angebot') +
                '</button>' +
                '<p class="pw-buchung__meta">' +
                esc(when) +
                (place ? ' · ' + esc(place) : '') +
                '</p></div>' +
                '<span class="fb-pill fb-pill--ok">gebucht</span>' +
                (url
                    ? '<a class="fb-btn fb-btn--ghost" href="' +
                      esc(url) +
                      '" target="_blank" rel="noopener">In Bookings öffnen</a>'
                    : '') +
                '</article>'
            );
        })
        .join('');
    const empty = quietEmpty(
        'Noch nichts gebucht',
        'Buchen Sie in Microsoft Bookings. Hier erscheinen danach die Termine, in denen Sie selbst stehen.'
    );
    return (
        '<section class="pw-view">' +
        '<header class="pw-view__head"><div>' +
        '<h1 class="fb-greet"><span class="fb-num">' +
        n +
        '</span> ' +
        greet +
        '</h1>' +
        (week ? '<p class="fb-sub pw-view__sub">' + week + '</p>' : '') +
        '</div></header>' +
        (n ? '<div class="pw-buchung-list">' + cards + '</div>' : empty) +
        (state.role === 'schueler' && !state.localDemoOnly
            ? '<div class="fb-actions-row"><button type="button" class="fb-btn fb-btn--ghost" id="pwBtnLoadMine"' +
              (state.bookingsBusy ? ' disabled' : '') +
              '>Buchungen laden</button></div>'
            : '') +
        '</section>'
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
            return renderPlan(Object.assign({}, state, { view: 'plan' }));
        case 'detail':
            return renderDetail(state);
        case 'buchung':
            return renderBuchung(state);
        case 'dashboard':
            return state.role === 'schueler' ? renderBuchung(state) : renderDashboard(state);
        default:
            return state.role === 'schueler' ? renderBuchung(state) : renderDashboard(state);
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

export function readFiltersFromDom(prev) {
    const p = prev || {};
    function keep(id, key) {
        const el = document.getElementById(id);
        return el ? String(el.value || '') : String(p[key] || '');
    }
    return {
        status: keep('pwFilterStatus', 'status'),
        kategorie: keep('pwFilterKat', 'kategorie'),
        tag: keep('pwFilterTag', 'tag'),
        lehrer: keep('pwFilterLehrer', 'lehrer'),
        klasse: keep('pwFilterKlasse', 'klasse'),
        buchung: keep('pwFilterBuchung', 'buchung'),
        q: keep('pwFilterQ', 'q')
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
