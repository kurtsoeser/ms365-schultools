/**
 * UI-Rendering Freistellungs-Planer.
 */
import {
    computeDashboardKpis,
    validateFreistellung,
    approvalPath,
    inclusiveDayCount,
    toIsoDateTimeLocal,
    toIsoDateOnly,
    STATUS_CHOICES
} from './freistellung-planer-logic.js';
import * as FrLogic from './freistellung-planer-logic.js';
import { htmlFreistellungKategorienPanel } from './freistellung-kategorien-ui.js';
import { htmlFreistellungKlassenPanel } from './freistellung-klassen-ui.js';
import { NACHWEISE_MAX_FILES, NACHWEISE_MAX_BYTES, formatNachweisSize } from './freistellung-planer-nachweise.js';
import { renderFreistellungAccessDebugEntry } from './freistellung-planer-access-debug-ui.js';
import { renderFreistellungDrawer } from './freistellung-planer-detail-ui.js';
import { buildFrCalendarEvents, renderFrCalendarShell } from './freistellung-planer-calendar.js';
import { startOfWeekMonday } from '../../shared/ms365-calendar-shell.js';
import { renderSystemAdminPage } from '../../shared/system-admin-shell.js';
import { loadEffectivePermissionsConfig } from './freistellung-planer-permissions.js';
import * as FrState from './freistellung-planer-state.js';
import {
    viewsForRole,
    useStudentPlanerChrome,
    useKvPlanerChrome,
    useDirektionPlanerChrome,
    useMinimalPlanerChrome,
    filterItems,
    formatDeDateTime,
    statusLabel,
    roleLabel,
    scopeFromState,
    canDeleteFreistellung,
    resolveStudentKlasseCode,
    isStudentKlasseLocked,
    findClassInStammdaten,
    classesForStudentPicker,
    prefillStudentFreistellungForm,
    applyKvFromClass
} from './freistellung-planer-state.js';
import {
    isPlanerDemoRoleUiEnabled,
    canUsePlanerRoleSwitcher,
    sortPlanerRoles,
    PLANER_ROLE_ORDER
} from './freistellung-planer-entra-role.js';
import { mergeKategorieChoices, loadExtraKategorien } from './freistellung-planer-kategorien.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/** Fallback bis parseCodeList in logic landet (identisch AKT). */
function parseCodeList(value) {
    if (typeof FrLogic.parseCodeList === 'function') return FrLogic.parseCodeList(value);
    if (Array.isArray(value)) return value.map(String).map((s) => s.trim()).filter(Boolean);
    return String(value || '')
        .split(/[;|,]/)
        .map((s) => s.trim())
        .filter(Boolean);
}

function joinCodeList(value) {
    if (typeof FrLogic.joinCodeList === 'function') return FrLogic.joinCodeList(value);
    return parseCodeList(value).join('; ');
}

/**
 * Sektions-Nav wie AKT; Fallback: eine Sektion aus viewsForRole.
 * @param {object} state
 */
function resolveNavSections(state) {
    if (typeof FrState.navSectionsForRole === 'function') {
        return FrState.navSectionsForRole(state.role, state.planerRoles, {
            planerAdminAccess: state.planerAdminAccess
        });
    }
    return [{ id: 'main', label: '', items: viewsForRole(state.role, state) }];
}

function statusBadge(status) {
    const s = String(status || 'Ausstehend').trim();
    const cls =
        s === 'Genehmigt'
            ? 'fr-badge fr-badge--ok'
            : s === 'Abgelehnt'
              ? 'fr-badge fr-badge--bad'
              : 'fr-badge fr-badge--warn';
    return `<span class="${cls}">${esc(statusLabel(s))}</span>`;
}

function optionList(items, selected, emptyLabel, valueKey, labelKey) {
    const vk = valueKey || 'code';
    const lk = labelKey || 'name';
    const opts = [`<option value="">${esc(emptyLabel)}</option>`];
    (items || []).forEach((it) => {
        const v = typeof it === 'string' ? it : it[vk] || it.value;
        const label = typeof it === 'string' ? it : it[lk] || it.label || v;
        opts.push(
            `<option value="${esc(v)}"${String(selected) === String(v) ? ' selected' : ''}>${esc(label)}</option>`
        );
    });
    return opts.join('');
}

export function isFreistellungAccountBusy(state) {
    return !!(state && (state.loading || state.authRolePending || state.kontoOffen));
}

export function isFreistellungSignedOutAfterBoot(state) {
    if (isFreistellungAccountBusy(state)) return false;
    if (!state || !state.authSettled) return false;
    if (state.demoRoleOverride || state.localDemoOnly) return false;
    return !String(state.accountEmail || '').trim();
}

function renderNavSession(state) {
    if (
        isFreistellungAccountBusy(state) &&
        !String((state && state.accountEmail) || '').trim() &&
        !(state && state.demoRoleOverride) &&
        !(state && state.localDemoOnly)
    ) {
        return `
        <div class="fr-nav__session" aria-label="Anmeldung">
          <span class="fr-nav__role-label">Anmeldung</span>
          <p class="fr-nav__hint muted">Konto wird geprüft …</p>
        </div>`;
    }
    if (isFreistellungSignedOutAfterBoot(state)) {
        return `
        <div class="fr-nav__session" aria-label="Anmeldung">
          <span class="fr-nav__role-label">Anmeldung</span>
          <p class="fr-nav__hint">Nicht angemeldet</p>
        </div>`;
    }
    const studentChrome = useStudentPlanerChrome(state);
    const demoRoleUi = isPlanerDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    const switchable = demoRoleUi ? PLANER_ROLE_ORDER : sortPlanerRoles(state.planerRoles || []);
    const showRoleSwitcher = canUsePlanerRoleSwitcher(state.planerRoles || [], demoRoleUi);
    const isSchueler = state.role === 'schueler';
    const klasseCode = resolveStudentKlasseCode(state);
    const klasseHit = findClassInStammdaten(state.stammdaten.classes, klasseCode);
    const klasseLabel = klasseCode ? (klasseHit && (klasseHit.name || klasseHit.code)) || klasseCode : '';
    const pickerClasses = classesForStudentPicker(state);

    const roleBlock = showRoleSwitcher
        ? `<div class="fr-nav__role-btns" role="group" aria-label="Rolle wählen">
            ${switchable
                .map(
                    (r) =>
                        `<button type="button" class="fr-chip${state.role === r ? ' is-active' : ''}" data-fr-role="${esc(
                            r
                        )}">${esc(roleLabel(r))}</button>`
                )
                .join('')}
          </div>`
        : '';

    const kvChrome = useKvPlanerChrome(state);
    if (studentChrome || kvChrome) {
        return `
        <div class="fr-nav__session" aria-label="Anmeldung">
          <span class="fr-nav__role-label">Anmeldung</span>
          ${
              state.accountEmail || state.accountName ||
              (typeof window !== 'undefined' &&
                  typeof window.ms365AuthIsLoggedIn === 'function' &&
                  window.ms365AuthIsLoggedIn())
                  ? `<p class="fr-nav__account-name">${esc(state.accountName || state.accountEmail || 'Angemeldet')}</p>`
                  : '<p class="fr-nav__hint">Mit Schul-Konto anmelden:</p>'
          }
          ${
              isSchueler && klasseCode
                  ? `<p class="fr-nav__hint">Klasse <span class="fr-badge fr-nav__account-klasse">${esc(klasseLabel)}</span></p>`
                  : ''
          }
          ${
              isSchueler && !klasseCode && pickerClasses.length
                  ? `<label class="fr-nav__klasse" for="frDemoKlasse">Klasse
              <select id="frDemoKlasse">
                <option value="">– wählen –</option>
                ${pickerClasses
                    .map(
                        (c) =>
                            `<option value="${esc(c.code)}">${esc(c.name || c.code)}</option>`
                    )
                    .join('')}
              </select>
            </label>
            <p class="fr-nav__hint">Ihre E-Mail ist in den Schul-Stammdaten noch keiner Klasse zugeordnet – bitte Klasse wählen.</p>`
                  : isSchueler && !klasseCode
                    ? `<p class="fr-nav__hint">Klasse konnte nicht ermittelt werden. Stammdaten (Schülerliste) fehlen auf diesem Gerät – bitte Sekretariat/IT.</p>`
                    : ''
          }
          ${
              kvChrome && showRoleSwitcher
                  ? `<div class="fr-nav__session-role" style="margin-top:10px;">
              <span class="fr-nav__session-role-label">Rolle</span>
              ${roleBlock}
            </div>`
                  : kvChrome
                    ? ''
                    : showRoleSwitcher
                      ? `<div class="fr-nav__session-role" style="margin-top:10px;">
              <span class="fr-nav__session-role-label">Ansicht</span>
              ${roleBlock}
            </div>`
                      : ''
          }
          ${
              kvChrome
                  ? `<p class="fr-nav__hint muted" style="margin-top:6px;">Genehmigung in Teams/Outlook (Approvals).</p>`
                  : ''
          }
        </div>`;
    }

    return `
        <div class="fr-nav__session" aria-label="Anmeldung und Rolle">
          <span class="fr-nav__role-label">Anmeldung</span>
          ${
              state.accountEmail || state.accountName ||
              (typeof window !== 'undefined' &&
                  typeof window.ms365AuthIsLoggedIn === 'function' &&
                  window.ms365AuthIsLoggedIn())
                  ? `<p class="fr-nav__account-name">${esc(state.accountName || state.accountEmail || 'Angemeldet')}</p>`
                  : '<p class="fr-nav__hint">Mit Schul-Konto anmelden:</p>'
          }
          ${
              state.accountEmail || switchable.length
                  ? `<div class="fr-nav__session-role">
              <span class="fr-nav__session-role-label">Rolle${showRoleSwitcher ? ' wählen' : ''}</span>
              ${roleBlock}
              ${
                  isSchueler && klasseCode
                      ? `<span class="fr-badge fr-nav__account-klasse">${esc(klasseLabel)}</span>`
                      : ''
              }
            </div>`
                  : ''
          }
          ${
              isSchueler
                  ? `<label class="fr-nav__klasse" for="frDemoKlasse">Klasse
              <select id="frDemoKlasse" ${state.studentMatch ? 'disabled' : ''}>
                <option value="">– wählen –</option>
                ${(state.stammdaten.classes || [])
                    .map(
                        (c) =>
                            `<option value="${esc(c.code)}"${klasseCode === c.code ? ' selected' : ''}>${esc(
                                c.name || c.code
                            )}</option>`
                    )
                    .join('')}
              </select>
            </label>
            <p class="fr-nav__hint">${
                state.studentMatch
                    ? 'Klasse aus Stammdaten (Schüler-E-Mail).'
                    : demoRoleUi
                      ? 'Demo: Klasse wählen, wenn keine Schüler-E-Mail in den Stammdaten.'
                      : state.planerRoleSources && state.planerRoleSources.schueler === 'direktion-schueler'
                        ? 'Als Direktion: Klasse wählen, um die Schüler-Ansicht zu testen.'
                        : 'Schüler-Entra-Gruppe oder E-Mail in Stammdaten (students).'
            }</p>`
                  : showRoleSwitcher && !demoRoleUi
                    ? '<p class="fr-nav__hint">Mehrere Berechtigungen – hier umschalten. Genehmigung selbst in Teams/Outlook.</p>'
                    : demoRoleUi
                      ? '<p class="fr-nav__hint">IT-Vorschau (?demoRole=1).</p>'
                      : '<p class="fr-nav__hint">Genehmigung läuft über Microsoft Approvals (Power Automate).</p>'
          }
        </div>`;
}

function showPlanerDevTools(state) {
    const demoRoleUi = isPlanerDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    if (demoRoleUi) return true;
    return state.role === 'direktion';
}

function viewTitle(state) {
    const id = state.view;
    if (id === 'administration') return 'Systemadministration';
    if (id === 'freigabe') return 'Entscheidungen';
    if (id === 'bericht') return 'Berichte';
    if (id === 'antrag') return 'Antrag stellen';
    if (id === 'meine') return 'Meine Anträge';
    if (id === 'liste') return state.role === 'direktion' ? 'Alle Anträge' : 'Klassenliste';
    if (id === 'kalender') return 'Kalender';
    if (id === 'dashboard') return state.role === 'kv' ? 'Meine Klasse' : 'Übersicht';
    return 'Freistellungen';
}

function viewLead(state) {
    const id = state.view;
    if (id === 'administration') {
        return 'Datenquelle, Berechtigungen, Kalender und Wartung. Das IT-Setup bleibt eine eigene Seite.';
    }
    if (state.role === 'schueler' || (state.role !== 'kv' && state.role !== 'direktion')) {
        if (id === 'antrag') return 'Zeitraum und Grund angeben. Die Freigabe läuft über Microsoft Approvals.';
        if (id === 'kalender') return 'Eigene Freistellungen im Kalender.';
        return 'Eingereichte Freistellungen und aktueller Status.';
    }
    if (state.role === 'kv') return 'Anträge Ihrer Klasse – Status und Übersicht.';
    if (id === 'freigabe') return 'Offene Anträge. Die Freigabe läuft in Microsoft Approvals.';
    if (id === 'kalender') return 'Genehmigte und offene Freistellungen.';
    return 'Übersicht, Anträge und Genehmigungen. Technik liegt unter Systemadministration.';
}

function renderAuthRequiredPanel() {
    return `<section class="tm-panel fr-status-panel" aria-live="polite">
      <h2 class="fr-h2">Anmeldung erforderlich</h2>
      <p>Bitte melden Sie sich mit Ihrem Schul-Konto (Microsoft 365) an, um Ihre Berechtigung und Freistellungsanträge zu laden.</p>
      <div class="fr-status-panel__actions">
        <button type="button" class="btn btn-primary" data-fr-auth-login><i class="bi bi-box-arrow-in-right"></i> Mit Microsoft 365 anmelden</button>
      </div>
    </section>`;
}

function renderAccessDeniedPanel(state) {
    const hint = state.roleHintPublic || state.roleHintStaff || state.roleHint || '';
    const student = useStudentPlanerChrome(state);
    const kv = useKvPlanerChrome(state);
    const deniedFallback = student
        ? 'Derzeit kein Zugriff auf Ihre Freistellungsanträge.'
        : kv
          ? 'Derzeit kein Zugriff auf die Freistellungsanträge Ihrer Klasse.'
          : 'Kein Zugriff auf den Freistellungs-Planer für dieses Konto.';
    return `<section class="tm-panel fr-status-panel" role="alert">
      <h2 class="fr-h2">Kein Zugriff auf den Freistellungs-Planer</h2>
      <p>${esc(hint || deniedFallback)}</p>
      ${
          state.accountEmail
              ? `<p class="muted fr-status-panel__meta">Geprüftes Konto: <strong>${esc(state.accountEmail)}</strong></p>`
              : ''
      }
      <div class="fr-status-panel__actions">
        <button type="button" class="btn" data-fr-auth-switch><i class="bi bi-arrow-left-right"></i> Anderes Konto wählen</button>
      </div>
    </section>`;
}

function renderWorkTop(state) {
    if (
        isFreistellungAccountBusy(state) ||
        isFreistellungSignedOutAfterBoot(state) ||
        state.planerAccessDenied
    ) {
        return '';
    }
    const onAdmin = state.view === 'administration';
    const kv = useKvPlanerChrome(state);
    const kvClasses = kv
        ? (state.stammdaten.classes || []).filter((c) => {
              const em = String(c.headEmail || c.klassenvorstandEmail || c.kvEmail || '')
                  .trim()
                  .toLowerCase();
              return em && em === String(state.accountEmail || '').toLowerCase();
          })
        : [];
    const classLabels = kvClasses.map((c) => c.name || c.code).filter(Boolean);
    const selfTitled =
        state.view === 'dashboard' ||
        state.view === 'freigabe' ||
        state.view === 'liste' ||
        state.view === 'meine' ||
        state.view === 'bericht' ||
        state.view === 'antrag' ||
        !state.view;
    if (selfTitled) {
        if (!state.accountEmail && !classLabels.length) return '';
        return `<div class="fr-top-refresh">
            ${
                classLabels.length
                    ? `<span class="fb-pill">Klasse ${esc(classLabels.join(', '))}</span>`
                    : ''
            }
            ${
                state.accountEmail
                    ? `<button type="button" class="fb-btn fb-btn--ghost" id="frBtnLoad" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>`
                    : ''
            }
        </div>`;
    }
    return `
        <header class="fr-top fr-top--student">
          <div class="fr-top__copy">
            <h2 class="fr-top__student-title">${esc(viewTitle(state))}</h2>
            <p class="fr-top__lead">${esc(viewLead(state))}</p>
            ${
                classLabels.length
                    ? `<p class="fr-top__lead">Klassen: <span class="fb-pill">${esc(
                          classLabels.join(', ')
                      )}</span></p>`
                    : ''
            }
          </div>
          <div class="fr-top__user">
            ${
                state.accountEmail && !onAdmin
                    ? `<button type="button" class="fb-btn fb-btn--ghost" id="frBtnLoad" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>`
                    : ''
            }
          </div>
        </header>`;
}

function renderAdministration(state) {
    if (state.role !== 'direktion') {
        return '<p class="muted">Nur für Direktion / Admin.</p>';
    }
    const perms = loadEffectivePermissionsConfig();
    const live = state.stammdatenLive || {};
    const liveLine = live.ok
        ? `${Number(live.classCount) || 0} Klassen / ${Number(live.teacherCount) || 0} Lehrkräfte (${live.source || 'live'})`
        : live.error
          ? String(live.error)
          : `${(state.stammdaten.classes || []).length} Klassen lokal`;
    const connected = !!(state.siteUrl && !state.error && !state.localDemoOnly);
    const devTools = showPlanerDevTools(state);
    const groupLine = (name, id) => esc(name || id || '–');
    return renderSystemAdminPage({
        prefix: 'fr',
        activeTab: state.systemTab,
        dataAttr: 'data-fr-sys-tab',
        ariaLabel: 'Systemadministration Freistellungen',
        panels: [
            {
                id: 'site',
                step: '1',
                title: 'Datenquelle',
                leadHtml: 'Liste „' + esc(state.listName || 'Freistellungen') + '“ auf der Schul-Site.',
                body: `
        <div class="ms-sys-health" role="status">
          <span class="ms-sys-status ms-sys-status--${state.error ? 'bad' : connected ? 'ok' : 'warn'}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${
              state.error ? 'Fehler' : connected ? 'Verbunden' : state.siteUrl ? 'URL gesetzt' : 'Nicht verbunden'
          }</span>
          <span class="ms-sys-status ms-sys-status--${live.ok ? 'ok' : 'warn'}"><span class="ms-sys-status__dot" aria-hidden="true"></span>${esc(
              liveLine
          )}</span>
          ${state.localDemoOnly ? '<span class="ms-sys-status ms-sys-status--warn"><span class="ms-sys-status__dot" aria-hidden="true"></span>Demo lokal</span>' : ''}
        </div>
        <div class="ms-sys-field">
          <label for="frSiteUrl">SharePoint-URL</label>
          <div class="ms-sys-field__row">
            <input type="url" id="frSiteUrl" class="ms-sys-input" value="${esc(
                state.siteUrl
            )}" placeholder="https://…sharepoint.com/sites/…" spellcheck="false" autocomplete="off">
            <button type="button" class="btn ms-sys-btn-primary" id="frBtnLoad"><i class="bi bi-arrow-repeat"></i>Verbindung prüfen</button>
          </div>
        </div>
        <div class="ms-sys-links">
          <a class="btn ms-sys-btn-secondary" href="freistellung-setup.html"><i class="bi bi-wrench"></i>IT-Setup</a>
        </div>`
            },
            {
                id: 'perms',
                step: '2',
                title: 'Berechtigungen',
                leadHtml: 'Überblick der gespeicherten Entra-Gruppen. Listenrechte setzt das IT-Setup.',
                body: `
        <dl class="ms-sys-dl">
          <dt>Direktion</dt><dd>${groupLine(perms.groupDirektion, perms.groupDirektionId)}</dd>
          <dt>Klassenvorstand</dt><dd>${groupLine(perms.groupKv, perms.groupKvId)}</dd>
          <dt>Schüler/innen</dt><dd>${groupLine(perms.groupSchueler, perms.groupSchuelerId)}</dd>
          <dt>Direktion-Mail (Flow)</dt><dd>${esc(state.emailDirektion || '–')}</dd>
        </dl>
        <div class="ms-sys-links">
          <a class="btn ms-sys-btn-secondary" href="freistellung-setup.html"><i class="bi bi-shield-check"></i>IT-Setup · Berechtigungen</a>
        </div>`
            },
            {
                id: 'kalender',
                step: '3',
                title: 'Kalender',
                leadHtml: 'Der Arbeitskalender bleibt die Ansicht „Kalender“. Hier nur der Export.',
                body: `
        <div class="ms-sys-actions">
          <button type="button" class="btn ms-sys-btn-secondary" data-fr-view-jump="kalender"><i class="bi bi-calendar3"></i>Kalender öffnen</button>
          <button type="button" class="btn ms-sys-btn-ghost" id="frBtnCsv"><i class="bi bi-download"></i>CSV</button>
        </div>`
            },
            {
                id: 'wartung',
                step: '4',
                title: 'Wartung',
                leadHtml: 'Klassen, Kategorien, Stammdaten und Hilfe – getrennt vom Genehmigungsalltag.',
                body: `
        <div class="ms-sys-links" style="margin-bottom:14px">
          <a class="btn ms-sys-btn-secondary" href="../tenant.html"><i class="bi bi-gear"></i>Stammdaten</a>
          <a class="btn ms-sys-btn-ghost" href="freistellung-konzept.html"><i class="bi bi-diagram-3"></i>Ablauf</a>
          <a class="btn ms-sys-btn-ghost" href="lehrer-freistellung-planer.html"><i class="bi bi-briefcase"></i>Lehrkräfte</a>
          <a class="btn ms-sys-btn-ghost" href="../hilfe.html#tool-freistellung-planer"><i class="bi bi-life-preserver"></i>Hilfe</a>
        </div>
        ${htmlFreistellungKlassenPanel({
            mountId: 'frPlanerKlassenList',
            syncBtnId: 'frPlanerKlassenSyncSp',
            reloadBtnId: 'frPlanerKlassenReload',
            cleanBtnId: 'frPlanerKlassenClean'
        })}
        ${htmlFreistellungKategorienPanel({
            listId: 'frPlanerKatExtraList',
            newInputId: 'frPlanerKatExtraNew',
            addId: 'frPlanerKatExtraAdd',
            syncListBtnId: 'frPlanerKatSyncSp'
        })}`
            },
            {
                id: 'danger',
                title: 'Danger Zone',
                danger: true,
                leadHtml: 'Lokale Demo verwerfen. SharePoint-Einträge bleiben, bis sie bewusst gelöscht werden.',
                body: devTools
                    ? `<div class="ms-sys-actions">
          <button type="button" class="btn ms-sys-btn-secondary" id="frBtnDemo"><i class="bi bi-database"></i>Demo laden</button>
          <button type="button" class="btn ms-sys-btn-ghost" id="frBtnJsonImport"><i class="bi bi-filetype-json"></i>JSON</button>
          <input type="file" id="frJsonImportFile" accept="application/json,.json" hidden>
          <button type="button" class="btn ms-sys-btn-danger" id="frBtnDemoReset"><i class="bi bi-trash"></i>Demo zurücksetzen</button>
        </div>`
                    : '<p class="muted">Keine Demo-Aktionen in dieser Sitzung.</p>'
            }
        ]
    });
}

export function renderApp(state, root) {
    if (!root) return;
    const sections = resolveNavSections(state);
    const studentChrome = useStudentPlanerChrome(state);
    const kvChrome = useKvPlanerChrome(state);
    const direktionChrome = useDirektionPlanerChrome(state);
    const minimalChrome =
        typeof useMinimalPlanerChrome === 'function'
            ? useMinimalPlanerChrome(state)
            : studentChrome || kvChrome || direktionChrome;
    const showAccessDebug = !minimalChrome || state.view === 'administration';
    const personaBanner = state.demoPersonaEmail
        ? `<div class="fr-alert fr-alert--info" role="status">Demo-Ansicht als <strong>${esc(
              state.demoPersonaName || 'Lehrkraft'
          )}</strong> (${esc(state.demoPersonaEmail)}). Angemeldet bleibt ${esc(
              state.accountName || state.accountEmail || 'IT'
          )}.</div>`
        : '';
    const openCount = (state.items || []).filter(
        (it) => String((it && it.status) || '').trim().toLowerCase() === 'ausstehend'
    ).length;
    const navHtml = sections
        .map((sec) => {
            const label = sec.label
                ? `<span class="fr-nav__section-label">${esc(sec.label)}</span>`
                : '';
            const btns = (sec.items || [])
                .map((v) => {
                    const badge =
                        v.id === 'freigabe' && openCount > 0
                            ? `<span class="fr-nav__badge">${openCount}</span>`
                            : '';
                    return `
            <button type="button" class="fr-nav__btn${state.view === v.id ? ' is-active' : ''}${
                v.id === 'antrag' ? ' fr-nav__btn--cta' : ''
            }" data-fr-view="${esc(v.id)}">
              <i class="bi ${v.icon}" aria-hidden="true"></i><span class="fr-nav__btn-label">${esc(
                  v.label
              )}</span>${badge}
            </button>`;
                })
                .join('');
            return `<div class="fr-nav__section">${label}${btns}</div>`;
        })
        .join('');
    root.innerHTML = `
    <div class="fr-shell fr-shell--student">
      <aside class="fr-nav" aria-label="Navigation">
        <div class="fr-nav__brand">
          <i class="bi bi-calendar2-check" aria-hidden="true"></i>
          <div>
            <strong>Freistellungen</strong>
            <span>${
                studentChrome ? 'Schüler/innen' : kvChrome ? 'Klassenübersicht' : 'Verwaltung'
            }</span>
          </div>
        </div>
        <nav class="fr-nav__list">
          ${navHtml}
        </nav>
        ${renderNavSession(state)}
      </aside>
      <div class="fr-main">
        ${renderWorkTop(state)}
        ${personaBanner}
        ${showAccessDebug ? renderFreistellungAccessDebugEntry(state) : ''}
        ${
            !isFreistellungAccountBusy(state) &&
            !isFreistellungSignedOutAfterBoot(state) &&
            !state.demoPersonaEmail &&
            state.localDemoOnly &&
            state.view === 'administration'
                ? '<div class="fr-alert fr-alert--info" role="status">Lokaler Demo-Modus – Anzeige ohne SharePoint. Mit „Laden“ echte Liste verbinden.</div>'
                : ''
        }
        ${
            !isFreistellungAccountBusy(state) &&
            !isFreistellungSignedOutAfterBoot(state) &&
            !state.planerAccessDenied &&
            state.info
                ? `<div class="fr-alert fr-alert--info" role="status">${esc(state.info)}</div>`
                : ''
        }
        ${
            !isFreistellungAccountBusy(state) && !isFreistellungSignedOutAfterBoot(state) && state.error
                ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(state.error)}</div>`
                : ''
        }
        ${
            isFreistellungAccountBusy(state)
                ? '<div class="fr-skel" role="status" aria-busy="true" aria-label="Konto und Berechtigungen werden geprüft"><div class="fr-skel__card"></div><div class="fr-skel__card"></div></div>'
                : ''
        }
        <div id="frView" class="fr-view${isFreistellungAccountBusy(state) ? ' is-loading' : ''}">
          ${
              isFreistellungAccountBusy(state)
                  ? ''
                  : isFreistellungSignedOutAfterBoot(state)
                    ? renderAuthRequiredPanel()
                    : state.planerAccessDenied
                      ? renderAccessDeniedPanel(state)
                      : renderView(state)
          }
        </div>
      </div>
    </div>
    ${state.detailId ? renderFreistellungDrawer(state) : ''}
  `;
}

export function htmlFreistellungUploadField(opts) {
    const student = !!(opts && opts.student);
    const maxMb = Math.round(NACHWEISE_MAX_BYTES / 1024 / 1024);
    return `
    <div class="fr-upload" id="frUploadWrap">
      <p class="fr-upload__hint">${
          student
              ? 'z. B. ärztliche Bestätigung oder Einladung – wird mit dem Antrag gespeichert.'
              : 'PDF, Fotos oder Word – max. ' +
                NACHWEISE_MAX_FILES +
                ' Dateien, je ' +
                maxMb +
                ' MB.'
      }</p>
      <div class="fr-upload__drop">
        <div class="fr-upload__actions">
          <label class="btn" for="frFormFiles"><i class="bi bi-cloud-arrow-up"></i>Dateien auswählen</label>
          <input type="file" id="frFormFiles" class="fr-upload__input" multiple accept=".pdf,.jpg,.jpeg,.png,.heic,.doc,.docx,.webp,image/*,application/pdf">
          <span class="muted fr-upload__or">oder Dateien hier ablegen</span>
        </div>
      </div>
      <ul id="frFormFilesList" class="fr-upload__list" aria-live="polite"></ul>
    </div>`;
}

/**
 * Dateiliste unter dem Upload-Feld aktualisieren.
 * @param {HTMLElement} root
 */
export function bindFreistellungUploadField(root) {
    const host = (root || document).querySelector('#frUploadWrap');
    const input = (root || document).querySelector('#frFormFiles');
    const list = (root || document).querySelector('#frFormFilesList');
    if (!host || !input || !list) return;
    if (host.dataset.frUploadBound === '1') {
        renderFreistellungUploadFileList(input, list);
        return;
    }
    host.dataset.frUploadBound = '1';
    input.addEventListener('change', () => renderFreistellungUploadFileList(input, list));
    renderFreistellungUploadFileList(input, list);
}

function renderFreistellungUploadFileList(input, list) {
    const files = input && input.files ? Array.from(input.files) : [];
    if (!files.length) {
        list.innerHTML = '<li class="fr-upload__empty muted">Noch keine Dateien gewählt.</li>';
        return;
    }
    list.innerHTML = files
        .map(
            (f) =>
                '<li class="fr-upload__file"><i class="bi bi-file-earmark" aria-hidden="true"></i><span>' +
                esc(f.name) +
                '</span><span class="muted">' +
                esc(formatNachweisSize(f.size)) +
                '</span></li>'
        )
        .join('');
}

export function clearFreistellungUploadField(root) {
    const input = (root || document).querySelector('#frFormFiles');
    const list = (root || document).querySelector('#frFormFilesList');
    if (input) input.value = '';
    if (input && list) renderFreistellungUploadFileList(input, list);
}

function kategorieChoicesForState(state) {
    const fromState = state && state.kategorieChoices;
    if (Array.isArray(fromState) && fromState.length) return fromState;
    return mergeKategorieChoices(loadExtraKategorien());
}

/**
 * Kompakte Mehrfachauswahl in der Filterleiste (wie AKT).
 * @param {string} key
 * @param {{code?:string,value?:string,name?:string,label?:string}[]} items
 * @param {string} selectedCsv
 * @param {boolean} [open]
 * @param {{ allLabel?: string, countLabel?: string, ariaLabel?: string, searchPlaceholder?: string, emptyHint?: string, showSearch?: boolean, query?: string }} [opts]
 */
function renderFilterMulti(key, items, selectedCsv, open, opts) {
    const o = opts || {};
    const selected = parseCodeList(selectedCsv);
    const selectedSet = new Set(selected.map(String));
    const n = selected.length;
    const allLabel = o.allLabel || 'Alle';
    const countLabel = o.countLabel || 'ausgewählt';
    const query = String(o.query || '')
        .trim()
        .toLowerCase();
    let summary = allLabel;
    if (n === 1) {
        const hit = (items || []).find((it) => String(it.code || it.value || '').trim() === selected[0]);
        summary = (hit && (hit.name || hit.label)) || selected[0];
    } else if (n > 1) {
        summary = n + ' ' + countLabel;
    }
    const inputName = 'frFilter' + key.charAt(0).toUpperCase() + key.slice(1);
    const rows = (items || [])
        .map((it) => {
            const v = String(it.code || it.value || '').trim();
            if (!v) return '';
            const label = it.name || it.label || v;
            const checked = selectedSet.has(v);
            const hay = String(label)
                .trim()
                .toLowerCase();
            const hidden = !!(query && hay.indexOf(query) === -1);
            return `<label class="fr-filter-multi__item${hidden ? ' is-filtered-out' : ''}">
        <input type="checkbox" name="${esc(inputName)}" value="${esc(v)}"${checked ? ' checked' : ''}>
        <span>${esc(label)}</span>
      </label>`;
        })
        .filter(Boolean)
        .join('');
    const showSearch = o.showSearch !== false;
    return `<details class="fr-filter-multi${n ? ' is-filtered' : ''}" data-fr-filter-multi="${esc(key)}"${
        open ? ' open' : ''
    }>
      <summary class="fr-filter-multi__summary" aria-label="${esc(o.ariaLabel || allLabel)}">${esc(
          summary
      )}</summary>
      <div class="fr-filter-multi__panel">
        ${
            showSearch
                ? `<input type="search" class="fr-control fr-filter-multi__search" data-fr-filter-multi-q="${esc(
                      key
                  )}" value="${esc(o.query || '')}" placeholder="${esc(
                      o.searchPlaceholder || 'Suchen …'
                  )}" autocomplete="off" aria-label="Liste filtern">`
                : ''
        }
        <div class="fr-filter-multi__list" role="group" aria-label="${esc(o.ariaLabel || allLabel)}">
          ${rows || `<p class="fr-filter-multi__empty">${esc(o.emptyHint || 'Keine Einträge')}</p>`}
        </div>
        ${
            n
                ? `<button type="button" class="fr-text-link fr-filter-multi__clear" data-fr-filter-multi-clear="${esc(
                      key
                  )}">Auswahl leeren</button>`
                : ''
        }
      </div>
    </details>`;
}

function filtersActive(filters) {
    const f = filters || {};
    return !!(f.q || f.klasse || f.status || f.kategorie || f.multiDay);
}

/** @param {object} filters @param {string} preset */
function statusPresetActive(filters, preset) {
    const list = parseCodeList(filters && filters.status).map((s) => String(s).trim().toLowerCase());
    const p = String(preset || '')
        .trim()
        .toLowerCase();
    if (!p) return !list.length;
    return list.length === 1 && list[0] === p;
}

function filterBar(state, opts) {
    const o = opts || {};
    const sd = state.stammdaten || {};
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion', ...(o.scope || {}) });
    const count = filterItems(state.items, state.filters, scope).length;
    const active = filtersActive(state.filters);
    const statusItems = (STATUS_CHOICES || []).map((s) => ({ code: s, name: s }));
    const katItems = kategorieChoicesForState(state).map((k) => ({ code: k, name: k }));
    const open = state.filterMultiOpen || {};
    const qMap = state.filterMultiQuery || {};
    return `
    ${
        o.calendarPresets
            ? `<div class="fr-seg" role="group" aria-label="Status-Voreinstellung">
        <button type="button" class="fr-seg__btn${statusPresetActive(state.filters, '') ? ' is-active' : ''}" data-fr-status-preset="">Alle</button>
        <button type="button" class="fr-seg__btn${statusPresetActive(state.filters, 'Genehmigt') ? ' is-active' : ''}" data-fr-status-preset="Genehmigt">Genehmigt</button>
        <button type="button" class="fr-seg__btn${statusPresetActive(state.filters, 'Ausstehend') ? ' is-active' : ''}" data-fr-status-preset="Ausstehend">Ausstehend</button>
      </div>`
            : ''
    }
    <div class="fr-filters${active ? ' fr-filters--active' : ''}" data-fr-filters>
      ${
          o.showToday
              ? '<button type="button" class="btn fr-btn-ghost fr-btn-compact" id="frCalToday">Heute</button>'
              : ''
      }
      <input type="search" id="frFilterQ" placeholder="Suche…" value="${esc(state.filters.q || '')}" aria-label="Suche">
      ${renderFilterMulti('klasse', sd.classes || [], state.filters.klasse, !!open.klasse, {
          allLabel: 'Alle Klassen',
          countLabel: 'Klassen',
          ariaLabel: 'Klassen filtern',
          searchPlaceholder: 'Klasse suchen …',
          emptyHint: 'Keine Klassen verfügbar',
          query: qMap.klasse
      })}
      ${renderFilterMulti('status', statusItems, state.filters.status, !!open.status, {
          allLabel: 'Alle Status',
          countLabel: 'Status',
          ariaLabel: 'Status filtern',
          emptyHint: 'Keine Status',
          showSearch: false
      })}
      ${renderFilterMulti('kategorie', katItems, state.filters.kategorie, !!open.kategorie, {
          allLabel: 'Alle Kategorien',
          countLabel: 'Kategorien',
          ariaLabel: 'Kategorie filtern',
          emptyHint: 'Keine Kategorien',
          showSearch: false
      })}
      <select id="frFilterMulti" aria-label="Dauer">
        <option value="">Alle Dauern</option>
        <option value="0"${state.filters.multiDay === '0' ? ' selected' : ''}>1 Tag</option>
        <option value="1"${state.filters.multiDay === '1' ? ' selected' : ''}>Mehrtägig</option>
      </select>
      ${
          active
              ? '<button type="button" class="fr-text-link fr-filter-reset" id="frFilterReset">Filter zurücksetzen</button>'
              : '<button type="button" class="btn" id="frFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Reset</button>'
      }
      <span class="fr-filters__meta">${count} ${count === 1 ? 'Eintrag' : 'Einträge'}</span>
    </div>`;
}

function renderView(state) {
    switch (state.view) {
        case 'liste':
            return renderListe(state);
        case 'kalender':
            return renderKalender(state);
        case 'antrag':
            return renderAntrag(state);
        case 'meine':
            return renderMeine(state);
        case 'freigabe':
            return renderFreigabe(state);
        case 'bericht':
            return renderBericht(state);
        case 'administration':
            return renderAdministration(state);
        default:
            return renderDashboard(state);
    }
}

function renderFrDecisionRow(it) {
    const id = it.itemId || it.antragId;
    const s = String(it.status || 'Ausstehend').trim();
    const pillMod = s === 'Genehmigt' ? 'ok' : s === 'Abgelehnt' ? 'bad' : 'warn';
    const pillLabel = s === 'Ausstehend' ? 'Offen' : statusLabel(s);
    const title = it.schuelerName || it.titel || 'Antrag';
    const metaParts = [it.klasse, it.kategorie].filter(Boolean);
    const range =
        it.ende && it.ende !== it.beginn
            ? `${formatDeDateTime(it.beginn)} – ${formatDeDateTime(it.ende)}`
            : formatDeDateTime(it.beginn);
    return `<button type="button" class="fb-rowbtn" data-fr-detail="${esc(id)}">
      <div class="fb-rowbtn__main">
        <div class="fb-rowbtn__title">${esc(title)}</div>
        <div class="fb-rowbtn__meta">${esc(metaParts.join(' · ') || 'Freistellung')}</div>
      </div>
      <div class="fb-rowbtn__date">${esc(range)}</div>
      <span class="fb-pill fb-pill--${pillMod}">${esc(pillLabel)}</span>
    </button>`;
}

function renderDashboard(state) {
    const kvChrome = useKvPlanerChrome(state);
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, {}, scope);
    const offen = items.filter((i) => String(i.status) === 'Ausstehend');
    const today = toIsoDateOnly(new Date()) || '';

    const headline =
        offen.length === 1
            ? '<span class="fb-num">1</span> Antrag ohne Entscheidung.'
            : `<span class="fb-num">${offen.length}</span> ohne Entscheidung.`;
    const sub = offen.length
        ? 'Zeile anklicken, Details prüfen. Freigaben laufen über Microsoft Approvals oder direkt im Detailfeld.'
        : 'Aktuell wartet keine Entscheidung auf Sie.';

    if (kvChrome) {
        const sorted = [...items].sort((a, b) => String(b.beginn || '').localeCompare(String(a.beginn || '')));
        return `
    <h1 class="fb-greet">${headline}</h1>
    <p class="fb-sub">${esc(sub)}</p>
    ${
        offen.length
            ? `<div class="fb-sec-label">OFFEN IN IHRER KLASSE</div>
               <div class="fb-rowlist">${offen.slice(0, 8).map(renderFrDecisionRow).join('')}</div>`
            : ''
    }
    <div class="fb-sec-label">ANTRÄGE IHRER KLASSE</div>
    ${
        sorted.length
            ? renderTable(sorted.slice(0, 12), state)
            : (state.items || []).length
              ? '<p class="muted">In der Liste liegen Anträge, aber keiner passt zu Ihrer Klasse/KV-Zuordnung. IT: Stammdaten <code>headEmail</code> prüfen.</p>'
              : '<p class="muted">Aktuell keine Anträge für Ihre Klasse.</p>'
    }
    <div class="fb-actions-row">
      <button type="button" class="fb-btn" data-fr-view-jump="antrag">Neuer Antrag</button>
      <button type="button" class="fb-btn fb-btn--ghost" data-fr-view-jump="kalender">Kalender</button>
    </div>`;
    }

    const nextItems = items
        .filter((i) => {
            const st = String(i.status || '').trim();
            if (st !== 'Genehmigt') return false;
            const start = toIsoDateOnly(i.beginn) || '';
            return start && today && start >= today;
        })
        .sort((a, b) => String(a.beginn || '').localeCompare(String(b.beginn || '')))
        .slice(0, 5);

    return `
    <h1 class="fb-greet">${headline}</h1>
    <p class="fb-sub">${esc(sub)}</p>
    ${
        offen.length
            ? `<div class="fb-sec-label">OFFEN</div>
               <div class="fb-rowlist">${offen.slice(0, 8).map(renderFrDecisionRow).join('')}</div>`
            : ''
    }
    <div class="fb-sec-label">DEMNÄCHST GENEHMIGT</div>
    ${
        nextItems.length
            ? `<div class="fb-rowlist">${nextItems.map(renderFrDecisionRow).join('')}</div>`
            : '<p class="muted" style="margin:6px 0 0;">Keine kommenden genehmigten Freistellungen.</p>'
    }
    <div class="fb-actions-row">
      ${
          offen.length
              ? '<button type="button" class="fb-btn" data-fr-view-jump="freigabe" data-fr-filter-status="Ausstehend">Entscheidungen öffnen</button>'
              : '<button type="button" class="fb-btn" data-fr-view-jump="antrag">Neuer Antrag</button>'
      }
      <button type="button" class="fb-btn fb-btn--ghost" data-fr-view-jump="liste">Alle Anträge</button>
      <button type="button" class="fb-btn fb-btn--ghost" data-fr-view-jump="kalender">Kalender</button>
    </div>`;
}

function renderListe(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope);
    const headline =
        items.length === 1
            ? '<span class="fb-num">1</span> Antrag im Überblick.'
            : `<span class="fb-num">${items.length}</span> Anträge im Überblick.`;
    const sub =
        state.role === 'direktion'
            ? 'Gefiltert oder durchsucht. Zeile anklicken öffnet das Detail.'
            : 'Anträge Ihrer Klasse. Zeile anklicken öffnet das Detail.';
    return `
    <section class="fr-panel fr-panel--calm fr-panel--list">
      <h1 class="fb-greet">${headline}</h1>
      <p class="fb-sub">${esc(sub)}</p>
      ${filterBar(state)}
      ${
          items.length
              ? renderTable(items, state)
              : `<p class="fr-quiet-status" role="status">
        <i class="bi bi-check-circle" aria-hidden="true"></i>
        <span>Keine Einträge.</span>
      </p>`
      }
    </section>`;
}

function renderKalender(state) {
    const mode = String(state.calMode || 'month');
    const y = Number(state.calYear) || new Date().getFullYear();
    const m = Number(state.calMonth) || new Date().getMonth() + 1;
    let weekAnchor = new Date(state.calWeekAnchor || Date.now());
    if (Number.isNaN(weekAnchor.getTime())) weekAnchor = startOfWeekMonday(new Date());
    else weekAnchor = startOfWeekMonday(weekAnchor);

    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope);
    const events = buildFrCalendarEvents(items);
    const kvChrome = useKvPlanerChrome(state);
    const shell = renderFrCalendarShell({
        mode,
        year: y,
        month: m,
        weekAnchor,
        events,
        hint: kvChrome
            ? 'Abwesenheiten Ihrer Klasse(n) · Antippen für Details · Agenda, Woche und Monat.'
            : 'Freistellung antippen für Details · Agenda, Woche und Monat.'
    });

    return `
    <section class="fr-panel fr-panel--calm fr-cal-layout">
      <h2 class="fr-sr-only">Kalender</h2>
      ${filterBar(state, { calendarPresets: true, showToday: true })}
      ${shell}
    </section>`;
}

function renderMeine(state) {
    const studentChrome = useStudentPlanerChrome(state);
    const mineEmail =
        (typeof FrState.effectiveAccountEmail === 'function'
            ? FrState.effectiveAccountEmail(state)
            : '') || state.accountEmail;
    const items = filterItems(state.items, state.filters, {
        onlyMine: true,
        accountEmail: mineEmail
    });
    const headline =
        items.length === 1
            ? '<span class="fb-num">1</span> eigener Antrag.'
            : `<span class="fb-num">${items.length}</span> eigene Anträge.`;
    const sub = studentChrome
        ? 'Eingereichte Freistellungen und aktueller Status.'
        : 'Ihre eingereichten Freistellungen. Zeile anklicken öffnet das Detail.';
    const empty = `<p class="fr-quiet-status" role="status">
      <i class="bi bi-check-circle" aria-hidden="true"></i>
      <span>${
          studentChrome
              ? 'Noch keine Anträge. Stellen Sie Ihren ersten Antrag über den Knopf unten.'
              : 'Noch keine eigenen Anträge.'
      }</span>
    </p>`;
    return `
    <section class="fr-panel fr-panel--calm fr-panel--list${studentChrome ? ' fr-panel--student-meine' : ''}">
      <h1 class="fb-greet">${headline}</h1>
      <p class="fb-sub">${esc(sub)}</p>
      ${studentChrome ? '' : filterBar(state)}
      ${items.length ? `<div class="fb-rowlist">${items.map(renderFrDecisionRow).join('')}</div>` : empty}
      <div class="fb-actions-row">
        <button type="button" class="fb-btn" data-fr-view-jump="antrag">Antrag stellen</button>
        ${studentChrome ? '<button type="button" class="fb-btn fb-btn--ghost" data-fr-view-jump="kalender">Kalender</button>' : ''}
      </div>
    </section>`;
}

function renderFreigabe(state) {
    const email =
        (typeof FrState.effectiveAccountEmail === 'function'
            ? FrState.effectiveAccountEmail(state)
            : '') || state.accountEmail;
    const scope =
        state.role === 'direktion' && !state.demoPersonaEmail
            ? { accountEmail: email }
            : { onlyKv: true, accountEmail: email };
    const items = filterItems(state.items, { status: 'Ausstehend' }, scope);
    const headline =
        items.length === 1
            ? '<span class="fb-num">1</span> Antrag ohne Entscheidung.'
            : `<span class="fb-num">${items.length}</span> ohne Entscheidung.`;
    return `
    <section class="fr-panel fr-panel--calm fr-panel--list">
      <h1 class="fb-greet">${headline}</h1>
      <p class="fb-sub">Zeile anklicken, Nachweise und Zeitraum prüfen, rechts im Detailfeld entscheiden. In Produktion übernimmt Microsoft Approvals automatisch.</p>
      ${
          items.length
              ? `<div class="fb-rowlist">${items.map(renderFrDecisionRow).join('')}</div>`
              : `<p class="fr-quiet-status" role="status">
        <i class="bi bi-check-circle" aria-hidden="true"></i>
        <span>Keine offenen Anträge.</span>
      </p>`
      }
    </section>`;
}

function renderBericht(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope);
    const kpi = computeDashboardKpis(items);
    const byKlasse = {};
    items.forEach((it) => {
        const k = it.klasse || '–';
        byKlasse[k] = (byKlasse[k] || 0) + 1;
    });
    const rows = Object.keys(byKlasse)
        .sort()
        .map((k) => `<tr><td>${esc(k)}</td><td>${byKlasse[k]}</td></tr>`)
        .join('');
    return `
    <section class="fr-panel">
      <h1 class="fb-greet"><span class="fb-num">${kpi.gesamt}</span> ${kpi.gesamt === 1 ? 'Antrag' : 'Anträge'} ausgewertet.</h1>
      <p class="fb-sub">${kpi.genehmigt} genehmigt · ${kpi.abgelehnt} abgelehnt · ${kpi.mehrtage} mehrtägig.</p>
      ${filterBar(state)}
      <div class="fb-sec-label">NACH KLASSE</div>
      <table class="fr-table">
        <thead><tr><th>Klasse</th><th>Anzahl</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="2" class="muted">Keine Daten</td></tr>'}</tbody>
      </table>
      <div class="fb-actions-row">
        <button type="button" class="fb-btn fb-btn--ghost" id="frBtnCsvInline">CSV exportieren</button>
      </div>
    </section>`;
}

function htmlAntragKlasseField(state, f, studentChrome) {
    const klasseCode = resolveStudentKlasseCode(state);
    const selected = String(klasseCode || f.klasse || '').trim();
    const classes = studentChrome ? classesForStudentPicker(state) : state.stammdaten.classes || [];
    const classOpts = optionList(classes, selected, 'Klasse wählen…');
    const klasseLocked = studentChrome && isStudentKlasseLocked(state);
    const src =
        state.studentMatch && state.studentMatch.klasseSource === 'entra-class-group'
            ? 'Microsoft-365-Klassengruppe'
            : state.studentMatch && state.studentMatch.klasseSource === 'entra-group-label'
              ? 'Microsoft-365-Gruppe (Klasse)'
              : state.studentMatch && state.studentMatch.klasseSource === 'local-pick'
                ? 'Ihre Auswahl (dieses Gerät)'
                : 'Schul-Stammdaten';
    const hintStamm =
        studentChrome && klasseLocked
            ? `<span class="fr-field-hint">Ihre Klasse (${esc(src)}) – nicht änderbar.</span>`
            : '';
    const hintPick =
        studentChrome && !klasseLocked && classes.length
            ? '<span class="fr-field-hint">Wählen Sie Ihre Klasse aus der Schul-Liste.</span>'
            : '';
    const hintEmpty =
        studentChrome && !klasseLocked && !classes.length
            ? '<p class="fr-field-hint fr-field-hint--warn">Noch keine Klassen aus den Stammdaten. Bitte als IT im <strong>Freistellungen-Setup → Gruppen speichern</strong>, danach hier <strong>Aktualisieren</strong>.</p>'
            : '';
    if (studentChrome) {
        if (klasseLocked && selected) {
            return `<label class="fr-field">Klasse *
            <input type="text" id="frFormKlasse" class="fr-input is-locked" required readonly aria-readonly="true" value="${esc(
                selected
            )}">
            ${hintStamm}
          </label>`;
        }
        const opts = classes.length
            ? classOpts
            : optionList([], selected, 'Klassenliste wird geladen…');
        return `<label class="fr-field fr-form-klasse-label">Klasse *
            <select id="frFormKlasse" class="fr-form-klasse-select" required aria-label="Klasse wählen">${opts}</select>
            ${classes.length ? hintPick : hintEmpty}
          </label>`;
    }
    return `<label class="fr-field fr-form-klasse-label">Klasse *
            <select id="frFormKlasse" class="fr-form-klasse-select" required aria-label="Klasse wählen">${classOpts}</select>
          </label>`;
}

/** Kurzer Zeitraum-Hinweis unter Beginn/Ende. */
function antragZeitraumSummary(beginn, ende, path) {
    const from = formatDeDateTime(beginn);
    const to = formatDeDateTime(ende);
    if (!from || from === '–' || !to || to === '–') {
        return 'Datum und Uhrzeit wählen – auch für einzelne Stunden (z. B. 08:00–09:00).';
    }
    const days = path && path.dayCount != null ? path.dayCount : null;
    if (days === 1) return `${from} → ${to} · 1 Tag / Stundenfreistellung`;
    if (days != null) return `${from} → ${to} · ${days} Tage`;
    return `${from} → ${to}`;
}

function htmlAntragPathSteps(path) {
    const steps = Array.isArray(path && path.steps) && path.steps.length ? path.steps : ['Klassenvorstand'];
    return steps
        .map(
            (step, i) =>
                `<li class="fr-path-step"><span class="fr-path-step__n">${i + 1}</span><span class="fr-path-step__l">${esc(
                    step
                )}</span></li>`
        )
        .join('<li class="fr-path-step__arrow" aria-hidden="true"><i class="bi bi-chevron-right"></i></li>');
}

function renderAntrag(state) {
    const studentChrome = useStudentPlanerChrome(state);
    if (studentChrome) prefillStudentFreistellungForm(state);
    const f = state.form || {};
    const path = approvalPath(f.beginn, f.ende);
    const check = validateFreistellung({
        draft: { ...f, _allowedKategorien: kategorieChoicesForState(state) }
    });
    const klasseCode = resolveStudentKlasseCode(state);
    const klasseChosen = !!String(klasseCode || f.klasse || '').trim();
    const nameLocked = studentChrome && !!String(state.accountName || '').trim();
    const kvPrefilled =
        studentChrome && klasseChosen && !!String(f.kvEmail || '').trim().includes('@');
    const klasseField = htmlAntragKlasseField(state, f, studentChrome);
    const kvField = kvPrefilled
        ? `<label class="fr-field fr-field--wide">Klassenvorstand
            <input type="text" class="fr-input is-locked" readonly value="${esc(
                (f.kvName ? f.kvName + ' · ' : '') + (f.kvEmail || '')
            )}">
            <input type="hidden" id="frFormKvEmail" value="${esc(f.kvEmail)}">
            <span class="fr-field-hint"><i class="bi bi-check2-circle" aria-hidden="true"></i> Automatisch aus Stammdaten Ihrer Klasse.</span>
          </label>`
        : studentChrome
          ? `<label class="fr-field fr-field--wide">Klassenvorstand *
            <input type="search" id="frFormKvSearch" placeholder="Lehrperson suchen (Name oder E-Mail)…" autocomplete="off" value="${esc(
                f.kvName && f.kvEmail ? f.kvName + ' · ' + f.kvEmail : f.kvEmail || ''
            )}">
            <input type="hidden" id="frFormKvEmail" value="${esc(f.kvEmail)}">
            <input type="hidden" id="frFormKvName" value="${esc(f.kvName || '')}">
            <ul id="frFormKvSearchHits" class="fr-kv-search-hits" hidden></ul>
            <span class="fr-field-hint">${
                klasseCode
                    ? 'Lehrperson suchen – wird dem Antrag zugeordnet.'
                    : 'Zuerst Klasse wählen, dann Lehrperson.'
            }</span>
          </label>`
          : `<label class="fr-field fr-field--wide">Klassenvorstand (E-Mail) *
            <input type="email" id="frFormKvEmail" required value="${esc(f.kvEmail)}" placeholder="kv@schule.at">
          </label>`;
    const zeitraumHint = antragZeitraumSummary(f.beginn, f.ende, path);
    const formStep = Math.min(4, Math.max(1, Number(state.formStep) || 1));
    const tocStep = (n, label) =>
        `<li>
          <button type="button" class="fr-antrag-toc__btn${formStep === n ? ' is-active' : ''}${
              n < formStep ? ' is-done' : ''
          }" data-fr-goto-step="${n}" aria-current="${formStep === n ? 'step' : 'false'}">
            <span>${n}</span> ${label}
          </button>
        </li>`;

    let stepBody = '';
    if (formStep === 1) {
        stepBody = `
        <fieldset class="fr-antrag-section fr-antrag-section--step">
          <legend><i class="bi bi-person" aria-hidden="true"></i> Person &amp; Klasse</legend>
          <div class="fr-form-grid">
            <label class="fr-field">Ihr Name *
              <input type="text" id="frFormName" required value="${esc(f.schuelerName)}" placeholder="Vor- und Nachname"${
                  nameLocked ? ' readonly class="fr-input is-locked"' : ''
              }>
              ${nameLocked ? '<span class="fr-field-hint">Aus Ihrem Microsoft-365-Konto.</span>' : ''}
            </label>
            ${klasseField}
          </div>
        </fieldset>`;
    } else if (formStep === 2) {
        stepBody = `
        <fieldset class="fr-antrag-section fr-antrag-section--step">
          <legend><i class="bi bi-calendar3" aria-hidden="true"></i> Zeitraum</legend>
          <div class="fr-form-grid">
            <label class="fr-field">Beginn *
              <input type="datetime-local" id="frFormBeginn" required step="300" value="${esc(
                  toIsoDateTimeLocal(f.beginn) || f.beginn || ''
              )}">
            </label>
            <label class="fr-field">Ende *
              <input type="datetime-local" id="frFormEnde" required step="300" value="${esc(
                  toIsoDateTimeLocal(f.ende) || f.ende || ''
              )}">
            </label>
          </div>
          <p class="fr-zeitraum-chip ${path.multiDay ? 'is-multi' : ''}" id="frZeitraumChip" role="status">
            <i class="bi bi-clock-history" aria-hidden="true"></i>
            <span>${esc(zeitraumHint)}</span>
          </p>
        </fieldset>`;
    } else if (formStep === 3) {
        stepBody = `
        <fieldset class="fr-antrag-section fr-antrag-section--step">
          <legend><i class="bi bi-chat-left-text" aria-hidden="true"></i> Grund</legend>
          <div class="fr-form-grid">
            <label class="fr-field">Kategorie *
              <select id="frFormKat" required>
                ${kategorieChoicesForState(state)
                    .map(
                        (k) =>
                            `<option value="${esc(k)}"${f.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
                    )
                    .join('')}
              </select>
            </label>
            <label class="fr-field fr-field--wide">Begründung
              <textarea id="frFormBeschreibung" rows="4" placeholder="Kurz den Grund beschreiben (hilft bei der Freigabe)…">${esc(
                  f.beschreibung
              )}</textarea>
            </label>
          </div>
        </fieldset>`;
    } else {
        const summaryBits = [
            f.schuelerName || '–',
            f.klasse || '',
            antragZeitraumSummary(f.beginn, f.ende, path),
            f.kategorie || ''
        ].filter(Boolean);
        stepBody = `
        <fieldset class="fr-antrag-section fr-antrag-section--step">
          <legend><i class="bi bi-person-badge" aria-hidden="true"></i> Freigabe &amp; Absenden</legend>
          <p class="fr-antrag-summary muted" role="status">${esc(summaryBits.join(' · '))}</p>
          <div class="fr-form-grid">
            ${kvField}
          </div>
        </fieldset>
        <fieldset class="fr-antrag-section fr-antrag-section--step fr-antrag-section--upload">
          <legend><i class="bi bi-paperclip" aria-hidden="true"></i> Nachweise <span class="fr-legend-opt">(optional)</span></legend>
          ${htmlFreistellungUploadField({ student: studentChrome })}
        </fieldset>
        <aside class="fr-antrag-footer ${path.multiDay ? 'is-multi' : ''}">
          <div class="fr-antrag-footer__path">
            <p class="fr-antrag-footer__kicker">Genehmigungspfad</p>
            <ul class="fr-path-steps">${htmlAntragPathSteps(path)}</ul>
            <p class="fr-antrag-footer__meta">${esc(path.label)}${
                path.dayCount != null
                    ? ` · ${path.dayCount} Tag${path.dayCount === 1 ? '' : 'e'}`
                    : ''
            }</p>
            ${
                path.multiDay
                    ? '<p class="fr-antrag-footer__note">Mehrtägig: zuerst Klassenvorstand, danach Direktion.</p>'
                    : '<p class="fr-antrag-footer__note">Kurzfristig / stundenweise: Freigabe durch den Klassenvorstand.</p>'
            }
          </div>
          ${
              check.warnings.length
                  ? `<ul class="fr-hints">${check.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul>`
                  : ''
          }
        </aside>`;
    }

    return `
    <section class="fr-panel fr-panel--antrag fr-panel--calm">
      <h1 class="fb-greet">Antrag stellen</h1>
      <p class="fb-sub">${
          studentChrome
              ? 'In wenigen Schritten absenden. Die Genehmigung läuft danach automatisch über Microsoft Approvals.'
              : 'Schritt für Schritt – Speichern schreibt in SharePoint und startet den Genehmigungsflow.'
      }</p>
      <form id="frAntragForm" class="fr-form fr-form--antrag fr-form--wizard" autocomplete="on" data-fr-step="${formStep}" novalidate>
        <ol class="fr-antrag-toc" aria-label="Antragsschritte">
          ${tocStep(1, 'Person')}
          ${tocStep(2, 'Zeitraum')}
          ${tocStep(3, 'Grund')}
          ${tocStep(4, 'Freigabe')}
        </ol>
        ${stepBody}
        <div class="fr-wizard-actions">
          ${
              formStep > 1
                  ? '<button type="button" class="fb-btn fb-btn--ghost" data-fr-step-prev>← Zurück</button>'
                  : '<span></span>'
          }
          <div class="fr-wizard-actions__primary">
            ${
                formStep < 4
                    ? '<button type="button" class="fb-btn" data-fr-step-next>Weiter →</button>'
                    : `<button type="submit" class="fb-btn" id="frBtnSubmit">Antrag einreichen</button>
                       <button type="button" class="fb-btn fb-btn--ghost" id="frBtnFormReset">Zurücksetzen</button>`
            }
          </div>
        </div>
      </form>
    </section>`;
}

function renderTable(items, _state, opts) {
    const o = opts || {};
    const rows = items
        .map((it) => {
            const days = inclusiveDayCount(it.beginn, it.ende);
            const decide =
                o.showDecide && String(it.status) === 'Ausstehend'
                    ? `<td class="fr-td-actions">
                       <button type="button" class="fb-btn" data-fr-approve="${esc(it.itemId)}" title="Status Genehmigt (Fallback)">Genehmigen</button>
                       <button type="button" class="fb-btn fb-btn--ghost" data-fr-reject="${esc(it.itemId)}" title="Status Abgelehnt (Fallback)">Ablehnen</button>
                       </td>`
                    : '';
            const id = it.itemId || it.antragId;
            return `<tr class="fr-table__row" tabindex="0" role="link" data-fr-detail="${esc(
                id
            )}" aria-label="${esc((it.schuelerName || it.titel || 'Antrag') + ' öffnen')}">
          <td><strong>${esc(it.schuelerName || it.titel)}</strong></td>
          <td>${esc(it.klasse)}</td>
          <td>${esc(formatDeDateTime(it.beginn))}${
                it.ende && it.ende !== it.beginn ? ' – ' + esc(formatDeDateTime(it.ende)) : ''
            }</td>
          <td>${days != null ? days : '–'}${it.multiDay ? ' <span class="fr-badge fr-badge--info">mehr</span>' : ''}</td>
          <td>${statusBadge(it.status)}</td>
          <td>${esc(it.kategorie)}</td>
          <td>${esc(it.kvName || it.kvEmail || '–')}</td>
          ${decide}
          <td class="fr-table__chevron"><i class="bi bi-chevron-right" aria-hidden="true"></i></td>
        </tr>`;
        })
        .join('');
    return `
    <div class="fr-table-wrap">
      <table class="fr-table fr-table--rows">
        <thead>
          <tr>
            <th>Schüler/in</th><th>Klasse</th><th>Zeitraum</th><th>Tage</th>
            <th>Status</th><th>Kategorie</th><th>KV</th>${o.showDecide ? '<th></th>' : ''}
            <th class="fr-table__chevron"><span class="fr-sr-only">Öffnen</span></th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

export function readFormFromDom(root) {
    const host = root || document;
    const $ = (id) => host.querySelector('#' + id);
    /** Nur vorhandene Felder – Wizard-Schritte dürfen state.form nicht leeren. */
    const out = { status: 'Ausstehend' };
    if ($('frFormName')) out.schuelerName = String($('frFormName').value || '').trim();
    if ($('frFormKlasse')) out.klasse = String($('frFormKlasse').value || '').trim();
    if ($('frFormBeginn')) out.beginn = String($('frFormBeginn').value || '').trim();
    if ($('frFormEnde')) out.ende = String($('frFormEnde').value || '').trim();
    if ($('frFormKat')) out.kategorie = String($('frFormKat').value || '').trim();
    if ($('frFormKvEmail')) {
        out.kvEmail = String($('frFormKvEmail').value || '')
            .trim()
            .toLowerCase();
    }
    if ($('frFormKvName')) out.kvName = String($('frFormKvName').value || '').trim();
    if ($('frFormBeschreibung')) {
        out.beschreibung = String($('frFormBeschreibung').value || '').trim();
    }
    return out;
}

/**
 * Schüler: Klassenvorstand per Microsoft-365-Suche wählen.
 * @param {HTMLElement} root
 */
export function wireFreistellungStudentKvSearch(root) {
    if (!root) return;
    const search = root.querySelector('#frFormKvSearch');
    const hits = root.querySelector('#frFormKvSearchHits');
    const emailEl = root.querySelector('#frFormKvEmail');
    const nameEl = root.querySelector('#frFormKvName');
    if (!search || !hits || !emailEl) return;
    if (search.dataset.frKvWired === '1') return;
    search.dataset.frKvWired = '1';

    let timer = null;
    let inFlight = 0;

    function pickUser(u) {
        const mail = String(u.mail || u.userPrincipalName || '')
            .trim()
            .toLowerCase();
        const name = String(u.displayName || '').trim();
        if (!mail.includes('@')) return;
        emailEl.value = mail;
        if (nameEl) nameEl.value = name;
        search.value = name ? name + ' · ' + mail : mail;
        hits.hidden = true;
        hits.innerHTML = '';
    }

    search.addEventListener('input', function () {
        clearTimeout(timer);
        const q = String(search.value || '').trim();
        if (q.length < 2) {
            hits.hidden = true;
            hits.innerHTML = '';
            return;
        }
        timer = setTimeout(async function () {
            const G = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
            if (!G || typeof G.searchUsers !== 'function' || typeof G.getGraphToken !== 'function') {
                return;
            }
            const flight = ++inFlight;
            try {
                const scopes = ['https://graph.microsoft.com/User.Read', 'https://graph.microsoft.com/User.ReadBasic.All'];
                let tok = null;
                if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
                    try {
                        tok = await window.ms365AuthAcquireTokenSilent(scopes);
                    } catch {
                        tok = null;
                    }
                }
                if (!tok) {
                    hits.innerHTML =
                        '<li class="muted">Lehrersuche braucht eine IT-Freigabe in Microsoft Entra – Klassenvorstand-E-Mail unten eintragen.</li>';
                    hits.hidden = false;
                    return;
                }
                const users = await G.searchUsers(tok, q);
                if (flight !== inFlight) return;
                if (!users.length) {
                    hits.innerHTML = '<li class="muted">Keine Treffer</li>';
                    hits.hidden = false;
                    return;
                }
                hits.innerHTML = users
                    .slice(0, 12)
                    .map(function (u) {
                        const mail = String(u.mail || u.userPrincipalName || '').trim();
                        const name = String(u.displayName || mail).trim();
                        return (
                            '<li><button type="button" class="fr-kv-search-hit" data-email="' +
                            esc(mail) +
                            '" data-name="' +
                            esc(name) +
                            '">' +
                            esc(name) +
                            (mail ? ' <span class="muted">' + esc(mail) + '</span>' : '') +
                            '</button></li>'
                        );
                    })
                    .join('');
                hits.hidden = false;
            } catch {
                if (flight === inFlight) {
                    hits.innerHTML = '<li class="muted">Suche nicht verfügbar</li>';
                    hits.hidden = false;
                }
            }
        }, 320);
    });

    hits.addEventListener('click', function (ev) {
        const btn = ev.target.closest('.fr-kv-search-hit');
        if (!btn) return;
        ev.preventDefault();
        pickUser({
            mail: btn.getAttribute('data-email'),
            displayName: btn.getAttribute('data-name')
        });
    });
}

function readMultiFilterFromDom(root, key, inputName) {
    const host = root || document;
    const multi = host.querySelector(`[data-fr-filter-multi="${key}"]`);
    if (!multi) return null;
    const checked = Array.from(
        multi.querySelectorAll(`input[type="checkbox"][name="${inputName}"]:checked`)
    )
        .map((el) => String(el.value || '').trim())
        .filter(Boolean);
    return joinCodeList(checked);
}

export function readFiltersFromDom(root) {
    const host = root || document;
    const val = (id) => {
        const el = host.querySelector('#' + id);
        return el ? String(el.value || '').trim() : '';
    };
    return {
        q: val('frFilterQ'),
        klasse: readMultiFilterFromDom(host, 'klasse', 'frFilterKlasse') ?? val('frFilterKlasse'),
        status: readMultiFilterFromDom(host, 'status', 'frFilterStatus') ?? val('frFilterStatus'),
        kategorie:
            readMultiFilterFromDom(host, 'kategorie', 'frFilterKategorie') ?? val('frFilterKat'),
        multiDay: val('frFilterMulti')
    };
}

export { applyKvFromClass } from './freistellung-planer-state.js';
