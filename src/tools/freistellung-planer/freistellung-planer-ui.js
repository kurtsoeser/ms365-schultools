/**
 * UI-Rendering Freistellungs-Planer.
 */
import {
    computeDashboardKpis,
    validateFreistellung,
    approvalPath,
    inclusiveDayCount,
    toIsoDateTimeLocal,
    STATUS_CHOICES,
    monthGridDates,
    itemCoversDay
} from './freistellung-planer-logic.js';
import { htmlFreistellungKategorienPanel } from './freistellung-kategorien-ui.js';
import { htmlFreistellungKlassenPanel } from './freistellung-klassen-ui.js';
import { NACHWEISE_MAX_FILES, NACHWEISE_MAX_BYTES, formatNachweisSize } from './freistellung-planer-nachweise.js';
import { renderFreistellungAccessDebugEntry } from './freistellung-planer-access-debug-ui.js';
import {
    viewsForRole,
    useStudentPlanerChrome,
    useKvPlanerChrome,
    useDirektionPlanerChrome,
    useMinimalPlanerChrome,
    filterItems,
    formatDeDate,
    formatDeDateTime,
    statusLabel,
    roleLabel,
    scopeFromState,
    canDecide,
    resolveStudentKlasseCode,
    isStudentKlasseLocked,
    findClassInStammdaten,
    classesForStudentPicker,
    studentKlasseFromRecord,
    prefillStudentFreistellungForm,
    applyKvFromClass,
    itemVisibleForRole
} from './freistellung-planer-state.js';
import {
    isPlanerDemoRoleUiEnabled,
    sortPlanerRoles,
    roleSourceLabel,
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

function renderNavSession(state) {
    const studentChrome = useStudentPlanerChrome(state);
    const demoRoleUi = isPlanerDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    const switchable = demoRoleUi ? PLANER_ROLE_ORDER : sortPlanerRoles(state.planerRoles || []);
    const showRoleSwitcher = demoRoleUi || switchable.length > 1;
    const activeSource = roleSourceLabel(
        (state.planerRoleSources && state.planerRoleSources[state.role]) || state.roleSource
    );
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
        : switchable.length === 1
          ? `<p class="fr-nav__role-fixed"><span class="fr-badge fr-badge--info">${esc(roleLabel(state.role))}</span>${
                activeSource ? ` <small class="fr-nav__account-src">(${esc(activeSource)})</small>` : ''
            }</p>`
          : state.accountEmail && !studentChrome
            ? `<p class="fr-nav__hint">Keine Planer-Rolle zugewiesen.</p>`
            : studentChrome && state.accountEmail
              ? ''
              : '';

    const kvChrome = useKvPlanerChrome(state);
    if (studentChrome || kvChrome) {
        return `
        <div class="fr-nav__session" aria-label="Anmeldung">
          <span class="fr-nav__role-label">Anmeldung</span>
          ${
              state.accountEmail
                  ? `<p class="fr-nav__account-name">${esc(state.accountName || state.accountEmail)}</p>`
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
              kvChrome && switchable.length > 1
                  ? `<div class="fr-nav__session-role" style="margin-top:10px;">
              <span class="fr-nav__session-role-label">Rolle</span>
              ${roleBlock}
            </div>`
                  : kvChrome
                    ? `<p class="fr-nav__role-fixed" style="margin-top:8px;"><span class="fr-badge fr-badge--info">${esc(
                          roleLabel('kv')
                      )}</span></p>`
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
              state.accountEmail
                  ? `<p class="fr-nav__account-name">${esc(state.accountName || state.accountEmail)}</p>`
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
                      ? '<p class="fr-nav__hint">Demo-Umschalter (ohne Entra-Gruppen oder ?demoRole=1).</p>'
                      : '<p class="fr-nav__hint">Genehmigung läuft über Microsoft Approvals (Power Automate).</p>'
          }
        </div>`;
}

function showPlanerDevTools(state) {
    const demoRoleUi = isPlanerDemoRoleUiEnabled(state.entraGroupsConfigured, state.demoRoleOverride);
    if (demoRoleUi) return true;
    return state.role === 'direktion';
}

function renderKvTop(state) {
    const denied = state.planerAccessDenied;
    const hint = state.roleHintPublic || state.roleHintStaff || state.roleHint || '';
    const kvClasses = (state.stammdaten.classes || []).filter((c) => {
        const em = String(c.headEmail || c.klassenvorstandEmail || c.kvEmail || '')
            .trim()
            .toLowerCase();
        return em && em === String(state.accountEmail || '').toLowerCase();
    });
    const classLabels = kvClasses.map((c) => c.name || c.code).filter(Boolean);
    return `
        <header class="fr-top fr-top--student">
          <div>
            <h2 class="fr-top__student-title">Freistellungen meiner Klasse</h2>
            <p class="fr-nav__hint">Anträge Ihrer Klasse – Status und Übersicht.</p>
            ${
                classLabels.length
                    ? `<p class="fr-nav__hint">Klassen: <span class="fr-badge fr-nav__account-klasse">${esc(
                          classLabels.join(', ')
                      )}</span></p>`
                    : ''
            }
          </div>
          ${
              state.accountEmail
                  ? `<button type="button" class="btn" id="frBtnLoad" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>`
                  : ''
          }
        </header>
        ${
            denied
                ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(
                      hint || 'Derzeit kein Zugriff auf die Freistellungsanträge Ihrer Klasse.'
                  )}</div>`
                : ''
        }`;
}

function renderStudentTop(state) {
    const denied = state.planerAccessDenied;
    const hint = state.roleHintPublic || '';
    const onAntrag = state.view === 'antrag';
    const title = onAntrag ? 'Neuer Antrag' : 'Meine Anträge';
    const lead = onAntrag
        ? 'Antrag absenden – Genehmigung über Microsoft Approvals.'
        : 'Eingereichte Freistellungen und Status.';
    return `
        <header class="fr-top fr-top--student">
          <div>
            <h2 class="fr-top__student-title">${esc(title)}</h2>
            <p class="fr-nav__hint">${esc(lead)}</p>
          </div>
          <div class="fr-top__student-actions" style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;">
          ${
              state.accountEmail && !denied && !onAntrag
                  ? `<button type="button" class="btn btn-success" data-fr-view-jump="antrag"><i class="bi bi-plus-lg"></i>Antrag stellen</button>`
                  : ''
          }
          ${
              state.accountEmail
                  ? `<button type="button" class="btn" id="frBtnLoad" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>`
                  : ''
          }
          </div>
        </header>
        ${
            denied
                ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(
                      hint || 'Derzeit kein Zugriff auf Ihre Freistellungsanträge.'
                  )}</div>`
                : ''
        }`;
}

function htmlAdminSiteControls(state, opts) {
    const o = opts || {};
    const devTools = showPlanerDevTools(state);
    const setupLinks = o.includeSetupLinks
        ? `<a class="btn btn-sm" href="freistellung-konzept.html" style="text-decoration:none" title="Ablauf und Genehmigungslogik"><i class="bi bi-diagram-3"></i>Ablauf</a>
              <a class="btn btn-sm" href="freistellung-setup.html" style="text-decoration:none"><i class="bi bi-gear"></i>Setup</a>`
        : '';
    const devBlock = devTools
        ? `<button type="button" class="btn btn-sm" id="frBtnDemo" title="SJ 2026/27 Demo laden"><i class="bi bi-database"></i>Demo</button>
              <button type="button" class="btn btn-sm" id="frBtnJsonImport" title="JSON-Testpaket"><i class="bi bi-filetype-json"></i>JSON</button>
              <input type="file" id="frJsonImportFile" accept="application/json,.json" hidden>
              <button type="button" class="btn btn-sm alt" id="frBtnDemoReset" title="Demo zurücksetzen"><i class="bi bi-trash"></i>Reset</button>`
        : '';
    return `
            <div class="fr-site-controls">
              <label class="fr-site-controls__label" for="frSiteUrl">SharePoint-Site</label>
              <input type="url" class="fr-site-controls__input" id="frSiteUrl" value="${esc(state.siteUrl)}" placeholder="https://ihre-schule.sharepoint.com/sites/Administration" spellcheck="false" autocomplete="off">
              <div class="fr-site-controls__actions">
                <button type="button" class="btn btn-sm btn-success" id="frBtnLoad"><i class="bi bi-arrow-repeat"></i>Laden</button>
                ${setupLinks}
                <button type="button" class="btn btn-sm" id="frBtnCsv"><i class="bi bi-download"></i>CSV</button>
                ${devBlock}
              </div>
            </div>`;
}

function renderAdminSitePanel(state) {
    const roleBadge =
        state.role === 'kv' || state.role === 'direktion' || state.role === 'schueler'
            ? roleLabel(state.role)
            : '–';
    return `
        <section class="fr-panel">
          <h2>SharePoint-Verbindung</h2>
          <p class="muted" style="margin:0 0 10px;font-size:0.9em;">Liste „${esc(state.listName || 'Freistellungen')}“ auf der gewählten Site.</p>
          ${htmlAdminSiteControls(state, { includeSetupLinks: false })}
          <p class="muted" style="margin:10px 0 0;font-size:0.88em;">
            Rolle <span class="fr-badge fr-badge--info">${esc(roleBadge)}</span>
            ${state.localDemoOnly ? ' <span class="fr-badge fr-badge--warn">Demo lokal</span>' : ''}
            ${
                state.roleHintStaff || state.roleHint
                    ? '<br><span>' + esc(state.roleHintStaff || state.roleHint) + '</span>'
                    : ''
            }
          </p>
        </section>`;
}

function renderAdministration(state) {
    if (state.role !== 'direktion') {
        return '<p class="muted">Nur für Direktion / Admin.</p>';
    }
    return `
    <section class="fr-panel fr-panel--compact">
      <h2>Administration</h2>
      <p class="muted fr-panel__lead">Einrichtung, Dokumentation und Technik – nicht für den täglichen Genehmigungsablauf.</p>
      <nav class="fr-admin-nav" aria-label="Weitere Seiten">
        <a class="fr-admin-nav__item" href="freistellung-konzept.html"><i class="bi bi-diagram-3" aria-hidden="true"></i><span>Ablauf</span></a>
        <a class="fr-admin-nav__item" href="freistellung-setup.html"><i class="bi bi-wrench" aria-hidden="true"></i><span>Setup</span></a>
        <a class="fr-admin-nav__item" href="../tenant.html"><i class="bi bi-gear" aria-hidden="true"></i><span>Stammdaten</span></a>
        <a class="fr-admin-nav__item" href="../hilfe.html#tool-freistellung-planer"><i class="bi bi-question-circle" aria-hidden="true"></i><span>Hilfe</span></a>
        <a class="fr-admin-nav__item" href="../index.html"><i class="bi bi-arrow-left" aria-hidden="true"></i><span>Dashboard</span></a>
      </nav>
    </section>
    ${renderAdminSitePanel(state)}
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
    })}`;
}

function renderDirektionTop(state) {
    const denied = state.planerAccessDenied;
    const hint = state.roleHintStaff || state.roleHint || '';
    const onAdmin = state.view === 'administration';
    return `
        <header class="fr-top fr-top--student">
          <div>
            <h2 class="fr-top__student-title">${onAdmin ? 'Administration' : 'Freistellungen – Verwaltung'}</h2>
            <p class="fr-nav__hint">${
                onAdmin
                    ? 'SharePoint, Klassen aus Stammdaten, Kategorien und Links zur Einrichtung.'
                    : 'Übersicht, Anträge und Genehmigungen. Technik: Menü „Administration“.'
            }</p>
          </div>
          ${
              state.accountEmail && !onAdmin
                  ? `<button type="button" class="btn" id="frBtnLoad" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>`
                  : ''
          }
        </header>
        ${
            denied
                ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(
                      hint || 'Kein Zugriff auf den Freistellungs-Planer für dieses Konto.'
                  )}</div>`
                : ''
        }`;
}

function renderStaffTop(state) {
    const roleBadge =
        state.role === 'kv' || state.role === 'direktion' || state.role === 'schueler'
            ? roleLabel(state.role)
            : '–';
    return `
        <header class="fr-top">
          <div class="fr-top__site">
            ${htmlAdminSiteControls(state, { includeSetupLinks: true })}
          </div>
          <div class="fr-top__user">
            <span>Rolle <span class="fr-badge fr-badge--info">${esc(roleBadge)}</span></span>
            ${state.localDemoOnly ? '<span class="fr-badge fr-badge--warn">Demo lokal</span>' : ''}
            ${
                state.roleHintStaff || state.roleHint
                    ? `<small class="muted" style="display:block;max-width:22rem;margin-top:4px;">${esc(
                          state.roleHintStaff || state.roleHint
                      )}</small>`
                    : ''
            }
          </div>
        </header>
        ${
            state.planerAccessDenied
                ? `<div class="fr-alert fr-alert--bad" role="alert">Kein Zugriff auf den Freistellungs-Planer für dieses Konto. ${esc(
                      state.roleHintStaff || state.roleHint || ''
                  )}</div>`
                : ''
        }`;
}

export function renderApp(state, root) {
    if (!root) return;
    const views = viewsForRole(state.role, state);
    const studentChrome = useStudentPlanerChrome(state);
    const kvChrome = useKvPlanerChrome(state);
    const direktionChrome = useDirektionPlanerChrome(state);
    const minimalChrome = studentChrome || kvChrome || direktionChrome;
    root.innerHTML = `
    <div class="fr-shell${minimalChrome ? ' fr-shell--student' : ''}">
      <aside class="fr-nav" aria-label="Navigation">
        <div class="fr-nav__brand">
          <i class="bi bi-calendar2-check" aria-hidden="true"></i>
          <div>
            <strong>Freistellungen</strong>
            <span>${
                studentChrome
                    ? 'Schüler'
                    : kvChrome
                      ? 'Klassenübersicht'
                      : direktionChrome
                        ? 'Verwaltung'
                        : 'Antrag &amp; Genehmigung'
            }</span>
          </div>
        </div>
        <nav class="fr-nav__list">
          ${views
              .map(
                  (v) => `
            <button type="button" class="fr-nav__btn${state.view === v.id ? ' is-active' : ''}" data-fr-view="${esc(v.id)}">
              <i class="bi ${v.icon}" aria-hidden="true"></i>${esc(v.label)}
            </button>`
              )
              .join('')}
        </nav>
        ${renderNavSession(state)}
      </aside>
      <div class="fr-main">
        ${
            studentChrome
                ? renderStudentTop(state)
                : kvChrome
                  ? renderKvTop(state)
                  : direktionChrome
                    ? renderDirektionTop(state)
                    : renderStaffTop(state)
        }
        ${renderFreistellungAccessDebugEntry(state)}
        ${
            !minimalChrome && state.localDemoOnly
                ? '<div class="fr-alert fr-alert--info" role="status">Lokaler Demo-Modus – Anzeige ohne SharePoint. Mit „Laden“ echte Liste verbinden.</div>'
                : ''
        }
        ${
            !minimalChrome && state.info
                ? `<div class="fr-alert fr-alert--info" role="status">${esc(state.info)}</div>`
                : ''
        }
        ${state.error ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(state.error)}</div>` : ''}
        ${state.loading ? '<div class="fr-skel" aria-busy="true"><div class="fr-skel__card"></div><div class="fr-skel__card"></div></div>' : ''}
        <div id="frView" class="fr-view${state.loading ? ' is-loading' : ''}">
          ${state.loading || state.planerAccessDenied ? '' : renderView(state)}
        </div>
      </div>
    </div>
    ${state.detailId ? renderDetailModal(state) : ''}
  `;
}

export function htmlFreistellungUploadField(opts) {
    const student = !!(opts && opts.student);
    const maxMb = Math.round(NACHWEISE_MAX_BYTES / 1024 / 1024);
    return `
    <div class="fr-form-span2 fr-upload" id="frUploadWrap">
      <span class="fr-upload__title">Anhänge / Nachweise (optional)</span>
      <p class="muted fr-upload__hint">${
          student
              ? 'z. B. ärztliche Bestätigung oder Einladung – werden mit dem Antrag gespeichert.'
              : 'PDF, Fotos oder Word – max. ' +
                NACHWEISE_MAX_FILES +
                ' Dateien, je ' +
                maxMb +
                ' MB.'
      }</p>
      <div class="fr-upload__actions">
        <label class="btn btn-sm" for="frFormFiles"><i class="bi bi-paperclip"></i>Dateien auswählen</label>
        <input type="file" id="frFormFiles" class="fr-upload__input" multiple accept=".pdf,.jpg,.jpeg,.png,.heic,.doc,.docx,.webp,image/*,application/pdf">
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

function filterBar(state) {
    const sd = state.stammdaten;
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const count = filterItems(state.items, state.filters, scope).length;
    return `
    <div class="fr-filters" data-fr-filters>
      <input type="search" id="frFilterQ" placeholder="Suche…" value="${esc(state.filters.q || '')}" aria-label="Suche">
      <select id="frFilterKlasse" aria-label="Klasse">${optionList(sd.classes, state.filters.klasse, 'Alle Klassen')}</select>
      <select id="frFilterStatus" aria-label="Status">
        <option value="">Alle Status</option>
        ${STATUS_CHOICES.map(
            (s) =>
                `<option value="${esc(s)}"${state.filters.status === s ? ' selected' : ''}>${esc(s)}</option>`
        ).join('')}
      </select>
      <select id="frFilterKat" aria-label="Kategorie">
        <option value="">Alle Kategorien</option>
        ${kategorieChoicesForState(state)
            .map(
                (k) =>
                    `<option value="${esc(k)}"${state.filters.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
            )
            .join('')}
      </select>
      <select id="frFilterMulti" aria-label="Dauer">
        <option value="">Alle Dauern</option>
        <option value="0"${state.filters.multiDay === '0' ? ' selected' : ''}>1 Tag</option>
        <option value="1"${state.filters.multiDay === '1' ? ' selected' : ''}>Mehrtägig</option>
      </select>
      <button type="button" class="btn" id="frFilterReset"><i class="bi bi-arrow-counterclockwise"></i>Reset</button>
      <span class="fr-filters__meta">${count} Einträge</span>
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

function renderDashboard(state) {
    const kvChrome = useKvPlanerChrome(state);
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, {}, scope);
    const kpi = computeDashboardKpis(items);
    const offen = items.filter((i) => String(i.status) === 'Ausstehend');
    if (kvChrome) {
        const sorted = [...items].sort((a, b) => String(b.beginn || '').localeCompare(String(a.beginn || '')));
        return `
    <section class="fr-panel">
      <h2>Übersicht</h2>
      <div class="fr-kpi-row">
        <div class="fr-kpi"><strong>${kpi.ausstehend}</strong><span>Ausstehend</span></div>
        <div class="fr-kpi"><strong>${kpi.genehmigt}</strong><span>Genehmigt</span></div>
        <div class="fr-kpi"><strong>${kpi.mehrtage}</strong><span>Mehrtägig</span></div>
        <div class="fr-kpi"><strong>${kpi.demnaechst}</strong><span>In 14 Tagen</span></div>
      </div>
    </section>
    <section class="fr-panel">
      <h3>Anträge Ihrer Schülerinnen und Schüler</h3>
      ${
          sorted.length
              ? renderTable(sorted, state)
              : (state.items || []).length
                ? '<p class="muted">In der Liste liegen Anträge, aber keiner passt zu Ihrer Klasse/KV-Zuordnung (Filter). IT: Stammdaten <code>headEmail</code> für Ihre Klasse(n); in Anträgen Feld <strong>Klassenvorstand</strong> pflegen. In SharePoint sehen KVs mit „Gestaltung“ oft <em>alle</em> Einträge – im Planer nur die eigenen Klassen.</p>'
                : '<p class="muted">Aktuell keine Anträge für Ihre Klasse.</p>'
      }
      ${
          offen.length
              ? '<p class="muted" style="margin-top:10px;">Offene Genehmigungen bearbeiten Sie in Microsoft Approvals (Teams oder Outlook).</p>'
              : ''
      }
      <div class="fr-actions" style="margin-top:12px">
        <button type="button" class="btn" data-fr-view-jump="kalender"><i class="bi bi-calendar3"></i>Kalender</button>
      </div>
    </section>`;
    }
    return `
    <section class="fr-panel">
      <h2>Übersicht</h2>
      <p class="muted">Anträge landen in der SharePoint-Liste; der Power-Automate-Flow startet Microsoft Approvals (KV, bei Mehrtagen sequentiell Direktion).</p>
      <div class="fr-kpi-row">
        <div class="fr-kpi"><strong>${kpi.ausstehend}</strong><span>Ausstehend</span></div>
        <div class="fr-kpi"><strong>${kpi.genehmigt}</strong><span>Genehmigt</span></div>
        <div class="fr-kpi"><strong>${kpi.mehrtage}</strong><span>Mehrtägig</span></div>
        <div class="fr-kpi"><strong>${kpi.demnaechst}</strong><span>In 14 Tagen</span></div>
      </div>
    </section>
    <section class="fr-panel">
      <h3>Offene Anträge</h3>
      ${offen.length ? renderTable(offen, state) : '<p class="muted">Keine offenen Anträge.</p>'}
      <div class="fr-actions" style="margin-top:12px">
        <button type="button" class="btn btn-success" data-fr-view-jump="antrag"><i class="bi bi-plus-lg"></i>Neuer Antrag</button>
        <button type="button" class="btn" data-fr-view-jump="kalender"><i class="bi bi-calendar3"></i>Kalender</button>
        ${canDecide(state) ? '<button type="button" class="btn" data-fr-view-jump="freigabe"><i class="bi bi-check2-square"></i>Offene Genehmigungen</button>' : ''}
      </div>
    </section>`;
}

function renderListe(state) {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope);
    return `
    <section class="fr-panel">
      <h2>Liste</h2>
      ${filterBar(state)}
      ${items.length ? renderTable(items, state) : '<p class="muted">Keine Einträge.</p>'}
    </section>`;
}

function calendarChipLabel(it) {
    const klasse = String(it.klasse || '').trim();
    const name = String(it.schuelerName || it.titel || '').trim();
    if (klasse && name) return klasse + ' · ' + name;
    return name || klasse || 'Antrag';
}

function calendarStatusClass(status) {
    const s = String(status || '').toLowerCase();
    if (s === 'genehmigt') return 'fr-cal__chip--ok';
    if (s === 'abgelehnt') return 'fr-cal__chip--bad';
    return 'fr-cal__chip--warn';
}

function renderKalender(state) {
    const y = state.calYear;
    const m = state.calMonth;
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope).filter(
        (it) => String(it.status || '').toLowerCase() !== 'abgelehnt'
    );
    const days = monthGridDates(y, m);
    const monthName = new Date(y, m - 1, 1).toLocaleDateString('de-AT', { month: 'long', year: 'numeric' });
    const kvChrome = useKvPlanerChrome(state);
    const cells = days
        .map((day) => {
            const inMonth = Number(day.slice(5, 7)) === m;
            const dayItems = items.filter((it) => itemCoversDay(it, day));
            return `<div class="fr-cal__day${!inMonth ? ' is-out' : ''}">
          <div class="fr-cal__num">${Number(day.slice(8))}</div>
          ${dayItems
              .slice(0, 4)
              .map(
                  (it) =>
                      `<button type="button" class="fr-cal__chip ${calendarStatusClass(it.status)}" data-fr-detail="${esc(
                          it.itemId
                      )}" title="${esc(calendarChipLabel(it))} – ${esc(statusLabel(it.status))}">${esc(
                          calendarChipLabel(it)
                      )}</button>`
              )
              .join('')}
          ${dayItems.length > 4 ? `<span class="fr-cal__more muted">+${dayItems.length - 4}</span>` : ''}
        </div>`;
        })
        .join('');
    return `
    <section class="fr-panel">
      <div class="fr-cal__head">
        <div>
          <h2>Kalender</h2>
          <p class="muted" style="margin:4px 0 0;line-height:1.45;">
            ${
                kvChrome
                    ? 'Abwesenheiten der Schülerinnen und Schüler Ihrer Klasse(n) – nach Genehmigung und offene Anträge.'
                    : 'Schulweite Übersicht: wer wann freigestellt ist (ohne abgelehnte Anträge).'
            }
          </p>
        </div>
        <div class="fr-cal__nav">
          <button type="button" class="btn" id="frCalPrev" aria-label="Vorheriger Monat"><i class="bi bi-chevron-left"></i></button>
          <strong>${esc(monthName)}</strong>
          <button type="button" class="btn" id="frCalNext" aria-label="Nächster Monat"><i class="bi bi-chevron-right"></i></button>
        </div>
      </div>
      ${filterBar(state)}
      <div class="fr-cal__legend muted">
        <span><span class="fr-cal__dot fr-cal__dot--warn"></span> Ausstehend</span>
        <span><span class="fr-cal__dot fr-cal__dot--ok"></span> Genehmigt</span>
      </div>
      <div class="fr-cal__weekdays" aria-hidden="true"><span>Mo</span><span>Di</span><span>Mi</span><span>Do</span><span>Fr</span><span>Sa</span><span>So</span></div>
      <div class="fr-cal__grid">${cells}</div>
    </section>`;
}

function renderMeine(state) {
    const studentChrome = useStudentPlanerChrome(state);
    const items = filterItems(state.items, state.filters, {
        onlyMine: true,
        accountEmail: state.accountEmail
    });
    return `
    <section class="fr-panel${studentChrome ? ' fr-panel--student-meine' : ''}">
      ${studentChrome ? '' : '<h2>Meine Anträge</h2>'}
      ${studentChrome ? '' : filterBar(state)}
      ${
          items.length
              ? renderTable(items, state)
              : `<p class="muted">${
                    studentChrome
                        ? 'Noch keine Anträge. Stellen Sie Ihren ersten Antrag über die Navigation oder den Button unten.'
                        : 'Noch keine eigenen Anträge.'
                }</p>`
      }
      ${
          studentChrome && !items.length
              ? `<div class="fr-actions" style="margin-top:12px">
        <button type="button" class="btn btn-success" data-fr-view-jump="antrag"><i class="bi bi-plus-lg"></i>Antrag stellen</button>
      </div>`
              : ''
      }
    </section>`;
}

function renderFreigabe(state) {
    const scope =
        state.role === 'direktion'
            ? { accountEmail: state.accountEmail }
            : { onlyKv: true, accountEmail: state.accountEmail };
    const items = filterItems(state.items, { status: 'Ausstehend' }, scope);
    return `
    <section class="fr-panel">
      <h2>Offene Genehmigungen</h2>
      <p class="muted">Freigabe in <strong>Microsoft Approvals</strong> (Teams/Handy). Nach Flow v2 stehen Genehmiger und Datum auch in der Liste (Detailansicht, CSV). Manuelles ✓/✗ nur Fallback ohne Flow.</p>
      ${items.length ? renderTable(items, state, { showDecide: true }) : '<p class="muted">Keine ausstehenden Anträge.</p>'}
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
      <h2>Berichte</h2>
      ${filterBar(state)}
      <div class="fr-kpi-row">
        <div class="fr-kpi"><strong>${kpi.gesamt}</strong><span>Gesamt</span></div>
        <div class="fr-kpi"><strong>${kpi.genehmigt}</strong><span>Genehmigt</span></div>
        <div class="fr-kpi"><strong>${kpi.abgelehnt}</strong><span>Abgelehnt</span></div>
        <div class="fr-kpi"><strong>${kpi.mehrtage}</strong><span>Mehrtägig</span></div>
      </div>
      <h3>Nach Klasse</h3>
      <table class="fr-table">
        <thead><tr><th>Klasse</th><th>Anzahl</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="2" class="muted">Keine Daten</td></tr>'}</tbody>
      </table>
      <div class="fr-actions" style="margin-top:12px">
        <button type="button" class="btn" id="frBtnCsvInline"><i class="bi bi-download"></i>CSV exportieren</button>
      </div>
    </section>`;
}

function htmlAntragKlasseField(state, f, studentChrome) {
    const klasseCode = resolveStudentKlasseCode(state);
    const selected = String(klasseCode || f.klasse || '').trim();
    const classes = studentChrome ? classesForStudentPicker(state) : state.stammdaten.classes || [];
    const classOpts = optionList(classes, selected, 'Klasse wählen…');
    const klasseLocked = studentChrome && isStudentKlasseLocked(state);
    const hit = findClassInStammdaten(classes, selected);
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
            ? '<span class="muted" style="font-size:0.85em;display:block;margin-top:4px;">Ihre Klasse (' +
              esc(src) +
              ') – nicht änderbar.</span>'
            : '';
    const hintPick =
        studentChrome && !klasseLocked && classes.length
            ? '<span class="muted" style="font-size:0.85em;display:block;margin-top:4px;">Wählen Sie Ihre Klasse aus der Schul-Liste.</span>'
            : '';
    const hintEmpty =
        studentChrome && !klasseLocked && !classes.length
            ? '<p class="muted" style="margin:6px 0 0;font-size:0.88em;">Noch keine Klassen aus den Stammdaten (1A, 1B, …). Bitte als IT im <strong>Freistellungen-Setup → Gruppen speichern</strong> – das überschreibt die alten Demo-Werte (1AHW …) in SharePoint. Danach hier <strong>Aktualisieren</strong>.</p>'
            : '';
    if (studentChrome) {
        if (klasseLocked && selected) {
            return `<label>Klasse *
            <input type="text" id="frFormKlasse" required readonly aria-readonly="true" value="${esc(
                selected
            )}" style="background:var(--surface-2,#f4f6f8);">
            ${hintStamm}
          </label>`;
        }
        const opts =
            classes.length
                ? classOpts
                : optionList([], selected, 'Klassenliste wird geladen…');
        return `<label class="fr-form-klasse-label">Klasse *
            <select id="frFormKlasse" class="fr-form-klasse-select" required aria-label="Klasse wählen">${opts}</select>
            ${classes.length ? hintPick : hintEmpty}
          </label>`;
    }
    return `<label class="fr-form-klasse-label">Klasse *
            <select id="frFormKlasse" class="fr-form-klasse-select" required aria-label="Klasse wählen">${classOpts}</select>
          </label>`;
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
    const klasseLocked = studentChrome && isStudentKlasseLocked(state);
    const klasseChosen = !!String(klasseCode || f.klasse || '').trim();
    const nameLocked = studentChrome && !!String(state.accountName || '').trim();
    const kvPrefilled =
        studentChrome && klasseChosen && !!String(f.kvEmail || '').trim().includes('@');
    const klasseField = htmlAntragKlasseField(state, f, studentChrome);
    const kvField = kvPrefilled
        ? `<label>Klassenvorstand
            <input type="text" readonly value="${esc(
                (f.kvName ? f.kvName + ' · ' : '') + (f.kvEmail || '')
            )}" style="background:var(--surface-2,#f4f6f8);">
            <input type="hidden" id="frFormKvEmail" value="${esc(f.kvEmail)}">
            <span class="muted" style="font-size:0.85em;display:block;margin-top:4px;">Automatisch aus Stammdaten Ihrer Klasse.</span>
          </label>`
        : studentChrome
          ? `<label>Klassenvorstand *
            <input type="search" id="frFormKvSearch" placeholder="Lehrperson in Microsoft 365 suchen (Name oder E-Mail)…" autocomplete="off" value="${esc(
                  f.kvName && f.kvEmail ? f.kvName + ' · ' + f.kvEmail : f.kvEmail || ''
              )}">
            <input type="hidden" id="frFormKvEmail" value="${esc(f.kvEmail)}">
            <input type="hidden" id="frFormKvName" value="${esc(f.kvName || '')}">
            <ul id="frFormKvSearchHits" class="fr-kv-search-hits" hidden></ul>
            <span class="muted" style="font-size:0.85em;display:block;margin-top:4px;">${
                klasseCode
                    ? 'Suche Ihre Lehrperson – wird dem Antrag zugeordnet.'
                    : 'Zuerst Klasse, dann Lehrperson wählen.'
            }</span>
          </label>`
          : `<label>Klassenvorstand (E-Mail) *
            <input type="email" id="frFormKvEmail" required value="${esc(f.kvEmail)}" placeholder="kv@schule.at">
          </label>`;
    return `
    <section class="fr-panel">
      ${
          studentChrome
              ? ''
              : `<h2>Neuer Freistellungsantrag</h2>
      <p class="muted">Nach dem Speichern schreibt die App in die SharePoint-Liste (Status <em>Ausstehend</em>). Der eingerichtete Flow startet dann den Genehmigungsprozess.</p>`
      }
      <form id="frAntragForm" class="fr-form" autocomplete="on">
        <div class="fr-form-grid">
          <label>Ihr Name *
            <input type="text" id="frFormName" required value="${esc(f.schuelerName)}" placeholder="Vor- und Nachname"${
                nameLocked ? ' readonly style="background:var(--surface-2,#f4f6f8);"' : ''
            }>
          </label>
          ${klasseField}
          <label>Beginn *
            <input type="datetime-local" id="frFormBeginn" required step="300" value="${esc(
                toIsoDateTimeLocal(f.beginn) || f.beginn || ''
            )}">
          </label>
          <label>Ende *
            <input type="datetime-local" id="frFormEnde" required step="300" value="${esc(
                toIsoDateTimeLocal(f.ende) || f.ende || ''
            )}">
          </label>
          <p class="muted fr-form-span2" style="margin:0;font-size:0.85em;">Datum und Uhrzeit – auch für einzelne Unterrichtsstunden (z. B. 08:00–09:00 am selben Tag).</p>
          <label>Kategorie *
            <select id="frFormKat" required>
              ${kategorieChoicesForState(state)
                  .map(
                      (k) =>
                          `<option value="${esc(k)}"${f.kategorie === k ? ' selected' : ''}>${esc(k)}</option>`
                  )
                  .join('')}
            </select>
          </label>
          ${kvField}
          ${htmlFreistellungUploadField({ student: studentChrome })}
          <label class="fr-form-span2">Begründung
            <textarea id="frFormBeschreibung" rows="3" placeholder="Kurz den Grund beschreiben…">${esc(f.beschreibung)}</textarea>
          </label>
        </div>
        <div class="fr-path ${path.multiDay ? 'is-multi' : ''}">
          <strong>Genehmigungspfad:</strong> ${esc(path.label)}
          ${path.dayCount != null ? ` <span class="muted">(${path.dayCount} Tag${path.dayCount === 1 ? '' : 'e'})</span>` : ''}
        </div>
        ${
            check.warnings.length
                ? `<ul class="fr-hints">${check.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul>`
                : ''
        }
        <div class="fr-actions">
          <button type="submit" class="btn btn-success" id="frBtnSubmit"><i class="bi bi-send"></i>Antrag einreichen</button>
          <button type="button" class="btn" id="frBtnFormReset">Zurücksetzen</button>
        </div>
      </form>
    </section>`;
}

function renderTable(items, state, opts) {
    const o = opts || {};
    const rows = items
        .map((it) => {
            const days = inclusiveDayCount(it.beginn, it.ende);
            const decide =
                o.showDecide && String(it.status) === 'Ausstehend'
                    ? `<button type="button" class="btn btn-sm" data-fr-approve="${esc(it.itemId)}" title="Status Genehmigt (Fallback)">✓</button>
                       <button type="button" class="btn btn-sm" data-fr-reject="${esc(it.itemId)}" title="Status Abgelehnt (Fallback)">✗</button>`
                    : '';
            return `<tr>
          <td><button type="button" class="fr-link" data-fr-detail="${esc(it.itemId)}">${esc(it.schuelerName || it.titel)}</button></td>
          <td>${esc(it.klasse)}</td>
          <td>${esc(formatDeDateTime(it.beginn))}${
                it.ende && it.ende !== it.beginn ? ' – ' + esc(formatDeDateTime(it.ende)) : ''
            }</td>
          <td>${days != null ? days : '–'}${it.multiDay ? ' <span class="fr-badge fr-badge--info">mehr</span>' : ''}</td>
          <td>${statusBadge(it.status)}</td>
          <td>${esc(it.kategorie)}</td>
          <td>${esc(it.kvName || it.kvEmail || '–')}</td>
          <td class="fr-td-actions">${decide}
            <button type="button" class="btn btn-sm" data-fr-detail="${esc(it.itemId)}"><i class="bi bi-eye"></i></button>
          </td>
        </tr>`;
        })
        .join('');
    return `
    <div class="fr-table-wrap">
      <table class="fr-table">
        <thead>
          <tr>
            <th>Schüler/in</th><th>Klasse</th><th>Zeitraum</th><th>Tage</th>
            <th>Status</th><th>Kategorie</th><th>KV</th><th></th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function renderDetailModal(state) {
    const it = (state.items || []).find((x) => String(x.itemId) === String(state.detailId));
    if (!it || !itemVisibleForRole(state, it, { scopeAll: state.role === 'direktion' })) return '';
    return `
    <div class="fr-modal" role="dialog" aria-modal="true" aria-labelledby="frDetailTitle">
      <div class="fr-modal__card">
        <header class="fr-modal__head">
          <h2 id="frDetailTitle">${esc(it.schuelerName || it.titel)}</h2>
          <button type="button" class="btn" data-fr-close-detail aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
        </header>
        <div class="fr-modal__body">
          <dl class="fr-dl">
            <dt>Status</dt><dd>${statusBadge(it.status)}</dd>
            <dt>Klasse</dt><dd>${esc(it.klasse)}</dd>
            <dt>Zeitraum</dt><dd>${esc(formatDeDateTime(it.beginn))} – ${esc(formatDeDateTime(it.ende))} (${it.dayCount ?? '–'} Tag${
                it.dayCount === 1 ? '' : 'e'
            })</dd>
            <dt>Genehmigung</dt><dd>${esc(it.approvalLabel)}</dd>
            <dt>Kategorie</dt><dd>${esc(it.kategorie)}</dd>
            <dt>KV</dt><dd>${esc(it.kvName || '–')} &lt;${esc(it.kvEmail || '')}&gt;</dd>
            <dt>Beschreibung</dt><dd>${esc(it.beschreibung || '–')}</dd>
            <dt>Nachweise</dt><dd>${
                (it.nachweise || []).length
                    ? (it.nachweise || [])
                          .map(
                              (n) =>
                                  `<a href="${esc(n.url)}" target="_blank" rel="noopener noreferrer">${esc(
                                      n.name
                                  )}</a>`
                          )
                          .join('<br>')
                    : '–'
            }</dd>
            <dt>Bemerkungen</dt><dd>${esc(it.bemerkungen || '–')}</dd>
            <dt>Beantragt von</dt><dd>${esc(it.authorName || '–')}${it.authorEmail ? ` &lt;${esc(it.authorEmail)}&gt;` : ''}</dd>
            ${
                it.genehmigtVonKv || it.genehmigtAmKv
                    ? `<dt>Freigabe KV</dt><dd>${esc(it.genehmigtVonKv || '–')}${it.genehmigtAmKv ? ' · ' + esc(formatDeDate(it.genehmigtAmKv)) : ''}</dd>`
                    : ''
            }
            ${
                it.genehmigtVonDirektion || it.genehmigtAmDirektion
                    ? `<dt>Freigabe Direktion</dt><dd>${esc(it.genehmigtVonDirektion || '–')}${it.genehmigtAmDirektion ? ' · ' + esc(formatDeDate(it.genehmigtAmDirektion)) : ''}</dd>`
                    : ''
            }
            ${
                it.abgelehntVon || it.abgelehntAm
                    ? `<dt>Abgelehnt</dt><dd>${esc(it.abgelehntVon || '–')}${it.abgelehntAm ? ' · ' + esc(formatDeDate(it.abgelehntAm)) : ''}</dd>`
                    : ''
            }
          </dl>
        </div>
      </div>
    </div>`;
}

export function readFormFromDom(root) {
    const $ = (id) => (root || document).querySelector('#' + id);
    return {
        schuelerName: String(($('frFormName') && $('frFormName').value) || '').trim(),
        klasse: String(($('frFormKlasse') && $('frFormKlasse').value) || '').trim(),
        beginn: String(($('frFormBeginn') && $('frFormBeginn').value) || '').trim(),
        ende: String(($('frFormEnde') && $('frFormEnde').value) || '').trim(),
        kategorie: String(($('frFormKat') && $('frFormKat').value) || '').trim(),
        kvEmail: String(($('frFormKvEmail') && $('frFormKvEmail').value) || '')
            .trim()
            .toLowerCase(),
        kvName: String(($('frFormKvName') && $('frFormKvName').value) || '').trim(),
        beschreibung: String(($('frFormBeschreibung') && $('frFormBeschreibung').value) || '').trim(),
        status: 'Ausstehend'
    };
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

export function readFiltersFromDom(root) {
    const $ = (id) => (root || document).querySelector('#' + id);
    return {
        q: String(($('frFilterQ') && $('frFilterQ').value) || '').trim(),
        klasse: String(($('frFilterKlasse') && $('frFilterKlasse').value) || ''),
        status: String(($('frFilterStatus') && $('frFilterStatus').value) || ''),
        kategorie: String(($('frFilterKat') && $('frFilterKat').value) || ''),
        multiDay: String(($('frFilterMulti') && $('frFilterMulti').value) || '')
    };
}

export { applyKvFromClass } from './freistellung-planer-state.js';
