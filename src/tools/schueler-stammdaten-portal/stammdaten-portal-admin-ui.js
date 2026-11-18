import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import { PRESET_OPTIONS } from './stammdaten-portal-fields.js';
import { buildEinschreiberWorkbookBlob } from './stammdaten-portal-export.js';
import { clearLocalDemoEnrollments, isDemoOptInRequested, seedLocalDemoEnrollments } from './stammdaten-portal-demo.js';
import { STATUS } from './stammdaten-portal-logic.js';
import {
    remoteEnroll,
    remoteEnsureList,
    remoteHealth,
    remoteList,
    remoteResetPin,
    remoteSetStatus
} from './stammdaten-portal-remote.js';
import { buildSelfHostZip } from './stammdaten-portal-selfhost.js';
import {
    buildParentFormUrl,
    createLocalRecord,
    ensurePilotRemoteDefaults,
    exportSettingsAsPublicConfig,
    getAdminRecordDetail,
    listRecords,
    loadPortalSettings,
    markGeprueft,
    markImported,
    resetLocalAccessPin,
    savePortalSettings
} from './stammdaten-portal-store.js';

const STATUS_FILTERS = [
    { id: 'all', label: 'Alle' },
    { id: STATUS.VORERFASST, label: 'Vorerfasst' },
    { id: STATUS.IN_BEARBEITUNG, label: 'In Bearbeitung' },
    { id: STATUS.VOLLSTAENDIG, label: 'Vollständig' },
    { id: STATUS.GEPRUEFT, label: 'Geprüft' },
    { id: STATUS.IMPORTIERT, label: 'Importiert' }
];

const VIEWS = [
    { id: 'uebersicht', label: 'Übersicht', icon: 'bi-speedometer2' },
    { id: 'anmeldungen', label: 'Anmeldungen', icon: 'bi-people' },
    { id: 'neu', label: 'Neu anlegen', icon: 'bi-plus-lg' },
    { id: 'einrichtung', label: 'Einrichtung', icon: 'bi-gear' }
];

/**
 * @param {HTMLElement} host
 */
export function mountAdminApp(host) {
    if (!host) return;
    let settings = ensurePilotRemoteDefaults();
    let view = 'uebersicht';
    let lastPin = null;
    let lastGuid = null;
    let detailGuid = null;
    let remoteCache = [];
    let message = '';
    let error = '';
    let loading = false;
    let filterQ = '';
    let filterStatus = 'all';
    /** @type {{ health?: string, store?: string, list?: string, detail?: string }} */
    let gwStatus = {};
    let softNotice = '';

    const root = document.createElement('div');
    root.className = 'ssp-admin-root';
    host.appendChild(root);

    bootstrapRemote()
        .then(() => maybeSeedDemo())
        .then(() => refreshRecords())
        .then(render);

    async function bootstrapRemote() {
        settings = ensurePilotRemoteDefaults();
        hydrateSiteFromSchoolTools();
        const hostName = typeof window !== 'undefined' ? window.location.hostname : '';
        const isLocal = hostName === 'localhost' || hostName === '127.0.0.1';
        if (isLocal && !String(settings.functionKey || '').trim()) {
            try {
                const res = await fetch('/__ssp-dev-admin-key', { cache: 'no-store' });
                if (res.ok) {
                    const data = await res.json();
                    if (data?.adminKey) {
                        settings = ensurePilotRemoteDefaults({ functionKey: data.adminKey });
                    }
                }
            } catch {
                /* optional */
            }
        }
        await probeGateway({ quiet: true });
    }

    /** Site-URL aus school-tool-sites übernehmen, ohne gesetzte Werte zu überschreiben. */
    function hydrateSiteFromSchoolTools() {
        settings = loadPortalSettings();
        if (String(settings.spSiteUrl || '').trim()) return;
        try {
            const fromShared = resolveSchoolToolSiteUrl('stammdatenPortal');
            if (fromShared) {
                settings = savePortalSettings({ spSiteUrl: fromShared });
            }
        } catch {
            /* optional */
        }
    }

    async function maybeSeedDemo() {
        if (!isDemoOptInRequested()) return;
        settings = loadPortalSettings();
        if (settings.mode === 'remote') {
            softNotice =
                'Demo (?demo=1) startet nur im lokalen Modus. Unter Einrichtung auf „local“ wechseln und Demo laden.';
            return;
        }
        const result = await seedLocalDemoEnrollments({ force: false });
        if (result.ok && result.seeded > 0) {
            message = `Demo geladen (${result.seeded} Anmeldungen). PINs nur lokal – siehe Konsole.`;
            if (result.pins?.length) console.info('[ssp-demo] PINs', result.pins);
        } else if (result.reason) {
            softNotice = result.reason;
        }
    }

    function currentActor() {
        try {
            const w = typeof window !== 'undefined' ? window : null;
            const account = w?.ms365Auth?.getAccount?.() || w?.ms365Msal?.getActiveAccount?.();
            const email = String(account?.username || account?.mail || '').trim();
            const name = String(account?.name || '').trim();
            return { actorEmail: email, actorName: name };
        } catch {
            return { actorEmail: '', actorName: '' };
        }
    }

    async function probeGateway({ quiet = false } = {}) {
        settings = loadPortalSettings();
        if (settings.mode !== 'remote' || !settings.apiBaseUrl) {
            gwStatus = { health: 'local', store: '', list: '', detail: 'Modus local – kein Gateway.' };
            return;
        }
        try {
            const h = await remoteHealth(settings);
            gwStatus = {
                health: 'ok',
                store: h.store || '?',
                list: gwStatus.list || '',
                detail: `Gateway OK · Store: ${h.store || '?'}`
            };
            if (!quiet) {
                message = gwStatus.detail;
                error = '';
            }
        } catch (e) {
            gwStatus = {
                health: 'bad',
                store: '',
                list: gwStatus.list === 'ok' ? gwStatus.list : 'bad',
                detail: e?.message || 'Gateway nicht erreichbar.'
            };
            if (!quiet) {
                softNotice =
                    gwStatus.detail +
                    ' – bestehende Anmeldungen bleiben im Cache; Demo/local ist weiter nutzbar.';
                error = '';
                message = '';
            }
            return;
        }
        if (settings.functionKey && settings.spSiteUrl) {
            try {
                const r = await remoteEnsureList(settings);
                if (r.mode === 'memory') {
                    gwStatus.list = 'warn';
                    gwStatus.detail = `${gwStatus.detail} · Liste: Store noch memory`;
                } else {
                    gwStatus.list = 'ok';
                    gwStatus.detail = `${gwStatus.detail} · Liste „${r.listName || 'SchuelerAnmeldungen'}“ OK`;
                }
            } catch (e) {
                gwStatus.list = 'bad';
                gwStatus.detail = `${gwStatus.detail} · Liste: ${e?.message || 'Fehler'}`;
            }
        } else if (!quiet) {
            gwStatus.list = 'warn';
        }
    }

    async function refreshRecords() {
        settings = loadPortalSettings();
        if (settings.mode === 'remote') {
            if (!settings.apiBaseUrl || !settings.functionKey) {
                softNotice =
                    softNotice ||
                    'Remote ohne Admin-Key/API – zuletzt geladene Anmeldungen bleiben sichtbar. Key unter Einrichtung setzen.';
                return;
            }
            try {
                const data = await remoteList(settings);
                remoteCache = (data.records || []).map((r) => ({
                    uniqueGuid: r.id,
                    title: r.title,
                    status: r.status,
                    updatedAt: r.updatedAt,
                    reviewedAt: r.reviewedAt || null,
                    reviewedBy: r.reviewedBy || '',
                    reviewedByName: r.reviewedByName || '',
                    payload: r.payload || {}
                }));
                softNotice = '';
            } catch (e) {
                // local-first: Cache nicht leeren bei 401/403/Gateway-down
                softNotice =
                    (e?.message || 'Remote-Liste fehlgeschlagen.') +
                    (remoteCache.length
                        ? ` · ${remoteCache.length} Anmeldung(en) aus dem letzten erfolgreichen Abruf bleiben sichtbar.`
                        : ' · Keine gecachten Anmeldungen.');
                error = '';
            }
        }
    }

    function currentRecords() {
        return settings.mode === 'remote' ? remoteCache : listRecords();
    }

    function detailFor(guid) {
        if (settings.mode === 'remote') {
            const r = remoteCache.find((x) => x.uniqueGuid === guid);
            if (!r) return null;
            return {
                id: r.uniqueGuid,
                title: r.title,
                status: r.status,
                payload: r.payload,
                reviewedAt: r.reviewedAt || null,
                reviewedBy: r.reviewedBy || '',
                reviewedByName: r.reviewedByName || '',
                formUrl: buildParentFormUrl(r.uniqueGuid, settings)
            };
        }
        return getAdminRecordDetail(guid);
    }

    function kpis(records) {
        const counts = {
            total: records.length,
            vorerfasst: 0,
            bearbeitung: 0,
            vollstaendig: 0,
            geprueft: 0,
            importiert: 0
        };
        for (const r of records) {
            if (r.status === STATUS.VORERFASST) counts.vorerfasst += 1;
            else if (r.status === STATUS.IN_BEARBEITUNG) counts.bearbeitung += 1;
            else if (r.status === STATUS.VOLLSTAENDIG) counts.vollstaendig += 1;
            else if (r.status === STATUS.GEPRUEFT) counts.geprueft += 1;
            else if (r.status === STATUS.IMPORTIERT) counts.importiert += 1;
        }
        return counts;
    }

    function topTitle() {
        const map = {
            uebersicht: 'Übersicht',
            anmeldungen: 'Anmeldungen',
            neu: 'Neuen Datensatz anlegen',
            einrichtung: 'Einrichtung'
        };
        return map[view] || 'Schüler-Stammdaten';
    }

    function topHint() {
        if (view === 'einrichtung') {
            return 'Gateway, SharePoint, Self-Host und Feldprofil. Technik nur hier.';
        }
        if (view === 'neu') {
            return 'GUID + PIN erzeugen. Link ohne PIN an Eltern; PIN getrennt übermitteln.';
        }
        if (view === 'anmeldungen') {
            return 'Zwischenspeicher anzeigen, Links kopieren, PIN neu setzen, Sokrates exportieren.';
        }
        return settings.mode === 'remote'
            ? 'Remote-Gateway aktiv. Technik unter „Einrichtung“.'
            : 'Lokaler Pilot. Für Produktion unter „Einrichtung“ auf remote umstellen.';
    }

    function render() {
        settings = loadPortalSettings();
        const records = currentRecords();
        const kpi = kpis(records);
        root.innerHTML = `
    <div class="fr-shell fr-shell--student ssp-shell">
      <aside class="fr-nav" aria-label="Navigation">
        <div class="fr-nav__brand">
          <i class="bi bi-person-vcard" aria-hidden="true"></i>
          <div>
            <strong>Stammdaten</strong>
            <span>Sekretariat / IT</span>
          </div>
        </div>
        <nav class="fr-nav__list">
          ${VIEWS.map(
              (v) =>
                  `<button type="button" class="fr-nav__btn${view === v.id ? ' is-active' : ''}" data-ssp-view="${esc(v.id)}">
                    <i class="bi ${esc(v.icon)}" aria-hidden="true"></i>${esc(v.label)}
                  </button>`
          ).join('')}
        </nav>
        <div class="fr-nav__session">
          <span class="fr-nav__role-label">Betrieb</span>
          <p class="fr-nav__account-name">${esc(settings.schoolName || 'Schule')}</p>
          <div class="fr-nav__role-btns" role="status">
            <span class="fr-chip is-active">${esc(settings.mode === 'remote' ? 'remote' : 'local')}</span>
            <span class="fr-chip">${esc(settings.presetId || 'kern')}</span>
            ${
                gwStatus.health === 'ok'
                    ? `<span class="fr-chip ssp-chip-ok" title="${esc(gwStatus.detail || '')}">Gateway</span>`
                    : gwStatus.health === 'bad'
                      ? `<span class="fr-chip ssp-chip-bad" title="${esc(gwStatus.detail || '')}">Gateway</span>`
                      : ''
            }
            ${
                gwStatus.list === 'ok'
                    ? `<span class="fr-chip ssp-chip-ok" title="${esc(gwStatus.detail || '')}">Liste</span>`
                    : gwStatus.list === 'bad'
                      ? `<span class="fr-chip ssp-chip-bad">Liste</span>`
                      : gwStatus.list === 'warn'
                        ? `<span class="fr-chip ssp-chip-warn">Liste</span>`
                        : ''
            }
            ${gwStatus.store ? `<span class="fr-chip">${esc(gwStatus.store)}</span>` : ''}
          </div>
          <p class="fr-nav__hint">${records.length} Anmeldung(en) · PIN nie in der URL</p>
          <div class="ssp-nav-links">
            <a class="fr-nav__btn" href="../index.html"><i class="bi bi-grid"></i>Dashboard</a>
            <a class="fr-nav__btn" href="schueler-stammdaten-formular.html?mode=${esc(settings.mode || 'local')}" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i>Eltern-Formular</a>
          </div>
        </div>
      </aside>
      <div class="fr-main">
        <header class="fr-top fr-top--student">
          <div>
            <h2 class="fr-top__student-title">${esc(topTitle())}</h2>
            <p class="fr-nav__hint" style="opacity:.75;margin:0">${esc(topHint())}</p>
          </div>
          <button type="button" class="btn" id="sspRefresh" title="Aktualisieren"><i class="bi bi-arrow-repeat"></i>Aktualisieren</button>
        </header>
        ${error ? `<div class="fr-alert fr-alert--bad" role="alert">${esc(error)}</div>` : ''}
        ${softNotice ? `<div class="fr-alert fr-alert--warn ssp-soft-notice" role="status">${esc(softNotice)}</div>` : ''}
        ${message ? `<div class="fr-alert fr-alert--info" role="status">${esc(message)}</div>` : ''}
        <div class="fr-view">
          ${
              view === 'uebersicht'
                  ? renderOverview(records, kpi)
                  : view === 'anmeldungen'
                    ? renderAnmeldungen(records)
                    : view === 'neu'
                      ? renderNeu()
                      : renderEinrichtung()
          }
        </div>
      </div>
    </div>`;
        bind();
    }

    function renderAttentionBanner(kpi) {
        const open = kpi.vorerfasst + kpi.bearbeitung;
        if (kpi.vollstaendig > 0) {
            return `
        <section class="ssp-calm-banner ssp-calm-banner--action" role="status">
          <div class="ssp-calm-banner__text">
            <p class="ssp-calm-banner__eyebrow">Was braucht Aufmerksamkeit?</p>
            <p><strong>${kpi.vollstaendig}</strong> Anmeldung(en) vollständig – Sekretariat kann prüfen und für Sokrates freigeben.</p>
          </div>
          <button type="button" class="btn btn-success" data-ssp-jump-status="${esc(STATUS.VOLLSTAENDIG)}">Zu prüfen öffnen</button>
        </section>`;
        }
        if (kpi.geprueft > 0) {
            return `
        <section class="ssp-calm-banner ssp-calm-banner--action" role="status">
          <div class="ssp-calm-banner__text">
            <p class="ssp-calm-banner__eyebrow">Was braucht Aufmerksamkeit?</p>
            <p><strong>${kpi.geprueft}</strong> geprüft – bereit für Sokrates-Einschreiber-Export.</p>
          </div>
          <button type="button" class="btn btn-success" id="sspExportQuick">Sokrates exportieren</button>
        </section>`;
        }
        if (open > 0) {
            return `
        <section class="ssp-calm-banner ssp-calm-banner--action" role="status">
          <div class="ssp-calm-banner__text">
            <p class="ssp-calm-banner__eyebrow">Was braucht Aufmerksamkeit?</p>
            <p><strong>${open}</strong> Anmeldung(en) noch offen (Vorerfasst / In Bearbeitung).</p>
          </div>
          <button type="button" class="btn" data-ssp-view="anmeldungen">Anmeldungen anzeigen</button>
        </section>`;
        }
        if (kpi.total === 0) {
            return `
        <section class="ssp-calm-banner ssp-calm-banner--action" role="status">
          <div class="ssp-calm-banner__text">
            <p class="ssp-calm-banner__eyebrow">Loslegen</p>
            <p>Noch keine Anmeldungen. Legen Sie den ersten Datensatz an oder laden Sie die lokale Demo.</p>
          </div>
          <div class="ssp-calm-banner__actions">
            <button type="button" class="btn btn-success" data-ssp-view="neu">Neu anlegen</button>
            <button type="button" class="btn" data-ssp-load-demo>Demo laden</button>
          </div>
        </section>`;
        }
        return '';
    }

    function renderQuietEmpty({ title, text, ctaLabel, ctaAttr }) {
        return `
      <div class="ssp-quiet-empty" role="status">
        <p class="ssp-quiet-empty__title">${esc(title)}</p>
        <p class="ssp-quiet-empty__text">${esc(text)}</p>
        ${
            ctaLabel
                ? `<button type="button" class="btn" ${ctaAttr || ''}>${esc(ctaLabel)}</button>`
                : ''
        }
      </div>`;
    }

    function renderOverview(records, kpi) {
        const recent = records.slice(0, 5);
        return `
      ${renderAttentionBanner(kpi)}
      <section class="fr-panel">
        <h2>Übersicht</h2>
        <p class="muted">Status der Schüler-Anmeldungen${settings.mode === 'remote' ? ' (Gateway)' : ' (lokal)'}.</p>
        <div class="fr-kpi-row">
          <div class="fr-kpi"><strong>${kpi.total}</strong><span>Gesamt</span></div>
          <div class="fr-kpi"><strong>${kpi.vorerfasst}</strong><span>Vorerfasst</span></div>
          <div class="fr-kpi"><strong>${kpi.bearbeitung}</strong><span>In Bearbeitung</span></div>
          <div class="fr-kpi"><strong>${kpi.vollstaendig}</strong><span>Vollständig</span></div>
          <div class="fr-kpi"><strong>${kpi.geprueft}</strong><span>Geprüft</span></div>
        </div>
      </section>
      <section class="fr-panel">
        <h3>Schnellstart</h3>
        <div class="fr-actions" style="margin-top:4px">
          <button type="button" class="btn btn-success" data-ssp-view="neu"><i class="bi bi-plus-lg"></i>Neu anlegen</button>
          <button type="button" class="btn" data-ssp-view="anmeldungen"><i class="bi bi-people"></i>Alle Anmeldungen</button>
          <button type="button" class="btn" id="sspExportQuick"><i class="bi bi-file-earmark-excel"></i>Sokrates-Export</button>
          <button type="button" class="btn" data-ssp-load-demo title="Nur lokal, opt-in"><i class="bi bi-magic"></i>Demo</button>
          <button type="button" class="btn" data-ssp-view="einrichtung"><i class="bi bi-gear"></i>Einrichtung</button>
        </div>
      </section>
      <section class="fr-panel">
        <h3>Zuletzt aktualisiert</h3>
        ${
            recent.length
                ? `<table class="ssp-table">
              <thead><tr><th>Titel</th><th>Status</th><th>Aktualisiert</th></tr></thead>
              <tbody>
                ${recent
                    .map(
                        (r) => `<tr>
                    <td><button type="button" class="ssp-linkish" data-open="${esc(r.uniqueGuid)}">${esc(r.title || '—')}</button></td>
                    <td>${statusBadge(r.status)}</td>
                    <td>${esc(formatTs(r.updatedAt))}</td>
                  </tr>`
                    )
                    .join('')}
              </tbody>
            </table>`
                : renderQuietEmpty({
                      title: 'Noch keine Anmeldungen',
                      text: 'Legen Sie den ersten Datensatz an oder laden Sie die lokale Demo (?demo=1).',
                      ctaLabel: 'Neu anlegen',
                      ctaAttr: 'data-ssp-view="neu"'
                  })
        }
      </section>`;
    }

    function renderAnmeldungen(records) {
        const q = filterQ.trim().toLowerCase();
        let filtered = records;
        if (filterStatus && filterStatus !== 'all') {
            filtered = filtered.filter((r) => String(r.status || '') === filterStatus);
        }
        if (q) {
            filtered = filtered.filter(
                (r) =>
                    String(r.title || '')
                        .toLowerCase()
                        .includes(q) ||
                    String(r.uniqueGuid || '')
                        .toLowerCase()
                        .includes(q) ||
                    String(r.status || '')
                        .toLowerCase()
                        .includes(q)
            );
        }
        const detail = detailGuid ? detailFor(detailGuid) : null;
        const kpi = kpis(records);
        return `
      ${renderAttentionBanner(kpi)}
      <section class="fr-panel">
        <h2>Anmeldungen (${records.length})</h2>
        <div class="ssp-status-chips" role="tablist" aria-label="Status filtern">
          ${STATUS_FILTERS.map((f) => {
              const count =
                  f.id === 'all'
                      ? records.length
                      : records.filter((r) => r.status === f.id).length;
              return `<button type="button" class="ssp-status-chip${filterStatus === f.id ? ' is-active' : ''}" data-ssp-status-filter="${esc(f.id)}" role="tab" aria-selected="${filterStatus === f.id ? 'true' : 'false'}">${esc(f.label)} <span>${count}</span></button>`;
          }).join('')}
        </div>
        <div class="fr-filters ssp-filters">
          <input type="search" id="sspFilterQ" placeholder="Suchen (Name, GUID, Status) …" value="${esc(filterQ)}" />
          <span class="fr-filters__meta muted">${filtered.length} sichtbar</span>
          <button type="button" class="btn" id="sspExportAll"><i class="bi bi-file-earmark-excel"></i>Sokrates XLSX (Geprüft)</button>
        </div>
        ${
            filtered.length
                ? `<div class="ssp-table-wrap"><table class="ssp-table">
              <thead><tr><th>Titel</th><th>Status</th><th>Aktualisiert</th><th></th></tr></thead>
              <tbody>
                ${filtered
                    .map(
                        (r) => `<tr class="${detailGuid === r.uniqueGuid ? 'ssp-row-active' : ''}">
                    <td>${esc(r.title)}<div class="ssp-muted" style="font-size:.75rem">${esc(r.uniqueGuid)}</div></td>
                    <td>${statusBadge(r.status)}</td>
                    <td>${esc(formatTs(r.updatedAt))}</td>
                    <td class="ssp-row-actions">
                      <button type="button" class="btn" data-view="${esc(r.uniqueGuid)}">Anzeigen</button>
                      <button type="button" class="btn" data-copy="${esc(r.uniqueGuid)}">Link</button>
                      <button type="button" class="btn" data-reset-pin="${esc(r.uniqueGuid)}">PIN neu</button>
                    </td>
                  </tr>`
                    )
                    .join('')}
              </tbody>
            </table></div>`
                : records.length
                  ? renderQuietEmpty({
                        title: 'Keine Treffer',
                        text: 'Filter oder Suche anpassen.',
                        ctaLabel: 'Filter zurücksetzen',
                        ctaAttr: 'id="sspResetFilters"'
                    })
                  : renderQuietEmpty({
                        title: 'Noch keine Anmeldungen',
                        text: 'Legen Sie den ersten Datensatz an. Einrichtung und Gateway bleiben unter „Einrichtung“.',
                        ctaLabel: 'Neu anlegen',
                        ctaAttr: 'data-ssp-view="neu"'
                    })
        }
      </section>
      ${
          detail
              ? `<section class="fr-panel" id="sspDetail">
            <h2>Datensatz: ${esc(detail.title || detail.id)}</h2>
            <p class="muted">Status: ${statusBadge(detail.status)}
              · <a href="${esc(detail.formUrl)}" target="_blank" rel="noopener">Eltern-Formular öffnen</a>
              ${
                  detail.reviewedAt || detail.reviewedBy
                      ? ` · Geprüft ${esc(formatTs(detail.reviewedAt))}${detail.reviewedBy || detail.reviewedByName ? ` von ${esc(detail.reviewedByName || detail.reviewedBy)}` : ''}`
                      : ''
              }</p>
            ${renderPayloadPreview(detail.payload)}
            <div class="fr-actions" style="margin-top:12px">
              <button type="button" class="btn" id="sspCloseDetail">Schließen</button>
              <button type="button" class="btn" data-copy="${esc(detail.id)}">Link kopieren</button>
              <button type="button" class="btn" data-reset-pin="${esc(detail.id)}">PIN neu setzen</button>
              ${
                  detail.status === STATUS.VOLLSTAENDIG
                      ? `<button type="button" class="btn btn-success" data-mark-geprueft="${esc(detail.id)}"><i class="bi bi-stamp"></i>Als geprüft markieren</button>`
                      : ''
              }
            </div>
          </section>`
              : ''
      }`;
    }

    function renderNeu() {
        return `
      <section class="fr-panel">
        <h2>Neuen Datensatz anlegen</h2>
        <p class="muted">Erzeugt GUID + PIN. Der Link enthält <strong>nur die GUID</strong> – PIN getrennt an Eltern übermitteln.</p>
        <div class="ssp-grid" style="max-width:520px">
          <div class="ssp-field"><label for="sspNewTitle">Anzeigename / Titel</label>
            <input id="sspNewTitle" placeholder="z. B. Mustermann, Max" /></div>
        </div>
        <div class="fr-actions" style="margin-top:12px">
          <button type="button" class="btn btn-success" id="sspCreate" ${loading ? 'disabled' : ''}>
            <i class="bi bi-plus-lg"></i>${loading ? 'Wird angelegt …' : 'Anlegen'}
          </button>
        </div>
        ${
            lastPin
                ? `<div class="ssp-pin-result">
              <p>PIN (nur jetzt sichtbar):</p>
              <div class="ssp-pin-box">${esc(lastPin)}</div>
              <p class="muted" style="margin-top:.75rem">Eltern-Link:</p>
              <code class="ssp-link-code">${esc(buildParentFormUrl(lastGuid, settings))}</code>
              <div class="fr-actions" style="margin-top:12px">
                <button type="button" class="btn" id="sspCopyLink"><i class="bi bi-link-45deg"></i>Link kopieren</button>
                <button type="button" class="btn" id="sspCopyPin"><i class="bi bi-key"></i>PIN kopieren</button>
                <a class="btn" href="${esc(buildParentFormUrl(lastGuid, settings))}" target="_blank" rel="noopener"><i class="bi bi-box-arrow-up-right"></i>Formular öffnen</a>
                <button type="button" class="btn" data-ssp-view="anmeldungen">Zur Liste</button>
              </div>
            </div>`
                : ''
        }
      </section>`;
    }

    function renderEinrichtung() {
        return `
      <section class="fr-panel ssp-setup-block ssp-setup-block--it">
        <h2>1 · Schule & Formular</h2>
        <p class="muted">Name und wo Eltern das Formular öffnen (App oder Schulhomepage). IT-Einrichtung – nicht Alltag Sekretariat.</p>
        <div class="ssp-grid">
          <div class="ssp-field"><label for="sspSchoolName">Schulname</label><input id="sspSchoolName" value="${esc(settings.schoolName)}" /></div>
          <div class="ssp-field"><label for="sspSchoolSlug">Schul-Slug</label><input id="sspSchoolSlug" value="${esc(settings.schoolSlug)}" placeholder="z. B. hak-steyr" /></div>
          <div class="ssp-field ssp-field--wide"><label for="sspFormBase">Formular-Basis-URL</label>
            <input id="sspFormBase" value="${esc(settings.formBaseUrl)}" placeholder="leer = App-Formular · oder https://schule.at/anmeldung/" /></div>
          <div class="ssp-field"><label for="sspPreset">Feldprofil</label>
            <select id="sspPreset">
              ${PRESET_OPTIONS.map((p) => `<option value="${p.id}" ${settings.presetId === p.id ? 'selected' : ''}>${esc(p.label)}</option>`).join('')}
            </select>
          </div>
        </div>
      </section>

      <section class="fr-panel ssp-setup-block ssp-setup-block--it">
        <h2>2 · Gateway (Azure)</h2>
        <p class="muted">Zentraler Dienst für PIN-Prüfung und Speichern. Admin-Key nie ins Eltern-Formular.</p>
        <div class="ssp-grid">
          <div class="ssp-field"><label for="sspMode">Modus</label>
            <select id="sspMode">
              <option value="local" ${settings.mode !== 'remote' ? 'selected' : ''}>local – Pilot nur dieser Browser</option>
              <option value="remote" ${settings.mode === 'remote' ? 'selected' : ''}>remote – zentrales Gateway</option>
            </select>
          </div>
          <div class="ssp-field ssp-field--wide"><label for="sspApiBase">API-Base-URL</label>
            <input id="sspApiBase" value="${esc(settings.apiBaseUrl || 'https://func-ms365-ssp-dev.azurewebsites.net/api')}" placeholder="https://func-ms365-ssp-dev.azurewebsites.net/api" />
            <span class="ssp-field-hint">Genau so: <code>https://func-ms365-ssp-dev.azurewebsites.net/api</code> (ohne <code>/health</code> am Ende).</span>
          </div>
          <div class="ssp-field ssp-field--wide"><label for="sspFuncKey">Admin-Key</label>
            <input id="sspFuncKey" type="password" value="${esc(settings.functionKey || '')}" autocomplete="off" placeholder="langer Geheimwert aus Azure" />
            <span class="ssp-field-hint">Wert aus Azure App Setting <code>STAMMDATEN_ADMIN_KEY</code> (nicht den Namen eintippen). Lokal auch in <code>backend/stammdaten-portal-api/.admin-key.local.txt</code>.</span>
          </div>
        </div>
        <div class="fr-actions" style="margin-top:12px">
          <button type="button" class="btn" id="sspHealth"><i class="bi bi-heart-pulse"></i>Gateway prüfen</button>
        </div>
        ${
            gwStatus.detail
                ? `<p class="ssp-gw-status ssp-gw-status--${
                      gwStatus.health === 'ok' && gwStatus.list !== 'bad'
                          ? 'ok'
                          : gwStatus.health === 'bad' || gwStatus.list === 'bad'
                            ? 'bad'
                            : 'warn'
                  }" role="status">${esc(gwStatus.detail)}</p>`
                : ''
        }
        <p class="muted ssp-field-hint" style="margin-top:10px">Lokal über Vite: Proxy <code>/__ssp-gw</code> + Dev-Admin-Key. Seite: <code id="sspPageOrigin"></code></p>
      </section>

      <section class="fr-panel ssp-setup-block ssp-setup-block--it">
        <h2>3 · SharePoint-Liste (Anmeldungen)</h2>
        <p class="muted">Hier landen die echten Anmeldungen. Site im <strong>Schul-Tenant</strong> – z. B. Sekretariats-Site. Listenname Standard: <code>SchuelerAnmeldungen</code>.</p>
        <div class="ssp-grid">
          <div class="ssp-field ssp-field--wide"><label for="sspSpSite">SharePoint-Site-URL</label>
            <input id="sspSpSite" value="${esc(settings.spSiteUrl || '')}" placeholder="https://schule.sharepoint.com/sites/Sekretariat" /></div>
          <div class="ssp-field"><label for="sspSpList">Listenname</label>
            <input id="sspSpList" value="${esc(settings.spListName || 'SchuelerAnmeldungen')}" /></div>
        </div>
        <ol class="ssp-setup-steps muted">
          <li>Site-URL eintragen und unten speichern.</li>
          <li>Modus <strong>remote</strong>, Gateway + Admin-Key ok.</li>
          <li>„Liste anlegen / prüfen“ – erzeugt die Liste inkl. Spalten (Gateway braucht Graph-Zugriff + <code>STAMMDATEN_STORE=sharepoint</code>).</li>
          <li>Dieselbe Site-URL zusätzlich in den Azure App Settings der Function App setzen, damit Eltern-Daten dort landen.</li>
        </ol>
        <div class="fr-actions" style="margin-top:12px">
          <button type="button" class="btn btn-success" id="sspEnsureList"><i class="bi bi-list-check"></i>Liste anlegen / prüfen</button>
        </div>
      </section>

      <section class="fr-panel ssp-setup-block ssp-setup-block--it">
        <h2>4 · Speichern & Self-Host</h2>
        <p class="muted">Config ohne Secrets für Schulhomepage; ZIP mit Formular-Dateien.</p>
        <div class="fr-actions">
          <button type="button" class="btn btn-success" id="sspSaveSetup"><i class="bi bi-check2"></i>Alles speichern</button>
          <button type="button" class="btn" id="sspDownloadConfig"><i class="bi bi-filetype-json"></i>Config JSON</button>
          <button type="button" class="btn" id="sspSelfHost"><i class="bi bi-file-zip"></i>Self-Host-ZIP</button>
          <button type="button" class="btn" data-ssp-load-demo><i class="bi bi-magic"></i>Lokale Demo</button>
          <button type="button" class="btn" id="sspClearDemo">Demo entfernen</button>
        </div>
        <p class="muted" style="margin-top:12px">Doku: <code>docs/schueler-stammdaten-portal.md</code> · Demo nur opt-in (?demo=1 / Button), nie still in Remote.</p>
      </section>`;
    }

    function bind() {
        root.querySelectorAll('[data-ssp-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                view = btn.getAttribute('data-ssp-view') || 'uebersicht';
                error = '';
                message = '';
                softNotice = '';
                render();
            });
        });
        root.querySelectorAll('[data-ssp-status-filter]').forEach((btn) => {
            btn.addEventListener('click', () => {
                filterStatus = btn.getAttribute('data-ssp-status-filter') || 'all';
                render();
            });
        });
        root.querySelectorAll('[data-ssp-jump-status]').forEach((btn) => {
            btn.addEventListener('click', () => {
                filterStatus = btn.getAttribute('data-ssp-jump-status') || 'all';
                view = 'anmeldungen';
                render();
            });
        });
        root.querySelector('#sspResetFilters')?.addEventListener('click', () => {
            filterQ = '';
            filterStatus = 'all';
            render();
        });
        root.querySelectorAll('[data-ssp-load-demo]').forEach((btn) => {
            btn.addEventListener('click', onLoadDemo);
        });
        root.querySelector('#sspRefresh')?.addEventListener('click', async () => {
            error = '';
            softNotice = '';
            await refreshRecords();
            message = 'Aktualisiert.';
            render();
        });
        root.querySelector('#sspExportQuick')?.addEventListener('click', onExport);
        root.querySelector('#sspExportAll')?.addEventListener('click', onExport);
        root.querySelector('#sspFilterQ')?.addEventListener('input', (e) => {
            filterQ = e.target.value || '';
            const wrap = root.querySelector('.ssp-table-wrap') || root.querySelector('.fr-panel');
            // Live-Filter ohne kompletten Shell-Rebuild: erneut rendern
            render();
            const input = root.querySelector('#sspFilterQ');
            if (input) {
                input.focus();
                const len = input.value.length;
                input.setSelectionRange(len, len);
            }
            void wrap;
        });
        root.querySelectorAll('[data-open]').forEach((btn) => {
            btn.addEventListener('click', () => {
                detailGuid = btn.getAttribute('data-open');
                view = 'anmeldungen';
                render();
            });
        });
        root.querySelectorAll('[data-view]').forEach((btn) => {
            btn.addEventListener('click', () => {
                detailGuid = btn.getAttribute('data-view');
                message = '';
                error = '';
                render();
                root.querySelector('#sspDetail')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
            });
        });
        root.querySelectorAll('[data-copy]').forEach((btn) => {
            btn.addEventListener('click', () => {
                copyText(buildParentFormUrl(btn.getAttribute('data-copy'), settings));
                message = 'Link kopiert.';
                error = '';
                render();
            });
        });
        root.querySelectorAll('[data-reset-pin]').forEach((btn) => {
            btn.addEventListener('click', () => onResetPin(btn.getAttribute('data-reset-pin')));
        });
        root.querySelectorAll('[data-mark-geprueft]').forEach((btn) => {
            btn.addEventListener('click', () => onMarkGeprueft(btn.getAttribute('data-mark-geprueft')));
        });
        root.querySelector('#sspCloseDetail')?.addEventListener('click', () => {
            detailGuid = null;
            render();
        });
        root.querySelector('#sspCreate')?.addEventListener('click', onCreate);
        root.querySelector('#sspCopyLink')?.addEventListener('click', () =>
            copyText(buildParentFormUrl(lastGuid, settings))
        );
        root.querySelector('#sspCopyPin')?.addEventListener('click', () => copyText(lastPin || ''));
        root.querySelector('#sspSaveSetup')?.addEventListener('click', onSaveSetup);
        root.querySelector('#sspHealth')?.addEventListener('click', onHealth);
        root.querySelector('#sspEnsureList')?.addEventListener('click', onEnsureList);
        const originEl = root.querySelector('#sspPageOrigin');
        if (originEl) originEl.textContent = window.location.origin;
        root.querySelector('#sspDownloadConfig')?.addEventListener('click', onDownloadConfig);
        root.querySelector('#sspSelfHost')?.addEventListener('click', onSelfHost);
        root.querySelector('#sspClearDemo')?.addEventListener('click', () => {
            const r = clearLocalDemoEnrollments();
            message = r.removed ? `${r.removed} Demo-Zeile(n) entfernt.` : 'Keine Demo-Zeilen gefunden.';
            softNotice = '';
            refreshRecords().then(render);
        });
    }

    function readSetupPatch() {
        return {
            schoolName: val('sspSchoolName'),
            schoolSlug: val('sspSchoolSlug'),
            formBaseUrl: val('sspFormBase'),
            apiBaseUrl: val('sspApiBase'),
            functionKey: val('sspFuncKey'),
            spSiteUrl: val('sspSpSite'),
            spListName: val('sspSpList') || 'SchuelerAnmeldungen',
            mode: val('sspMode'),
            presetId: val('sspPreset')
        };
    }

    async function onSaveSetup() {
        settings = savePortalSettings(readSetupPatch());
        message = 'Setup gespeichert.';
        error = '';
        await refreshRecords();
        render();
    }

    async function onHealth() {
        settings = savePortalSettings(readSetupPatch());
        settings = loadPortalSettings();
        await probeGateway({ quiet: false });
        if (gwStatus.health === 'ok') {
            message = gwStatus.detail || 'Gateway OK';
            error = '';
        }
        render();
    }

    async function onEnsureList() {
        settings = savePortalSettings(readSetupPatch());
        if (!settings.spSiteUrl) {
            error = 'Bitte zuerst die SharePoint-Site-URL eintragen und speichern.';
            message = '';
            render();
            return;
        }
        try {
            const r = await remoteEnsureList(settings);
            if (r.mode === 'memory') {
                gwStatus.list = 'warn';
                message =
                    r.hint ||
                    'Gateway läuft noch auf memory – Liste nicht in SharePoint. Azure: STAMMDATEN_STORE=sharepoint setzen.';
                error = '';
            } else {
                gwStatus.list = 'ok';
                gwStatus.store = gwStatus.store || 'sharepoint';
                gwStatus.health = gwStatus.health || 'ok';
                message = `SharePoint-Liste „${r.listName || 'SchuelerAnmeldungen'}“ OK${r.created ? ' (neu angelegt)' : ''}${r.columnsAdded ? `, +${r.columnsAdded} Spalten` : ''}.`;
                error = '';
            }
            gwStatus.detail = message;
        } catch (e) {
            gwStatus.list = 'bad';
            error = e?.message || 'Liste anlegen fehlgeschlagen.';
            message = '';
        }
        render();
    }

    function onDownloadConfig() {
        settings = savePortalSettings(readSetupPatch());
        const cfg = exportSettingsAsPublicConfig(settings);
        downloadBlob(
            new Blob([JSON.stringify(cfg, null, 2)], { type: 'application/json' }),
            `stammdaten-portal-${cfg.schoolSlug || 'schule'}.json`
        );
        message = 'Öffentliche Config (ohne Admin-Key) heruntergeladen.';
        render();
    }

    async function onSelfHost() {
        settings = savePortalSettings(readSetupPatch());
        try {
            const blob = await buildSelfHostZip({ settings });
            downloadBlob(blob, `schueler-stammdaten-selfhost-${settings.schoolSlug || 'schule'}.zip`);
            message = 'Self-Host-ZIP erzeugt.';
            error = '';
        } catch (e) {
            error = e?.message || 'ZIP fehlgeschlagen (JSZip geladen?).';
            message = '';
        }
        render();
    }

    async function onCreate() {
        error = '';
        message = '';
        loading = true;
        render();
        try {
            settings = loadPortalSettings();
            if (settings.mode === 'remote') {
                const data = await remoteEnroll(settings, { title: val('sspNewTitle') });
                lastPin = data.pin;
                lastGuid = data.id;
            } else {
                const { record, accessPin } = await createLocalRecord({ title: val('sspNewTitle') });
                lastPin = accessPin;
                lastGuid = record.id;
            }
            message = 'Datensatz angelegt. PIN jetzt notieren.';
            await refreshRecords();
        } catch (e) {
            error = e?.message || 'Anlegen fehlgeschlagen.';
        }
        loading = false;
        render();
    }

    async function onResetPin(guid) {
        error = '';
        message = '';
        try {
            if (
                !window.confirm(
                    'Neuen PIN erzeugen? Der bisherige PIN wird ungültig. Formulardaten bleiben erhalten.'
                )
            ) {
                return;
            }
            settings = loadPortalSettings();
            if (settings.mode === 'remote') {
                const data = await remoteResetPin(settings, guid);
                lastPin = data.pin;
                lastGuid = data.id;
            } else {
                const result = await resetLocalAccessPin(guid);
                lastPin = result.accessPin;
                lastGuid = result.record.id;
            }
            detailGuid = lastGuid;
            view = 'neu';
            message = 'Neuer PIN erzeugt – bitte notieren (Anzeige unter „Neu anlegen“).';
            await refreshRecords();
            render();
        } catch (e) {
            error = e?.message || 'PIN konnte nicht erneuert werden.';
            render();
        }
    }

    async function onMarkGeprueft(guid) {
        const id = String(guid || '').trim();
        if (!id) return;
        error = '';
        message = '';
        softNotice = '';
        loading = true;
        render();
        try {
            const actor = currentActor();
            if (settings.mode === 'remote') {
                await remoteSetStatus(settings, [id], STATUS.GEPRUEFT, actor);
            } else {
                markGeprueft([id], actor);
            }
            message = 'Als geprüft markiert – bereit für Sokrates-Export.';
            await refreshRecords();
            detailGuid = id;
            render();
        } catch (e) {
            softNotice = e?.message || 'Stempel fehlgeschlagen.';
            error = '';
            render();
        } finally {
            loading = false;
        }
    }

    async function onLoadDemo() {
        error = '';
        message = '';
        softNotice = '';
        settings = loadPortalSettings();
        if (settings.mode === 'remote') {
            const ok = window.confirm(
                'Remote-Modus aktiv. Demo wechselt auf lokalen Modus und legt nur lokale Demo-Zeilen an (keine SharePoint-Daten). Fortfahren?'
            );
            if (!ok) return;
        }
        const result = await seedLocalDemoEnrollments({
            force: true,
            switchToLocal: settings.mode === 'remote'
        });
        settings = loadPortalSettings();
        if (!result.ok) {
            softNotice = result.reason || 'Demo konnte nicht geladen werden.';
            render();
            return;
        }
        if (result.pins?.length) console.info('[ssp-demo] PINs', result.pins);
        message = `Demo geladen: ${result.seeded} Anmeldung(en) (local-first).`;
        view = 'anmeldungen';
        filterStatus = 'all';
        await refreshRecords();
        render();
    }

    async function onExport() {
        const approved = currentRecords().filter((r) => r.status === STATUS.GEPRUEFT);
        if (!approved.length) {
            error =
                'Keine Datensätze mit Status „Geprüft“. Bitte zuerst im Detail „Als geprüft markieren“ (Sekretariat).';
            message = '';
            render();
            return;
        }
        try {
            const blob = buildEinschreiberWorkbookBlob(approved);
            downloadBlob(blob, `sokrates-einschreiber-${stamp()}.xlsx`);
            const ids = approved.map((r) => r.uniqueGuid);
            if (settings.mode === 'remote') {
                await remoteSetStatus(settings, ids, STATUS.IMPORTIERT);
            } else {
                markImported(ids);
            }
            message = `${approved.length} Zeile(n) exportiert; Status → Importiert.`;
            error = '';
            await refreshRecords();
            render();
        } catch (e) {
            error = e?.message || 'Export fehlgeschlagen.';
            render();
        }
    }

    function val(id) {
        return String(root.querySelector(`#${id}`)?.value || '').trim();
    }
}

function statusBadge(status) {
    const s = String(status || '');
    let cls = 'fr-badge--info';
    if (s === STATUS.GEPRUEFT) cls = 'fr-badge--ok';
    else if (s === STATUS.VOLLSTAENDIG) cls = 'fr-badge--warn';
    else if (s === STATUS.IN_BEARBEITUNG) cls = 'fr-badge--warn';
    else if (s === STATUS.IMPORTIERT) cls = 'fr-badge--info';
    return `<span class="fr-badge ${cls}">${esc(s)}</span>`;
}

function downloadBlob(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

async function copyText(text) {
    try {
        await navigator.clipboard.writeText(String(text || ''));
    } catch {
        /* ignore */
    }
}

function formatTs(iso) {
    if (!iso) return '—';
    try {
        return new Date(iso).toLocaleString('de-AT');
    } catch {
        return String(iso);
    }
}

function stamp() {
    const d = new Date();
    return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

function esc(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function renderPayloadPreview(payload) {
    const p = payload || {};
    const s = p.schueler || {};
    const w = p.wohnadresse || {};
    const g = Array.isArray(p.erziehungsberechtigte) ? p.erziehungsberechtigte : [];
    const rows = [
        ['Schüler:in', [s.familienname, s.vorname].filter(Boolean).join(', ') || '—'],
        ['Geburtsdatum', s.geburtsdatum || '—'],
        ['SVNR', s.svnr || '—'],
        ['Adresse', [w.strasse, w.hausnummer, w.plz, w.ort].filter(Boolean).join(' ') || '—'],
        [
            'ErzBer 1',
            g[0]
                ? [g[0].adressart, g[0].vorname, g[0].familienname, g[0].mail, g[0].mobil]
                      .filter(Boolean)
                      .join(' · ')
                : '—'
        ],
        [
            'ErzBer 2',
            g[1]
                ? [g[1].adressart, g[1].vorname, g[1].familienname, g[1].mail, g[1].mobil]
                      .filter(Boolean)
                      .join(' · ')
                : '—'
        ]
    ];
    return `
      <table class="ssp-table">
        <tbody>
          ${rows.map(([k, v]) => `<tr><th style="width:8rem">${esc(k)}</th><td>${esc(v)}</td></tr>`).join('')}
        </tbody>
      </table>
      <details style="margin-top:.75rem">
        <summary class="muted">Rohdaten (JSON)</summary>
        <pre class="ssp-json">${esc(JSON.stringify(p, null, 2))}</pre>
      </details>`;
}
