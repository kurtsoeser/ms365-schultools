/**
 * Admin-UI: Klassen aus Stammdaten → SharePoint-Spalte „Klasse“ (wie Kategorien-Panel).
 */
import { patchFreistellungKlasseColumn } from './freistellung-planer-graph.js';
import {
    loadEffectivePermissionsConfig,
    savePermissionsConfig,
    normalizeClassCatalog
} from './freistellung-planer-permissions.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') {
        window.ms365ToastOrAlert(msg);
    } else {
        window.alert(msg);
    }
}

function canonicalApi() {
    return typeof window !== 'undefined' ? window.ms365StammdatenCanonical : null;
}

/**
 * Klassenzeilen aus der kanonischen Stammdaten-Quelle (App-Daten v2).
 * @returns {{ code: string, name: string, year: string, headName: string, headEmail: string }[]}
 */
export function listStammdatenKlassenForFreistellung() {
    const api = canonicalApi();
    if (api && typeof api.reconcileStammdatenStorage === 'function') {
        try {
            api.reconcileStammdatenStorage();
        } catch {
            /* ignore */
        }
    }
    if (api && typeof api.listAllClasses === 'function') {
        return api.listAllClasses();
    }
    return [];
}

/**
 * @param {{ mountId: string, syncBtnId?: string, reloadBtnId?: string }} spec
 * @param {() => { siteUrl: string, listId: string }} getSiteCtx
 */
export function wireFreistellungKlassenAdmin(spec, getSiteCtx) {
    const mount = document.getElementById(spec.mountId);
    if (!mount) return { render() {} };

    function render() {
        const api = canonicalApi();
        if (api && typeof api.reconcileStammdatenStorage === 'function') {
            try {
                api.reconcileStammdatenStorage();
            } catch {
                /* ignore */
            }
        }
        const rows = listStammdatenKlassenForFreistellung();
        const year = api && typeof api.currentYearLabel === 'function' ? api.currentYearLabel() : '';
        if (!rows.length) {
            mount.innerHTML =
                '<p class="muted fr-kl-admin__empty">Keine Klassen in den Stammdaten (App-Daten v2). Bitte unter <a href="../tenant.html">Stammdaten → Klassen</a> speichern (z.&nbsp;B. 1A, 1B, …), dann hier „Speicher bereinigen“.</p>';
            return;
        }
        const body = rows
            .map(
                (r) =>
                    `<tr>
              <td><strong>${escapeHtml(r.code)}</strong></td>
              <td>${escapeHtml(r.name)}</td>
              <td>${escapeHtml(r.year || '–')}</td>
              <td>${escapeHtml(r.headName || '–')}</td>
              <td class="muted">${escapeHtml(r.headEmail || '–')}</td>
            </tr>`
            )
            .join('');
        mount.innerHTML =
            '<p class="fr-kl-admin__source muted">Quelle: <strong>App-Daten v2</strong>' +
            (year ? ' · Schuljahr <strong>' + escapeHtml(year) + '</strong>' : '') +
            '</p>' +
            '<div class="fr-kl-admin__table-wrap">' +
            '<table class="fr-kl-admin__table" aria-label="Klassen aus Stammdaten">' +
            '<thead><tr><th>Kürzel</th><th>Name</th><th>Jahrgang</th><th>Klassenvorstand</th><th>E-Mail</th></tr></thead>' +
            '<tbody>' +
            body +
            '</tbody></table></div>' +
            '<p class="fr-kl-admin__meta muted">Aus Stammdaten: <strong>' +
            rows.length +
            '</strong> Klassen – genau diese werden als Auswahl in der SharePoint-Spalte <strong>Klasse</strong> gespeichert.</p>';
    }

    const reloadBtn = spec.reloadBtnId ? document.getElementById(spec.reloadBtnId) : null;
    if (reloadBtn && !reloadBtn.dataset.frKlWired) {
        reloadBtn.dataset.frKlWired = '1';
        reloadBtn.addEventListener('click', () => {
            render();
            toast('Klassenliste aus Stammdaten aktualisiert.');
        });
    }

    const cleanBtn = spec.cleanBtnId ? document.getElementById(spec.cleanBtnId) : null;
    if (cleanBtn && !cleanBtn.dataset.frKlWired) {
        cleanBtn.dataset.frKlWired = '1';
        cleanBtn.addEventListener('click', () => {
            const api = canonicalApi();
            if (!api || typeof api.reconcileStammdatenStorage !== 'function') {
                toast('Stammdaten-Modul nicht geladen – Seite neu laden (Strg+F5).');
                return;
            }
            const result = api.reconcileStammdatenStorage();
            render();
            toast(result.message || 'Speicher bereinigt.');
        });
    }

    const syncBtn = spec.syncBtnId ? document.getElementById(spec.syncBtnId) : null;
    if (syncBtn && !syncBtn.dataset.frKlWired) {
        syncBtn.dataset.frKlWired = '1';
        syncBtn.addEventListener('click', async () => {
            const rows = listStammdatenKlassenForFreistellung();
            if (!rows.length) {
                toast('Keine Klassen in den Stammdaten.');
                return;
            }
            const ctx = typeof getSiteCtx === 'function' ? getSiteCtx() : {};
            const siteUrl = String((ctx && ctx.siteUrl) || '').trim();
            const listId = String((ctx && ctx.listId) || '').trim();
            if (!siteUrl || !listId) {
                toast('Site-URL und Listen-ID fehlen (zuerst Freistellungen-Liste anlegen/wählen).');
                return;
            }
            const G = window.ms365SpoGraph;
            if (!G) {
                toast('SharePoint-Hilfen nicht geladen.');
                return;
            }
            const codes = rows.map((r) => r.code);
            const catalog = normalizeClassCatalog(
                rows.map((r) => ({ code: r.code, name: r.name || r.code }))
            );
            const prev = syncBtn.innerHTML;
            try {
                syncBtn.disabled = true;
                syncBtn.innerHTML = '<i class="bi bi-hourglass-split"></i>Aktualisiere …';
                const cfg = loadEffectivePermissionsConfig();
                savePermissionsConfig({ ...cfg, classCatalog: catalog });
                const tok = await G.getGraphToken([
                    'https://graph.microsoft.com/User.Read',
                    'https://graph.microsoft.com/Sites.ReadWrite.All'
                ]);
                const site = await G.resolveSiteFromWebUrl(tok, siteUrl);
                const result = await patchFreistellungKlasseColumn(site.id, listId, codes, {
                    webUrl: siteUrl
                });
                if (!result || !result.ok) {
                    throw new Error(
                        (result && result.reason) ||
                            'Spalte „Klasse“ konnte nicht aktualisiert werden.'
                    );
                }
                const verified = Array.isArray(result.verified) ? result.verified : codes;
                toast(
                    'SharePoint-Spalte „Klasse“ ist jetzt: ' +
                        verified.join(', ') +
                        ' (' +
                        verified.length +
                        ', via ' +
                        (result.via || '?') +
                        '). Bitte Spalte in SharePoint neu öffnen (F5).'
                );
            } catch (e) {
                toast(e && e.message ? e.message : String(e));
            } finally {
                syncBtn.disabled = false;
                syncBtn.innerHTML = prev;
            }
        });
    }

    render();
    return { render, list: listStammdatenKlassenForFreistellung };
}

/**
 * @param {{ mountId?: string, syncBtnId?: string, reloadBtnId?: string }} ids
 */
export function htmlFreistellungKlassenPanel(ids) {
    const id = ids || {};
    const mountId = id.mountId || 'frKlassenStammList';
    const syncBtnId = id.syncBtnId || 'frKlassenSyncSp';
    const reloadBtnId = id.reloadBtnId || 'frKlassenReload';
    const cleanBtnId = id.cleanBtnId || 'frKlassenClean';
    return `
    <section class="fr-panel fr-kl-admin" aria-labelledby="frKlAdminTitle">
      <h2 id="frKlAdminTitle"><i class="bi bi-grid-3x3-gap" aria-hidden="true"></i> Klassen aus Stammdaten</h2>
      <p class="muted fr-panel__lead">Eine Quelle: <strong>App-Daten v2</strong> (wie unter Stammdaten → Klassen). Diese Liste wird in die SharePoint-Spalte <strong>Klasse</strong> geschrieben – für das Schüler-Dropdown.</p>
      <div id="${escapeHtml(mountId)}" class="fr-kl-admin__mount"></div>
      <div class="fr-kl-admin__foot">
        <button type="button" class="btn btn-sm alt" id="${escapeHtml(cleanBtnId)}"><i class="bi bi-database-check"></i>Speicher bereinigen</button>
        <button type="button" class="btn btn-sm alt" id="${escapeHtml(reloadBtnId)}"><i class="bi bi-arrow-clockwise"></i>Aus Stammdaten neu laden</button>
        <button type="button" class="btn btn-sm" id="${escapeHtml(syncBtnId)}"><i class="bi bi-arrow-repeat"></i>SharePoint-Spalte „Klasse“ aktualisieren</button>
      </div>
      <p class="muted" style="margin:8px 0 0;font-size:0.85em;">Pflege: <a href="../tenant.html">Stammdaten → Klassen</a>. Zuerst „Speicher bereinigen“, dann SharePoint aktualisieren.</p>
    </section>`;
}
