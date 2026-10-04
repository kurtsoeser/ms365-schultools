/**
 * Admin-UI: zusätzliche Freistellungs-Kategorien (Setup + Planer Direktion).
 */
import {
    mergeKategorieChoices,
    loadExtraKategorien,
    saveExtraKategorien,
    addExtraKategorie,
    removeExtraKategorieAt,
    KATEGORIE_CHOICES
} from './freistellung-planer-kategorien.js';
import { patchFreistellungKategorieColumn } from './freistellung-planer-graph.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {{ listId: string, addId: string, newInputId: string, saveBtnId?: string, syncListBtnId?: string }} spec
 * @param {() => { siteUrl: string, listId: string }} getSiteCtx
 */
export function wireFreistellungKategorienAdmin(spec, getSiteCtx) {
    const listEl = document.getElementById(spec.listId);
    const addBtn = document.getElementById(spec.addId);
    const input = document.getElementById(spec.newInputId);

    function render() {
        if (!listEl) return;
        const extra = loadExtraKategorien();
        const all = mergeKategorieChoices(extra);
        const defaults = new Set(KATEGORIE_CHOICES.map((k) => k.toLowerCase()));
        listEl.innerHTML =
            '<p class="muted" style="margin:0 0 8px;font-size:0.88em;">Standard: ' +
            escapeHtml(KATEGORIE_CHOICES.join(' · ')) +
            '</p>' +
            (extra.length
                ? extra
                      .map(
                          (k, idx) =>
                              `<li class="fr-setup-user-row"><span class="fr-setup-user-row__label">${escapeHtml(
                                  k
                              )}</span><button type="button" class="fr-setup-user-row__rm btn btn-sm alt" data-fr-kat-rm="${idx}" title="Entfernen"><i class="bi bi-x"></i></button></li>`
                      )
                      .join('')
                : '<li class="fr-setup-user-row fr-setup-user-row--empty muted">Noch keine zusätzlichen Kategorien.</li>') +
            '<li class="fr-setup-user-row fr-setup-user-row--empty muted" style="border:none;background:transparent;padding:6px 0 0;">Gesamt im Formular: ' +
            all.length +
            ' Kategorien</li>';
        listEl.querySelectorAll('[data-fr-kat-rm]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const i = Number(btn.getAttribute('data-fr-kat-rm'));
                removeExtraKategorieAt(i);
                render();
            });
        });
    }

    if (addBtn && input && !addBtn.dataset.frKatWired) {
        addBtn.dataset.frKatWired = '1';
        addBtn.addEventListener('click', () => {
            const label = String(input.value || '').trim();
            if (!label) return;
            addExtraKategorie(label);
            input.value = '';
            render();
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert('Kategorie „' + label + '“ hinzugefügt (lokal).');
            }
        });
    }

    const syncBtn = spec.syncListBtnId ? document.getElementById(spec.syncListBtnId) : null;
    if (syncBtn && !syncBtn.dataset.frKatWired) {
        syncBtn.dataset.frKatWired = '1';
        syncBtn.addEventListener('click', () => {
            const ctx = getSiteCtx();
            const siteUrl = String(ctx.siteUrl || '').trim();
            const listId = String(ctx.listId || spec.listIdFallback || '').trim();
            if (!siteUrl || !listId) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert('Site-URL und Listen-ID fehlen.');
                }
                return;
            }
            const G = window.ms365SpoGraph;
            if (!G) return;
            G.getGraphToken([
                'https://graph.microsoft.com/User.Read',
                'https://graph.microsoft.com/Sites.ReadWrite.All'
            ])
                .then((tok) => G.resolveSiteFromWebUrl(tok, siteUrl))
                .then((site) =>
                    patchFreistellungKategorieColumn(site.id, listId, loadExtraKategorien())
                )
                .then(() => {
                    if (typeof window.ms365ToastOrAlert === 'function') {
                        window.ms365ToastOrAlert('SharePoint-Spalte „Kategorie“ aktualisiert.');
                    }
                })
                .catch((e) => {
                    if (typeof window.ms365ToastOrAlert === 'function') {
                        window.ms365ToastOrAlert(e && e.message ? e.message : String(e));
                    }
                });
        });
    }

    render();
    return { render, getMerged: () => mergeKategorieChoices(loadExtraKategorien()) };
}

export function htmlFreistellungKategorienPanel(ids) {
    const id = ids || {};
    return `
    <section class="tm-panel" aria-labelledby="frKatAdminTitle">
      <div class="tm-panel__head">
        <i class="bi bi-tags" aria-hidden="true"></i>
        <div>
          <h3 id="frKatAdminTitle">Antrags-Kategorien</h3>
          <p>Zusätzliche Kategorien für Schüler- und KV-Formulare (werden mit „Gruppen speichern“ auf die Site veröffentlicht).</p>
        </div>
      </div>
      <ul id="${escapeHtml(id.listId || 'frKatExtraList')}" class="fr-setup-user-list" style="max-width:520px;"></ul>
      <div class="fr-setup-egp fr-setup-egp--tight" style="max-width:520px;margin-top:8px;">
        <input type="text" id="${escapeHtml(id.newInputId || 'frKatExtraNew')}" placeholder="Neue Kategorie …" maxlength="120" style="flex:1;min-width:160px;">
        <button type="button" class="btn btn-sm" id="${escapeHtml(id.addId || 'frKatExtraAdd')}"><i class="bi bi-plus-lg"></i>Hinzufügen</button>
      </div>
      ${
          id.syncListBtnId
              ? `<button type="button" class="btn btn-sm" id="${escapeHtml(
                    id.syncListBtnId
                )}" style="margin-top:10px;"><i class="bi bi-arrow-repeat"></i>Kategorie-Spalte in SharePoint aktualisieren</button>`
              : ''
      }
    </section>`;
}
