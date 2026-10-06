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
        const stdChips = KATEGORIE_CHOICES.map(
            (k) => `<span class="fr-kat-chip fr-kat-chip--std">${escapeHtml(k)}</span>`
        ).join('');
        const extraBlock = extra.length
            ? `<ul class="fr-kat-extra-list">${extra
                  .map(
                      (k, idx) =>
                          `<li class="fr-kat-extra-row"><span class="fr-kat-extra-row__label">${escapeHtml(
                              k
                          )}</span><button type="button" class="fr-kat-extra-row__rm btn btn-sm alt" data-fr-kat-rm="${idx}" title="Entfernen" aria-label="Entfernen"><i class="bi bi-x-lg" aria-hidden="true"></i></button></li>`
                  )
                  .join('')}</ul>`
            : '<p class="fr-kat-extra-empty muted">Noch keine zusätzlichen Kategorien.</p>';
        listEl.innerHTML =
            '<div class="fr-kat-admin__grid">' +
            '<div class="fr-kat-admin__block">' +
            '<span class="fr-kat-admin__block-title">Standard</span>' +
            '<div class="fr-kat-chips" role="list">' +
            stdChips +
            '</div></div>' +
            '<div class="fr-kat-admin__block">' +
            '<span class="fr-kat-admin__block-title">Zusätzlich <span class="fr-kat-admin__count">(' +
            extra.length +
            ')</span></span>' +
            extraBlock +
            '</div></div>' +
            '<p class="fr-kat-admin__meta muted">Im Antragsformular: <strong>' +
            all.length +
            '</strong> Kategorien</p>';
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
        const commitNew = () => {
            const label = String(input.value || '').trim();
            if (!label) return;
            addExtraKategorie(label);
            input.value = '';
            render();
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert('Kategorie „' + label + '“ hinzugefügt (lokal).');
            }
        };
        addBtn.addEventListener('click', commitNew);
        input.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                commitNew();
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
    <section class="fr-panel fr-kat-admin" aria-labelledby="frKatAdminTitle">
      <h2 id="frKatAdminTitle">Antrags-Kategorien</h2>
      <p class="muted fr-panel__lead">Zusätzliche Kategorien für Schüler- und KV-Formulare – mit „Gruppen speichern“ (Setup) auf die Site.</p>
      <div id="${escapeHtml(id.listId || 'frKatExtraList')}" class="fr-kat-admin__mount"></div>
      <div class="fr-kat-admin__add">
        <input type="text" id="${escapeHtml(id.newInputId || 'frKatExtraNew')}" placeholder="Neue Kategorie …" maxlength="120" autocomplete="off">
        <button type="button" class="btn btn-sm" id="${escapeHtml(id.addId || 'frKatExtraAdd')}"><i class="bi bi-plus-lg"></i>Hinzufügen</button>
      </div>
      ${
          id.syncListBtnId
              ? `<div class="fr-kat-admin__foot"><button type="button" class="btn btn-sm" id="${escapeHtml(
                    id.syncListBtnId
                )}"><i class="bi bi-arrow-repeat"></i>SharePoint-Spalte „Kategorie“ aktualisieren</button></div>`
              : ''
      }
    </section>`;
}
