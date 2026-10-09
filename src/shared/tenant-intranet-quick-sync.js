/**
 * Stammdaten: zentrale Intranet-Site-URL (Stammdaten) + Schnell-Buttons in Listen-Tabs.
 */
import {
    DEFAULT_INTRANET_LIST_TITLES,
    INTRANET_LIST_KIND_OPTIONS,
    resolveIntranetListTitle
} from './intranet-list-title-logic.js';
import {
    readIntranetSiteUrl,
    restoreIntranetListTitlesToTenantFields,
    runQuickIntranetSync,
    syncIntranetListTitlesToMounts,
    writeIntranetListTitles,
    writeIntranetSiteUrl
} from './stammdaten-intranet-listen-ui.js';
import { mountTenantItLibraryEmbeddedUi } from './tenant-it-library-embedded-ui.js';
import {
    refreshIntranetListOpenLinks,
    saveIntranetListLink,
    showIntranetSyncToast,
    wrapIntranetSyncButtons
} from './tenant-intranet-list-links.js';

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadSiteUrlIntoField() {
    const input = document.getElementById('tenantIntranetSiteUrl');
    if (!input || String(input.value || '').trim()) return;
    const saved = readIntranetSiteUrl();
    if (saved) input.value = saved;
}

function bindSiteUrlField() {
    const input = document.getElementById('tenantIntranetSiteUrl');
    if (!input || input.dataset.intranetSiteBound === '1') return;
    input.dataset.intranetSiteBound = '1';
    let timer = null;
    input.addEventListener('input', function () {
        if (timer) clearTimeout(timer);
        timer = setTimeout(function () {
            writeIntranetSiteUrl(String(input.value || '').trim());
        }, 400);
    });
    input.addEventListener('change', function () {
        writeIntranetSiteUrl(String(input.value || '').trim());
    });
}

const KIND_TO_INPUT_ID = {
    schueler: 'tenantIntranetListSchueler',
    faecher: 'tenantIntranetListFaecher',
    fachgruppen: 'tenantIntranetListFachgruppen',
    arges: 'tenantIntranetListArges',
    klassen: 'tenantIntranetListKlassen',
    lehrer: 'tenantIntranetListLehrer'
};

function inputForKind(kind) {
    const id = KIND_TO_INPUT_ID[kind];
    return id ? document.getElementById(id) : null;
}

function flushListTitleInput(kind) {
    if (!kind) return;
    const input = inputForKind(kind);
    if (!input) return;
    const v = String(input.value || '').trim();
    writeIntranetListTitles({ [kind]: v });
    syncIntranetListTitlesToMounts();
}

function bindIntranetListTitleRows() {
    INTRANET_LIST_KIND_OPTIONS.forEach(function (opt) {
        const kind = opt.kind;
        const input = inputForKind(kind);
        if (!input || input.dataset.intranetListTitleBound === '1') return;
        input.dataset.intranetListTitleBound = '1';
        const def = DEFAULT_INTRANET_LIST_TITLES[kind] || '';
        if (!input.placeholder && def) input.placeholder = def;
        let timer = null;
        input.addEventListener('input', function () {
            const stored = String(input.value || '').trim();
            const effective = resolveIntranetListTitle(kind, stored);
            input.title = stored ? 'Sync: „' + effective + '“' : 'Standard, wenn leer: „' + def + '“';
            if (timer) clearTimeout(timer);
            timer = setTimeout(function () {
                flushListTitleInput(kind);
            }, 400);
        });
        input.addEventListener('change', function () {
            flushListTitleInput(kind);
        });
    });
}

function bindQuickSyncButtons() {
    document.querySelectorAll('[data-tenant-intranet-sync]').forEach(function (btn) {
        if (btn.dataset.intranetSyncBound === '1') return;
        btn.dataset.intranetSyncBound = '1';
        btn.addEventListener('click', function () {
            const kind = String(btn.getAttribute('data-tenant-intranet-sync') || '').trim();
            const prevLabel = btn.innerHTML;
            btn.disabled = true;
            runQuickIntranetSync(kind)
                .then(function (result) {
                    if (!result) return;
                    if (result.webUrl) {
                        saveIntranetListLink(result.kind, {
                            url: result.webUrl,
                            title: result.listTitle,
                            count: result.count
                        });
                        refreshIntranetListOpenLinks();
                    }
                    showIntranetSyncToast(result.listTitle, result.count, !!result.webUrl);
                })
                .catch(function (e) {
                    toast(e && e.message ? e.message : String(e));
                })
                .finally(function () {
                    btn.disabled = false;
                    btn.innerHTML = prevLabel;
                });
        });
    });
}

export function mountTenantIntranetQuickSync() {
    loadSiteUrlIntoField();
    bindSiteUrlField();
    restoreIntranetListTitlesToTenantFields();
    bindIntranetListTitleRows();
    syncIntranetListTitlesToMounts();
    wrapIntranetSyncButtons();
    bindQuickSyncButtons();
    mountTenantItLibraryEmbeddedUi();
}
