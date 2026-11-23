/**
 * Admin-Tab „Module“: Master-Detail-Katalog mit Werkzeug-Zuordnung.
 */
import {
    LICENSE_PACKAGES,
    allKnownToolIds,
    packageIcon,
    packageLabel
} from './license-feature-catalog.js';
import { toolLabel } from './dashboard-audience-catalog.js';
import { packageStatsHtml } from './admin-license-packages-ui.js';
import { renderSalesBundlesPanel } from './admin-sales-bundles-ui.js';
import { matchSalesBundle } from './sales-bundles-catalog.js';
import { adoptionCounts } from './admin-license-overview-logic.js';

function $(id) {
    return document.getElementById(id);
}

function escapeHtml(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

/** @type {Record<string, string[]>} */
let draftByPackage = {};
/** @type {string|null} */
let selectedPackageId = null;
/** @type {boolean} */
let dirty = false;
let searchQuery = '';

function cloneDefaults() {
    /** @type {Record<string, string[]>} */
    const out = {};
    LICENSE_PACKAGES.forEach((p) => {
        out[p.id] = (p.toolIds || []).slice();
    });
    return out;
}

function packageMeta(id) {
    return LICENSE_PACKAGES.find((p) => p.id === id) || null;
}

function findPackageForTool(toolId) {
    const id = String(toolId || '');
    for (const pkgId of Object.keys(draftByPackage)) {
        if ((draftByPackage[pkgId] || []).indexOf(id) >= 0) return pkgId;
    }
    return null;
}

function setDirty(on) {
    dirty = !!on;
    const save = $('adminModSaveBtn');
    const badge = $('adminModDirtyBadge');
    if (save) save.disabled = !dirty;
    if (badge) badge.hidden = !dirty;
}

function allToolsSorted() {
    return Array.from(allKnownToolIds()).sort((a, b) =>
        toolLabel(a).localeCompare(toolLabel(b), 'de')
    );
}

function matchesSearch(toolId, q) {
    if (!q) return true;
    const hay = (toolLabel(toolId) + ' ' + toolId).toLowerCase();
    return hay.indexOf(q) >= 0;
}

function renderStats() {
    const stats = $('adminModCatalogStats');
    if (!stats) return;
    let toolCount = 0;
    Object.keys(draftByPackage).forEach((id) => {
        toolCount += (draftByPackage[id] || []).length;
    });
    stats.innerHTML = packageStatsHtml(LICENSE_PACKAGES.length, toolCount, {
        packageCaption: 'Pakete',
        toolCaption: 'Zuordnungen',
        hideCatalog: true
    });
}

function renderPackageRail() {
    const host = $('adminModPackageList');
    const headCount = document.querySelector('.admin-mod-rail__head span');
    if (headCount) headCount.textContent = String(LICENSE_PACKAGES.length);
    if (!host) return;
    host.replaceChildren();
    LICENSE_PACKAGES.forEach((pkg) => {
        const count = (draftByPackage[pkg.id] || []).length;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className =
            'admin-mod-rail__item' + (selectedPackageId === pkg.id ? ' is-active' : '');
        btn.setAttribute('aria-current', selectedPackageId === pkg.id ? 'true' : 'false');
        btn.innerHTML =
            '<span class="admin-mod-rail__icon" aria-hidden="true"><i class="bi ' +
            escapeHtml(packageIcon(pkg.id)) +
            '"></i></span>' +
            '<span class="admin-mod-rail__text">' +
            '<strong>' +
            escapeHtml(pkg.label) +
            '</strong>' +
            '<small>' +
            escapeHtml(String(count)) +
            ' Werkzeuge' +
            (pkg.priceHint ? ' · ' + escapeHtml(pkg.priceHint) : '') +
            '</small>' +
            '</span>';
        btn.addEventListener('click', () => selectPackage(pkg.id));
        host.appendChild(btn);
    });
}

function assignTool(toolId, packageId, on) {
    const tid = String(toolId || '');
    const pid = String(packageId || '');
    if (!tid || !pid || !draftByPackage[pid]) return;

    Object.keys(draftByPackage).forEach((key) => {
        draftByPackage[key] = (draftByPackage[key] || []).filter((x) => x !== tid);
    });
    if (on) {
        draftByPackage[pid] = (draftByPackage[pid] || []).concat([tid]);
        draftByPackage[pid].sort((a, b) => toolLabel(a).localeCompare(toolLabel(b), 'de'));
    }
    setDirty(true);
    renderStats();
    renderPackageRail();
    renderDetail();
}

function renderDetail() {
    const empty = $('adminModDetailEmpty');
    const body = $('adminModDetailBody');
    const title = $('adminModDetailTitle');
    const desc = $('adminModDetailDesc');
    const req = $('adminModDetailReq');
    const list = $('adminModToolAssign');
    const countEl = $('adminModDetailCount');
    const pkg = selectedPackageId ? packageMeta(selectedPackageId) : null;

    if (!pkg) {
        if (empty) empty.hidden = false;
        if (body) body.hidden = true;
        return;
    }
    if (empty) empty.hidden = true;
    if (body) body.hidden = false;

    const inPkg = new Set(draftByPackage[pkg.id] || []);
    if (title) title.textContent = pkg.label;
    if (desc) desc.textContent = pkg.description || '';
    const price = $('adminModDetailPrice');
    if (price) {
        if (pkg.priceHint) {
            price.hidden = false;
            price.innerHTML =
                '<i class="bi bi-tag" aria-hidden="true"></i>' +
                escapeHtml(pkg.priceHint) +
                ' <span class="admin-mod-detail__price-note">(nur Admin-Orientierung)</span>';
        } else {
            price.hidden = true;
            price.textContent = '';
        }
    }
    if (req) {
        if (Array.isArray(pkg.requires) && pkg.requires.length) {
            req.hidden = false;
            req.innerHTML =
                '<i class="bi bi-link-45deg" aria-hidden="true"></i>Benötigt ' +
                escapeHtml(pkg.requires.map(packageLabel).join(', '));
        } else {
            req.hidden = true;
            req.textContent = '';
        }
    }
    if (countEl) {
        countEl.textContent = inPkg.size + ' von ' + allKnownToolIds().size + ' Werkzeugen in diesem Paket';
    }

    if (!list) return;
    list.replaceChildren();
    const q = searchQuery.trim().toLowerCase();
    const tools = allToolsSorted()
        .filter((id) => matchesSearch(id, q))
        .sort((a, b) => {
            const aIn = inPkg.has(a) ? 0 : 1;
            const bIn = inPkg.has(b) ? 0 : 1;
            if (aIn !== bIn) return aIn - bIn;
            return toolLabel(a).localeCompare(toolLabel(b), 'de');
        });

    if (!tools.length) {
        const emptyRow = document.createElement('p');
        emptyRow.className = 'admin-mod-assign__empty';
        emptyRow.textContent = 'Keine Werkzeuge für diese Suche.';
        list.appendChild(emptyRow);
        return;
    }

    tools.forEach((tid) => {
        const owner = findPackageForTool(tid);
        const checked = owner === pkg.id;
        const elsewhere = owner && owner !== pkg.id;
        const row = document.createElement('label');
        row.className =
            'admin-mod-assign__row' +
            (checked ? ' is-in' : '') +
            (elsewhere ? ' is-elsewhere' : '');
        row.innerHTML =
            '<input type="checkbox" ' +
            (checked ? 'checked ' : '') +
            'data-tool="' +
            escapeHtml(tid) +
            '">' +
            '<span class="admin-mod-assign__main">' +
            '<strong>' +
            escapeHtml(toolLabel(tid)) +
            '</strong>' +
            '<code>' +
            escapeHtml(tid) +
            '</code>' +
            '</span>' +
            (elsewhere
                ? '<span class="admin-mod-assign__owner">in ' +
                  escapeHtml(packageLabel(owner)) +
                  '</span>'
                : checked
                  ? '<span class="admin-mod-assign__owner is-here">in diesem Paket</span>'
                  : '<span class="admin-mod-assign__owner is-free">nicht zugeordnet</span>');
        const cb = row.querySelector('input[type="checkbox"]');
        if (cb) {
            cb.addEventListener('change', () => {
                if (elsewhere && cb.checked) {
                    const ok = window.confirm(
                        '„' +
                            toolLabel(tid) +
                            '“ liegt aktuell in „' +
                            packageLabel(owner) +
                            '“. Hierher verschieben?'
                    );
                    if (!ok) {
                        cb.checked = false;
                        return;
                    }
                }
                assignTool(tid, pkg.id, !!cb.checked);
            });
        }
        list.appendChild(row);
    });
}

function selectPackage(id) {
    selectedPackageId = id;
    renderPackageRail();
    renderDetail();
}

function resetDraft() {
    draftByPackage = cloneDefaults();
    setDirty(false);
    if (!selectedPackageId) selectedPackageId = LICENSE_PACKAGES[0] ? LICENSE_PACKAGES[0].id : null;
    renderStats();
    renderPackageRail();
    renderDetail();
}

function saveDraftLocal() {
    try {
        sessionStorage.setItem(
            'ms365-admin-license-packages-draft-v1',
            JSON.stringify({
                v: 1,
                packages: Object.keys(draftByPackage).map((id) => ({
                    id,
                    toolIds: draftByPackage[id] || []
                })),
                savedAt: Date.now()
            })
        );
    } catch {
        /* ignore */
    }
    setDirty(false);
    toast(
        'Zuordnung als Entwurf gespeichert. Zentrale SharePoint-Persistenz folgt als Nächstes – bis dahin gilt weiter der Code-Katalog für Schulen.'
    );
}

function loadDraftLocal() {
    try {
        const raw = sessionStorage.getItem('ms365-admin-license-packages-draft-v1');
        if (!raw) return false;
        const data = JSON.parse(raw);
        if (!data || !Array.isArray(data.packages)) return false;
        const next = cloneDefaults();
        data.packages.forEach((p) => {
            if (!p || !p.id || !next[p.id]) return;
            next[p.id] = Array.isArray(p.toolIds) ? p.toolIds.map(String) : [];
        });
        draftByPackage = next;
        setDirty(false);
        return true;
    } catch {
        return false;
    }
}

function renderAdoption(schools) {
    const host = $('adminModAdoption');
    if (!host) return;
    if (!Array.isArray(schools)) {
        host.textContent = 'Paket-Mix erscheint, sobald die Schulliste geladen ist.';
        return;
    }
    const counts = adoptionCounts(schools, LICENSE_PACKAGES, matchSalesBundle);
    if (!counts.withPackages) {
        host.textContent = schools.length
            ? 'Noch keine Schule mit gesetzten Paketen.'
            : 'Paket-Mix erscheint, sobald die Schulliste geladen ist.';
        return;
    }
    host.replaceChildren();
    const title = document.createElement('p');
    title.className = 'admin-mod-adoption__title';
    title.textContent = 'Freigeschaltet an Schulen';
    host.appendChild(title);
    const pkgs = document.createElement('p');
    pkgs.className = 'admin-mod-adoption__line';
    pkgs.textContent = LICENSE_PACKAGES.map((pkg) => pkg.label + ' ' + (counts.byPackage[pkg.id] || 0)).join(' · ');
    host.appendChild(pkgs);
    const bundles = Object.keys(counts.byBundle).map((id) => counts.byBundle[id]);
    const bundleLine = document.createElement('p');
    bundleLine.className = 'admin-mod-adoption__line';
    const bundleBits = bundles.map((bundle) => bundle.label + ' ' + bundle.count);
    if (counts.custom) bundleBits.push('ohne Bündel ' + counts.custom);
    bundleLine.textContent = bundleBits.length ? bundleBits.join(' · ') : 'Kein Verkaufsbündel erkannt.';
    host.appendChild(bundleLine);
}

function renderSalesReference() {
    const host = $('adminModSalesBundles');
    if (!host) return;
    renderSalesBundlesPanel(host, { mode: 'reference' });
}

function renderCatalog() {
    if (!loadDraftLocal()) {
        draftByPackage = cloneDefaults();
        setDirty(false);
    }
    if (!selectedPackageId) {
        selectedPackageId = LICENSE_PACKAGES[0] ? LICENSE_PACKAGES[0].id : null;
    }
    renderSalesReference();
    renderStats();
    renderPackageRail();
    renderDetail();
    renderAdoption(window.ms365AdminSchoolSnapshot);
}

function wire() {
    if (!$('adminPanelModules')) return;

    const search = $('adminModToolSearch');
    if (search && !search.dataset.bound) {
        search.dataset.bound = '1';
        search.addEventListener('input', () => {
            searchQuery = String(search.value || '');
            renderDetail();
        });
    }

    const save = $('adminModSaveBtn');
    if (save && !save.dataset.bound) {
        save.dataset.bound = '1';
        save.addEventListener('click', saveDraftLocal);
    }

    const reset = $('adminModResetBtn');
    if (reset && !reset.dataset.bound) {
        reset.dataset.bound = '1';
        reset.addEventListener('click', () => {
            if (!window.confirm('Zuordnung auf den Standard-Katalog zurücksetzen?')) return;
            try {
                sessionStorage.removeItem('ms365-admin-license-packages-draft-v1');
            } catch {
                /* ignore */
            }
            resetDraft();
            toast('Standard-Katalog wiederhergestellt.');
        });
    }

    renderCatalog();
    window.addEventListener('ms365-admin-schools', (event) => {
        const schools = event.detail && event.detail.schools;
        renderAdoption(Array.isArray(schools) ? schools : []);
    });
    window.addEventListener('hashchange', () => {
        if (String(location.hash || '').replace(/^#/, '').split('/')[0] === 'modules') {
            renderCatalog();
        }
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wire);
} else {
    wire();
}
