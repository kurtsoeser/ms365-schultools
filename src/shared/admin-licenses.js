/**
 * Schulen/Lizenzen: Übersicht + Schildetail mit Modul-Paketen.
 */
import {
    LICENSE_PACKAGES,
    DEFAULT_NEW_SCHOOL_PACKAGES,
    packageLabel,
    resolveFeatureEntitlements,
    defaultNewSchoolEntitlements,
    applyToolEntitlementToggle,
    pruneToolOverrides
} from './license-feature-catalog.js';
import {
    packageStatsHtml,
    renderPackageGrid,
    renderToolEntitlementGroups
} from './admin-license-packages-ui.js';
import { entitlementsForSalesBundle } from './sales-bundles-catalog.js';
import { renderSalesBundlesPanel } from './admin-sales-bundles-ui.js';
import {
    SCHOOL_STAMMDATEN_FIELDS,
    SCHOOL_STAMMDATEN_GROUPS,
    assertTenantIdAllowed,
    isPendingTenantId,
    isRealTenantGuid
} from './school-stammdaten-catalog.js';

/** @type {Array<object>} */
var schoolsCache = [];
/** @type {Array<object>} */
var columnsCache = [];
/** @type {string|null} null = keine, '' = neu, id = bearbeiten */
var editingId = null;
/** @type {string|null} */
var detailSchoolId = null;
/** @type {{ packages: string[], toolsInclude: string[], toolsExclude: string[] } | null} */
var detailEntitlementsDraft = null;
/** @type {'packages'|'tools'} */
var detailEntitlementView = 'packages';
/** @type {string} */
var detailToolsSearch = '';
/** @type {{ key: string, dir: 'asc'|'desc' }} */
var sortState = { key: 'schoolName', dir: 'asc' };
/** @type {{ q: string, status: string, tenant: ''|'pending'|'set', quick: 'all'|'pending'|'trial'|'active' }} */
var listFilter = { q: '', status: '', tenant: '', quick: 'all' };
/** @type {'stammdaten'|'lizenz'|'module'|'quellen'} */
var detailTab = 'stammdaten';
/** @type {boolean} */
var reloadInFlight = false;
/** @type {boolean|null} */
var lastKnownLoggedIn = null;
var FILTER_STORAGE_KEY = 'ms365-admin-lic-filter-v1';

/** Schlanke Übersicht – Details nur in der Schulseite. */
var CORE_HEADS = [
    { key: 'schoolName', label: 'Schule' },
    { key: 'status', label: 'Status' },
    { key: 'tenantId', label: 'Tenant' },
    { key: 'validUntil', label: 'Gültig bis' },
    { key: 'modules', label: 'Module' }
];

function $(id) {
    return document.getElementById(id);
}

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    async function getToken() {
        var acquire =
            window.ms365LicenseApi && typeof window.ms365LicenseApi.acquireLicenseToken === 'function'
                ? window.ms365LicenseApi.acquireLicenseToken()
                : null;
        if (!acquire) {
            throw new Error('Bitte zuerst mit MS365 anmelden.');
        }
        return Promise.race([
            acquire,
            new Promise(function (_, reject) {
                setTimeout(function () {
                    reject(new Error('Anmeldung/Token dauert zu lange – bitte erneut anmelden.'));
                }, 20000);
            })
        ]);
    }

    function setStatus(msg) {
        var el = $('adminLicStatus');
        if (el) el.textContent = msg || '';
    }

    function escapeHtml(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function stammdatenMetaByName() {
        /** @type {Record<string, (typeof SCHOOL_STAMMDATEN_FIELDS)[number]>} */
        var map = {};
        SCHOOL_STAMMDATEN_FIELDS.forEach(function (f) {
            map[String(f.name).toLowerCase()] = f;
        });
        return map;
    }

    /** Übersicht zeigt keine Extra-Spalten mehr – nur Kernfelder. */
    function tableExtraColumns() {
        return [];
    }

    function colCount() {
        return CORE_HEADS.length + 1;
    }

    function packageCountHtml(school) {
        var pkgs = resolvedPackages(school);
        var n = pkgs.length;
        if (!n) return '<span class="admin-lic-muted">–</span>';
        var legacy = !schoolEntitlementDoc(school);
        return (
            '<span class="admin-lic-count" title="' +
            escapeHtml(pkgs.map(packageLabel).join(', ')) +
            '">' +
            escapeHtml(String(n)) +
            ' Paket' +
            (n === 1 ? '' : 'e') +
            (legacy ? ' · Legacy' : '') +
            '</span>'
        );
    }

    function tenantIdCellHtml(tenantId) {
        var tid = String(tenantId || '').trim();
        if (!tid) {
            return '<span class="admin-lic-muted">–</span>';
        }
        if (isPendingTenantId(tid) && !isRealTenantGuid(tid)) {
            return (
                '<span class="admin-lic-tenant-open" title="' +
                escapeHtml(tid) +
                '">offen</span>'
            );
        }
        var short =
            tid.length > 12 ? tid.slice(0, 8) + '…' + tid.slice(-4) : tid;
        return (
            '<code class="admin-lic-tid" title="' +
            escapeHtml(tid) +
            '">' +
            escapeHtml(short) +
            '</code>'
        );
    }

    function statusBadge(status) {
        var s = String(status || '').toLowerCase();
        return (
            '<span class="admin-lic-badge admin-lic-badge--' +
            escapeHtml(s || 'unknown') +
            '">' +
            escapeHtml(s || '–') +
            '</span>'
        );
    }

    function domainsText(school) {
        if (Array.isArray(school.domains) && school.domains.length) {
            return school.domains.join('\n');
        }
        var parts = [];
        if (school.primaryDomain) parts.push(school.primaryDomain);
        if (school.additionalDomains) parts.push(String(school.additionalDomains));
        return parts.join('\n');
    }

    function schoolEntitlementDoc(school) {
        if (school && school.entitlementsDoc && typeof school.entitlementsDoc === 'object') {
            return {
                packages: Array.isArray(school.entitlementsDoc.packages)
                    ? school.entitlementsDoc.packages.slice()
                    : [],
                toolsInclude: Array.isArray(school.entitlementsDoc.toolsInclude)
                    ? school.entitlementsDoc.toolsInclude.slice()
                    : [],
                toolsExclude: Array.isArray(school.entitlementsDoc.toolsExclude)
                    ? school.entitlementsDoc.toolsExclude.slice()
                    : []
            };
        }
        if (school && school.entitlements && Array.isArray(school.entitlements.packages)) {
            return {
                packages: school.entitlements.packages.slice(),
                toolsInclude: Array.isArray(school.entitlements.toolsInclude)
                    ? school.entitlements.toolsInclude.slice()
                    : [],
                toolsExclude: Array.isArray(school.entitlements.toolsExclude)
                    ? school.entitlements.toolsExclude.slice()
                    : []
            };
        }
        return null;
    }

    function resolvedPackages(school) {
        var resolved = resolveFeatureEntitlements(schoolEntitlementDoc(school), {
            legacyFullAccess: true
        });
        return resolved.packages || [];
    }

    function packageChipsHtml(school) {
        var pkgs = resolvedPackages(school);
        var legacy = !schoolEntitlementDoc(school);
        if (!pkgs.length) {
            return '<span class="admin-lic-muted">–</span>';
        }
        var shown = pkgs.slice(0, 4);
        var html =
            '<div class="admin-lic-chips admin-lic-chips--packages">' +
            shown
                .map(function (id) {
                    return (
                        '<span class="admin-lic-chip admin-lic-chip--pkg" title="' +
                        escapeHtml(packageLabel(id)) +
                        '">' +
                        escapeHtml(id) +
                        '</span>'
                    );
                })
                .join('');
        if (pkgs.length > 4) {
            html +=
                '<span class="admin-lic-chip">+' + escapeHtml(String(pkgs.length - 4)) + '</span>';
        }
        if (legacy) {
            html += '<span class="admin-lic-chip admin-lic-chip--legacy" title="Noch kein Modul-JSON – Legacy: alle Pakete">alle</span>';
        }
        html += '</div>';
        return html;
    }

    function domainChips(school) {
        var list =
            Array.isArray(school.domains) && school.domains.length
                ? school.domains
                : school.primaryDomain
                  ? [school.primaryDomain]
                  : [];
        if (!list.length) return '<span class="admin-lic-muted">–</span>';
        return (
            '<div class="admin-lic-chips">' +
            list
                .map(function (d) {
                    return '<span class="admin-lic-chip">' + escapeHtml(d) + '</span>';
                })
                .join('') +
            '</div>'
        );
    }

    function parseDomainsInput(raw) {
        var lines = String(raw || '')
            .split(/[\n,;]+/)
            .map(function (s) {
                return s.trim();
            })
            .filter(Boolean);
        return {
            primaryDomain: lines[0] || '',
            additionalDomains: lines.slice(1).join('\n')
        };
    }

    function extraOf(school) {
        return (school && school.extra && typeof school.extra === 'object' && school.extra) || {};
    }

    function formatExtraView(col, value) {
        if (value == null || value === '') return '<span class="admin-lic-muted">–</span>';
        if (col.type === 'boolean') return value ? 'Ja' : 'Nein';
        return escapeHtml(String(value));
    }

    function extraInputHtml(col, value) {
        var name = escapeHtml(col.name);
        var v = value == null ? '' : value;
        if (col.type === 'multiline') {
            return (
                '<textarea class="admin-lic-input admin-lic-input--area" data-extra="' +
                name +
                '" rows="2">' +
                escapeHtml(v) +
                '</textarea>'
            );
        }
        if (col.type === 'number') {
            return (
                '<input class="admin-lic-input" data-extra="' +
                name +
                '" type="number" step="any" value="' +
                escapeHtml(v) +
                '">'
            );
        }
        if (col.type === 'date') {
            return (
                '<input class="admin-lic-input" data-extra="' +
                name +
                '" type="date" value="' +
                escapeHtml(v) +
                '">'
            );
        }
        if (col.type === 'boolean') {
            return (
                '<label class="admin-lic-check"><input type="checkbox" data-extra="' +
                name +
                '"' +
                (v ? ' checked' : '') +
                '> Ja</label>'
            );
        }
        if (col.type === 'choice') {
            var opts = Array.isArray(col.choices) ? col.choices : [];
            return (
                '<select class="admin-lic-input" data-extra="' +
                name +
                '">' +
                '<option value="">–</option>' +
                opts
                    .map(function (o) {
                        var s = String(o);
                        return (
                            '<option value="' +
                            escapeHtml(s) +
                            '"' +
                            (String(v) === s ? ' selected' : '') +
                            '>' +
                            escapeHtml(s) +
                            '</option>'
                        );
                    })
                    .join('') +
                '</select>'
            );
        }
        return (
            '<input class="admin-lic-input" data-extra="' +
            name +
            '" type="text" maxlength="255" value="' +
            escapeHtml(v) +
            '">'
        );
    }

    function bodyFromEditRow(tr) {
        var domains = parseDomainsInput(tr.querySelector('[data-f="domains"]').value);
        var extra = {};
        tr.querySelectorAll('[data-extra]').forEach(function (el) {
            var key = el.getAttribute('data-extra');
            if (!key) return;
            if (el.type === 'checkbox') {
                extra[key] = !!el.checked;
            } else if (el.type === 'number') {
                var raw = String(el.value || '').trim();
                extra[key] = raw === '' ? null : Number(raw);
            } else {
                extra[key] = String(el.value || '').trim();
            }
        });
        return {
            schoolName: tr.querySelector('[data-f="schoolName"]').value.trim(),
            tenantId: tr.querySelector('[data-f="tenantId"]').value.trim(),
            contactEmail: tr.querySelector('[data-f="contactEmail"]').value.trim(),
            status: tr.querySelector('[data-f="status"]').value,
            validUntil: tr.querySelector('[data-f="validUntil"]').value || null,
            primaryDomain: domains.primaryDomain,
            additionalDomains: domains.additionalDomains,
            notes: tr.querySelector('[data-f="notes"]').value.trim(),
            extra: extra
        };
    }

    function statusSelectHtml(selected, fallback) {
        var opts = ['trial', 'active', 'expired', 'blocked'];
        var cur = String(selected || fallback || 'trial').toLowerCase();
        return (
            '<select class="admin-lic-input" data-f="status">' +
            opts
                .map(function (o) {
                    return (
                        '<option value="' +
                        o +
                        '"' +
                        (o === cur ? ' selected' : '') +
                        '>' +
                        o +
                        '</option>'
                    );
                })
                .join('') +
            '</select>'
        );
    }

    function schoolHasPendingTenant(school) {
        return !!(school && isPendingTenantId(school.tenantId) && !isRealTenantGuid(school.tenantId));
    }

    function schoolMetaLine(school) {
        var ex = extraOf(school);
        var bits = [];
        var ort = String(ex.Ort || ex.Standort || '').trim();
        var domain = String(
            school.primaryDomain ||
                (Array.isArray(school.domains) && school.domains[0]) ||
                ''
        ).trim();
        if (ort) bits.push(ort);
        if (domain) bits.push(domain);
        return bits.join(' · ');
    }

    function schoolMetaLineDetail(school) {
        var ex = extraOf(school);
        var bits = [];
        var ort = String(ex.Ort || ex.Standort || '').trim();
        var tel = String(ex.Telefon || '').trim();
        var nr = String(ex.Schulnummer || '').trim();
        if (ort) bits.push(ort);
        if (tel) bits.push(tel);
        if (nr) bits.push('Nr. ' + nr);
        return bits.join(' · ');
    }

    function loadStoredFilter() {
        try {
            var raw = localStorage.getItem(FILTER_STORAGE_KEY);
            if (!raw) return;
            var parsed = JSON.parse(raw);
            if (!parsed || typeof parsed !== 'object') return;
            listFilter.q = String(parsed.q || '');
            listFilter.status = String(parsed.status || '');
            listFilter.tenant = String(parsed.tenant || '');
            listFilter.quick = String(parsed.quick || 'all');
            if (
                listFilter.quick !== 'all' &&
                listFilter.quick !== 'pending' &&
                listFilter.quick !== 'trial' &&
                listFilter.quick !== 'active'
            ) {
                listFilter.quick = 'all';
            }
        } catch (e) {
            /* ignore */
        }
    }

    function saveStoredFilter() {
        try {
            localStorage.setItem(
                FILTER_STORAGE_KEY,
                JSON.stringify({
                    q: listFilter.q,
                    status: listFilter.status,
                    tenant: listFilter.tenant,
                    quick: listFilter.quick
                })
            );
        } catch (e) {
            /* ignore */
        }
    }

    function syncFilterControlsFromState() {
        var searchEl = $('adminLicSearch');
        var statusEl = $('adminLicStatusFilter');
        var tenantEl = $('adminLicTenantFilter');
        if (searchEl && searchEl.value !== listFilter.q) searchEl.value = listFilter.q;
        if (statusEl) statusEl.value = listFilter.status || '';
        if (tenantEl) tenantEl.value = listFilter.tenant || '';
        document.querySelectorAll('[data-quick-filter]').forEach(function (btn) {
            var id = btn.getAttribute('data-quick-filter') || 'all';
            btn.classList.toggle('is-active', id === (listFilter.quick || 'all'));
        });
    }

    function applyQuickFilter(quick) {
        listFilter.quick = quick || 'all';
        if (listFilter.quick === 'pending') {
            listFilter.tenant = 'pending';
            listFilter.status = '';
        } else if (listFilter.quick === 'trial') {
            listFilter.status = 'trial';
            listFilter.tenant = '';
        } else if (listFilter.quick === 'active') {
            listFilter.status = 'active';
            listFilter.tenant = '';
        } else {
            listFilter.status = '';
            listFilter.tenant = '';
        }
        syncFilterControlsFromState();
        saveStoredFilter();
        renderTable(schoolsCache);
    }

    function resetFilters() {
        listFilter = { q: '', status: '', tenant: '', quick: 'all' };
        syncFilterControlsFromState();
        saveStoredFilter();
        renderTable(schoolsCache);
    }

    function schoolMatchesFilter(school) {
        var statusWant = String(listFilter.status || '')
            .trim()
            .toLowerCase();
        if (statusWant && String(school.status || '').toLowerCase() !== statusWant) return false;
        if (listFilter.tenant === 'pending' && !schoolHasPendingTenant(school)) return false;
        if (listFilter.tenant === 'set' && schoolHasPendingTenant(school)) return false;
        var q = String(listFilter.q || '')
            .trim()
            .toLowerCase();
        if (!q) return true;
        var ex = extraOf(school);
        var hay = [
            school.schoolName,
            school.tenantId,
            school.status,
            school.notes,
            school.primaryDomain,
            schoolMetaLine(school),
            Array.isArray(school.contactEmails) ? school.contactEmails.join(' ') : school.contactEmail,
            Array.isArray(school.domains) ? school.domains.join(' ') : '',
            ex.Direktion,
            ex.ItM365Admin,
            ex.Schuladministration,
            ex.Datenstatus
        ]
            .filter(Boolean)
            .join(' ')
            .toLowerCase();
        return hay.indexOf(q) >= 0;
    }

    function filteredSchools(schools) {
        return (schools || []).filter(schoolMatchesFilter);
    }

    function updateQuickFilterCounts() {
        var pending = 0;
        var trial = 0;
        var active = 0;
        schoolsCache.forEach(function (s) {
            if (schoolHasPendingTenant(s)) pending += 1;
            var st = String(s.status || '').toLowerCase();
            if (st === 'trial') trial += 1;
            if (st === 'active') active += 1;
        });
        var map = { pending: pending, trial: trial, active: active };
        Object.keys(map).forEach(function (k) {
            document.querySelectorAll('[data-qf-count="' + k + '"]').forEach(function (el) {
                el.textContent = String(map[k]);
            });
        });
    }

    function updateFilterMeta() {
        var el = $('adminLicFilterMeta');
        if (!el) return;
        var total = schoolsCache.length;
        var shown = filteredSchools(schoolsCache).length;
        var pending = schoolsCache.filter(schoolHasPendingTenant).length;
        var parts = [];
        if (listFilter.q || listFilter.status || listFilter.tenant || listFilter.quick !== 'all') {
            parts.push(shown + ' von ' + total);
        } else {
            parts.push(total + ' Schulen');
        }
        if (pending) parts.push(pending + ' Tenant offen');
        el.textContent = parts.join(' · ');
        updateQuickFilterCounts();
        syncFilterControlsFromState();
    }

    function extraCellsHtml(school, editing) {
        var ex = extraOf(school);
        return tableExtraColumns()
            .map(function (col) {
                if (editing) {
                    return '<td>' + extraInputHtml(col, ex[col.name]) + '</td>';
                }
                return '<td>' + formatExtraView(col, ex[col.name]) + '</td>';
            })
            .join('');
    }

    function buildEditRow(school) {
        var isNew = !school || !school.id;
        var tr = document.createElement('tr');
        tr.className = 'admin-lic-row admin-lic-row--edit';
        tr.dataset.editId = isNew ? '' : String(school.id);
        tr.innerHTML =
            '<td>' +
            '<input class="admin-lic-input" data-f="schoolName" type="text" placeholder="Schulname" value="' +
            escapeHtml((school && school.schoolName) || '') +
            '" maxlength="200">' +
            '<input class="admin-lic-input admin-lic-input--sub" data-f="domains" type="text" placeholder="Domain (z. B. schule.at)" value="' +
            escapeHtml(domainsText(school || {}).split('\n')[0] || '') +
            '">' +
            '<input type="hidden" data-f="contactEmail" value="' +
            escapeHtml(
                (school && school.contactEmails && school.contactEmails.join('\n')) ||
                    (school && school.contactEmail) ||
                    ''
            ) +
            '">' +
            '<input type="hidden" data-f="notes" value="' +
            escapeHtml((school && school.notes) || '') +
            '">' +
            '</td>' +
            '<td>' +
            statusSelectHtml(school && school.status, isNew ? 'trial' : 'active') +
            '</td>' +
            '<td><input class="admin-lic-input admin-lic-input--mono" data-f="tenantId" type="text" placeholder="GUID oder leer → PENDING" value="' +
            escapeHtml((school && school.tenantId) || '') +
            '" maxlength="64" spellcheck="false"></td>' +
            '<td><input class="admin-lic-input" data-f="validUntil" type="date" value="' +
            escapeHtml((school && school.validUntil) || '') +
            '"></td>' +
            '<td><span class="admin-lic-muted">' +
            (isNew ? 'Core + Personen + Schuljahr' : 'in Detail') +
            '</span></td>' +
            '<td class="admin-lic-actions"></td>';

        var actions = tr.querySelector('.admin-lic-actions');
        var btnSave = document.createElement('button');
        btnSave.type = 'button';
        btnSave.className = 'btn btn-sm btn-success admin-lic-icon-btn';
        btnSave.title = 'Speichern';
        btnSave.innerHTML = '<i class="bi bi-check-lg"></i>';
        btnSave.addEventListener('click', function () {
            saveEditRow(tr);
        });
        var btnCancel = document.createElement('button');
        btnCancel.type = 'button';
        btnCancel.className = 'btn btn-sm alt admin-lic-icon-btn';
        btnCancel.title = 'Abbrechen';
        btnCancel.innerHTML = '<i class="bi bi-x-lg"></i>';
        btnCancel.addEventListener('click', function () {
            editingId = null;
            renderTable(schoolsCache);
        });
        actions.appendChild(btnSave);
        actions.appendChild(btnCancel);
        return tr;
    }

    function buildViewRow(s) {
        var tr = document.createElement('tr');
        tr.className =
            'admin-lic-row admin-lic-row--clickable' +
            (schoolHasPendingTenant(s) ? ' admin-lic-row--pending' : '');
        var meta = schoolMetaLine(s);
        tr.innerHTML =
            '<td><div class="admin-lic-school">' +
            '<button type="button" class="admin-lic-school__open" data-open-school="' +
            escapeHtml(String(s.id || '')) +
            '">' +
            '<span class="admin-lic-school__name">' +
            escapeHtml(s.schoolName || '–') +
            '</span>' +
            '</button>' +
            (meta
                ? '<span class="admin-lic-school__meta">' + escapeHtml(meta) + '</span>'
                : '') +
            '</div></td>' +
            '<td>' +
            statusBadge(s.status) +
            '</td>' +
            '<td>' +
            tenantIdCellHtml(s.tenantId) +
            '</td>' +
            '<td>' +
            escapeHtml(s.validUntil || '–') +
            '</td>' +
            '<td>' +
            packageCountHtml(s) +
            '</td>' +
            '<td class="admin-lic-actions"></td>';

        function open() {
            openSchoolDetail(String(s.id));
        }
        var openBtn = tr.querySelector('[data-open-school]');
        if (openBtn) openBtn.addEventListener('click', open);
        tr.addEventListener('click', function (e) {
            if (e.target && e.target.closest && e.target.closest('button, a, input, select, textarea')) {
                return;
            }
            open();
        });

        var actions = tr.querySelector('.admin-lic-actions');
        var btnOpen = document.createElement('button');
        btnOpen.type = 'button';
        btnOpen.className = 'btn btn-sm admin-lic-icon-btn';
        btnOpen.innerHTML = '<i class="bi bi-box-arrow-up-right"></i>';
        btnOpen.title = 'Öffnen';
        btnOpen.addEventListener('click', open);
        var btnDel = document.createElement('button');
        btnDel.type = 'button';
        btnDel.className = 'btn btn-sm btn-danger admin-lic-icon-btn';
        btnDel.innerHTML = '<i class="bi bi-trash"></i>';
        btnDel.title = 'Löschen';
        btnDel.addEventListener('click', function () {
            removeSchool(s);
        });
        actions.appendChild(btnOpen);
        actions.appendChild(btnDel);
        return tr;
    }

    function sortValue(school, key) {
        if (!school) return '';
        if (key.indexOf('extra:') === 0) {
            var name = key.slice(6);
            var v = extraOf(school)[name];
            if (v == null) return '';
            if (typeof v === 'boolean') return v ? '1' : '0';
            return String(v).toLowerCase();
        }
        if (key === 'domains') {
            if (Array.isArray(school.domains) && school.domains.length) {
                return school.domains.join(' ').toLowerCase();
            }
            return String(school.primaryDomain || '').toLowerCase();
        }
        if (key === 'schoolName') {
            return String(school.schoolName || school.contactEmail || '').toLowerCase();
        }
        if (key === 'modules') {
            return resolvedPackages(school).join(' ');
        }
        if (key === 'validUntil') {
            return String(school.validUntil || '');
        }
        return String(school[key] || '').toLowerCase();
    }

    function sortedSchools(list) {
        var rows = (list || []).slice();
        var key = sortState.key;
        var dir = sortState.dir === 'desc' ? -1 : 1;
        rows.sort(function (a, b) {
            var va = sortValue(a, key);
            var vb = sortValue(b, key);
            if (va < vb) return -1 * dir;
            if (va > vb) return 1 * dir;
            return 0;
        });
        return rows;
    }

    function renderHead() {
        var tr = $('adminLicTableHead');
        if (!tr) return;
        tr.replaceChildren();
        CORE_HEADS.forEach(function (h) {
            var th = document.createElement('th');
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'admin-lic-sort';
            btn.setAttribute('data-sort', h.key);
            btn.textContent = h.label;
            th.appendChild(btn);
            tr.appendChild(th);
        });
        tableExtraColumns().forEach(function (col) {
            var th = document.createElement('th');
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'admin-lic-sort';
            btn.setAttribute('data-sort', 'extra:' + col.name);
            btn.textContent = col.displayName || col.name;
            th.appendChild(btn);
            tr.appendChild(th);
        });
        var thAct = document.createElement('th');
        thAct.className = 'admin-lic-table__actions-col';
        tr.appendChild(thAct);
        wireSortButtons();
        updateSortHeaders();
    }

    function updateSortHeaders() {
        document.querySelectorAll('.admin-lic-sort').forEach(function (btn) {
            var key = btn.getAttribute('data-sort') || '';
            var active = key === sortState.key;
            btn.classList.toggle('is-active', active);
            btn.setAttribute(
                'aria-sort',
                active ? (sortState.dir === 'asc' ? 'ascending' : 'descending') : 'none'
            );
            var icon = btn.querySelector('.admin-lic-sort__icon');
            if (!icon) {
                icon = document.createElement('i');
                icon.className = 'bi admin-lic-sort__icon';
                icon.setAttribute('aria-hidden', 'true');
                btn.appendChild(icon);
            }
            icon.className =
                'bi admin-lic-sort__icon ' +
                (active
                    ? sortState.dir === 'asc'
                        ? 'bi-caret-up-fill'
                        : 'bi-caret-down-fill'
                    : 'bi-arrow-down-up');
        });
    }

    function renderTable(schools) {
        var tbody = $('adminLicTableBody');
        if (!tbody) return;
        renderHead();
        tbody.replaceChildren();
        var source = filteredSchools(schools || []);
        var list = sortedSchools(source);
        updateFilterMeta();

        if (editingId === '') {
            tbody.appendChild(buildEditRow(null));
        }

        if (!list.length && editingId !== '') {
            var empty = document.createElement('tr');
            var emptyMsg =
                schoolsCache.length && (listFilter.q || listFilter.status || listFilter.tenant)
                    ? 'Keine Schule entspricht dem Filter.'
                    : 'Noch keine Schulen – mit „+ Neue Schule“ anlegen.';
            empty.innerHTML =
                '<td colspan="' +
                colCount() +
                '"><div class="admin-lic-empty">' +
                escapeHtml(emptyMsg) +
                '</div></td>';
            tbody.appendChild(empty);
            return;
        }

        list.forEach(function (s) {
            if (editingId && String(s.id) === String(editingId)) {
                tbody.appendChild(buildEditRow(s));
            } else {
                tbody.appendChild(buildViewRow(s));
            }
        });
    }

    function typeLabel(type) {
        var map = {
            text: 'Text',
            multiline: 'Mehrzeilig',
            number: 'Zahl',
            date: 'Datum',
            boolean: 'Ja/Nein',
            choice: 'Auswahl'
        };
        return map[type] || type || 'Text';
    }

    function renderColumnsPanel() {
        var list = $('adminLicColumnsList');
        if (!list) return;
        list.replaceChildren();
        if (!columnsCache.length) {
            var empty = document.createElement('li');
            empty.className = 'admin-lic-cols__empty';
            empty.textContent =
                'Noch keine Extra-Spalten – „Schema aktualisieren“ ausführen oder oben anlegen.';
            list.appendChild(empty);
            return;
        }
        var metaMap = stammdatenMetaByName();
        columnsCache.forEach(function (col) {
            var key = String((col && col.name) || '')
                .trim()
                .toLowerCase();
            var stam = metaMap[key];
            var li = document.createElement('li');
            li.className =
                'admin-lic-cols__item' + (stam ? ' admin-lic-cols__item--stammdaten' : '');
            var badges = '';
            if (stam) {
                badges +=
                    '<span class="admin-lic-chip admin-lic-chip--stammdaten">Stammdaten</span>';
                if (stam.tableDefault) {
                    badges +=
                        '<span class="admin-lic-chip admin-lic-chip--table">in Tabelle</span>';
                }
            }
            li.innerHTML =
                '<div class="admin-lic-cols__meta">' +
                '<strong>' +
                escapeHtml(col.displayName || col.name) +
                '</strong>' +
                '<span class="admin-lic-muted">' +
                escapeHtml(typeLabel(col.type)) +
                ' · ' +
                escapeHtml(col.name) +
                '</span>' +
                (badges ? '<div class="admin-lic-cols__badges">' + badges + '</div>' : '') +
                '</div>';
            if (!stam) {
                var btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn btn-sm btn-danger admin-lic-icon-btn';
                btn.title = 'Spalte in SharePoint löschen';
                btn.innerHTML = '<i class="bi bi-trash"></i>';
                btn.addEventListener('click', function () {
                    removeColumn(col);
                });
                li.appendChild(btn);
            } else {
                var lock = document.createElement('span');
                lock.className = 'admin-lic-muted admin-lic-cols__lock';
                lock.title = 'Kuratierte Stammdaten-Spalte – nicht löschen';
                lock.innerHTML = '<i class="bi bi-lock" aria-hidden="true"></i>';
                li.appendChild(lock);
            }
            list.appendChild(li);
        });
    }

    async function ensureSchemaFromPanel() {
        setStatus('Aktualisiere Schema …');
        try {
            var token = await getToken();
            var result = await window.ms365LicenseApi.adminEnsureSchema(token);
            var created =
                result && result.stammdaten && Array.isArray(result.stammdaten.created)
                    ? result.stammdaten.created.slice()
                    : [];
            var extras = [];
            if (result && result.toolSiteUrlColumn && result.toolSiteUrlColumn.created) {
                extras.push('ToolSiteUrl');
            }
            if (result && result.itContactsColumn && result.itContactsColumn.created) {
                extras.push('ItContacts');
            }
            if (result && result.entitlementsColumn && result.entitlementsColumn.created) {
                extras.push('LicenseEntitlements');
            }
            var allNew = created.concat(extras);
            toast(
                allNew.length
                    ? 'Schema aktualisiert – neu: ' + allNew.join(', ')
                    : 'Schema aktuell – inkl. ToolSiteUrl vorhanden.'
            );
            await reload();
            renderColumnsPanel();
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Schema-Update fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }

    function setColumnsPanelOpen(open) {
        var panel = $('adminLicColumnsPanel');
        if (!panel) return;
        panel.hidden = !open;
        if (open) renderColumnsPanel();
    }

    async function saveEditRow(tr) {
        var body = bodyFromEditRow(tr);
        if (!body.schoolName) {
            toast('Schulname ist erforderlich.');
            return;
        }
        try {
            body.tenantId = assertTenantIdAllowed(body.tenantId, body.schoolName);
        } catch (e) {
            toast((e && e.message) || 'Tenant-ID ungültig.');
            return;
        }
        var editId = tr.dataset.editId || '';
        setStatus('Speichere …');
        try {
            var token = await getToken();
            if (editId) {
                await window.ms365LicenseApi.adminUpdateSchool(token, editId, body);
                toast('Aktualisiert.');
            } else {
                if (!body.status) body.status = 'trial';
                body.entitlements = defaultNewSchoolEntitlements();
                await window.ms365LicenseApi.adminCreateSchool(token, body);
                toast(
                    isPendingTenantId(body.tenantId)
                        ? 'Angelegt (Trial, Tenant pending).'
                        : 'Angelegt (Core + Personen + Schuljahr).'
                );
            }
            editingId = null;
            await reload();
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Speichern fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }

    async function reload() {
        if (reloadInFlight) return;
        reloadInFlight = true;
        setStatus('Lade Schulen …');
        try {
            var token = await getToken();
            var data = await window.ms365LicenseApi.adminListSchools(token);
            schoolsCache = data.schools || [];
            columnsCache = data.columns || [];
            if ($('adminLicColumnsPanel') && !$('adminLicColumnsPanel').hidden) {
                renderColumnsPanel();
            }
            var pendingN = schoolsCache.filter(schoolHasPendingTenant).length;
            setStatus(
                schoolsCache.length +
                    ' Schule(n)' +
                    (pendingN ? ' · ' + pendingN + ' Tenant offen' : '') +
                    (columnsCache.length ? ' · ' + columnsCache.length + ' Extra-Spalte(n)' : '')
            );
            updateFilterMeta();
            syncViewFromHash();
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Laden fehlgeschlagen: ' + ((e && e.message) || e));
        } finally {
            reloadInFlight = false;
        }
    }

    function onAuthChanged() {
        var loggedIn =
            typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
        if (loggedIn === lastKnownLoggedIn) return;
        lastKnownLoggedIn = loggedIn;
        if (loggedIn) {
            var start = function () {
                reload();
            };
            if (
                window.ms365OperatorAccess &&
                typeof window.ms365OperatorAccess.refreshOperatorStatus === 'function'
            ) {
                window.ms365OperatorAccess.refreshOperatorStatus().then(start).catch(start);
            } else {
                start();
            }
        } else {
            if (window.ms365OperatorAccess && window.ms365OperatorAccess.clearOperatorCache) {
                window.ms365OperatorAccess.clearOperatorCache();
            }
            schoolsCache = [];
            renderTable(schoolsCache);
            setStatus('Bitte oben rechts mit dem Betreiber-Konto anmelden.');
        }
    }

    async function removeSchool(school) {
        if (!school || !school.id) return;
        if (
            !window.confirm(
                'Schule „' + (school.schoolName || school.tenantId) + '“ wirklich löschen?'
            )
        ) {
            return;
        }
        setStatus('Lösche …');
        try {
            var token = await getToken();
            await window.ms365LicenseApi.adminDeleteSchool(token, school.id);
            toast('Gelöscht.');
            if (editingId && String(editingId) === String(school.id)) editingId = null;
            await reload();
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Löschen fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }

    async function removeColumn(col) {
        if (!col || !col.name) return;
        if (
            !window.confirm(
                'Spalte „' +
                    (col.displayName || col.name) +
                    '“ in SharePoint wirklich löschen? Vorhandene Werte gehen verloren.'
            )
        ) {
            return;
        }
        setStatus('Lösche Spalte …');
        try {
            var token = await getToken();
            await window.ms365LicenseApi.adminDeleteColumn(token, col.name);
            toast('Spalte gelöscht.');
            await reload();
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Spalte löschen fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }

    async function createColumnFromForm(ev) {
        if (ev) ev.preventDefault();
        var displayName = ($('adminLicColDisplayName') && $('adminLicColDisplayName').value.trim()) || '';
        var type = ($('adminLicColType') && $('adminLicColType').value) || 'text';
        var choicesRaw = ($('adminLicColChoices') && $('adminLicColChoices').value) || '';
        if (!displayName) {
            toast('Bitte einen Spaltennamen angeben.');
            return;
        }
        var body = { displayName: displayName, type: type };
        if (type === 'choice') {
            body.choices = choicesRaw
                .split(/[\n,;]+/)
                .map(function (s) {
                    return s.trim();
                })
                .filter(Boolean);
            if (!body.choices.length) {
                toast('Choice braucht mindestens eine Option.');
                return;
            }
        }
        setStatus('Lege Spalte an …');
        try {
            var token = await getToken();
            await window.ms365LicenseApi.adminCreateColumn(token, body);
            toast('Spalte „' + displayName + '“ angelegt.');
            if ($('adminLicColDisplayName')) $('adminLicColDisplayName').value = '';
            if ($('adminLicColChoices')) $('adminLicColChoices').value = '';
            await reload();
            setColumnsPanelOpen(true);
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Spalte anlegen fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }

    function syncChoiceFieldVisibility() {
        var type = ($('adminLicColType') && $('adminLicColType').value) || 'text';
        var wrap = $('adminLicColChoicesWrap');
        if (wrap) wrap.hidden = type !== 'choice';
    }

    function startAdd() {
        editingId = '';
        showListView();
        renderTable(schoolsCache);
        var first = document.querySelector('.admin-lic-row--edit [data-f="schoolName"]');
        if (first) first.focus();
    }

    function parseLicensesHash() {
        try {
            var hash = String(location.hash || '')
                .replace(/^#/, '')
                .trim();
            var m = hash.match(/^licenses\/school\/([^/?#]+)/i);
            if (m) return { view: 'detail', id: decodeURIComponent(m[1]) };
            return { view: 'list', id: null };
        } catch (e) {
            return { view: 'list', id: null };
        }
    }

    function setLicensesHash(schoolId) {
        try {
            var next = schoolId
                ? '#licenses/school/' + encodeURIComponent(schoolId)
                : '#licenses';
            if (String(location.hash || '') !== next) {
                history.pushState(null, '', next);
            }
        } catch (e) {
            /* ignore */
        }
    }

    function setLicensesViewMode(mode) {
        var list = $('adminLicListView');
        var detail = $('adminLicDetailView');
        var isDetail = mode === 'detail';
        if (list) {
            list.hidden = isDetail;
            list.setAttribute('aria-hidden', isDetail ? 'true' : 'false');
        }
        if (detail) {
            detail.hidden = !isDetail;
            detail.setAttribute('aria-hidden', isDetail ? 'false' : 'true');
        }
        var panel = $('adminPanelLicenses');
        if (panel) panel.setAttribute('data-lic-view', isDetail ? 'detail' : 'list');
    }

    function showListView() {
        detailSchoolId = null;
        detailEntitlementsDraft = null;
        setLicensesViewMode('list');
        var title = $('adminAppTitle');
        var subtitle = $('adminAppSubtitle');
        if (title && (!$('adminTabLicenses') || $('adminTabLicenses').getAttribute('aria-selected') === 'true')) {
            title.textContent = 'Schulen & Lizenzen';
        }
        if (subtitle) {
            subtitle.textContent = 'Schulen finden, öffnen und freischalten.';
        }
    }

    function findSchoolById(id) {
        var want = String(id || '');
        for (var i = 0; i < schoolsCache.length; i++) {
            if (String(schoolsCache[i].id) === want) return schoolsCache[i];
        }
        return null;
    }

    function openSchoolDetail(id) {
        var school = findSchoolById(id);
        if (!school) {
            toast('Schule nicht gefunden – Liste aktualisieren.');
            return;
        }
        editingId = null;
        detailSchoolId = String(school.id);
        detailToolsSearch = '';
        var searchEl = $('adminLicDetailToolsSearch');
        if (searchEl) searchEl.value = '';
        var doc = schoolEntitlementDoc(school);
        detailEntitlementsDraft = doc
            ? {
                  packages: doc.packages.slice(),
                  toolsInclude: doc.toolsInclude.slice(),
                  toolsExclude: doc.toolsExclude.slice()
              }
            : defaultNewSchoolEntitlements();
        /* Legacy ohne JSON: alle Pakete anzeigen, speichern schreibt explizit */
        if (!doc) {
            detailEntitlementsDraft = {
                packages: resolvedPackages(school).slice(),
                toolsInclude: [],
                toolsExclude: []
            };
        }
        detailTab = 'stammdaten';
        setLicensesHash(detailSchoolId);
        renderDetailView(school);
    }

    function appendDetailExtraField(host, col, school) {
        var wrap = document.createElement('label');
        var wide = col.type === 'multiline' || col.type === 'note';
        wrap.className =
            'admin-lic-detail__field' + (wide ? ' admin-lic-detail__field--wide' : '');
        wrap.innerHTML =
            '<span>' +
            escapeHtml(col.displayName || col.name) +
            '</span>' +
            extraInputHtml(col, extraOf(school)[col.name]).replace(
                'data-extra="' + escapeHtml(col.name) + '"',
                'data-df-extra="' + escapeHtml(col.name) + '"'
            );
        host.appendChild(wrap);
    }

    function appendStammdatenGroups(host, school, groupIds) {
        if (!host) return;
        host.replaceChildren();
        var byName = {};
        columnsCache.forEach(function (col) {
            byName[String(col.name || '').toLowerCase()] = col;
        });
        var used = {};
        (groupIds || []).forEach(function (groupId) {
            var fields = SCHOOL_STAMMDATEN_FIELDS.filter(function (f) {
                return f.group === groupId && byName[String(f.name).toLowerCase()];
            });
            if (!fields.length) return;
            var section = document.createElement('div');
            section.className = 'admin-lic-detail__group';
            section.innerHTML =
                '<h4 class="admin-lic-detail__group-title">' +
                escapeHtml(SCHOOL_STAMMDATEN_GROUPS[groupId] || groupId) +
                '</h4>';
            var grid = document.createElement('div');
            grid.className = 'admin-lic-detail__group-grid';
            fields.forEach(function (f) {
                var col = byName[String(f.name).toLowerCase()];
                used[String(f.name).toLowerCase()] = true;
                appendDetailExtraField(grid, col, school);
            });
            section.appendChild(grid);
            host.appendChild(section);
        });
        return used;
    }

    function setDetailTab(tab) {
        var allowed = { stammdaten: 1, lizenz: 1, module: 1, quellen: 1 };
        detailTab = allowed[tab] ? tab : 'stammdaten';
        document.querySelectorAll('[data-lic-detail-tab]').forEach(function (btn) {
            var on = btn.getAttribute('data-lic-detail-tab') === detailTab;
            btn.classList.toggle('is-active', on);
            btn.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        document.querySelectorAll('[data-lic-detail-panel]').forEach(function (panel) {
            var on = panel.getAttribute('data-lic-detail-panel') === detailTab;
            panel.hidden = !on;
        });
        if (detailTab === 'module') {
            setDetailEntitlementView(detailEntitlementView);
        }
    }

    function renderDetailForm(school) {
        var hostLizenz = $('adminLicDetailFormLizenz');
        var hostStamm = $('adminLicDetailFormStammdaten');
        var hostQuellen = $('adminLicDetailFormQuellen');
        if (!hostLizenz || !hostStamm || !hostQuellen) return;

        var contact =
            (school.contactEmails && school.contactEmails.length
                ? school.contactEmails.join('\n')
                : school.contactEmail) || '';
        var pendingHint =
            isPendingTenantId(school.tenantId) && !isRealTenantGuid(school.tenantId)
                ? '<p class="admin-lic-detail__hint">Tenant noch offen – sobald die Schul-IT sich anmeldet, erscheint die echte GUID im Konto-Menü bzw. im Lizenz-Overlay.</p>'
                : '';

        hostLizenz.innerHTML =
            '<div class="admin-lic-detail__group-grid">' +
            '<label class="admin-lic-detail__field"><span>Schulname</span>' +
            '<input class="admin-lic-input" data-df="schoolName" type="text" maxlength="200" value="' +
            escapeHtml(school.schoolName || '') +
            '"></label>' +
            '<label class="admin-lic-detail__field"><span>Tenant-ID</span>' +
            '<input class="admin-lic-input admin-lic-input--mono" data-df="tenantId" type="text" maxlength="64" spellcheck="false" placeholder="GUID oder leer → PENDING" value="' +
            escapeHtml(school.tenantId || '') +
            '"></label>' +
            '<label class="admin-lic-detail__field"><span>Status</span>' +
            statusSelectHtml(school.status, 'trial').replace('data-f="status"', 'data-df="status"') +
            '</label>' +
            '<label class="admin-lic-detail__field"><span>Gültig bis</span>' +
            '<input class="admin-lic-input" data-df="validUntil" type="date" value="' +
            escapeHtml(school.validUntil || '') +
            '"></label>' +
            '<label class="admin-lic-detail__field admin-lic-detail__field--wide"><span>IT-Kontakt(e)</span>' +
            '<textarea class="admin-lic-input admin-lic-input--area" data-df="contactEmail" rows="2">' +
            escapeHtml(contact) +
            '</textarea></label>' +
            '<label class="admin-lic-detail__field admin-lic-detail__field--wide"><span>Domains</span>' +
            '<textarea class="admin-lic-input admin-lic-input--area" data-df="domains" rows="2">' +
            escapeHtml(domainsText(school)) +
            '</textarea></label>' +
            '<label class="admin-lic-detail__field admin-lic-detail__field--wide"><span>Tool-Site-URL (Schul-SharePoint)</span>' +
            '<input class="admin-lic-input admin-lic-input--mono" data-df="toolSiteUrl" type="text" maxlength="500" spellcheck="false" placeholder="https://schule.sharepoint.com/sites/MS365-Schultools" value="' +
            escapeHtml(school.toolSiteUrl || '') +
            '">' +
            '<small class="admin-lic-detail__hint">Nur Zeiger – keine Stammdaten-Inhalte. SharePoint-Spalte intern <code>ToolSiteUrl</code> (einzeiliger Text). Danach License-API deployen, sonst bleibt der Wert leer.</small></label>' +
            '<label class="admin-lic-detail__field admin-lic-detail__field--wide"><span>Notizen</span>' +
            '<textarea class="admin-lic-input admin-lic-input--area" data-df="notes" rows="2">' +
            escapeHtml(school.notes || '') +
            '</textarea></label>' +
            pendingHint +
            '</div>';

        var used = appendStammdatenGroups(hostStamm, school, ['adresse', 'rollen', 'kontakt']) || {};
        appendStammdatenGroups(hostQuellen, school, ['meta', 'workshop']);

        var custom = columnsCache.filter(function (col) {
            var key = String(col.name || '').toLowerCase();
            if (used[key]) return false;
            var meta = stammdatenMetaByName()[key];
            if (meta && (meta.group === 'meta' || meta.group === 'workshop')) return false;
            if (meta && (meta.group === 'adresse' || meta.group === 'rollen' || meta.group === 'kontakt')) {
                return false;
            }
            return true;
        });
        if (custom.length) {
            var other = document.createElement('div');
            other.className = 'admin-lic-detail__group';
            other.innerHTML = '<h4 class="admin-lic-detail__group-title">Weitere Spalten</h4>';
            var otherGrid = document.createElement('div');
            otherGrid.className = 'admin-lic-detail__group-grid';
            custom.forEach(function (col) {
                appendDetailExtraField(otherGrid, col, school);
            });
            other.appendChild(otherGrid);
            hostQuellen.appendChild(other);
        }

        setDetailTab(detailTab || 'stammdaten');
    }

    function togglePackage(packageId, on, opts) {
        if (!detailEntitlementsDraft) return;
        var id = String(packageId || '');
        var resetToolOverrides = !!(opts && opts.resetToolOverrides);
        var set = {};
        detailEntitlementsDraft.packages.forEach(function (p) {
            set[p] = true;
        });
        var pkg = LICENSE_PACKAGES.find(function (p) {
            return p.id === id;
        });
        if (on) {
            set[id] = true;
            if (pkg && Array.isArray(pkg.requires)) {
                pkg.requires.forEach(function (req) {
                    set[req] = true;
                });
            }
        } else {
            delete set[id];
            LICENSE_PACKAGES.forEach(function (p) {
                if (Array.isArray(p.requires) && p.requires.indexOf(id) >= 0 && set[p.id]) {
                    delete set[p.id];
                }
            });
        }
        detailEntitlementsDraft.packages = Object.keys(set);
        if (resetToolOverrides && pkg && Array.isArray(pkg.toolIds)) {
            var toolSet = {};
            pkg.toolIds.forEach(function (tid) {
                toolSet[tid] = true;
            });
            if (on) {
                detailEntitlementsDraft.toolsExclude = (
                    detailEntitlementsDraft.toolsExclude || []
                ).filter(function (tid) {
                    return !toolSet[tid];
                });
            } else {
                detailEntitlementsDraft.toolsInclude = (
                    detailEntitlementsDraft.toolsInclude || []
                ).filter(function (tid) {
                    return !toolSet[tid];
                });
            }
        }
        pruneToolOverrides(detailEntitlementsDraft);
        renderDetailEntitlements();
    }

    function togglePackageAllTools(packageId, on) {
        togglePackage(packageId, on, { resetToolOverrides: true });
    }

    function toggleTool(toolId, on) {
        if (!detailEntitlementsDraft) return;
        var result = applyToolEntitlementToggle(detailEntitlementsDraft, toolId, on);
        if (result && result.packagesEnabled && result.packagesEnabled.length) {
            toast(
                'Abhängigkeit mitgeschaltet: ' +
                    result.packagesEnabled.map(packageLabel).join(', ')
            );
        }
        renderDetailEntitlements();
    }

    function setDetailEntitlementView(view) {
        detailEntitlementView = view === 'tools' ? 'tools' : 'packages';
        document.querySelectorAll('[data-lic-entitlement-view]').forEach(function (btn) {
            var active = btn.getAttribute('data-lic-entitlement-view') === detailEntitlementView;
            btn.classList.toggle('is-active', active);
            btn.setAttribute('aria-selected', active ? 'true' : 'false');
        });
        var titleEl = $('adminLicDetailModulesTitle');
        if (titleEl) {
            titleEl.textContent =
                detailEntitlementView === 'tools'
                    ? 'Freigeschaltete Werkzeuge'
                    : 'Freigeschaltete Module';
        }
        var pkgHost = $('adminLicDetailPackages');
        var toolsHost = $('adminLicDetailTools');
        if (pkgHost) pkgHost.hidden = detailEntitlementView !== 'packages';
        if (toolsHost) toolsHost.hidden = detailEntitlementView !== 'tools';
        renderDetailEntitlements();
    }

    function renderDetailEntitlements() {
        var hint = $('adminLicDetailModulesHint');
        var stats = $('adminLicDetailStats');
        var toolsHint = $('adminLicDetailToolsHint');
        if (!detailEntitlementsDraft) return;
        var resolved = resolveFeatureEntitlements(detailEntitlementsDraft, {
            legacyFullAccess: false
        });
        var includeN = (detailEntitlementsDraft.toolsInclude || []).length;
        var excludeN = (detailEntitlementsDraft.toolsExclude || []).length;
        if (hint) {
            hint.textContent =
                resolved.packages.length +
                ' von ' +
                LICENSE_PACKAGES.length +
                ' Paketen aktiv · ' +
                resolved.featureCount +
                ' Werkzeuge freigeschaltet' +
                (includeN || excludeN
                    ? ' · ' + includeN + ' einzeln an / ' + excludeN + ' einzeln aus'
                    : '');
        }
        if (stats) {
            stats.innerHTML = packageStatsHtml(resolved.packages.length, resolved.featureCount);
            var detailSchool = detailSchoolId ? findSchoolById(detailSchoolId) : null;
            var statsWrap = stats.querySelector('.admin-mod-stats');
            if (statsWrap && detailSchool && schoolHasPendingTenant(detailSchool)) {
                statsWrap.insertAdjacentHTML(
                    'beforeend',
                    '<div class="admin-mod-stats__item admin-mod-stats__item--warn"><strong>pending</strong><span>Tenant-ID</span></div>'
                );
            }
        }
        if (toolsHint) {
            toolsHint.textContent =
                includeN || excludeN
                    ? includeN +
                      ' einzeln freigeschaltet, ' +
                      excludeN +
                      ' einzeln gesperrt – unabhängig vom Paket.'
                    : 'Paket-Schalter für alle Werkzeuge, darunter einzeln nachjustieren.';
        }

        if (detailEntitlementView === 'tools') {
            var list = $('adminLicDetailToolsList');
            if (list) {
                renderToolEntitlementGroups(list, {
                    draft: detailEntitlementsDraft,
                    search: detailToolsSearch,
                    onToggle: toggleTool,
                    onPackageToggle: togglePackageAllTools
                });
            }
            return;
        }

        var salesHost = $('adminLicDetailSalesBundles');
        if (salesHost) {
            renderSalesBundlesPanel(salesHost, {
                mode: 'apply',
                activePackageIds: detailEntitlementsDraft.packages,
                onApply: applySalesBundle
            });
        }

        var host = $('adminLicDetailPackages');
        if (!host) return;
        renderPackageGrid(host, {
            interactive: true,
            enabledIds: detailEntitlementsDraft.packages,
            onToggle: togglePackage
        });
    }

    function applySalesBundle(bundleId) {
        var next = entitlementsForSalesBundle(bundleId);
        if (!next || !detailEntitlementsDraft) return;
        var label =
            bundleId === 'starter-it'
                ? 'Starter IT (390 €)'
                : bundleId === 'schulalltag'
                  ? 'Schulalltag (890 €)'
                  : bundleId === 'schule-plus'
                    ? 'Schule Plus (1.690 €)'
                    : bundleId;
        if (
            !window.confirm(
                'Pakete dieser Schule auf „' +
                    label +
                    '“ setzen?\nEinzelne Werkzeug-Overrides werden geleert.'
            )
        ) {
            return;
        }
        detailEntitlementsDraft.packages = next.packages.slice();
        detailEntitlementsDraft.toolsInclude = [];
        detailEntitlementsDraft.toolsExclude = [];
        renderDetailEntitlements();
        toast('Verkaufsbündel „' + label + '“ vorbereitet – Speichern nicht vergessen.');
    }

    function renderDetailPackages() {
        renderDetailEntitlements();
    }

    function renderDetailView(school) {
        setLicensesViewMode('detail');
        var nameEl = $('adminLicDetailSchoolName');
        if (nameEl) nameEl.textContent = school.schoolName || 'Schule';
        var lead = $('adminLicDetailLead');
        if (lead) {
            var bits = [];
            bits.push(school.status ? String(school.status) : '–');
            if (school.validUntil) bits.push('gültig bis ' + school.validUntil);
            var meta = schoolMetaLineDetail(school);
            if (meta) bits.push(meta);
            if (schoolHasPendingTenant(school)) bits.push('Tenant offen');
            else if (school.tenantId) bits.push(String(school.tenantId));
            lead.textContent = bits.join(' · ');
        }
        var eyebrow = document.querySelector('#adminLicDetailHero .admin-lic-detail__eyebrow');
        if (eyebrow) {
            eyebrow.textContent = schoolHasPendingTenant(school)
                ? 'Schullizenz · Tenant pending'
                : 'Schullizenz';
        }
        var title = $('adminAppTitle');
        if (title) title.textContent = school.schoolName || 'Schildetail';
        var subtitle = $('adminAppSubtitle');
        if (subtitle) {
            subtitle.textContent = 'Stammdaten prüfen, Lizenz und Module freischalten.';
        }
        renderDetailForm(school);
        if (detailTab === 'module') {
            setDetailEntitlementView(detailEntitlementView);
        }
    }

    function readDetailBody() {
        var root = $('adminLicDetailView') || document;
        var domainsEl = root.querySelector('[data-df="domains"]');
        var domains = parseDomainsInput(domainsEl ? domainsEl.value : '');
        var extra = {};
        root.querySelectorAll('[data-df-extra]').forEach(function (el) {
            var key = el.getAttribute('data-df-extra');
            if (!key) return;
            if (el.type === 'checkbox') extra[key] = !!el.checked;
            else if (el.type === 'number') {
                var raw = String(el.value || '').trim();
                extra[key] = raw === '' ? null : Number(raw);
            } else extra[key] = String(el.value || '').trim();
        });
        function df(name) {
            var el = root.querySelector('[data-df="' + name + '"]');
            return el ? String(el.value || '').trim() : '';
        }
        return {
            schoolName: df('schoolName'),
            tenantId: df('tenantId'),
            contactEmail: df('contactEmail'),
            status: df('status') || 'trial',
            validUntil: df('validUntil') || null,
            primaryDomain: domains.primaryDomain,
            additionalDomains: domains.additionalDomains,
            notes: df('notes'),
            toolSiteUrl: df('toolSiteUrl'),
            extra: extra,
            entitlements: {
                packages: (detailEntitlementsDraft && detailEntitlementsDraft.packages) || [],
                toolsInclude: (detailEntitlementsDraft && detailEntitlementsDraft.toolsInclude) || [],
                toolsExclude: (detailEntitlementsDraft && detailEntitlementsDraft.toolsExclude) || []
            }
        };
    }

    async function saveDetailView() {
        if (!detailSchoolId) return;
        var body = readDetailBody();
        if (!body.schoolName) {
            toast('Schulname ist erforderlich.');
            return;
        }
        try {
            body.tenantId = assertTenantIdAllowed(body.tenantId, body.schoolName);
        } catch (e) {
            toast((e && e.message) || 'Tenant-ID ungültig.');
            return;
        }
        var wantToolUrl = String(body.toolSiteUrl || '').trim();
        setStatus('Speichere Schildetail …');
        try {
            var token = await getToken();
            if (wantToolUrl && typeof window.ms365LicenseApi.adminEnsureSchema === 'function') {
                try {
                    await window.ms365LicenseApi.adminEnsureSchema(token);
                } catch (schemaErr) {
                    console.warn('[admin-licenses] ensure-schema vor Speichern:', schemaErr);
                }
            }
            var saved = await window.ms365LicenseApi.adminUpdateSchool(token, detailSchoolId, body);
            var savedSchool = saved && saved.school ? saved.school : saved;
            await reload();
            var school = findSchoolById(detailSchoolId) || savedSchool;
            if (school) {
                var doc = schoolEntitlementDoc(school);
                detailEntitlementsDraft = doc
                    ? {
                          packages: doc.packages.slice(),
                          toolsInclude: doc.toolsInclude.slice(),
                          toolsExclude: doc.toolsExclude.slice()
                      }
                    : defaultNewSchoolEntitlements();
                renderDetailView(school);
            }
            var gotToolUrl = String(
                (school && school.toolSiteUrl) ||
                    (savedSchool && savedSchool.toolSiteUrl) ||
                    ''
            ).trim();
            if (wantToolUrl && !gotToolUrl) {
                setStatus('Gespeichert – Tool-Site-URL fehlt in der Antwort.');
                toast(
                    'Tool-Site-URL nicht übernommen. Richtige Liste ist die License-Liste (z. B. „MS365-Schultools-Lizenzen“), Spalte ToolSiteUrl – nicht „Übersicht-Tenants“.'
                );
            } else if (wantToolUrl) {
                toast('Schule gespeichert – Tool-Site-URL gesetzt.');
            } else {
                toast('Schule gespeichert.');
            }
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Speichern fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }

    function syncViewFromHash() {
        var route = parseLicensesHash();
        if (route.view === 'detail' && route.id) {
            if (!schoolsCache.length) return;
            var school = findSchoolById(route.id);
            if (!school) {
                showListView();
                setLicensesHash(null);
                return;
            }
            if (detailSchoolId !== route.id) {
                openSchoolDetail(route.id);
            } else {
                var doc = schoolEntitlementDoc(school);
                detailEntitlementsDraft = doc
                    ? {
                          packages: doc.packages.slice(),
                          toolsInclude: doc.toolsInclude.slice(),
                          toolsExclude: doc.toolsExclude.slice()
                      }
                    : {
                          packages: resolvedPackages(school).slice(),
                          toolsInclude: [],
                          toolsExclude: []
                      };
                renderDetailView(school);
            }
            return;
        }
        showListView();
        renderTable(schoolsCache);
    }

    function wireSortButtons() {
        document.querySelectorAll('#adminLicTableHead .admin-lic-sort').forEach(function (btn) {
            if (btn.dataset.sortWired === '1') return;
            btn.dataset.sortWired = '1';
            btn.addEventListener('click', function () {
                var key = btn.getAttribute('data-sort') || 'schoolName';
                if (sortState.key === key) {
                    sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
                } else {
                    sortState.key = key;
                    sortState.dir = 'asc';
                }
                renderTable(schoolsCache);
            });
        });
    }

    function wire() {
        if (!$('adminLicTableBody')) return;
        // Eventuell hängendes Soft-Gate aus älterer Session entsperren
        try {
            document.documentElement.removeAttribute('data-ms365-admin-boot');
            var stale = document.getElementById('ms365AdminOperatorBlock');
            if (stale) stale.remove();
        } catch (e) {
            /* ignore */
        }
        var btnReload = $('adminLicReloadBtn');
        if (btnReload) btnReload.addEventListener('click', reload);
        var btnMigrate = $('adminLicMigrateBtn');
        if (btnMigrate) {
            btnMigrate.addEventListener('click', async function () {
                if (
                    !window.confirm(
                        'Schulen ohne Modul-JSON jetzt explizit auf alle Pakete setzen? (Legacy-Vollzugriff materialisieren)'
                    )
                ) {
                    return;
                }
                setStatus('Migriere Module …');
                try {
                    var token = await getToken();
                    var result;
                    try {
                        var preview = await window.ms365LicenseApi.adminMigrateEntitlements(token, {
                            mode: 'legacy-full',
                            dryRun: true
                        });
                        var n = (preview && preview.candidates) || 0;
                        if (!n) {
                            toast('Keine Schulen ohne Modul-JSON.');
                            setStatus(schoolsCache.length + ' Schule(n)');
                            return;
                        }
                        if (
                            !window.confirm(
                                n +
                                    ' Schule(n) ohne Modul-JSON gefunden. Jetzt alle Pakete schreiben?'
                            )
                        ) {
                            setStatus(schoolsCache.length + ' Schule(n)');
                            return;
                        }
                        result = await window.ms365LicenseApi.adminMigrateEntitlements(token, {
                            mode: 'legacy-full',
                            dryRun: false
                        });
                    } catch (apiErr) {
                        var msg = String((apiErr && apiErr.message) || apiErr || '');
                        if (!/404|not found|nicht gefunden/i.test(msg)) throw apiErr;
                        /* Fallback: API noch ohne migrate-Endpoint → einzeln per PATCH */
                        toast('Migration-API fehlt (404) – schreibe Pakete einzeln …');
                        var full = {
                            v: 1,
                            packages: LICENSE_PACKAGES.map(function (p) {
                                return p.id;
                            }),
                            toolsInclude: [],
                            toolsExclude: [],
                            notes: 'migration:legacy-full'
                        };
                        var updated = 0;
                        for (var i = 0; i < schoolsCache.length; i++) {
                            var s = schoolsCache[i];
                            if (!s || !s.id || schoolEntitlementDoc(s)) continue;
                            await window.ms365LicenseApi.adminUpdateSchool(token, s.id, {
                                entitlements: full
                            });
                            updated += 1;
                        }
                        result = { updated: updated };
                        if (!updated) {
                            toast(
                                'Keine Schule ohne Modul-JSON – oder die License-API unterstützt Entitlements noch nicht. Bitte license-api neu deployen.'
                            );
                            setStatus(schoolsCache.length + ' Schule(n)');
                            return;
                        }
                    }
                    toast((result && result.updated) + ' Schule(n) aktualisiert.');
                    await reload();
                } catch (e) {
                    setStatus('Fehler: ' + ((e && e.message) || e));
                    toast('Migration fehlgeschlagen: ' + ((e && e.message) || e));
                }
            });
        }
        var btnAdd = $('adminLicAddBtn');
        if (btnAdd) btnAdd.addEventListener('click', startAdd);
        var btnCols = $('adminLicColumnsBtn');
        if (btnCols) {
            btnCols.addEventListener('click', function () {
                var panel = $('adminLicColumnsPanel');
                setColumnsPanelOpen(!(panel && !panel.hidden));
            });
        }
        var btnColsClose = $('adminLicColumnsCloseBtn');
        if (btnColsClose) {
            btnColsClose.addEventListener('click', function () {
                setColumnsPanelOpen(false);
            });
        }
        var btnEnsureSchema = $('adminLicEnsureSchemaBtn');
        if (btnEnsureSchema) {
            btnEnsureSchema.addEventListener('click', function () {
                ensureSchemaFromPanel();
            });
        }
        loadStoredFilter();
        syncFilterControlsFromState();

        var searchEl = $('adminLicSearch');
        if (searchEl && searchEl.dataset.wired !== '1') {
            searchEl.dataset.wired = '1';
            searchEl.addEventListener('input', function () {
                listFilter.q = String(searchEl.value || '');
                saveStoredFilter();
                renderTable(schoolsCache);
            });
        }
        var statusFilterEl = $('adminLicStatusFilter');
        if (statusFilterEl && statusFilterEl.dataset.wired !== '1') {
            statusFilterEl.dataset.wired = '1';
            statusFilterEl.addEventListener('change', function () {
                listFilter.status = String(statusFilterEl.value || '');
                listFilter.quick = 'all';
                saveStoredFilter();
                syncFilterControlsFromState();
                renderTable(schoolsCache);
            });
        }
        var tenantFilterEl = $('adminLicTenantFilter');
        if (tenantFilterEl && tenantFilterEl.dataset.wired !== '1') {
            tenantFilterEl.dataset.wired = '1';
            tenantFilterEl.addEventListener('change', function () {
                listFilter.tenant = String(tenantFilterEl.value || '');
                listFilter.quick = 'all';
                saveStoredFilter();
                syncFilterControlsFromState();
                renderTable(schoolsCache);
            });
        }
        document.querySelectorAll('[data-quick-filter]').forEach(function (btn) {
            if (btn.dataset.wired === '1') return;
            btn.dataset.wired = '1';
            btn.addEventListener('click', function () {
                applyQuickFilter(btn.getAttribute('data-quick-filter') || 'all');
            });
        });
        var resetBtn = $('adminLicFilterReset');
        if (resetBtn && resetBtn.dataset.wired !== '1') {
            resetBtn.dataset.wired = '1';
            resetBtn.addEventListener('click', resetFilters);
        }
        document.querySelectorAll('[data-lic-detail-tab]').forEach(function (btn) {
            if (btn.dataset.tabWired === '1') return;
            btn.dataset.tabWired = '1';
            btn.addEventListener('click', function () {
                setDetailTab(btn.getAttribute('data-lic-detail-tab') || 'stammdaten');
            });
        });
        var form = $('adminLicColumnForm');
        if (form) form.addEventListener('submit', createColumnFromForm);
        var typeSel = $('adminLicColType');
        if (typeSel) typeSel.addEventListener('change', syncChoiceFieldVisibility);
        var btnBack = $('adminLicDetailBack');
        if (btnBack) {
            btnBack.addEventListener('click', function () {
                showListView();
                setLicensesHash(null);
                renderTable(schoolsCache);
            });
        }
        if (!window.__ms365AdminLicEscWired) {
            window.__ms365AdminLicEscWired = true;
            document.addEventListener('keydown', function (e) {
                if (e.key !== 'Escape') return;
                var detail = $('adminLicDetailView');
                if (!detail || detail.hidden) return;
                if (e.target && e.target.closest && e.target.closest('input, textarea, select')) return;
                showListView();
                setLicensesHash(null);
                renderTable(schoolsCache);
            });
        }
        var btnSaveDetail = $('adminLicDetailSave');
        if (btnSaveDetail) btnSaveDetail.addEventListener('click', saveDetailView);
        document.querySelectorAll('[data-lic-entitlement-view]').forEach(function (btn) {
            if (btn.dataset.viewWired === '1') return;
            btn.dataset.viewWired = '1';
            btn.addEventListener('click', function () {
                setDetailEntitlementView(btn.getAttribute('data-lic-entitlement-view') || 'packages');
            });
        });
        var toolsSearch = $('adminLicDetailToolsSearch');
        if (toolsSearch && toolsSearch.dataset.searchWired !== '1') {
            toolsSearch.dataset.searchWired = '1';
            toolsSearch.addEventListener('input', function () {
                detailToolsSearch = String(toolsSearch.value || '');
                if (detailEntitlementView === 'tools') renderDetailEntitlements();
            });
        }
        syncChoiceFieldVisibility();
        setLicensesViewMode('list');
        renderHead();
        renderTable(schoolsCache);
        setStatus('Bitte oben rechts mit dem Betreiber-Konto anmelden.');

        window.addEventListener('hashchange', syncViewFromHash);
        window.addEventListener('ms365-auth-state-changed', onAuthChanged);
        setTimeout(function () {
            lastKnownLoggedIn = null;
            onAuthChanged();
        }, 500);
    }

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wire);
} else {
    wire();
}
