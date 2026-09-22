/**
 * Phase 5 – Schulen/Lizenzen: Inline-Tabelle mit +-Zeile.
 */
(function () {
    'use strict';

    /** @type {Array<object>} */
    var schoolsCache = [];
    /** @type {string|null} null = keine, '' = neu, id = bearbeiten */
    var editingId = null;
    /** @type {{ key: string, dir: 'asc'|'desc' }} */
    var sortState = { key: 'schoolName', dir: 'asc' };

    function $(id) {
        return document.getElementById(id);
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    async function getToken() {
        if (typeof window.ms365AuthAcquireIdToken === 'function') {
            return window.ms365AuthAcquireIdToken(['https://graph.microsoft.com/User.Read']);
        }
        if (typeof window.ms365AuthAcquireIdTokenPopup === 'function') {
            return window.ms365AuthAcquireIdTokenPopup(['https://graph.microsoft.com/User.Read']);
        }
        if (typeof window.ms365AuthAcquireToken === 'function') {
            return window.ms365AuthAcquireToken(['https://graph.microsoft.com/User.Read']);
        }
        throw new Error('Bitte mit kurt@kurtsoeser.at anmelden (MSAL).');
    }

    function setStatus(msg) {
        var el = $('adminLicStatus');
        if (el) el.textContent = msg || '';
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
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

    function bodyFromEditRow(tr) {
        var domains = parseDomainsInput(tr.querySelector('[data-f="domains"]').value);
        return {
            schoolName: tr.querySelector('[data-f="schoolName"]').value.trim(),
            tenantId: tr.querySelector('[data-f="tenantId"]').value.trim(),
            contactEmail: tr.querySelector('[data-f="contactEmail"]').value.trim(),
            status: tr.querySelector('[data-f="status"]').value,
            validUntil: tr.querySelector('[data-f="validUntil"]').value || null,
            primaryDomain: domains.primaryDomain,
            additionalDomains: domains.additionalDomains,
            notes: tr.querySelector('[data-f="notes"]').value.trim()
        };
    }

    function statusSelectHtml(selected) {
        var opts = ['trial', 'active', 'expired', 'blocked'];
        var cur = String(selected || 'active').toLowerCase();
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
            '<input class="admin-lic-input admin-lic-input--sub" data-f="contactEmail" type="email" placeholder="Kontakt-E-Mail" value="' +
            escapeHtml((school && school.contactEmail) || '') +
            '">' +
            '</td>' +
            '<td><input class="admin-lic-input admin-lic-input--mono" data-f="tenantId" type="text" placeholder="Tenant-ID" value="' +
            escapeHtml((school && school.tenantId) || '') +
            '" maxlength="64" spellcheck="false"></td>' +
            '<td>' +
            statusSelectHtml(school && school.status) +
            '</td>' +
            '<td><input class="admin-lic-input" data-f="validUntil" type="date" value="' +
            escapeHtml((school && school.validUntil) || '') +
            '"></td>' +
            '<td><textarea class="admin-lic-input admin-lic-input--area" data-f="domains" rows="2" placeholder="Domains, eine pro Zeile">' +
            escapeHtml(domainsText(school || {})) +
            '</textarea></td>' +
            '<td><textarea class="admin-lic-input admin-lic-input--area" data-f="notes" rows="2" placeholder="Notizen">' +
            escapeHtml((school && school.notes) || '') +
            '</textarea></td>' +
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
        tr.className = 'admin-lic-row';
        tr.innerHTML =
            '<td><div class="admin-lic-school">' +
            '<span class="admin-lic-school__name">' +
            escapeHtml(s.schoolName || '–') +
            '</span>' +
            (s.contactEmail
                ? '<span class="admin-lic-school__mail">' + escapeHtml(s.contactEmail) + '</span>'
                : '') +
            '</div></td>' +
            '<td><code class="admin-lic-tid">' +
            escapeHtml(s.tenantId || '') +
            '</code></td>' +
            '<td>' +
            statusBadge(s.status) +
            '</td>' +
            '<td>' +
            escapeHtml(s.validUntil || '–') +
            '</td>' +
            '<td>' +
            domainChips(s) +
            '</td>' +
            '<td><div class="admin-lic-notes">' +
            (s.notes ? escapeHtml(s.notes) : '<span class="admin-lic-muted">–</span>') +
            '</div></td>' +
            '<td class="admin-lic-actions"></td>';

        tr.addEventListener('dblclick', function (e) {
            if (e.target && e.target.closest && e.target.closest('button, a, input, select, textarea')) {
                return;
            }
            editingId = String(s.id);
            renderTable(schoolsCache);
        });

        var actions = tr.querySelector('.admin-lic-actions');
        var btnEdit = document.createElement('button');
        btnEdit.type = 'button';
        btnEdit.className = 'btn btn-sm admin-lic-icon-btn';
        btnEdit.innerHTML = '<i class="bi bi-pencil"></i>';
        btnEdit.title = 'Bearbeiten';
        btnEdit.addEventListener('click', function () {
            editingId = String(s.id);
            renderTable(schoolsCache);
        });
        var btnDel = document.createElement('button');
        btnDel.type = 'button';
        btnDel.className = 'btn btn-sm btn-danger admin-lic-icon-btn';
        btnDel.innerHTML = '<i class="bi bi-trash"></i>';
        btnDel.title = 'Löschen';
        btnDel.addEventListener('click', function () {
            removeSchool(s);
        });
        actions.appendChild(btnEdit);
        actions.appendChild(btnDel);
        return tr;
    }

    function sortValue(school, key) {
        if (!school) return '';
        if (key === 'domains') {
            if (Array.isArray(school.domains) && school.domains.length) {
                return school.domains.join(' ').toLowerCase();
            }
            return String(school.primaryDomain || '').toLowerCase();
        }
        if (key === 'schoolName') {
            return String(school.schoolName || school.contactEmail || '').toLowerCase();
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

    function updateSortHeaders() {
        document.querySelectorAll('.admin-lic-sort').forEach(function (btn) {
            var key = btn.getAttribute('data-sort') || '';
            var active = key === sortState.key;
            btn.classList.toggle('is-active', active);
            btn.setAttribute('aria-sort', active ? (sortState.dir === 'asc' ? 'ascending' : 'descending') : 'none');
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
        tbody.replaceChildren();
        updateSortHeaders();
        var list = sortedSchools(schools || []);

        if (editingId === '') {
            tbody.appendChild(buildEditRow(null));
        }

        if (!list.length && editingId !== '') {
            var empty = document.createElement('tr');
            empty.innerHTML =
                '<td colspan="7"><div class="admin-lic-empty">Noch keine Schulen – mit „+ Neue Schule“ anlegen.</div></td>';
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

    async function saveEditRow(tr) {
        var body = bodyFromEditRow(tr);
        if (!body.schoolName || !body.tenantId) {
            toast('Schulname und Tenant-ID sind erforderlich.');
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
                await window.ms365LicenseApi.adminCreateSchool(token, body);
                toast('Angelegt.');
            }
            editingId = null;
            await reload();
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Speichern fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }

    async function reload() {
        setStatus('Lade Schulen …');
        try {
            var token = await getToken();
            var data = await window.ms365LicenseApi.adminListSchools(token);
            schoolsCache = data.schools || [];
            renderTable(schoolsCache);
            setStatus(schoolsCache.length + ' Schule(n)');
        } catch (e) {
            setStatus('Fehler: ' + ((e && e.message) || e));
            toast('Laden fehlgeschlagen: ' + ((e && e.message) || e));
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

    function startAdd() {
        editingId = '';
        renderTable(schoolsCache);
        var first = document.querySelector('.admin-lic-row--edit [data-f="schoolName"]');
        if (first) first.focus();
    }

    function wire() {
        if (!$('adminLicTableBody')) return;
        var btnReload = $('adminLicReloadBtn');
        if (btnReload) btnReload.addEventListener('click', reload);
        var btnAdd = $('adminLicAddBtn');
        if (btnAdd) btnAdd.addEventListener('click', startAdd);
        document.querySelectorAll('.admin-lic-sort').forEach(function (btn) {
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
        window.addEventListener('ms365-auth-state-changed', function () {
            if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
                reload();
            }
        });
        setTimeout(function () {
            if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
                reload();
            } else {
                setStatus('Bitte oben rechts mit dem Betreiber-Konto anmelden.');
            }
        }, 400);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wire);
    } else {
        wire();
    }
})();
