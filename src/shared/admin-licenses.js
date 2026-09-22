/**
 * Schulen/Lizenzen: Inline-Tabelle inkl. Extra-Spalten (SharePoint).
 */
(function () {
    'use strict';

    /** @type {Array<object>} */
    var schoolsCache = [];
    /** @type {Array<object>} */
    var columnsCache = [];
    /** @type {string|null} null = keine, '' = neu, id = bearbeiten */
    var editingId = null;
    /** @type {{ key: string, dir: 'asc'|'desc' }} */
    var sortState = { key: 'schoolName', dir: 'asc' };
    /** @type {boolean} */
    var reloadInFlight = false;
    /** @type {boolean|null} */
    var lastKnownLoggedIn = null;

    var CORE_HEADS = [
        { key: 'schoolName', label: 'Schule / Kontakt' },
        { key: 'tenantId', label: 'Tenant-ID' },
        { key: 'status', label: 'Status' },
        { key: 'validUntil', label: 'Gültig bis' },
        { key: 'domains', label: 'Domains' },
        { key: 'notes', label: 'Notizen' }
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

    function colCount() {
        return CORE_HEADS.length + columnsCache.length + 1;
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

    function extraCellsHtml(school, editing) {
        var ex = extraOf(school);
        return columnsCache
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
            extraCellsHtml(school || {}, true) +
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
            extraCellsHtml(s, false) +
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
        columnsCache.forEach(function (col) {
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
        var list = sortedSchools(schools || []);

        if (editingId === '') {
            tbody.appendChild(buildEditRow(null));
        }

        if (!list.length && editingId !== '') {
            var empty = document.createElement('tr');
            empty.innerHTML =
                '<td colspan="' +
                colCount() +
                '"><div class="admin-lic-empty">Noch keine Schulen – mit „+ Neue Schule“ anlegen.</div></td>';
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
            empty.textContent = 'Noch keine Extra-Spalten – oben anlegen.';
            list.appendChild(empty);
            return;
        }
        columnsCache.forEach(function (col) {
            var li = document.createElement('li');
            li.className = 'admin-lic-cols__item';
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
                '</div>';
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-sm btn-danger admin-lic-icon-btn';
            btn.title = 'Spalte in SharePoint löschen';
            btn.innerHTML = '<i class="bi bi-trash"></i>';
            btn.addEventListener('click', function () {
                removeColumn(col);
            });
            li.appendChild(btn);
            list.appendChild(li);
        });
    }

    function setColumnsPanelOpen(open) {
        var panel = $('adminLicColumnsPanel');
        if (!panel) return;
        panel.hidden = !open;
        if (open) renderColumnsPanel();
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
        if (reloadInFlight) return;
        reloadInFlight = true;
        setStatus('Lade Schulen …');
        try {
            var token = await getToken();
            var data = await window.ms365LicenseApi.adminListSchools(token);
            schoolsCache = data.schools || [];
            columnsCache = data.columns || [];
            renderTable(schoolsCache);
            if ($('adminLicColumnsPanel') && !$('adminLicColumnsPanel').hidden) {
                renderColumnsPanel();
            }
            setStatus(
                schoolsCache.length +
                    ' Schule(n)' +
                    (columnsCache.length ? ' · ' + columnsCache.length + ' Extra-Spalte(n)' : '')
            );
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
        renderTable(schoolsCache);
        var first = document.querySelector('.admin-lic-row--edit [data-f="schoolName"]');
        if (first) first.focus();
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
        var form = $('adminLicColumnForm');
        if (form) form.addEventListener('submit', createColumnFromForm);
        var typeSel = $('adminLicColType');
        if (typeSel) typeSel.addEventListener('change', syncChoiceFieldVisibility);
        syncChoiceFieldVisibility();
        renderHead();
        renderTable(schoolsCache);
        setStatus('Bitte oben rechts mit dem Betreiber-Konto anmelden.');

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
})();
