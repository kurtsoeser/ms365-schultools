/**
 * UI: Lehrkräfte und Schüler:innen anhand Education-Lizenz aus dem Tenant einlesen.
 * Einmal binden pro Seite (Einrichtung Schritt 4/5, Schul-Einstellungen).
 */

import {
    applyStudentImportSelection,
    applyTeacherImportSelection,
    buildStudentImportPreview,
    buildTeacherImportPreview,
    facultyUserPlanSkuIds,
    studentUserPlanSkuIds
} from './graph-licenses.js';
import { escapeHtml, normStr } from './utils/strings.js';

/** Textfilter: alle Wörter müssen in Name, E-Mail, Kürzel/Klasse oder Lizenz vorkommen. */
export function rowMatchesTextFilter(row, query, colKey) {
    const tokens = String(query || '')
        .trim()
        .toLowerCase()
        .split(/\s+/)
        .filter(Boolean);
    if (!tokens.length) return true;
    const hay = [row && row.name, row && row.email, row && colKey ? row[colKey] : '', row && row.licenseLabel]
        .map(function (v) {
            return String(v || '').toLowerCase();
        })
        .join('\n');
    return tokens.every(function (t) {
        return hay.includes(t);
    });
}

function el(id) {
    return document.getElementById(id);
}

function defaultToast(msg) {
    if (typeof window.ms365ShowToast === 'function') window.ms365ShowToast(msg);
    else if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function graphApi() {
    const x = window.ms365GraphUnifiedGroups;
    if (!x) throw new Error('graph-unified-groups.js fehlt.');
    return x;
}

/**
 * @param {object} cfg
 * @param {'teachers'|'students'} [cfg.kind]
 */
export function bindTeacherTenantImport(cfg) {
    return bindLicenseTenantImport(Object.assign({ kind: 'teachers' }, cfg));
}

export function bindStudentTenantImport(cfg) {
    return bindLicenseTenantImport(Object.assign({ kind: 'students' }, cfg));
}

function kindSpec(kind) {
    if (kind === 'students') {
        return {
            skuIds: studentUserPlanSkuIds,
            buildPreview: buildStudentImportPreview,
            apply: applyStudentImportSelection,
            getExisting: function (cfg) {
                return typeof cfg.getExistingStudents === 'function'
                    ? cfg.getExistingStudents()
                    : [];
            },
            familyWho: 'Schüler:innen',
            personWord: 'Schüler-Konto/Konten',
            noneMsg: 'Keine Konten mit A1/A3/A5 für Schüler:innen gefunden.',
            colKey: 'klasse',
            colLabel: 'Klasse',
            emptyHint: 'Keine Treffer für die gewählten Lizenzen (oder alle Konten sind inaktiv).'
        };
    }
    return {
        skuIds: facultyUserPlanSkuIds,
        buildPreview: buildTeacherImportPreview,
        apply: applyTeacherImportSelection,
        getExisting: function (cfg) {
            return typeof cfg.getExistingTeachers === 'function' ? cfg.getExistingTeachers() : [];
        },
        familyWho: 'Lehrpersonal',
        personWord: 'Lehrkraft-Konto/Konten',
        noneMsg: 'Keine Konten mit A1/A3/A5 für Lehrpersonal gefunden.',
        colKey: 'code',
        colLabel: 'Kürzel',
        emptyHint: 'Keine Treffer für die gewählten Lizenzen (oder alle Konten sind inaktiv).'
    };
}

function bindLicenseTenantImport(cfg) {
    const toast = cfg.toast || defaultToast;
    const spec = kindSpec(cfg.kind || 'teachers');
    const state = {
        users: [],
        skuLookup: null,
        preview: [],
        families: { a1: true, a3: true, a5: true }
    };

    function selectedFamilies() {
        const out = [];
        if (state.families.a1) out.push('a1');
        if (state.families.a3) out.push('a3');
        if (state.families.a5) out.push('a5');
        return out;
    }

    function activeOnly() {
        const box = cfg.activeOnlyId ? el(cfg.activeOnlyId) : null;
        return !box || box.checked;
    }

    function filterQuery() {
        const inp = cfg.textFilterId ? el(cfg.textFilterId) : null;
        return inp ? inp.value : '';
    }

    function visibleRows() {
        return state.preview.filter(function (row) {
            return rowMatchesTextFilter(row, filterQuery(), spec.colKey);
        });
    }

    function rebuildPreview() {
        state.preview = spec.buildPreview(state.users, spec.getExisting(cfg) || [], state.skuLookup, {
            activeOnly: activeOnly(),
            guests: false,
            families: selectedFamilies()
        });
        renderPanel();
    }

    function setVisibleSelected(predicate) {
        visibleRows().forEach(function (r) {
            r.selected = predicate(r);
        });
        renderPanel();
    }

    function syncHeaderCheckbox(visible) {
        const box = cfg.selectAllRowsId ? el(cfg.selectAllRowsId) : null;
        if (!box) return;
        const selectable = visible.filter(function (r) {
            return !!r.email;
        });
        const selectedVis = selectable.filter(function (r) {
            return r.selected;
        });
        box.disabled = !selectable.length;
        if (!selectable.length) {
            box.checked = false;
            box.indeterminate = false;
            return;
        }
        box.checked = selectedVis.length === selectable.length;
        box.indeterminate = selectedVis.length > 0 && selectedVis.length < selectable.length;
    }

    function renderSkuFilters() {
        const host = el(cfg.skuFiltersId);
        if (!host) return;
        host.replaceChildren();
        const fams = [
            { key: 'a1', label: 'A1 für ' + spec.familyWho },
            { key: 'a3', label: 'A3 für ' + spec.familyWho },
            { key: 'a5', label: 'A5 für ' + spec.familyWho }
        ];
        fams.forEach(function (f) {
            const lab = document.createElement('label');
            lab.className = 'sw-lic-chip';
            const inp = document.createElement('input');
            inp.type = 'checkbox';
            inp.checked = !!state.families[f.key];
            inp.addEventListener('change', function () {
                state.families[f.key] = inp.checked;
                rebuildPreview();
            });
            lab.appendChild(inp);
            lab.appendChild(document.createTextNode(' ' + f.label));
            host.appendChild(lab);
        });
    }

    function renderPanel() {
        const status = el(cfg.statusId);
        const tbody = el(cfg.tbodyId);
        const applyBtn = el(cfg.applyBtnId);
        const rows = state.preview;
        const q = filterQuery();
        const visible = visibleRows();
        const neu = rows.filter(function (r) {
            return !r.alreadyInList;
        }).length;
        const vorh = rows.length - neu;
        const selected = rows.filter(function (r) {
            return r.selected;
        }).length;
        if (status) {
            if (!state.users.length) {
                status.textContent = '';
            } else {
                let text =
                    rows.length +
                    ' ' +
                    spec.personWord +
                    ' mit gewählter Lizenz · ' +
                    neu +
                    ' neu · ' +
                    vorh +
                    ' bereits in der Liste · ' +
                    selected +
                    ' ausgewählt';
                if (String(q).trim()) {
                    text += ' · ' + visible.length + ' angezeigt';
                }
                status.textContent = text;
            }
        }
        if (!tbody) return;
        tbody.replaceChildren();
        if (!rows.length || !visible.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 6;
            td.style.color = '#6c757d';
            if (!state.users.length) td.textContent = 'Noch nicht eingelesen.';
            else if (!rows.length) td.textContent = spec.emptyHint;
            else td.textContent = 'Keine Treffer für den Textfilter.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            if (applyBtn) applyBtn.disabled = selected === 0;
            syncHeaderCheckbox(visible);
            return;
        }
        visible.forEach(function (row) {
            const tr = document.createElement('tr');
            if (row.alreadyInList) tr.classList.add('sw-lic-row-exists');

            const tdChk = document.createElement('td');
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.checked = !!row.selected;
            chk.title = row.alreadyInList ? 'Bereits in der Liste – Name kann aktualisiert werden' : 'Übernehmen';
            chk.addEventListener('change', function () {
                row.selected = chk.checked;
                renderPanel();
            });
            tdChk.appendChild(chk);

            const tdEdit = document.createElement('td');
            const editInp = document.createElement('input');
            editInp.type = 'text';
            editInp.className = 'cell-editor';
            editInp.value = row[spec.colKey] || '';
            editInp.maxLength = spec.colKey === 'klasse' ? 16 : 12;
            editInp.setAttribute('aria-label', spec.colLabel);
            editInp.placeholder = spec.colKey === 'klasse' ? 'Klasse' : '';
            editInp.addEventListener('input', function () {
                row[spec.colKey] = editInp.value;
            });
            tdEdit.appendChild(editInp);

            const tdName = document.createElement('td');
            tdName.textContent = row.name || '';

            const tdEmail = document.createElement('td');
            tdEmail.textContent = row.email || '–';

            const tdLic = document.createElement('td');
            tdLic.innerHTML = '<span class="sw-lic-pill">' + escapeHtml(row.licenseLabel || '') + '</span>';

            const tdSt = document.createElement('td');
            tdSt.innerHTML = row.alreadyInList
                ? '<span class="sw-lic-status sw-lic-status--exists">in der Liste</span>'
                : '<span class="sw-lic-status sw-lic-status--new">neu</span>';

            tr.appendChild(tdChk);
            tr.appendChild(tdEdit);
            tr.appendChild(tdName);
            tr.appendChild(tdEmail);
            tr.appendChild(tdLic);
            tr.appendChild(tdSt);
            tbody.appendChild(tr);
        });
        if (applyBtn) applyBtn.disabled = selected === 0;
        syncHeaderCheckbox(visible);
    }

    async function runLoad() {
        const btn = el(cfg.btnId);
        const panel = el(cfg.panelId);
        const status = el(cfg.statusId);
        if (btn) {
            btn.disabled = true;
            btn.setAttribute('aria-busy', 'true');
        }
        if (panel) panel.hidden = false;
        if (status) status.textContent = 'Lese Lizenzen und Benutzer aus Microsoft 365 …';
        try {
            const G = graphApi();
            const token = await G.getGraphToken();
            let skuLookup = null;
            if (typeof G.fetchSubscribedSkus === 'function') {
                const sub = await G.fetchSubscribedSkus(token);
                if (sub && sub.ok && typeof G.skuLookupFromSubscribed === 'function') {
                    skuLookup = G.skuLookupFromSubscribed(sub.skus);
                }
            }
            state.skuLookup = skuLookup;
            const skuIds = spec.skuIds();
            let users = [];
            if (typeof G.fetchUsersByAssignedSkuIds === 'function') {
                users = await G.fetchUsersByAssignedSkuIds(token, skuIds, function (n, i, total) {
                    if (status) {
                        status.textContent =
                            'Lese ' + spec.familyWho + ' nach Lizenz … ' + n + ' Konten (Lizenz ' + i + '/' + total + ')';
                    }
                });
            }
            if (!users.length && typeof G.fetchUsersWithAssignedLicenses === 'function') {
                if (status) status.textContent = 'Lese Verzeichnisbenutzer mit Lizenzen …';
                users = await G.fetchUsersWithAssignedLicenses(token, function (n) {
                    if (status) status.textContent = 'Gelesen: ' + n + ' Person(en) …';
                });
            }
            state.users = users;
            rebuildPreview();
            if (!state.preview.length) {
                toast(spec.noneMsg);
            } else {
                toast('Eingelesen: ' + state.preview.length + ' ' + spec.personWord + '.');
            }
        } catch (e) {
            if (status) status.textContent = '';
            toast('Einlesen: ' + (e && e.message ? e.message : e));
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

    function closePanel() {
        const panel = el(cfg.panelId);
        if (panel) panel.hidden = true;
    }

    function applySelected() {
        const selected = state.preview.filter(function (r) {
            return r.selected;
        });
        if (!selected.length) {
            toast('Bitte mindestens eine Person auswählen.');
            return;
        }
        const result = spec.apply(spec.getExisting(cfg) || [], state.preview);
        if (typeof cfg.onApply === 'function') cfg.onApply(result);
        const bits = [];
        if (result.added.length) bits.push(result.added.length + ' neu');
        if (result.updated.length) bits.push(result.updated.length + ' Name(n) aktualisiert');
        toast(
            bits.length
                ? 'Übernommen: ' + bits.join(', ') + (cfg.saveHint ? ' ' + cfg.saveHint : '')
                : 'Keine Änderungen (Auswahl war bereits in der Liste).'
        );
        rebuildPreview();
        closePanel();
    }

    const btn = el(cfg.btnId);
    if (btn) btn.addEventListener('click', function () {
        runLoad();
    });
    const applyBtn = el(cfg.applyBtnId);
    if (applyBtn) {
        applyBtn.disabled = true;
        applyBtn.addEventListener('click', applySelected);
    }
    function bindClose(id) {
        const closeBtn = id ? el(id) : null;
        if (closeBtn) closeBtn.addEventListener('click', closePanel);
    }
    bindClose(cfg.closeBtnId);
    bindClose(cfg.closeFooterBtnId);
    const selNew = el(cfg.selectNewBtnId);
    if (selNew) {
        selNew.addEventListener('click', function () {
            setVisibleSelected(function (r) {
                return !r.alreadyInList && !!r.email;
            });
        });
    }
    const selAll = el(cfg.selectAllBtnId);
    if (selAll) {
        selAll.addEventListener('click', function () {
            setVisibleSelected(function (r) {
                return !!r.email;
            });
        });
    }
    const selNone = el(cfg.selectNoneBtnId);
    if (selNone) {
        selNone.addEventListener('click', function () {
            setVisibleSelected(function () {
                return false;
            });
        });
    }
    const headerChk = cfg.selectAllRowsId ? el(cfg.selectAllRowsId) : null;
    if (headerChk) {
        headerChk.addEventListener('change', function () {
            const on = !!headerChk.checked;
            setVisibleSelected(function (r) {
                return on && !!r.email;
            });
        });
    }
    const textFilter = cfg.textFilterId ? el(cfg.textFilterId) : null;
    if (textFilter) {
        textFilter.addEventListener('input', function () {
            if (state.users.length) renderPanel();
        });
    }
    const activeBox = cfg.activeOnlyId ? el(cfg.activeOnlyId) : null;
    if (activeBox) {
        activeBox.addEventListener('change', function () {
            if (state.users.length) rebuildPreview();
        });
    }
    renderSkuFilters();
    syncHeaderCheckbox([]);
}

function teachersToLines(rows) {
    return (rows || [])
        .map(function (x) {
            return (
                String(x.code || '').trim().toUpperCase() +
                ';' +
                normStr(x.name || '') +
                ';' +
                String(x.email || '').trim().toLowerCase()
            ).trim();
        })
        .filter(Boolean)
        .join('\n');
}

function studentsToLines(rows) {
    return (rows || [])
        .map(function (x) {
            return (
                normStr(x.klasse || '') +
                ';' +
                normStr(x.name || '') +
                ';' +
                String(x.email || '').trim().toLowerCase()
            );
        })
        .filter(function (s) {
            return normStr(s.replace(/;/g, ''));
        })
        .join('\n');
}

function patchDirectoryMatches(matches) {
    if (!matches || !Object.keys(matches).length) return;
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: matches });
        }
    } catch {
        // ignore
    }
}

function applyToTextarea(textareaId, result) {
    const ta = el(textareaId);
    if (!ta) return;
    if (result.students) ta.value = studentsToLines(result.students);
    else ta.value = teachersToLines(result.teachers);
    ta.dispatchEvent(new Event('input', { bubbles: true }));
    patchDirectoryMatches(result.directoryMatches);
}

function autoBind() {
    if (el('swBtnImportTeachersFromTenant')) {
        bindTeacherTenantImport({
            btnId: 'swBtnImportTeachersFromTenant',
            panelId: 'swTeacherTenantImportPanel',
            statusId: 'swTeacherTenantImportStatus',
            skuFiltersId: 'swTeacherTenantImportSkuFilters',
            tbodyId: 'swTeacherTenantImportBody',
            applyBtnId: 'swBtnTeacherTenantApply',
            closeBtnId: 'swBtnTeacherTenantClose',
            closeFooterBtnId: 'swBtnTeacherTenantCloseFooter',
            selectNewBtnId: 'swBtnTeacherTenantSelectNew',
            selectAllBtnId: 'swBtnTeacherTenantSelectAll',
            selectNoneBtnId: 'swBtnTeacherTenantSelectNone',
            selectAllRowsId: 'swTeacherTenantSelectAllRows',
            textFilterId: 'swTeacherTenantImportTextFilter',
            activeOnlyId: 'swTeacherTenantImportActiveOnly',
            getExistingTeachers: function () {
                const ta = el('swTeachersLines');
                if (!ta || typeof window.ms365TenantSettingsParseTeachersLines !== 'function') return [];
                return window.ms365TenantSettingsParseTeachersLines(ta.value);
            },
            onApply: function (result) {
                applyToTextarea('swTeachersLines', result);
            },
            saveHint: 'Bitte noch „Lehrerliste speichern“ wählen.'
        });
    }
    if (el('tenantBtnImportTeachersFromTenant')) {
        bindTeacherTenantImport({
            btnId: 'tenantBtnImportTeachersFromTenant',
            panelId: 'tenantTeacherTenantImportPanel',
            statusId: 'tenantTeacherTenantImportStatus',
            skuFiltersId: 'tenantTeacherTenantImportSkuFilters',
            tbodyId: 'tenantTeacherTenantImportBody',
            applyBtnId: 'tenantBtnTeacherTenantApply',
            closeBtnId: 'tenantBtnTeacherTenantClose',
            closeFooterBtnId: 'tenantBtnTeacherTenantCloseFooter',
            selectNewBtnId: 'tenantBtnTeacherTenantSelectNew',
            selectAllBtnId: 'tenantBtnTeacherTenantSelectAll',
            selectNoneBtnId: 'tenantBtnTeacherTenantSelectNone',
            selectAllRowsId: 'tenantTeacherTenantSelectAllRows',
            textFilterId: 'tenantTeacherTenantImportTextFilter',
            activeOnlyId: 'tenantTeacherTenantImportActiveOnly',
            getExistingTeachers: function () {
                const ta = el('tenantTeachersLines');
                if (!ta || typeof window.ms365TenantSettingsParseTeachersLines !== 'function') return [];
                return window.ms365TenantSettingsParseTeachersLines(ta.value);
            },
            onApply: function (result) {
                applyToTextarea('tenantTeachersLines', result);
            }
        });
    }
    if (el('swBtnImportStudentsFromTenant')) {
        bindStudentTenantImport({
            btnId: 'swBtnImportStudentsFromTenant',
            panelId: 'swStudentTenantImportPanel',
            statusId: 'swStudentTenantImportStatus',
            skuFiltersId: 'swStudentTenantImportSkuFilters',
            tbodyId: 'swStudentTenantImportBody',
            applyBtnId: 'swBtnStudentTenantApply',
            closeBtnId: 'swBtnStudentTenantClose',
            closeFooterBtnId: 'swBtnStudentTenantCloseFooter',
            selectNewBtnId: 'swBtnStudentTenantSelectNew',
            selectAllBtnId: 'swBtnStudentTenantSelectAll',
            selectNoneBtnId: 'swBtnStudentTenantSelectNone',
            selectAllRowsId: 'swStudentTenantSelectAllRows',
            textFilterId: 'swStudentTenantImportTextFilter',
            activeOnlyId: 'swStudentTenantImportActiveOnly',
            getExistingStudents: function () {
                const ta = el('swStudentsLines');
                if (!ta || typeof window.ms365TenantSettingsParseStudentsLines !== 'function') return [];
                return window.ms365TenantSettingsParseStudentsLines(ta.value);
            },
            onApply: function (result) {
                applyToTextarea('swStudentsLines', result);
            },
            saveHint: 'Bitte noch „Schülerliste speichern“ wählen.'
        });
    }
    if (el('tenantBtnImportStudentsFromTenant')) {
        bindStudentTenantImport({
            btnId: 'tenantBtnImportStudentsFromTenant',
            panelId: 'tenantStudentTenantImportPanel',
            statusId: 'tenantStudentTenantImportStatus',
            skuFiltersId: 'tenantStudentTenantImportSkuFilters',
            tbodyId: 'tenantStudentTenantImportBody',
            applyBtnId: 'tenantBtnStudentTenantApply',
            closeBtnId: 'tenantBtnStudentTenantClose',
            closeFooterBtnId: 'tenantBtnStudentTenantCloseFooter',
            selectNewBtnId: 'tenantBtnStudentTenantSelectNew',
            selectAllBtnId: 'tenantBtnStudentTenantSelectAll',
            selectNoneBtnId: 'tenantBtnStudentTenantSelectNone',
            selectAllRowsId: 'tenantStudentTenantSelectAllRows',
            textFilterId: 'tenantStudentTenantImportTextFilter',
            activeOnlyId: 'tenantStudentTenantImportActiveOnly',
            getExistingStudents: function () {
                const ta = el('tenantStudentsLines');
                if (!ta || typeof window.ms365TenantSettingsParseStudentsLines !== 'function') return [];
                return window.ms365TenantSettingsParseStudentsLines(ta.value);
            },
            onApply: function (result) {
                applyToTextarea('tenantStudentsLines', result);
            }
        });
    }
}

if (typeof window !== 'undefined') {
    window.ms365TeacherTenantImport = {
        bindTeacherTenantImport,
        bindStudentTenantImport,
        rowMatchesTextFilter
    };
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', autoBind);
    } else {
        autoBind();
    }
}
