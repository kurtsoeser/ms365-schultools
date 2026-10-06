/**
 * UI: Schüler-Tab „Nach Klassen“ (Phase 4).
 */
import { groupStudentsByClass, describeClassM365Match } from './tenant-klassen-hub-logic.js';

const VIEW_KEY = 'ms365-tenant-students-view-v1';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function readView() {
    try {
        const v = localStorage.getItem(VIEW_KEY);
        if (v === 'liste' || v === 'klassen') return v;
    } catch {
        /* ignore */
    }
    return 'liste';
}

function writeView(v) {
    try {
        localStorage.setItem(VIEW_KEY, v === 'liste' ? 'liste' : 'klassen');
    } catch {
        /* ignore */
    }
}

function applyViewMode(mode) {
    const hub = document.getElementById('tenantKlassenHub');
    const list = document.getElementById('tenantStudentsListWrap');
    const btnK = document.getElementById('tenantStudentsViewKlassen');
    const btnL = document.getElementById('tenantStudentsViewListe');
    const isKlassen = mode !== 'liste';
    if (hub) hub.hidden = !isKlassen;
    if (list) list.hidden = isKlassen;
    if (hub && isKlassen) hub.removeAttribute('hidden');
    if (btnK) btnK.classList.toggle('is-active', isKlassen);
    if (btnL) btnL.classList.toggle('is-active', !isKlassen);
}

function wireViewToggle() {
    const btnK = document.getElementById('tenantStudentsViewKlassen');
    const btnL = document.getElementById('tenantStudentsViewListe');
    if (!btnK || btnK.dataset.bound === '1') return;
    btnK.dataset.bound = '1';
    btnL && (btnL.dataset.bound = '1');
    applyViewMode(readView());
    btnK.addEventListener('click', function () {
        writeView('klassen');
        applyViewMode('klassen');
    });
    if (btnL) {
        btnL.addEventListener('click', function () {
            writeView('liste');
            applyViewMode('liste');
        });
    }
}

/**
 * @param {HTMLElement} mount
 * @param {object} api ms365TenantKlassenHubApi
 */
export function renderTenantKlassenHub(mount, api) {
    if (!mount || !api) return;
    const classes = api.getClasses() || [];
    const students = api.getStudents() || [];
    const grouped = groupStudentsByClass(classes, students, api.normClassCode);
    let selectedKey = mount.dataset.selectedClass || '';

    if (!selectedKey && grouped.byClass.length) {
        selectedKey = grouped.byClass[0].classKey;
    }
    if (selectedKey && !grouped.byClass.some(function (b) { return b.classKey === selectedKey; })) {
        selectedKey = grouped.byClass[0] ? grouped.byClass[0].classKey : '';
    }
    mount.dataset.selectedClass = selectedKey;

    const selected = grouped.byClass.find(function (b) {
        return b.classKey === selectedKey;
    });
    const m365 = selected
        ? describeClassM365Match(api.getClassM365(selected.classRow.code))
        : describeClassM365Match(null);

    let html = '<div class="ts-klassen-hub__layout">';
    html += '<aside class="ts-klassen-hub__classes" aria-label="Klassen">';
    html += '<p class="ts-klassen-hub__aside-title">Klassen <span class="muted">(' + grouped.byClass.length + ')</span></p>';
    html +=
        '<button type="button" class="btn btn-sm ts-klassen-hub__add-class" id="tenantKlassenHubAddClass" title="Neue Klasse in der Klassenliste anlegen"><i class="bi bi-plus-lg" aria-hidden="true"></i> Klasse anlegen</button>';
    html += '<ul class="ts-klassen-hub__class-list">';
    grouped.byClass.forEach(function (b) {
        const n = b.students.length;
        const active = b.classKey === selectedKey ? ' is-active' : '';
        const code = escapeHtml(b.classRow.code);
        html +=
            '<li><button type="button" class="ts-klassen-hub__class-btn' +
            active +
            '" data-class-key="' +
            escapeHtml(b.classKey) +
            '">' +
            '<span class="ts-klassen-hub__class-code">' +
            code +
            '</span>' +
            '<span class="ts-klassen-hub__class-meta">' +
            n +
            ' Schüler</span></button></li>';
    });
    if (!grouped.byClass.length) {
        html += '<li class="muted" style="padding:8px;">Keine Klassen – Tab <strong>Klassen</strong> pflegen.</li>';
    }
    html += '</ul>';
    if (grouped.unassigned.length) {
        html +=
            '<p class="ts-klassen-hub__unassigned muted">' +
            grouped.unassigned.length +
            ' Schüler ohne passende Klasse in der Klassenliste.</p>';
    }
    html += '</aside>';

    html += '<div class="ts-klassen-hub__main">';
    if (selected) {
        const c = selected.classRow;
        html += '<header class="ts-klassen-hub__head">';
        html += '<h3 class="ts-klassen-hub__title">' + escapeHtml(c.code) + (c.name ? ' · ' + escapeHtml(c.name) : '') + '</h3>';
        html += '<span class="ts-status-chip ts-status-chip--' + m365.kind + ' ts-klassen-hub__m365" title="' + escapeHtml(m365.title) + '">' + escapeHtml(m365.label) + '</span>';
        html += '</header>';
        if (c.headName || c.headEmail) {
            html +=
                '<p class="ts-klassen-hub__kv muted">KV: ' +
                escapeHtml(c.headName || '–') +
                (c.headEmail ? ' · ' + escapeHtml(c.headEmail) : '') +
                '</p>';
        }
        html += '<div class="teachers-table-wrap"><table class="teachers-table ts-klassen-hub__table"><thead><tr>';
        html += '<th>Name</th><th>E-Mail</th><th></th></tr></thead><tbody>';
        if (!selected.students.length) {
            html += '<tr><td colspan="3" class="muted">Keine Schüler in dieser Klasse.</td></tr>';
        } else {
            selected.students.forEach(function (item) {
                html += '<tr data-student-index="' + item.index + '">';
                html +=
                    '<td><input type="text" class="ts-klassen-hub__input" data-field="name" value="' +
                    escapeHtml(item.row.name || '') +
                    '" aria-label="Name"></td>';
                html +=
                    '<td><input type="email" class="ts-klassen-hub__input" data-field="email" value="' +
                    escapeHtml(item.row.email || '') +
                    '" aria-label="E-Mail"></td>';
                html +=
                    '<td class="action-cell"><button type="button" class="btn btn-sm" data-action="remove" title="Zeile entfernen"><i class="bi bi-trash"></i></button></td>';
                html += '</tr>';
            });
        }
        html += '</tbody></table></div>';
        html +=
            '<div class="settings-actions" style="margin-top:10px;"><button type="button" class="btn btn-success" id="tenantKlassenHubAddStudent"><i class="bi bi-plus-circle"></i>Schüler in ' +
            escapeHtml(c.code) +
            '</button>';
        html +=
            '<button type="button" class="btn" id="tenantKlassenHubOpenList">Gesamtliste mit Filter</button></div>';
    } else {
        html += '<p class="muted">Wählen Sie links eine Klasse oder legen Sie Klassen im Tab <strong>Klassen</strong> an.</p>';
    }
    html += '</div></div>';

    mount.innerHTML = html;

    const addClassBtn = mount.querySelector('#tenantKlassenHubAddClass');
    if (addClassBtn && typeof api.addClass === 'function') {
        addClassBtn.addEventListener('click', function () {
            const raw =
                typeof window !== 'undefined' && typeof window.prompt === 'function'
                    ? window.prompt('Kürzel der neuen Klasse (z. B. 3CK):', '')
                    : '';
            const code = String(raw || '').trim();
            if (!code) return;
            const nameRaw =
                typeof window.prompt === 'function'
                    ? window.prompt('Optional: Klassenname / Bezeichnung:', code)
                    : code;
            const ok = api.addClass(code, nameRaw != null ? nameRaw : code);
            if (!ok) return;
            mount.dataset.selectedClass = api.normClassCode(code);
            renderTenantKlassenHub(mount, api);
        });
    }

    mount.querySelectorAll('.ts-klassen-hub__class-btn').forEach(function (btn) {
        btn.addEventListener('click', function () {
            mount.dataset.selectedClass = btn.getAttribute('data-class-key') || '';
            renderTenantKlassenHub(mount, api);
        });
    });

    mount.querySelectorAll('.ts-klassen-hub__input').forEach(function (input) {
        input.addEventListener('change', function () {
            const tr = input.closest('tr');
            if (!tr) return;
            const idx = Number(tr.getAttribute('data-student-index'));
            const field = input.getAttribute('data-field');
            if (!field || !Number.isFinite(idx)) return;
            const patch = {};
            patch[field] = field === 'email' ? String(input.value || '').trim().toLowerCase() : String(input.value || '').trim();
            api.updateStudent(idx, patch);
            renderTenantKlassenHub(mount, api);
        });
    });

    mount.querySelectorAll('[data-action="remove"]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const tr = btn.closest('tr');
            const idx = Number(tr && tr.getAttribute('data-student-index'));
            if (!Number.isFinite(idx)) return;
            api.removeStudent(idx);
            renderTenantKlassenHub(mount, api);
        });
    });

    const addBtn = mount.querySelector('#tenantKlassenHubAddStudent');
    if (addBtn) {
        addBtn.addEventListener('click', function () {
            api.addStudent({ klasse: selected.classRow.code, name: '', email: '' });
            renderTenantKlassenHub(mount, api);
        });
    }
    const openList = mount.querySelector('#tenantKlassenHubOpenList');
    if (openList && selected) {
        openList.addEventListener('click', function () {
            api.focusClassInListView(selected.classRow.code);
            writeView('liste');
            applyViewMode('liste');
        });
    }
}

export function initTenantKlassenHub() {
    wireViewToggle();

    function boot() {
        const api = window.ms365TenantKlassenHubApi;
        const mount = document.getElementById('tenantKlassenHub');
        if (!api || !mount) return;
        renderTenantKlassenHub(mount, api);
    }

    window.addEventListener('ms365-tenant-klassen-hub-ready', boot);
    window.addEventListener('ms365-tenant-settings-changed', function (ev) {
        const d = ev && ev.detail;
        if (d && (d.reason === 'render' || d.reason === 'autosave')) return;
        boot();
    });
    window.addEventListener('ms365-tenant-klassen-hub-refresh', boot);
    boot();
}
