/**
 * Stammdaten-Tab „Lehrer je Klasse“ – Hub wie Schüler „Nach Klassen“.
 */
import {
    teacherPickOptions,
    listClassTeachingBoards,
    summarizeClassTeaching,
    addTeacherToClassBelegung,
    listBelegungRowsForKlasse,
    patchBelegungRowAt,
    removeBelegungRowAt,
    belegungSnapshotFromRows
} from './klassenlehrer-stammdaten-logic.js';
import { pickEntraUser } from './entra-user-picker.js';

const VIEW_KEY = 'ms365-tenant-klassenlehrer-view-v1';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function normStr(v) {
    return String(v ?? '').trim();
}

function readView() {
    try {
        const v = localStorage.getItem(VIEW_KEY);
        if (v === 'liste' || v === 'klassen') return v;
    } catch {
        /* ignore */
    }
    return 'klassen';
}

function writeView(v) {
    try {
        localStorage.setItem(VIEW_KEY, v === 'liste' ? 'liste' : 'klassen');
    } catch {
        /* ignore */
    }
}

/**
 * @param {{
 *   getClasses: () => object[],
 *   getTeachers: () => object[],
 *   getBelegung: () => object|null,
 *   saveBelegung: (snapshot: object) => void,
 *   toast?: (msg: string, kind?: string) => void,
 *   onChanged?: () => void
 * }} opts
 */
export function mountKlassenlehrerPanel(opts) {
    const o = opts || {};
    const root = document.getElementById('tenantKlassenlehrerHub');
    if (!root) return { refresh: function () {} };
    if (root.dataset.klassenlehrerMounted === '1') {
        return {
            refresh: function () {
                if (typeof root._ms365KlassenlehrerRefresh === 'function') {
                    root._ms365KlassenlehrerRefresh();
                }
            }
        };
    }
    root.dataset.klassenlehrerMounted = '1';

    const summaryEl = document.getElementById('tenantKlassenlehrerSummary');
    const searchEl = document.getElementById('tenantKlassenlehrerSearch');
    const filterEl = document.getElementById('tenantKlassenlehrerFilter');
    const btnViewKlassen = document.getElementById('tenantKlassenlehrerViewKlassen');
    const btnViewListe = document.getElementById('tenantKlassenlehrerViewListe');
    const wrapKlassen = document.getElementById('tenantKlassenlehrerKlassenWrap');
    const wrapListe = document.getElementById('tenantKlassenlehrerListeWrap');

    function toast(msg, kind) {
        if (typeof o.toast === 'function') o.toast(msg, kind);
        else if (typeof window.ms365Toast === 'function') window.ms365Toast(msg, kind);
    }

    function loadRows() {
        const snap = typeof o.getBelegung === 'function' ? o.getBelegung() : null;
        return snap && Array.isArray(snap.rows) ? snap.rows.slice() : [];
    }

    function persistRows(rows) {
        const prev = typeof o.getBelegung === 'function' ? o.getBelegung() : null;
        const snap = belegungSnapshotFromRows(rows, prev);
        if (typeof o.saveBelegung === 'function') o.saveBelegung(snap);
        if (typeof o.onChanged === 'function') o.onChanged();
        refresh();
    }

    function applyViewMode(mode) {
        const isKlassen = mode !== 'liste';
        if (wrapKlassen) wrapKlassen.hidden = !isKlassen;
        if (wrapListe) wrapListe.hidden = isKlassen;
        if (btnViewKlassen) btnViewKlassen.classList.toggle('is-active', isKlassen);
        if (btnViewListe) btnViewListe.classList.toggle('is-active', !isKlassen);
    }

    function updateSummary(classes, rows) {
        const sum = summarizeClassTeaching(classes, rows);
        if (!summaryEl) return;
        summaryEl.textContent =
            sum.withTeachers +
            ' von ' +
            sum.classTotal +
            ' Klassen mit Lehrkräften · ' +
            sum.rowCount +
            ' Belegungszeilen' +
            (sum.withoutTeachers ? ' · ' + sum.withoutTeachers + ' ohne Belegung' : '');
    }

    function renderGesamtliste() {
        const tbody = document.getElementById('tenantKlassenlehrerListeBody');
        if (!tbody) return;
        const classes = typeof o.getClasses === 'function' ? o.getClasses() : [];
        const rows = loadRows();
        const q = searchEl ? normStr(searchEl.value).toLowerCase() : '';
        const filter = filterEl ? filterEl.value : 'all';
        updateSummary(classes, rows);
        tbody.replaceChildren();
        const boards = listClassTeachingBoards(classes, rows, 'all', '');
        const withoutSet = new Set(
            boards.filter(function (b) {
                return !b.teachers.length;
            }).map(function (b) {
                return b.klasse;
            })
        );
        /** @type {Array<{ index: number, row: object }>} */
        const listed = [];
        rows.forEach(function (row, index) {
            const klasse = normStr(row.klasse).toUpperCase() || normStr(row.klasse);
            if (filter === 'without') return;
            if (filter === 'with' && withoutSet.has(klasse) && !row.lehrerEmail && !row.lehrerCode) {
                return;
            }
            if (q) {
                const hay = [row.klasse, row.lehrerName, row.lehrerEmail, row.lehrerCode, row.fach]
                    .map(function (x) {
                        return normStr(x).toLowerCase();
                    })
                    .join(' ');
                if (hay.indexOf(q) === -1) return;
            }
            listed.push({ index: index, row: row });
        });
        listed.sort(function (a, b) {
            const c = normStr(a.row.klasse).localeCompare(normStr(b.row.klasse), 'de', { numeric: true });
            if (c !== 0) return c;
            const na = normStr(a.row.lehrerName) || normStr(a.row.lehrerEmail);
            const nb = normStr(b.row.lehrerName) || normStr(b.row.lehrerEmail);
            return na.localeCompare(nb, 'de');
        });
        if (!listed.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 6;
            td.className = 'muted';
            td.textContent = rows.length
                ? 'Keine Treffer für Filter/Suche.'
                : 'Noch keine Unterrichtsbelegung – links Klasse wählen und Lehrkräfte hinzufügen oder aus Kursteams übernehmen.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }
        listed.forEach(function (entry) {
            tbody.appendChild(buildEditableRow(entry.index, entry.row, true));
        });
    }

    /**
     * @param {number} index
     * @param {object} row
     * @param {boolean} showKlasse
     */
    function buildEditableRow(index, row, showKlasse) {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = String(index);

        if (showKlasse) {
            const tdK = document.createElement('td');
            tdK.innerHTML = '<code>' + escapeHtml(row.klasse || '') + '</code>';
            tr.appendChild(tdK);
        }

        function cellInput(field, type, value, aria) {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = type || 'text';
            input.className = 'ts-klassen-hub__input';
            input.value = value || '';
            input.setAttribute('aria-label', aria);
            input.dataset.field = field;
            input.addEventListener('change', function () {
                const patch = {};
                let v = normStr(input.value);
                if (field === 'lehrerEmail') v = v.toLowerCase();
                if (field === 'lehrerCode') v = v.toUpperCase();
                patch[field] = v;
                persistRows(patchBelegungRowAt(loadRows(), index, patch));
            });
            td.appendChild(input);
            return td;
        }

        tr.appendChild(cellInput('lehrerName', 'text', row.lehrerName || '', 'Name'));
        tr.appendChild(cellInput('lehrerEmail', 'email', row.lehrerEmail || '', 'E-Mail'));
        tr.appendChild(cellInput('lehrerCode', 'text', row.lehrerCode || '', 'Account / Kurzzeichen'));
        tr.appendChild(cellInput('fach', 'text', row.fach || '', 'Fach'));

        const tdAction = document.createElement('td');
        tdAction.className = 'action-cell';
        const btnDel = document.createElement('button');
        btnDel.type = 'button';
        btnDel.className = 'btn btn-sm';
        btnDel.title = 'Zeile entfernen';
        btnDel.innerHTML = '<i class="bi bi-trash" aria-hidden="true"></i>';
        btnDel.addEventListener('click', function () {
            persistRows(removeBelegungRowAt(loadRows(), index));
            toast('Belegungszeile entfernt.', 'ok');
        });
        tdAction.appendChild(btnDel);
        tr.appendChild(tdAction);
        return tr;
    }

    function renderKlassenHub() {
        const mount = document.getElementById('tenantKlassenlehrerKlassenMount');
        if (!mount) return;
        const classes = typeof o.getClasses === 'function' ? o.getClasses() : [];
        const teachers = typeof o.getTeachers === 'function' ? o.getTeachers() : [];
        const rows = loadRows();
        const filter = filterEl ? filterEl.value : 'all';
        const query = searchEl ? searchEl.value : '';
        const boards = listClassTeachingBoards(classes, rows, filter, query);
        updateSummary(classes, rows);

        let selectedKey = mount.dataset.selectedClass || '';
        if (!selectedKey && boards.length) selectedKey = boards[0].klasse;
        if (selectedKey && !boards.some(function (b) {
            return b.klasse === selectedKey;
        })) {
            selectedKey = boards[0] ? boards[0].klasse : '';
        }
        mount.dataset.selectedClass = selectedKey;
        const selected = boards.find(function (b) {
            return b.klasse === selectedKey;
        });

        let html = '<div class="ts-klassen-hub__layout">';
        html += '<aside class="ts-klassen-hub__classes" aria-label="Klassen">';
        html +=
            '<p class="ts-klassen-hub__aside-title">Klassen <span class="muted">(' +
            boards.length +
            ')</span></p>';
        html += '<ul class="ts-klassen-hub__class-list">';
        if (!boards.length) {
            html +=
                '<li class="muted" style="padding:8px;">Keine Klassen – Tab <strong>Klassen</strong> pflegen.</li>';
        }
        boards.forEach(function (b) {
            const n = b.teachers.length;
            const active = b.klasse === selectedKey ? ' is-active' : '';
            html +=
                '<li><button type="button" class="ts-klassen-hub__class-btn' +
                active +
                '" data-class-key="' +
                escapeHtml(b.klasse) +
                '">' +
                '<span class="ts-klassen-hub__class-code">' +
                escapeHtml(b.klasse) +
                '</span>' +
                '<span class="ts-klassen-hub__class-meta">' +
                n +
                ' Lehrkraft' +
                (n === 1 ? '' : 'kräfte') +
                '</span></button></li>';
        });
        html += '</ul></aside>';

        html += '<div class="ts-klassen-hub__main">';
        if (selected) {
            const c = selected.classRow || { code: selected.klasse };
            const label = escapeHtml(selected.klasse) + (c.name && normStr(c.name).toUpperCase() !== selected.klasse ? ' · ' + escapeHtml(c.name) : '');
            html += '<header class="ts-klassen-hub__head">';
            html += '<h3 class="ts-klassen-hub__title">' + label + '</h3>';
            html +=
                '<span class="muted">' +
                selected.teachers.length +
                ' Lehrkraft' +
                (selected.teachers.length === 1 ? '' : 'kräfte') +
                '</span>';
            html += '</header>';
            if (selected.kvName || selected.kvEmail) {
                html +=
                    '<p class="ts-klassen-hub__kv muted">KV: ' +
                    escapeHtml(selected.kvName || '–') +
                    (selected.kvEmail ? ' · ' + escapeHtml(selected.kvEmail) : '') +
                    ' <span class="muted">(Tab Klassen)</span></p>';
            }
            html +=
                '<div class="teachers-table-wrap"><table class="teachers-table ts-klassen-hub__table" aria-label="Unterrichtende Lehrkräfte"><thead><tr>';
            html +=
                '<th>Name</th><th>E-Mail</th><th>Account</th><th>Fach</th><th class="action-cell"></th></tr></thead>';
            html += '<tbody id="tenantKlassenlehrerClassBody"></tbody></table></div>';
            html += '<div class="settings-actions ts-klassenlehrer-addbar" style="margin-top:10px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;">';
            html +=
                '<select id="tenantKlassenlehrerAddSelect" title="Lehrkraft aus Lehrerliste" aria-label="Lehrkraft hinzufügen"></select>';
            html +=
                '<input type="text" id="tenantKlassenlehrerAddFach" class="ts-klassen-hub__input" placeholder="Fach (optional)" style="max-width:120px;" aria-label="Fach">';
            html +=
                '<button type="button" class="btn btn-success" id="tenantKlassenlehrerAddBtn"><i class="bi bi-plus-circle"></i>Hinzufügen</button>';
            html +=
                '<button type="button" class="btn" id="tenantKlassenlehrerAddEntra" title="In Microsoft 365 suchen"><i class="bi bi-microsoft" aria-hidden="true"></i> M365</button>';
            html += '</div>';
        } else {
            html +=
                '<p class="muted">Wählen Sie links eine Klasse oder legen Sie Klassen im Tab <strong>Klassen</strong> an.</p>';
        }
        html += '</div></div>';
        mount.innerHTML = html;

        mount.querySelectorAll('.ts-klassen-hub__class-btn').forEach(function (btn) {
            btn.addEventListener('click', function () {
                mount.dataset.selectedClass = btn.getAttribute('data-class-key') || '';
                renderKlassenHub();
            });
        });

        if (!selected) return;

        const tbody = mount.querySelector('#tenantKlassenlehrerClassBody');
        const classRows = listBelegungRowsForKlasse(rows, selected.klasse);
        if (tbody) {
            if (!classRows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 5;
                td.className = 'muted';
                td.textContent = 'Noch keine unterrichtenden Lehrkräfte in dieser Klasse.';
                tr.appendChild(td);
                tbody.appendChild(tr);
            } else {
                classRows.forEach(function (entry) {
                    tbody.appendChild(buildEditableRow(entry.index, entry.row, false));
                });
            }
        }

        const pickOpts = teacherPickOptions(teachers);
        const sel = mount.querySelector('#tenantKlassenlehrerAddSelect');
        const fachInp = mount.querySelector('#tenantKlassenlehrerAddFach');
        const btnAdd = mount.querySelector('#tenantKlassenlehrerAddBtn');
        const btnEntra = mount.querySelector('#tenantKlassenlehrerAddEntra');
        if (sel) {
            const opt0 = document.createElement('option');
            opt0.value = '';
            opt0.textContent = pickOpts.length ? 'Lehrer wählen…' : 'Keine Lehrerliste';
            sel.appendChild(opt0);
            pickOpts.forEach(function (p) {
                const opt = document.createElement('option');
                opt.value = p.email;
                opt.textContent = p.label;
                sel.appendChild(opt);
            });
            sel.disabled = !pickOpts.length;
        }
        function addFromPick(email, name, code, fach) {
            const next = addTeacherToClassBelegung(loadRows(), selected.klasse, {
                email: email,
                name: name,
                code: code,
                fach: fach
            });
            persistRows(next);
            toast((name || email) + ' → ' + selected.klasse, 'ok');
        }
        if (btnAdd && sel) {
            btnAdd.addEventListener('click', function () {
                const em = normStr(sel.value).toLowerCase();
                if (!em) {
                    toast('Bitte eine Lehrkraft wählen.', 'warn');
                    return;
                }
                const hit = pickOpts.find(function (p) {
                    return p.email === em;
                });
                addFromPick(em, hit ? hit.name : '', hit ? hit.code : '', fachInp ? fachInp.value : '');
                if (fachInp) fachInp.value = '';
                sel.value = '';
            });
        }
        if (btnEntra) {
            btnEntra.addEventListener('click', async function () {
                let picked;
                try {
                    picked = await pickEntraUser({
                        title: 'Lehrkraft für Klasse ' + selected.klasse,
                        hint: 'Als unterrichtende Lehrkraft in „' + selected.klasse + '“ übernehmen.',
                        initialQuery: selected.klasse
                    });
                } catch (e) {
                    toast('M365-Suche: ' + (e && e.message ? e.message : String(e)), 'warn');
                    return;
                }
                if (!picked) return;
                const mail = normStr(picked.mail || picked.userPrincipalName || '').toLowerCase();
                if (!mail || mail.indexOf('@') === -1) {
                    toast('Gewählte Person hat keine E-Mail.', 'warn');
                    return;
                }
                addFromPick(
                    mail,
                    normStr(picked.displayName) || mail,
                    '',
                    fachInp ? fachInp.value : ''
                );
                if (fachInp) fachInp.value = '';
            });
        }
    }

    function refresh() {
        const mode = readView();
        applyViewMode(mode);
        if (mode === 'liste') renderGesamtliste();
        else renderKlassenHub();
    }

    root._ms365KlassenlehrerRefresh = refresh;

    if (btnViewKlassen && !btnViewKlassen.dataset.bound) {
        btnViewKlassen.dataset.bound = '1';
        btnViewKlassen.addEventListener('click', function () {
            writeView('klassen');
            refresh();
        });
    }
    if (btnViewListe && !btnViewListe.dataset.bound) {
        btnViewListe.dataset.bound = '1';
        btnViewListe.addEventListener('click', function () {
            writeView('liste');
            refresh();
        });
    }
    if (filterEl && !filterEl.dataset.bound) {
        filterEl.dataset.bound = '1';
        filterEl.addEventListener('change', refresh);
    }
    if (searchEl && !searchEl.dataset.bound) {
        searchEl.dataset.bound = '1';
        searchEl.addEventListener('input', refresh);
    }

    refresh();
    return { refresh: refresh };
}

export default { mountKlassenlehrerPanel };
