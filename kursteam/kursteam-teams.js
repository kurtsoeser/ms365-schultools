(function () {
    'use strict';

    const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

    function normalizeSubjectToken(s) {
        return String(s || '').trim().toUpperCase();
    }

    function parseExcludeSubjectsFromInput() {
        const el = document.getElementById('excludeSubjects');
        if (!el) return [];
        return String(el.value || '')
            .split(',')
            .map(normalizeSubjectToken)
            .filter(s => s.length > 0);
    }

    function setExcludeSubjectsInput(tokens) {
        const el = document.getElementById('excludeSubjects');
        if (!el) return;
        const uniq = Array.from(new Set((tokens || []).map(normalizeSubjectToken).filter(Boolean)));
        uniq.sort((a, b) => a.localeCompare(b, 'de'));
        el.value = uniq.join(',');
    }

    function collectAvailableSubjectsFromRawData() {
        const set = new Set();
        (ns.rawData || []).forEach(r => {
            const t = normalizeSubjectToken(r && r.fach);
            if (t) set.add(t);
        });
        return Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
    }

    function updateSubjectFilterSummary(available, excluded) {
        const el = document.getElementById('subjectFilterSummary');
        if (!el) return;
        const a = available.length;
        const e = excluded.length;
        if (!a) {
            el.textContent = 'Noch keine Daten: Importieren Sie zuerst Zeilen in Schritt 1 oder fügen Sie manuell Unterrichtszeilen hinzu.';
            return;
        }
        el.textContent = `${a} Fach/Fächer erkannt. ${e} ausgeschlossen.`;
    }

    function applySearchToSubjectList(query) {
        const q = normalizeSubjectToken(query);
        const list = document.getElementById('subjectFilterList');
        if (!list) return;
        Array.from(list.querySelectorAll('[data-subject]')).forEach(node => {
            const subj = String(node.getAttribute('data-subject') || '');
            node.style.display = !q || subj.includes(q) ? '' : 'none';
        });
    }

    function wireSubjectFilterEventsOnce() {
        if (wireSubjectFilterEventsOnce._wired) return;
        wireSubjectFilterEventsOnce._wired = true;

        const search = document.getElementById('subjectFilterSearch');
        if (search) {
            search.addEventListener('input', () => applySearchToSubjectList(search.value));
        }

        const btnNone = document.getElementById('subjectFilterExcludeNone');
        if (btnNone) {
            btnNone.addEventListener('click', () => {
                setExcludeSubjectsInput([]);
                ns.refreshSubjectFilterUI();
            });
        }

        const btnAll = document.getElementById('subjectFilterExcludeAll');
        if (btnAll) {
            btnAll.addEventListener('click', () => {
                setExcludeSubjectsInput(collectAvailableSubjectsFromRawData());
                ns.refreshSubjectFilterUI();
            });
        }

        const btnDefault = document.getElementById('subjectFilterResetDefault');
        if (btnDefault) {
            btnDefault.addEventListener('click', () => {
                setExcludeSubjectsInput(['ORD', 'DIR', 'KV']);
                ns.refreshSubjectFilterUI();
            });
        }

        const input = document.getElementById('excludeSubjects');
        if (input) {
            input.addEventListener('input', () => ns.refreshSubjectFilterUI());
        }
    }

    ns.refreshSubjectFilterUI = function refreshSubjectFilterUI() {
        const list = document.getElementById('subjectFilterList');
        const search = document.getElementById('subjectFilterSearch');
        if (!list) return;

        wireSubjectFilterEventsOnce();

        const available = collectAvailableSubjectsFromRawData();
        const excluded = new Set(parseExcludeSubjectsFromInput());

        list.replaceChildren();
        available.forEach(subj => {
            const label = document.createElement('label');
            label.className = 'subject-filter-item';
            label.setAttribute('data-subject', subj);

            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = excluded.has(subj); // checked = ausgeschlossen
            cb.addEventListener('change', () => {
                const current = new Set(parseExcludeSubjectsFromInput());
                if (cb.checked) current.add(subj);
                else current.delete(subj);
                setExcludeSubjectsInput(Array.from(current));
                updateSubjectFilterSummary(available, Array.from(current));
            });

            const text = document.createElement('code');
            text.textContent = subj;

            label.append(cb, text);
            list.appendChild(label);
        });

        updateSubjectFilterSummary(available, Array.from(excluded));
        if (search) applySearchToSubjectList(search.value);
    };

    ns.applyFilters = function applyFilters() {
        const excludeSubjects = parseExcludeSubjectsFromInput();
        const removeDuplicates = document.getElementById('removeDuplicates').checked;

        let filtered = ns.rawData.filter(row => {
            if (!row.fach || !row.lehrer) return false;
            const fach = row.fach.toUpperCase().trim();
            if (excludeSubjects.includes(fach)) return false;
            if (!row.klasse || row.klasse.trim() === '') return false;
            return true;
        });

        const originalCount = filtered.length;
        const removedByFilter = ns.rawData.length - originalCount;

        if (removeDuplicates) {
            const seen = new Set();
            filtered = filtered.filter(row => {
                const key = `${row.klasse}-${row.fach}-${row.lehrer}-${row.gruppe}`;
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        }

        ns.filteredData = filtered;
        ns.invalidateTeams();
        document.getElementById('filteredRecords').textContent = filtered.length;
        document.getElementById('removedDuplicates').textContent = removedByFilter + (originalCount - filtered.length);
        document.getElementById('filterStats').style.display = 'block';
        ns.displayFilteredData();
    };

    ns.displayFilteredData = function displayFilteredData() {
        const tbody = document.getElementById('dataTableBody');
        tbody.replaceChildren();
        ns.filteredData.forEach((row, index) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = row.klasse;
            const td2 = document.createElement('td');
            td2.textContent = row.fach;
            const td3 = document.createElement('td');
            td3.textContent = row.lehrer;
            const td4 = document.createElement('td');
            td4.textContent = row.gruppe || '-';
            const tdAction = document.createElement('td');
            tdAction.className = 'kt-action-col';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger kt-delete-btn';
            btn.textContent = 'X';
            btn.title = 'Zeile löschen';
            btn.setAttribute('aria-label', 'Zeile löschen');
            btn.addEventListener('click', () => ns.removeRow(index));
            tdAction.appendChild(btn);

            tr.append(td1, td2, td3, td4, tdAction);
            tbody.appendChild(tr);
        });
        const hasRows = ns.filteredData.length > 0;
        document.getElementById('dataTableContainer').style.display = hasRows ? 'block' : 'none';
        document.getElementById('continueBtn2').style.display = hasRows ? 'inline-block' : 'none';
    };

    function setCellEditMode(td, rowId, field) {
        if (!td || td.dataset.editing === '1') return;
        td.dataset.editing = '1';

        const originalText = td.textContent;
        const input = document.createElement('input');
        input.type = 'text';
        input.value = originalText === '-' ? '' : originalText;
        input.style.width = '100%';
        input.style.padding = '6px 8px';
        input.style.border = '1px solid #ced4da';
        input.style.borderRadius = '6px';
        input.style.fontSize = '0.95em';
        input.style.boxSizing = 'border-box';

        td.replaceChildren(input);
        input.focus();
        input.select();

        const commit = () => {
            const val = input.value.trim();
            ns.updateDataRowField(rowId, field, val);
            td.dataset.editing = '0';
            td.textContent = val || (field === 'gruppe' ? '-' : '');
            if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        };
        const cancel = () => {
            td.dataset.editing = '0';
            td.textContent = originalText;
        };

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                commit();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancel();
            }
        });
        input.addEventListener('blur', commit);
    }

    ns.updateDataRowField = function updateDataRowField(rowId, field, value) {
        const idx = ns.filteredData.findIndex(r => r && r.id === rowId);
        if (idx >= 0) {
            ns.filteredData[idx][field] = value;
        }
        const ridx = ns.rawData.findIndex(r => r && r.id === rowId);
        if (ridx >= 0) {
            ns.rawData[ridx][field] = value;
        }
        ns.invalidateTeams();
    };

    function normFilterToken(s) {
        return String(s || '').trim().toUpperCase();
    }

    function getManualFilterState() {
        return {
            klasse: normFilterToken(document.getElementById('manualFilterKlasse')?.value),
            fach: normFilterToken(document.getElementById('manualFilterFach')?.value),
            lehrer: normFilterToken(document.getElementById('manualFilterLehrer')?.value)
        };
    }

    function clearManualFilterInputs() {
        const k = document.getElementById('manualFilterKlasse');
        const f = document.getElementById('manualFilterFach');
        const l = document.getElementById('manualFilterLehrer');
        if (k) k.value = '';
        if (f) f.value = '';
        if (l) l.value = '';
    }

    function ensureManualSortState() {
        if (!ns.manualSort) ns.manualSort = { key: 'klasse', dir: 1 };
    }

    function applyManualFiltersAndSort() {
        ensureManualSortState();
        const filters = getManualFilterState();
        const out = [];
        ns.filteredData.forEach((row, index) => {
            const klasse = normFilterToken(row.klasse);
            const fach = normFilterToken(row.fach);
            const lehrer = normFilterToken(row.lehrer);
            if (filters.klasse && !klasse.includes(filters.klasse)) return;
            if (filters.fach && !fach.includes(filters.fach)) return;
            if (filters.lehrer && !lehrer.includes(filters.lehrer)) return;
            out.push({ row, index });
        });

        const key = ns.manualSort?.key;
        const dir = ns.manualSort?.dir || 1;
        if (key) {
            out.sort((a, b) => {
                const av = normFilterToken(a.row[key] || '');
                const bv = normFilterToken(b.row[key] || '');
                const cmp = av.localeCompare(bv, 'de');
                if (cmp !== 0) return cmp * dir;
                // stabile Zweitsortierung: id
                return (a.row.id > b.row.id ? 1 : a.row.id < b.row.id ? -1 : 0) * dir;
            });
        }
        return out;
    }

    function updateManualSortIndicators() {
        const table = document.getElementById('editableDataTable');
        if (!table) return;
        const ths = table.querySelectorAll('th[data-sort-key]');
        ths.forEach(th => {
            const label = th.dataset.label || th.textContent.replace(/[▲▼]\s*$/, '').trim();
            th.dataset.label = label;
            th.textContent = label;
            if (ns.manualSort && th.dataset.sortKey === ns.manualSort.key) {
                const ind = document.createElement('span');
                ind.className = 'kt-sort-indicator';
                ind.textContent = ns.manualSort.dir === -1 ? '▼' : '▲';
                th.appendChild(ind);
            }
        });
    }

    function wireManualSortAndFilterOnce() {
        if (wireManualSortAndFilterOnce._wired) return;
        wireManualSortAndFilterOnce._wired = true;

        const table = document.getElementById('editableDataTable');
        if (table) {
            table.querySelectorAll('th[data-sort-key]').forEach(th => {
                th.addEventListener('click', () => {
                    ensureManualSortState();
                    const k = th.dataset.sortKey;
                    if (ns.manualSort.key === k) ns.manualSort.dir = ns.manualSort.dir === 1 ? -1 : 1;
                    else ns.manualSort = { key: k, dir: 1 };
                    ns.displayEditableData();
                });
            });
        }

        const onInput = () => ns.displayEditableData();
        const k = document.getElementById('manualFilterKlasse');
        const f = document.getElementById('manualFilterFach');
        const l = document.getElementById('manualFilterLehrer');
        if (k) k.addEventListener('input', onInput);
        if (f) f.addEventListener('input', onInput);
        if (l) l.addEventListener('input', onInput);

        const reset = document.getElementById('manualFilterReset');
        if (reset) {
            reset.addEventListener('click', () => {
                clearManualFilterInputs();
                ns.displayEditableData();
            });
        }
    }

    ns.displayEditableData = function displayEditableData() {
        const container = document.getElementById('editableDataTableContainer');
        const tbody = document.getElementById('editableDataTableBody');
        if (!container || !tbody) return;

        wireManualSortAndFilterOnce();

        tbody.replaceChildren();
        const view = applyManualFiltersAndSort();
        view.forEach(({ row, index }) => {
            const tr = document.createElement('tr');

            const tdKlasse = document.createElement('td');
            tdKlasse.textContent = row.klasse || '';
            tdKlasse.addEventListener('dblclick', () => setCellEditMode(tdKlasse, row.id, 'klasse'));

            const tdFach = document.createElement('td');
            tdFach.textContent = row.fach || '';
            tdFach.addEventListener('dblclick', () => setCellEditMode(tdFach, row.id, 'fach'));

            const tdLehrer = document.createElement('td');
            tdLehrer.textContent = row.lehrer || '';
            tdLehrer.addEventListener('dblclick', () => setCellEditMode(tdLehrer, row.id, 'lehrer'));

            const tdGruppe = document.createElement('td');
            tdGruppe.textContent = row.gruppe ? row.gruppe : '-';
            tdGruppe.addEventListener('dblclick', () => setCellEditMode(tdGruppe, row.id, 'gruppe'));

            const tdAction = document.createElement('td');
            tdAction.className = 'kt-action-col';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger kt-delete-btn';
            btn.textContent = 'X';
            btn.title = 'Zeile löschen';
            btn.setAttribute('aria-label', 'Zeile löschen');
            btn.addEventListener('click', () => {
                ns.removeRow(index);
                ns.displayEditableData();
            });
            tdAction.appendChild(btn);

            tr.append(tdKlasse, tdFach, tdLehrer, tdGruppe, tdAction);
            tbody.appendChild(tr);
        });

        container.style.display = view.length ? 'block' : 'none';
        updateManualSortIndicators();
    };

    ns.displayManualTeamsPreview = function displayManualTeamsPreview() {
        const wrap = document.getElementById('manualTeamsPreviewContainer');
        const body = document.getElementById('manualTeamsPreviewBody');
        if (!wrap || !body) return;

        if (!ns.teamsGenerated || !Array.isArray(ns.teamsData) || ns.teamsData.length === 0) {
            wrap.style.display = 'none';
            body.replaceChildren();
            return;
        }

        body.replaceChildren();
        ns.teamsData.forEach(team => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = team.teamName;
            const td2 = document.createElement('td');
            td2.textContent = team.gruppenmail;
            const td3 = document.createElement('td');
            td3.textContent = team.besitzer;
            const td4 = document.createElement('td');
            td4.textContent = team.isValid ? '✅' : '❌';
            tr.append(td1, td2, td3, td4);
            body.appendChild(tr);
        });
        wrap.style.display = 'block';
    };

    ns.addManualDataRowInline = function addManualDataRowInline() {
        const id = Date.now() + Math.random();
        const row = {
            id,
            klasse: '',
            fach: '',
            lehrer: '',
            gruppe: '',
            original: { manualInline: true }
        };
        ns.rawData.push(row);
        ns.filteredData.push(row);
        ns.kursteamEntryMode = ns.kursteamEntryMode === 'unset' ? 'manual' : ns.kursteamEntryMode;
        ns.invalidateTeams();
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        clearManualFilterInputs();
        if (typeof ns.displayEditableData === 'function') ns.displayEditableData();

        // Fokus: erste bearbeitbare Zelle der neu hinzugefügten Zeile
        try {
            const tbody = document.getElementById('editableDataTableBody');
            const lastRow = tbody ? tbody.lastElementChild : null;
            if (lastRow && lastRow.children && lastRow.children.length >= 1) {
                const tdKlasse = lastRow.children[0];
                setCellEditMode(tdKlasse, id, 'klasse');
            }
        } catch (e) {
            /* ignore */
        }
    };

    ns.removeRow = function removeRow(index) {
        const row = ns.filteredData[index];
        ns.filteredData.splice(index, 1);
        if (row && row.id !== undefined && row.id !== null) {
            const ri = ns.rawData.findIndex(r => r.id === row.id);
            if (ri >= 0) ns.rawData.splice(ri, 1);
        }
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        ns.invalidateTeams();
        ns.displayFilteredData();
        if (ns.currentStep === 2.5 && typeof ns.displayEditableData === 'function') ns.displayEditableData();
        document.getElementById('filteredRecords').textContent = ns.filteredData.length;
    };

    ns.startKursteamFromWebuntis = function startKursteamFromWebuntis() {
        ns.kursteamEntryMode = 'webuntis';
        ns.goToStep(1);
    };

    ns.startKursteamManual = function startKursteamManual() {
        ns.kursteamEntryMode = 'manual';
        ns.rawData = [];
        ns.filteredData = [];
        document.getElementById('totalRecords').textContent = '0';
        document.getElementById('uniqueSubjects').textContent = '0';
        document.getElementById('uniqueTeachers').textContent = '0';
        document.getElementById('importStats').style.display = 'none';
        const fi = document.getElementById('fileInput');
        if (fi) fi.value = '';
        ns.invalidateTeams();
        ns.goToStep(2);
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        document.getElementById('filterStats').style.display = 'none';
        document.getElementById('dataTableContainer').style.display = 'none';
        document.getElementById('continueBtn2').style.display = 'none';
        const tbody = document.getElementById('dataTableBody');
        if (tbody) tbody.replaceChildren();
    };

    ns.addManualDataRow = function addManualDataRow() {
        ns.openModal(
            'Unterrichtszeile hinzufügen',
            '<label for="manualKlasse">Klasse</label><input type="text" id="manualKlasse" autocomplete="off" placeholder="z. B. 5A">' +
                '<label for="manualFach">Fach</label><input type="text" id="manualFach" autocomplete="off" placeholder="z. B. D">' +
                '<label for="manualLehrer">Lehrkraft (Kürzel)</label><input type="text" id="manualLehrer" autocomplete="off" placeholder="z. B. MEI">' +
                '<label for="manualGruppe">Schülergruppe (optional)</label><input type="text" id="manualGruppe" autocomplete="off" placeholder="leer oder z. B. G1">',
            () => {
                const klasse = document.getElementById('manualKlasse').value.trim();
                const fach = document.getElementById('manualFach').value.trim();
                const lehrer = document.getElementById('manualLehrer').value.trim();
                const gruppe = document.getElementById('manualGruppe').value.trim();
                if (!klasse || !fach || !lehrer) {
                    ns.showToast('Bitte Klasse, Fach und Lehrkraft ausfüllen.');
                    return;
                }
                const id = Date.now() + Math.random();
                const row = {
                    id,
                    klasse,
                    fach,
                    lehrer,
                    gruppe: gruppe || '',
                    original: {}
                };
                ns.rawData.push(row);
                ns.filteredData.push(row);
                ns.kursteamEntryMode = ns.kursteamEntryMode === 'unset' ? 'manual' : ns.kursteamEntryMode;
                if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
                ns.invalidateTeams();
                ns.closeModal();
                document.getElementById('filteredRecords').textContent = ns.filteredData.length;
                document.getElementById('filterStats').style.display = 'block';
                ns.displayFilteredData();
            }
        );
    };

    ns.resetFilters = function resetFilters() {
        ns.filteredData = [...ns.rawData];
        setExcludeSubjectsInput(['ORD', 'DIR', 'KV']);
        document.getElementById('removeDuplicates').checked = true;
        if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
        ns.applyFilters();
    };

    function defaultTeamNamePattern() {
        return [
            { type: 'yearPrefix' },
            { type: 'text', value: ' | ' },
            { type: 'klasse' },
            { type: 'text', value: ' | ' },
            { type: 'fach' }
        ];
    }

    function normalizePattern(pattern) {
        const arr = Array.isArray(pattern) ? pattern : [];
        const out = [];
        arr.forEach(p => {
            if (!p || typeof p !== 'object') return;
            const type = String(p.type || '').trim();
            if (!type) return;
            if (type === 'text') {
                out.push({ type: 'text', value: String(p.value ?? '') });
            } else if (type === 'yearPrefix' || type === 'klasse' || type === 'fach' || type === 'gruppe') {
                out.push({ type });
            }
        });
        return out.length ? out : defaultTeamNamePattern();
    }

    function buildTeamNameFromPattern(pattern, ctx) {
        const parts = [];
        normalizePattern(pattern).forEach(p => {
            if (p.type === 'text') parts.push(String(p.value ?? ''));
            else if (p.type === 'yearPrefix') parts.push(String(ctx.yearPrefix ?? ''));
            else if (p.type === 'klasse') parts.push(String(ctx.klasse ?? ''));
            else if (p.type === 'fach') parts.push(String(ctx.fach ?? ''));
            else if (p.type === 'gruppe') parts.push(String(ctx.gruppe ?? ''));
        });
        return parts.join('');
    }

    function tokenLabel(t) {
        if (t.type === 'yearPrefix') return 'Schuljahr';
        if (t.type === 'klasse') return 'Klasse';
        if (t.type === 'fach') return 'Fach';
        if (t.type === 'gruppe') return 'Gruppe';
        if (t.type === 'text') return `Text`;
        return t.type;
    }

    function getPatternFromBuilder() {
        const zone = document.getElementById('teamNameBuilder');
        if (!zone) return normalizePattern(ns.teamNamePattern);
        const tokens = [];
        zone.querySelectorAll('[data-token-type]').forEach(el => {
            const type = String(el.getAttribute('data-token-type') || '');
            if (type === 'text') tokens.push({ type: 'text', value: String(el.getAttribute('data-token-value') || '') });
            else tokens.push({ type });
        });
        return normalizePattern(tokens);
    }

    function setPreviewFromPattern(pattern) {
        const el = document.getElementById('teamNamePreview');
        if (!el) return;
        const yearPrefix = document.getElementById('yearPrefix')?.value || 'WS24';
        const preview = buildTeamNameFromPattern(pattern, { yearPrefix, klasse: '1AK', fach: 'D', gruppe: 'G1' });
        el.textContent = 'Vorschau: ' + preview;
    }

    function wireBuilderDnD(zone) {
        let dragEl = null;
        zone.addEventListener('dragstart', (e) => {
            const target = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
            if (!target) return;
            dragEl = target;
            target.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        });
        zone.addEventListener('dragend', () => {
            if (dragEl) dragEl.classList.remove('dragging');
            dragEl = null;
        });
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            const over = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
            if (!dragEl || !over || over === dragEl) return;
            const rect = over.getBoundingClientRect();
            const after = e.clientX > rect.left + rect.width / 2;
            if (after) over.after(dragEl);
            else over.before(dragEl);
        });
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            ns.teamNamePattern = getPatternFromBuilder();
            setPreviewFromPattern(ns.teamNamePattern);
        });
    }

    function addChip(zone, token) {
        const chip = document.createElement('span');
        chip.className = 'name-chip';
        chip.draggable = true;
        chip.setAttribute('data-token-type', token.type);
        if (token.type === 'text') chip.setAttribute('data-token-value', String(token.value ?? ''));

        const txt = document.createElement('span');
        if (token.type === 'text') {
            const v = String(token.value ?? '');
            txt.textContent = v === '' ? '(leer)' : v;
        } else {
            txt.textContent = tokenLabel(token);
        }

        const x = document.createElement('button');
        x.type = 'button';
        x.className = 'chip-x';
        x.textContent = '✕';
        x.title = 'Baustein entfernen';
        x.addEventListener('click', () => {
            chip.remove();
            ns.teamNamePattern = getPatternFromBuilder();
            setPreviewFromPattern(ns.teamNamePattern);
        });

        chip.append(txt, x);
        zone.appendChild(chip);
    }

    function wireNameBuilderOnce() {
        if (wireNameBuilderOnce._wired) return;
        wireNameBuilderOnce._wired = true;

        const zone = document.getElementById('teamNameBuilder');
        if (!zone) return;

        wireBuilderDnD(zone);

        const btnSep = document.getElementById('teamNameAddSep');
        if (btnSep) {
            btnSep.addEventListener('click', () => {
                const v = document.getElementById('teamNameSepValue')?.value;
                addChip(zone, { type: 'text', value: String(v ?? '') });
                ns.teamNamePattern = getPatternFromBuilder();
                setPreviewFromPattern(ns.teamNamePattern);
            });
        }
        const btnText = document.getElementById('teamNameAddText');
        if (btnText) {
            btnText.addEventListener('click', () => {
                const v = document.getElementById('teamNameTextValue')?.value;
                addChip(zone, { type: 'text', value: String(v ?? '') });
                ns.teamNamePattern = getPatternFromBuilder();
                setPreviewFromPattern(ns.teamNamePattern);
            });
        }
        const btnReset = document.getElementById('teamNameResetDefault');
        if (btnReset) {
            btnReset.addEventListener('click', () => {
                ns.teamNamePattern = defaultTeamNamePattern();
                ns.renderTeamNameBuilder();
            });
        }

        document.querySelectorAll('[data-teamname-token]').forEach(btn => {
            btn.addEventListener('click', () => {
                const type = String(btn.getAttribute('data-teamname-token') || '').trim();
                if (!type) return;
                addChip(zone, { type });
                ns.teamNamePattern = getPatternFromBuilder();
                setPreviewFromPattern(ns.teamNamePattern);
            });
        });

        const yp = document.getElementById('yearPrefix');
        if (yp) yp.addEventListener('input', () => setPreviewFromPattern(getPatternFromBuilder()));
    }

    ns.renderTeamNameBuilder = function renderTeamNameBuilder() {
        const zone = document.getElementById('teamNameBuilder');
        if (!zone) return;
        wireNameBuilderOnce();
        const pattern = normalizePattern(ns.teamNamePattern);
        zone.replaceChildren();
        pattern.forEach(t => addChip(zone, t));
        setPreviewFromPattern(pattern);
    };

    ns.generateTeamNames = function generateTeamNames() {
        const yearPrefix = document.getElementById('yearPrefix').value;
        const emailDomain =
            typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
                ? window.ms365GetTeacherEmailDomainSuffix()
                : '@';
        const separator = document.getElementById('teamSeparator') ? document.getElementById('teamSeparator').value : ' | ';
        // Pattern: aus Builder lesen (falls vorhanden), ansonsten Fallback auf alten Separator-Ansatz
        const pattern = document.getElementById('teamNameBuilder') ? getPatternFromBuilder() : null;

        ns.teamsData = ns.filteredData.map(row => {
            let klasseForName = row.klasse;
            if (row.klasse.includes(',')) klasseForName = ns.combineClassNames(row.klasse);

            const teamName = pattern
                ? buildTeamNameFromPattern(pattern, { yearPrefix, klasse: klasseForName, fach: row.fach, gruppe: row.gruppe })
                : `${yearPrefix}${separator}${klasseForName}${separator}${row.fach}`;
            const gruppenmailRaw = ns.buildGruppenmailBase(yearPrefix, klasseForName, row.fach, row.gruppe).replace(/\s+/g, '-');

            const originalGruppenmail = gruppenmailRaw;
            let gruppenmail = gruppenmailRaw.replace(ns.INVALID_CHARS_REPLACE, '');

            let besitzer = '';
            const lehrerCode = row.lehrer.toUpperCase().trim();
            if (ns.teacherEmailMapping[lehrerCode]) {
                besitzer = ns.teacherEmailMapping[lehrerCode];
            } else {
                besitzer = row.lehrer.toLowerCase().trim().replace(/\s+/g, '.');
                besitzer = besitzer.replace(ns.INVALID_CHARS_REPLACE, '');
                if (!besitzer.includes('@')) besitzer += emailDomain;
            }

            const hasInvalidChars = ns.INVALID_CHARS_TEST.test(originalGruppenmail);
            const isValid = !hasInvalidChars && teamName && gruppenmail && besitzer && gruppenmail.length > 0;
            const mappingUsed = !!ns.teacherEmailMapping[lehrerCode];

            return {
                teamName,
                gruppenmail,
                besitzer,
                isValid,
                error: hasInvalidChars ? 'Ungültige Zeichen in Gruppenmail' : (!isValid ? 'Unvollständige Daten' : null),
                originalClass: row.klasse,
                gruppe: row.gruppe,
                mappingUsed,
                lehrerCode,
                mailNicknameAdjusted: false
            };
        });

        const dupCount = ns.resolveDuplicateGruppenmails(ns.teamsData);
        document.getElementById('duplicateMailAdjustments').textContent = dupCount;
        ns.teamsGenerated = true;
        ns.displayTeamsData();
    };

    ns.displayTeamsData = function displayTeamsData() {
        const tbody = document.getElementById('teamsTableBody');
        tbody.replaceChildren();

        const validCount = ns.teamsData.filter(t => t.isValid).length;
        const invalidCount = ns.teamsData.length - validCount;
        const mappedCount = ns.teamsData.filter(t => t.mappingUsed).length;
        const dupAdj = ns.teamsData.filter(t => t.mailNicknameAdjusted).length;
        document.getElementById('duplicateMailAdjustments').textContent = dupAdj;

        if (!ns.teamsSort) ns.teamsSort = { key: 'teamName', dir: 1 };

        const getSortVal = (t, key) => {
            if (key === 'status') return t.isValid ? '1' : '0';
            return String(t[key] ?? '').toUpperCase();
        };

        const view = ns.teamsData
            .map((team, index) => ({ team, index }))
            .sort((a, b) => {
                const ak = getSortVal(a.team, ns.teamsSort.key);
                const bk = getSortVal(b.team, ns.teamsSort.key);
                const cmp = ak.localeCompare(bk, 'de');
                if (cmp !== 0) return cmp * ns.teamsSort.dir;
                return (a.index - b.index) * ns.teamsSort.dir;
            });

        const table = document.getElementById('teamsTableContainer');
        const ths = table ? table.querySelectorAll('th[data-teams-sort-key]') : [];
        ths.forEach(th => {
            const base = th.dataset.label || th.textContent.replace(/[▲▼]\s*$/, '').trim();
            th.dataset.label = base;
            th.textContent = base;
            if (th.dataset.teamsSortKey === ns.teamsSort.key) {
                const ind = document.createElement('span');
                ind.className = 'kt-sort-indicator';
                ind.textContent = ns.teamsSort.dir === -1 ? '▼' : '▲';
                th.appendChild(ind);
            }
            if (!th.dataset.wired) {
                th.dataset.wired = '1';
                th.addEventListener('click', () => {
                    const k = th.dataset.teamsSortKey;
                    if (ns.teamsSort.key === k) ns.teamsSort.dir = ns.teamsSort.dir === 1 ? -1 : 1;
                    else ns.teamsSort = { key: k, dir: 1 };
                    ns.displayTeamsData();
                });
            }
        });

        view.forEach(({ team, index }) => {
            const tr = document.createElement('tr');
            if (!team.isValid) tr.classList.add('error-row');

            const td1 = document.createElement('td');
            td1.appendChild(document.createTextNode(team.teamName));
            td1.addEventListener('dblclick', () => ns.editTeam(index));
            if (team.originalClass && team.originalClass.includes(',')) {
                td1.appendChild(document.createElement('br'));
                const small = document.createElement('small');
                small.style.color = '#6c757d';
                small.textContent = '(Original: ' + team.originalClass + ')';
                td1.appendChild(small);
            }

            const td2 = document.createElement('td');
            td2.appendChild(document.createTextNode(team.gruppenmail));
            td2.addEventListener('dblclick', () => ns.editTeam(index));
            if (team.mailNicknameAdjusted) {
                td2.appendChild(document.createElement('br'));
                const small = document.createElement('small');
                small.style.color = '#ff9800';
                small.textContent = '(Mail-Nickname wegen Duplikat angepasst)';
                td2.appendChild(small);
            }
            if (team.gruppe) {
                td2.appendChild(document.createElement('br'));
                const small = document.createElement('small');
                small.style.color = '#6c757d';
                small.textContent = 'Gruppe: ' + team.gruppe;
                td2.appendChild(small);
            }

            const td3 = document.createElement('td');
            td3.appendChild(document.createTextNode(team.besitzer));
            td3.addEventListener('dblclick', () => ns.editTeam(index));
            td3.appendChild(document.createElement('br'));
            const smallM = document.createElement('small');
            smallM.style.color = team.mappingUsed ? '#28a745' : '#ffc107';
            smallM.textContent = team.mappingUsed ? '✓ Mapping' : '⚠ Generiert (' + team.lehrerCode + ')';
            td3.appendChild(smallM);

            const td4 = document.createElement('td');
            td4.textContent = team.isValid ? '✅' : '❌ ' + (team.error || 'Fehler');
            td4.addEventListener('dblclick', () => ns.editTeam(index));

            const td5 = document.createElement('td');
            const b1 = document.createElement('button');
            b1.type = 'button';
            b1.className = 'btn btn-small';
            b1.textContent = '✏️';
            b1.addEventListener('click', () => ns.editTeam(index));
            const b2 = document.createElement('button');
            b2.type = 'button';
            b2.className = 'btn btn-small btn-danger';
            b2.textContent = '🗑️';
            b2.addEventListener('click', () => ns.deleteTeam(index));
            td5.append(b1, b2);

            tr.append(td1, td2, td3, td4, td5);
            tbody.appendChild(tr);
        });

        document.getElementById('validTeams').textContent = validCount;
        document.getElementById('invalidTeams').textContent = invalidCount;

        const existingWarning = document.getElementById('mappingWarning');
        if (existingWarning) existingWarning.remove();
        if (mappedCount < ns.teamsData.length) {
            const unmappedCount = ns.teamsData.length - mappedCount;
            const warning = document.createElement('div');
            warning.id = 'mappingWarning';
            warning.className = 'alert alert-warning';
            const strong = document.createElement('strong');
            strong.textContent = '⚠️ Achtung: ';
            warning.appendChild(strong);
            warning.appendChild(
                document.createTextNode(
                    unmappedCount + ' Lehrer haben keine E-Mail-Zuordnung. Die E-Mail-Adressen wurden automatisch generiert.'
                )
            );
            document.getElementById('validationResults').appendChild(warning);
        }

        document.getElementById('teamsTableContainer').style.display = 'block';
        document.getElementById('validationResults').style.display = 'block';
        const cont = document.getElementById('continueBtn4') || document.getElementById('continueBtn3');
        if (cont) cont.style.display = 'inline-block';
    };

    ns.editTeam = function editTeam(index) {
        const team = ns.teamsData[index];
        ns.openModal(
            'Team bearbeiten',
            '<label for="editName">Team-Name</label><input type="text" id="editName" value="' +
                ns.attrEscape(team.teamName) +
                '">' +
                '<label for="editMail">Gruppenmail</label><input type="text" id="editMail" value="' +
                ns.attrEscape(team.gruppenmail) +
                '">' +
                '<label for="editOwner">Besitzer</label><input type="email" id="editOwner" value="' +
                ns.attrEscape(team.besitzer) +
                '">',
            () => {
                const newName = document.getElementById('editName').value.trim();
                const newMail = document.getElementById('editMail').value.trim();
                const newOwner = document.getElementById('editOwner').value.trim();
                if (!newName || !newMail || !newOwner) {
                    ns.showToast('Bitte alle Felder ausfüllen.');
                    return;
                }
                ns.teamsData[index] = { ...team, teamName: newName, gruppenmail: newMail, besitzer: newOwner, isValid: true, error: null };
                ns.closeModal();
                ns.displayTeamsData();
            }
        );
    };

    ns.deleteTeam = function deleteTeam(index) {
        ns.confirmModal('Team löschen', 'Dieses Team wirklich aus der Liste entfernen?', () => {
            ns.teamsData.splice(index, 1);
            if (ns.teamsData.length === 0) ns.teamsGenerated = false;
            ns.displayTeamsData();
        });
    };

    ns.addManualKursteamTeam = function addManualKursteamTeam() {
        ns.openModal(
            'Team manuell hinzufügen',
            '<label for="addKtName">Team-Name</label><input type="text" id="addKtName" autocomplete="off" placeholder="z. B. WS26 | 1A | D">' +
                '<label for="addKtMail">Gruppenmail (Nickname)</label><input type="text" id="addKtMail" autocomplete="off" placeholder="z. B. WS26-1A-D">' +
                '<label for="addKtOwner">Besitzer (E-Mail)</label><input type="email" id="addKtOwner" autocomplete="off">',
            () => {
                const teamName = document.getElementById('addKtName').value.trim();
                const gruppenmailRaw = document.getElementById('addKtMail').value.trim();
                const besitzer = document.getElementById('addKtOwner').value.trim().toLowerCase();
                if (!teamName || !gruppenmailRaw || !besitzer) {
                    ns.showToast('Bitte alle Felder ausfüllen.');
                    return;
                }
                const originalGruppenmail = gruppenmailRaw;
                const gruppenmail = gruppenmailRaw.replace(ns.INVALID_CHARS_REPLACE, '');
                const hasInvalidChars = ns.INVALID_CHARS_TEST.test(originalGruppenmail);
                const isValid = !hasInvalidChars && gruppenmail.length > 0;
                ns.teamsData.push({
                    teamName,
                    gruppenmail,
                    besitzer,
                    isValid,
                    error: hasInvalidChars ? 'Ungültige Zeichen in Gruppenmail' : !isValid ? 'Unvollständige Daten' : null,
                    originalClass: '',
                    gruppe: '',
                    mappingUsed: true,
                    lehrerCode: '',
                    mailNicknameAdjusted: false
                });
                ns.teamsGenerated = true;
                ns.closeModal();
                ns.displayTeamsData();
                ns.showToast('Team hinzugefügt.');
            }
        );
    };

    // Global exports für HTML onclick
    window.startKursteamFromWebuntis = ns.startKursteamFromWebuntis;
    window.startKursteamManual = ns.startKursteamManual;
    window.addManualDataRow = ns.addManualDataRow;
    window.addManualDataRowInline = ns.addManualDataRowInline;
    window.applyFilters = ns.applyFilters;
    window.resetFilters = ns.resetFilters;
    window.generateTeamNames = ns.generateTeamNames;
    window.addManualKursteamTeam = ns.addManualKursteamTeam;

    // Beim initialen Laden einmal aufbauen (falls Schritt 2 schon gerendert ist)
    if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
    if (typeof ns.renderTeamNameBuilder === 'function') ns.renderTeamNameBuilder();
})();

