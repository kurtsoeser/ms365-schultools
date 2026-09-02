(function () {
    'use strict';

    function getEl(id) {
        return document.getElementById(id);
    }

    function toast(msg, kind) {
        const el = getEl('toast');
        if (!el) return;
        el.textContent = String(msg || '');
        el.className = 'toast' + (kind === 'ok' ? ' ok' : kind === 'err' ? ' err' : '');
        el.style.display = 'block';
        clearTimeout(toast._t);
        toast._t = setTimeout(function () {
            el.style.display = 'none';
        }, 3200);
    }

    function compareDe(a, b) {
        return String(a || '').localeCompare(String(b || ''), 'de', { numeric: true, sensitivity: 'base' });
    }

    const state = {
        tab: 'class', // class | year
        selectedKey: '',
        selectedKeys: new Set(),
        lastScriptLabel: '',
        classRows: [],
        yearRows: [],
        patterns: {
            classAlias: [],
            classDisplay: [],
            yearAlias: [],
            yearDisplay: []
        }
    };

    const PATTERN_KEYS = {
        classAlias: 'elternClassAliasPattern',
        classDisplay: 'elternClassDisplayPattern',
        yearAlias: 'elternYearAliasPattern',
        yearDisplay: 'elternYearDisplayPattern'
    };

    function api() {
        return window.ms365AppDataV2 || null;
    }

    function eg() {
        return window.ms365ElternGuardians || null;
    }

    function loadPatternsFromSetup() {
        const g = eg();
        const naming = g && g.getNaming ? g.getNaming() : null;
        if (naming) {
            state.patterns.classAlias = naming.classAliasPattern;
            state.patterns.classDisplay = naming.classDisplayPattern;
            state.patterns.yearAlias = naming.yearAliasPattern;
            state.patterns.yearDisplay = naming.yearDisplayPattern;
            return;
        }
        state.patterns.classAlias = [
            { type: 'text', value: 'eltern' },
            { type: 'klasse' }
        ];
        state.patterns.classDisplay = [
            { type: 'text', value: 'Eltern ' },
            { type: 'klasse' }
        ];
        state.patterns.yearAlias = [
            { type: 'text', value: 'elternjg' },
            { type: 'year' }
        ];
        state.patterns.yearDisplay = [
            { type: 'text', value: 'Eltern JG ' },
            { type: 'year' }
        ];
    }

    function namingFromState() {
        return {
            classAliasPattern: state.patterns.classAlias,
            classDisplayPattern: state.patterns.classDisplay,
            yearAliasPattern: state.patterns.yearAlias,
            yearDisplayPattern: state.patterns.yearDisplay
        };
    }

    function patternFromZone(key) {
        const zone = document.querySelector('[data-ev-zone="' + key + '"]');
        const g = eg();
        if (!zone || !g) return state.patterns[key] || [];
        const tokens = [];
        zone.querySelectorAll('[data-token-type]').forEach(function (el) {
            const type = String(el.getAttribute('data-token-type') || '');
            if (type === 'text') tokens.push({ type: 'text', value: String(el.getAttribute('data-token-value') || '') });
            else tokens.push({ type: type });
        });
        return g.normalizeNamePattern(tokens, state.patterns[key] || []);
    }

    function syncPatternFromZone(key) {
        state.patterns[key] = patternFromZone(key);
        updatePatternPreview(key);
    }

    function updatePatternPreview(key) {
        const el = document.querySelector('[data-ev-preview="' + key + '"]');
        const g = eg();
        if (!el || !g) return;
        const pattern = state.patterns[key] || [];
        const forAlias = key.indexOf('Alias') !== -1;
        const sample =
            key.indexOf('year') === 0
                ? g.buildNameFromPattern(pattern, { year: '2030', forAlias: forAlias })
                : g.buildNameFromPattern(pattern, { klasse: '1A', year: '2030', forAlias: forAlias });
        el.textContent = 'Vorschau: ' + sample;
    }

    function addChip(zone, token) {
        const g = eg();
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
            txt.textContent = g && g.tokenLabel ? g.tokenLabel(token) : token.type;
        }

        const x = document.createElement('button');
        x.type = 'button';
        x.className = 'chip-x';
        x.textContent = '✕';
        x.title = 'Baustein entfernen';
        x.addEventListener('click', function () {
            const key = zone.getAttribute('data-ev-zone');
            chip.remove();
            if (key) syncPatternFromZone(key);
        });

        chip.append(txt, x);
        zone.appendChild(chip);
    }

    function wireZoneDnD(zone) {
        let dragEl = null;
        zone.addEventListener('dragstart', function (e) {
            const target = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
            if (!target) return;
            dragEl = target;
            target.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        });
        zone.addEventListener('dragend', function () {
            if (dragEl) dragEl.classList.remove('dragging');
            dragEl = null;
            const key = zone.getAttribute('data-ev-zone');
            if (key) syncPatternFromZone(key);
        });
        zone.addEventListener('dragover', function (e) {
            e.preventDefault();
            const over = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
            if (!dragEl || !over || over === dragEl) return;
            const rect = over.getBoundingClientRect();
            const after = e.clientX > rect.left + rect.width / 2;
            if (after) over.after(dragEl);
            else over.before(dragEl);
        });
        zone.addEventListener('drop', function (e) {
            e.preventDefault();
            const key = zone.getAttribute('data-ev-zone');
            if (key) syncPatternFromZone(key);
        });
    }

    function renderPatternBuilder(key) {
        const zone = document.querySelector('[data-ev-zone="' + key + '"]');
        if (!zone) return;
        const pattern = state.patterns[key] || [];
        zone.replaceChildren();
        pattern.forEach(function (t) {
            addChip(zone, t);
        });
        updatePatternPreview(key);
    }

    function renderAllPatternBuilders() {
        Object.keys(PATTERN_KEYS).forEach(renderPatternBuilder);
    }

    function defaultPatternFor(key) {
        const g = eg();
        if (!g) return [];
        if (key === 'classAlias') return g.defaultClassAliasPattern();
        if (key === 'classDisplay') return g.defaultClassDisplayPattern();
        if (key === 'yearAlias') return g.defaultYearAliasPattern();
        if (key === 'yearDisplay') return g.defaultYearDisplayPattern();
        return [];
    }

    function saveNamingSchema() {
        const a = api();
        if (!a || typeof a.patchSetup !== 'function') {
            toast('Speichern nicht möglich', 'err');
            return;
        }
        Object.keys(PATTERN_KEYS).forEach(function (key) {
            syncPatternFromZone(key);
        });
        const patch = {};
        Object.keys(PATTERN_KEYS).forEach(function (key) {
            patch[PATTERN_KEYS[key]] = state.patterns[key];
        });
        a.patchSetup(patch);
        refresh();
        toast('Namensschema gespeichert', 'ok');
    }

    function applySisRecords(records) {
        const a = api();
        if (!a || typeof a.mergeStudentsImport !== 'function' || typeof a.saveYearBucket !== 'function') {
            throw new Error('Stammdaten-API fehlt.');
        }
        const { year, bucket } = currentBucket();
        const gById = new Map(
            (bucket.guardians || []).map(function (g) {
                return [g.id, g];
            })
        );
        function keyOf(r) {
            const em = String(r.email || '')
                .trim()
                .toLowerCase();
            if (em) return 'e:' + em;
            const ext = String(r.externalId || '')
                .trim()
                .toLowerCase();
            if (ext) return 'x:' + ext;
            return (
                'n:' +
                String(r.klasse || '')
                    .trim()
                    .toLowerCase() +
                '|' +
                String(r.name || '')
                    .trim()
                    .toLowerCase()
            );
        }
        const map = new Map();
        (bucket.students || []).forEach(function (s) {
            const pairs = (s.guardianIds || [])
                .map(function (id) {
                    const g = gById.get(id);
                    return g ? { name: g.name || '', email: g.email || '' } : null;
                })
                .filter(Boolean);
            const row = {
                id: s.id,
                klasse: s.klasse,
                name: s.name,
                email: s.email,
                guardianIds: s.guardianIds,
                parentPairs: pairs
            };
            map.set(keyOf(row), row);
        });
        (records || []).forEach(function (r) {
            const row = {
                klasse: r.klasse,
                name: r.name,
                email: r.email,
                externalId: r.externalId,
                parentPairs: r.parentPairs || []
            };
            const k = keyOf(row);
            const prev = map.get(k);
            if (prev) {
                map.set(k, {
                    id: prev.id,
                    klasse: row.klasse || prev.klasse,
                    name: row.name || prev.name,
                    email: row.email || prev.email,
                    guardianIds: prev.guardianIds,
                    parentPairs: row.parentPairs && row.parentPairs.length ? row.parentPairs : prev.parentPairs
                });
            } else {
                map.set(k, row);
            }
        });
        const next = a.mergeStudentsImport(
            {
                students: [],
                classes: bucket.classes || [],
                guardians: [],
                parentLists: bucket.parentLists || []
            },
            Array.from(map.values())
        );
        next.classes = bucket.classes || [];
        next.parentLists = bucket.parentLists || [];
        a.saveYearBucket(year, next);
        const sis = window.ms365SchoolSisImport;
        if (sis && typeof sis.diffSisImport === 'function') {
            const existing = (bucket.students || []).map(function (s) {
                const gBy = new Map((bucket.guardians || []).map(function (g) { return [g.id, g]; }));
                const pairs = (s.guardianIds || [])
                    .map(function (id) {
                        const g = gBy.get(id);
                        return g ? { name: g.name || '', email: g.email || '' } : null;
                    })
                    .filter(Boolean);
                return {
                    klasse: s.klasse,
                    name: s.name,
                    email: s.email,
                    parentPairs: pairs
                };
            });
            const diff = sis.diffSisImport(existing, records || []);
            if (sis.summarizeSisDiff) return { next: next, diff: diff, summary: sis.summarizeSisDiff(diff) };
            return { next: next, diff: diff };
        }
        return { next: next };
    }

    function domainFromCore() {
        try {
            const a = api();
            const c = a && typeof a.getContainer === 'function' ? a.getContainer() : null;
            return String((c && c.core && c.core.domain) || '').trim();
        } catch {
            return '';
        }
    }

    function renderDiagnose() {
        const eg = window.ms365ElternGuardians;
        const summary = getEl('evDiagnoseSummary');
        const ul = getEl('evDiagnoseIssues');
        const hints = getEl('evDiagnoseHints');
        if (!eg || typeof eg.buildElternDiagnoseReport !== 'function') return;
        const { bucket } = currentBucket();
        const report = eg.buildElternDiagnoseReport(bucket, eg.getNaming(), domainFromCore());
        if (summary) {
            const c = report.counts || {};
            summary.textContent =
                (report.ok ? 'Keine Warnungen. ' : 'Bitte prüfen: ') +
                (c.withParents || 0) +
                ' von ' +
                (c.lists || 0) +
                ' Listen mit Elternmails' +
                (c.exported ? ' · ' + c.exported + ' zuletzt exportiert' : '');
        }
        if (ul) {
            ul.innerHTML = '';
            (report.issues || []).forEach(function (iss) {
                const li = document.createElement('li');
                li.textContent = (iss.level === 'warn' ? 'Warnung: ' : '') + iss.summary;
                ul.appendChild(li);
            });
            if (!ul.childNodes.length) {
                const li = document.createElement('li');
                li.textContent = 'Alias-Kollisionen: keine. GAL: Listen sichtbar, Contacts versteckt (vom Sync-Skript).';
                ul.appendChild(li);
            }
        }
        if (hints && report.hints) {
            hints.textContent = report.hints.gal + ' ' + report.hints.contacts + ' ' + report.hints.naming;
        }
        try {
            const a = api();
            if (a && typeof a.patchSetup === 'function' && typeof a.getSetup === 'function') {
                const cur = a.getSetup() || {};
                const es = cur.elternSetup && typeof cur.elternSetup === 'object' ? cur.elternSetup : {};
                a.patchSetup({
                    elternSetup: {
                        completedSteps: Array.isArray(es.completedSteps) ? es.completedSteps : [],
                        lastDiagnoseAt: new Date().toISOString()
                    }
                });
            }
        } catch {
            /* ignore */
        }
        return report;
    }

    function setSetupStep(n) {
        const step = String(n || '1');
        document.querySelectorAll('#evSetupSteps [data-ev-step]').forEach(function (btn) {
            btn.setAttribute('aria-current', btn.getAttribute('data-ev-step') === step ? 'step' : 'false');
        });
        const importPanel = getEl('evImportPanel');
        const namingPanel = getEl('evNamingPanel');
        if (importPanel) importPanel.open = step === '1';
        if (namingPanel) namingPanel.open = step === '2';
        const diag = getEl('evDiagnosePanel');
        if (diag && step === '3') {
            renderDiagnose();
            if (typeof diag.scrollIntoView === 'function') diag.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        const listTitle = getEl('evListTitle');
        if (step === '4' && listTitle && typeof listTitle.scrollIntoView === 'function') {
            listTitle.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }

    function wireSisImport() {
        const fileInput = getEl('evImportFile');
        const status = getEl('evImportStatus');
        const sourceEl = getEl('evImportSource');
        function setStatus(msg) {
            if (status) status.textContent = msg || '';
        }
        if (getEl('evTplXlsx')) {
            getEl('evTplXlsx').addEventListener('click', function () {
                const sis = window.ms365SchoolSisImport;
                if (!sis || !sis.downloadXlsxTemplates || !sis.downloadXlsxTemplates()) {
                    toast('XLSX-Vorlage: Bibliothek fehlt – Seite neu laden', 'err');
                    return;
                }
                toast('XLSX-Vorlage heruntergeladen', 'ok');
            });
        }
        if (getEl('evTplCsv')) {
            getEl('evTplCsv').addEventListener('click', function () {
                const sis = window.ms365SchoolSisImport;
                if (!sis || !sis.downloadCsv) {
                    toast('CSV-Vorlage nicht verfügbar', 'err');
                    return;
                }
                sis.downloadCsv('Schueler-Eltern-Vorlage.csv', sis.ms365TemplateAoa());
                toast('CSV-Vorlage heruntergeladen', 'ok');
            });
        }
        if (!fileInput) return;
        fileInput.addEventListener('change', function () {
            const file = fileInput.files && fileInput.files[0];
            fileInput.value = '';
            if (!file) return;
            if (typeof XLSX === 'undefined') {
                toast('Excel-Bibliothek nicht geladen', 'err');
                return;
            }
            const sis = window.ms365SchoolSisImport;
            if (!sis) {
                toast('Import-Modul fehlt', 'err');
                return;
            }
            const sourceHint = sourceEl ? String(sourceEl.value || 'auto') : 'auto';
            setStatus('Lese Datei …');
            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const name = String(file.name || '').toLowerCase();
                    let wb;
                    if (name.endsWith('.csv') || name.endsWith('.txt')) {
                        let s = String(e.target.result || '');
                        if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
                        wb = XLSX.read(s, { type: 'string', FS: ';' });
                        let aoaProbe = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
                        if (!aoaProbe || aoaProbe.length < 2) wb = XLSX.read(s, { type: 'string', FS: ',' });
                    } else {
                        wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                    }
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                    const objectRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                    const result = sis.importStudentsAndGuardians({
                        aoa: aoa,
                        objectRows: objectRows,
                        source: sourceHint
                    });
                    const applied = applySisRecords(result.records);
                    refresh();
                    const extra = applied && applied.summary ? ' · ' + applied.summary : '';
                    setStatus(
                        'Import OK: ' +
                            result.meta.studentCount +
                            ' Schüler, ' +
                            result.meta.withParents +
                            ' mit Elternmails (Quelle: ' +
                            result.source +
                            ')' +
                            extra +
                            '.'
                    );
                    toast('Import übernommen', 'ok');
                } catch (err) {
                    setStatus('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)));
                    toast('Import fehlgeschlagen', 'err');
                }
            };
            reader.onerror = function () {
                setStatus('Datei konnte nicht gelesen werden.');
            };
            const n = String(file.name || '').toLowerCase();
            if (n.endsWith('.csv') || n.endsWith('.txt')) reader.readAsText(file);
            else reader.readAsArrayBuffer(file);
        });
    }

    function wireNamingBuilders() {
        document.querySelectorAll('[data-ev-zone]').forEach(function (zone) {
            wireZoneDnD(zone);
        });
        document.querySelectorAll('[data-ev-add-token]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const key = btn.getAttribute('data-ev-for');
                const type = btn.getAttribute('data-ev-add-token');
                const zone = document.querySelector('[data-ev-zone="' + key + '"]');
                if (!zone || !type) return;
                addChip(zone, { type: type });
                syncPatternFromZone(key);
            });
        });
        document.querySelectorAll('[data-ev-add-sep]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const key = btn.getAttribute('data-ev-add-sep');
                const inp = document.querySelector('[data-ev-sep="' + key + '"]');
                const zone = document.querySelector('[data-ev-zone="' + key + '"]');
                if (!zone) return;
                addChip(zone, { type: 'text', value: String((inp && inp.value) ?? '') });
                syncPatternFromZone(key);
            });
        });
        document.querySelectorAll('[data-ev-reset]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const key = btn.getAttribute('data-ev-reset');
                state.patterns[key] = defaultPatternFor(key);
                renderPatternBuilder(key);
            });
        });
        const saveBtn = getEl('evNamingSave');
        if (saveBtn) saveBtn.addEventListener('click', saveNamingSchema);
    }

    function currentBucket() {
        const a = api();
        if (!a || typeof a.getYearBucket !== 'function') {
            return { year: '', bucket: { students: [], classes: [], guardians: [], parentLists: [] } };
        }
        return a.getYearBucket();
    }

    function domain() {
        try {
            if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
                return String(window.ms365GetSchoolDomainNoAt() || '').replace(/^@+/, '');
            }
        } catch {
            /* ignore */
        }
        const a = api();
        const c = a && a.getContainer ? a.getContainer() : null;
        return String((c && c.core && c.core.domain) || '').replace(/^@+/, '');
    }

    function schoolName() {
        const a = api();
        const c = a && a.getContainer ? a.getContainer() : null;
        return String((c && c.core && c.core.schoolName) || '').trim();
    }

    function ensureScriptPrerequisites() {
        const dom = domain();
        if (!dom) {
            toast('Schul-Domain fehlt in den Stammdaten – bitte zuerst setzen.', 'err');
            return false;
        }
        return true;
    }

    function reloadSoll() {
        const { bucket } = currentBucket();
        const g = eg();
        const naming = namingFromState();
        state.classRows = g && g.buildClassParentSoll ? g.buildClassParentSoll(bucket, { naming: naming }) : [];
        state.yearRows = g && g.buildYearParentSoll ? g.buildYearParentSoll(bucket, { naming: naming }) : [];
    }

    function rowsForTab() {
        return state.tab === 'year' ? state.yearRows : state.classRows;
    }

    function rowKey(row) {
        return String(row.scope || 'class') + ':' + String(row.code || '');
    }

    function findRow(key) {
        return rowsForTab().find(function (r) {
            return rowKey(r) === key;
        }) || null;
    }

    function pruneSelectedKeys() {
        const valid = new Set(rowsForTab().map(rowKey));
        Array.from(state.selectedKeys).forEach(function (key) {
            if (!valid.has(key)) state.selectedKeys.delete(key);
        });
    }

    function selectedCountWithParents() {
        let n = 0;
        state.selectedKeys.forEach(function (key) {
            const r = findRow(key);
            if (r && r.guardianCount > 0) n += 1;
        });
        return n;
    }

    function setListSelected(key, on) {
        if (on) state.selectedKeys.add(key);
        else state.selectedKeys.delete(key);
        const sel = getEl('evSelectList');
        if (sel && key === state.selectedKey) sel.checked = !!on;
        updateListMeta();
    }

    function selectAllWithParents() {
        rowsForTab().forEach(function (r) {
            if (r.guardianCount > 0) state.selectedKeys.add(rowKey(r));
        });
        renderList();
        renderDetail();
        toast(selectedCountWithParents() + ' Listen ausgewählt', 'ok');
    }

    function clearSelection() {
        state.selectedKeys.clear();
        renderList();
        renderDetail();
        toast('Auswahl geleert', 'ok');
    }

    function updateListMeta() {
        const meta = getEl('evListMeta');
        if (!meta) return;
        const withParents = rowsForTab().filter(function (r) {
            return r.guardianCount > 0;
        }).length;
        const selected = selectedCountWithParents();
        meta.textContent =
            rowsForTab().length +
            ' Listen · ' +
            withParents +
            ' mit Elternmails · ' +
            selected +
            ' ausgewählt · Schuljahr ' +
            (currentBucket().year || '–');
    }

    function showScriptForLists(lists, toastMsg) {
        if (!lists.length) return;
        if (!ensureScriptPrerequisites()) return;
        const keepKey = lists.some(function (r) {
            return rowKey(r) === state.selectedKey;
        });
        if (!keepKey) state.selectedKey = rowKey(lists[0]);
        renderList();
        renderDetail();
        const script = buildScriptForLists(lists);
        setScript(script);
        markExported(lists);
        state.lastScriptLabel =
            lists.length > 1
                ? 'sammel-' + lists.length
                : lists[0].mailNickname || lists[0].code || 'sync';
        toast(toastMsg || 'Skript erzeugt', 'ok');
    }

    function renderList() {
        const ul = getEl('evList');
        const title = getEl('evListTitle');
        if (!ul) return;
        pruneSelectedKeys();
        const q = String((getEl('evSearch') && getEl('evSearch').value) || '')
            .trim()
            .toLowerCase();
        const rows = rowsForTab().filter(function (r) {
            if (!q) return true;
            const hay = [r.code, r.displayName, r.mailNickname, ...(r.classCodes || [])].join(' ').toLowerCase();
            return hay.indexOf(q) !== -1;
        });
        ul.replaceChildren();
        if (title) {
            title.innerHTML =
                state.tab === 'year'
                    ? '<i class="bi bi-calendar3" style="margin-right:8px;"></i>Jahrgangs-Elternlisten'
                    : '<i class="bi bi-list-ul" style="margin-right:8px;"></i>Klassen-Elternlisten';
        }
        updateListMeta();
        if (!rows.length) {
            const li = document.createElement('li');
            li.innerHTML =
                '<div class="muted" style="padding:12px;">Keine Einträge. Schüler und optional Eltern in den <a href="../tenant.html">Stammdaten</a> pflegen.</div>';
            ul.appendChild(li);
            return;
        }
        rows.forEach(function (r) {
            const li = document.createElement('li');
            const row = document.createElement('div');
            row.className = 'tree-row';
            const key = rowKey(r);
            const isCurrent = key === state.selectedKey;
            if (isCurrent) row.setAttribute('data-current', 'true');

            const checkLabel = document.createElement('label');
            checkLabel.className = 'tree-check';
            checkLabel.title =
                r.guardianCount > 0
                    ? 'Für Sammel-Skript auswählen'
                    : 'Keine Elternmails – nicht auswählbar';
            const check = document.createElement('input');
            check.type = 'checkbox';
            check.setAttribute('aria-label', 'Auswählen: ' + (r.displayName || r.code));
            check.disabled = !(r.guardianCount > 0);
            check.checked = state.selectedKeys.has(key);
            check.addEventListener('click', function (ev) {
                ev.stopPropagation();
            });
            check.addEventListener('change', function () {
                setListSelected(key, check.checked);
            });
            checkLabel.appendChild(check);

            const btn = document.createElement('button');
            btn.type = 'button';
            if (isCurrent) btn.setAttribute('aria-current', 'true');
            const pillClass = r.guardianCount > 0 ? 'ok' : 'warn';
            btn.innerHTML =
                '<span style="font-weight:900;color:#32325d;">' +
                escapeHtml(r.displayName || r.code) +
                '</span>' +
                '<span class="pill ' +
                pillClass +
                '">' +
                r.guardianCount +
                ' Eltern</span>' +
                (state.tab === 'class'
                    ? '<span class="pill">' + r.studentCount + ' Schüler</span>'
                    : '<span class="pill">' + (r.classCodes || []).length + ' Klassen</span>') +
                '<span class="muted" style="font-size:0.82em;">' +
                escapeHtml(r.mailNickname || '') +
                '</span>';
            btn.addEventListener('click', function () {
                state.selectedKey = key;
                renderList();
                renderDetail();
            });

            row.appendChild(checkLabel);
            row.appendChild(btn);
            li.appendChild(row);
            ul.appendChild(li);
        });
    }

    function escapeHtml(s) {
        return String(s || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function guardiansForStudent(student, byId) {
        return (student.guardianIds || [])
            .map(function (id) {
                return byId.get(String(id));
            })
            .filter(Boolean);
    }

    function renderDetail() {
        const empty = getEl('evDetailEmpty');
        const detail = getEl('evDetail');
        const wrap = getEl('evStudentsWrap');
        const summary = getEl('evDetailSummary');
        const title = getEl('evDetailTitle');
        const sel = getEl('evSelectList');
        const row = findRow(state.selectedKey);
        if (!row) {
            if (empty) empty.hidden = false;
            if (detail) detail.hidden = true;
            return;
        }
        if (empty) empty.hidden = true;
        if (detail) detail.hidden = false;
        if (title) title.textContent = row.displayName || row.code;
        if (summary) {
            summary.textContent =
                'Alias ' +
                (row.mailNickname || '–') +
                ' · ' +
                row.guardianCount +
                ' eindeutige Elternmail(s)' +
                (state.tab === 'year' ? ' · Klassen: ' + (row.classCodes || []).join(', ') : '');
        }
        if (sel) {
            sel.checked = state.selectedKeys.has(state.selectedKey);
            sel.disabled = !(row.guardianCount > 0);
            sel.onchange = function () {
                setListSelected(state.selectedKey, sel.checked);
                renderList();
            };
        }

        if (!wrap) return;
        wrap.replaceChildren();

        if (state.tab === 'year') {
            const p = document.createElement('p');
            p.className = 'muted';
            p.textContent =
                'Jahrgangslisten aggregieren die Elternmails aller Klassen mit diesem Abschlussjahr. Zuordnung der Eltern erfolgt pro Schüler in der Klassenansicht.';
            wrap.appendChild(p);
            const ul = document.createElement('ul');
            (row.guardians || []).forEach(function (g) {
                const li = document.createElement('li');
                li.textContent = (g.name ? g.name + ' – ' : '') + g.email;
                ul.appendChild(li);
            });
            if (!(row.guardians || []).length) {
                const li = document.createElement('li');
                li.className = 'muted';
                li.textContent = 'Noch keine Elternmails.';
                ul.appendChild(li);
            }
            wrap.appendChild(ul);
            return;
        }

        const { bucket } = currentBucket();
        const byId = new Map(
            (bucket.guardians || []).map(function (g) {
                return [g.id, g];
            })
        );
        const students = (bucket.students || [])
            .filter(function (s) {
                return String(s.klasse || '').trim().toUpperCase() === String(row.code).toUpperCase();
            })
            .slice()
            .sort(function (a, b) {
                return compareDe(a.name, b.name);
            });

        if (!students.length) {
            const p = document.createElement('p');
            p.className = 'muted';
            p.textContent = 'Keine Schüler in dieser Klasse.';
            wrap.appendChild(p);
            return;
        }

        students.forEach(function (stu) {
            const card = document.createElement('div');
            card.className = 'student-card';
            const h = document.createElement('h3');
            h.textContent = (stu.name || 'Ohne Name') + (stu.email ? ' · ' + stu.email : '');
            card.appendChild(h);

            const gs = guardiansForStudent(stu, byId);
            gs.forEach(function (g) {
                const rowEl = document.createElement('div');
                rowEl.className = 'guardian-row';
                const nameIn = document.createElement('input');
                nameIn.type = 'text';
                nameIn.value = g.name || '';
                nameIn.placeholder = 'Name';
                nameIn.setAttribute('aria-label', 'Name Erziehungsberechtigte');
                const mailIn = document.createElement('input');
                mailIn.type = 'email';
                mailIn.value = g.email || '';
                mailIn.placeholder = 'E-Mail';
                mailIn.setAttribute('aria-label', 'E-Mail Erziehungsberechtigte');
                const del = document.createElement('button');
                del.type = 'button';
                del.className = 'btn';
                del.innerHTML = '<i class="bi bi-x-lg"></i>';
                del.title = 'Zuordnung entfernen';
                del.addEventListener('click', function () {
                    try {
                        api().unlinkGuardianFromStudent(stu.id, g.id);
                        refresh();
                        toast('Zuordnung entfernt', 'ok');
                    } catch (e) {
                        toast(String(e && e.message ? e.message : e), 'err');
                    }
                });
                let saveTimer;
                function saveGuardian() {
                    clearTimeout(saveTimer);
                    saveTimer = setTimeout(function () {
                        try {
                            api().upsertGuardian({ id: g.id, name: nameIn.value, email: mailIn.value });
                            refresh(true);
                        } catch (e) {
                            toast(String(e && e.message ? e.message : e), 'err');
                        }
                    }, 400);
                }
                nameIn.addEventListener('input', saveGuardian);
                mailIn.addEventListener('change', saveGuardian);
                rowEl.append(nameIn, mailIn, del);
                card.appendChild(rowEl);
            });

            const add = document.createElement('div');
            add.className = 'add-guardian';
            const an = document.createElement('input');
            an.type = 'text';
            an.placeholder = 'Neuer Name';
            const am = document.createElement('input');
            am.type = 'email';
            am.placeholder = 'Neue E-Mail';
            const ab = document.createElement('button');
            ab.type = 'button';
            ab.className = 'btn btn-success';
            ab.innerHTML = '<i class="bi bi-plus-lg"></i>';
            ab.title = 'Elternkontakt zuordnen';
            ab.addEventListener('click', function () {
                try {
                    api().linkGuardianToStudent(stu.id, { name: an.value, email: am.value });
                    an.value = '';
                    am.value = '';
                    refresh();
                    toast('Elternkontakt zugeordnet', 'ok');
                } catch (e) {
                    toast(String(e && e.message ? e.message : e), 'err');
                }
            });
            add.append(an, am, ab);
            card.appendChild(add);
            wrap.appendChild(card);
        });
    }

    function buildScriptForLists(lists) {
        const g = eg();
        if (!g || typeof g.buildElternSyncScript !== 'function') {
            return '# Eltern-Hilfsmodul nicht geladen.';
        }
        const payload = lists.map(function (r) {
            return {
                displayName: r.displayName,
                mailNickname: r.mailNickname,
                guardians: r.guardians || [],
                primarySmtp: ''
            };
        });
        return g.buildElternSyncScript({
            lists: payload,
            domain: domain(),
            schoolName: schoolName()
        });
    }

    function setScript(text) {
        const ta = getEl('evPsScript');
        if (ta) ta.value = text || '';
    }

    function markExported(lists) {
        const a = api();
        if (!a || typeof a.upsertParentList !== 'function') return;
        const now = new Date().toISOString();
        lists.forEach(function (r) {
            try {
                a.upsertParentList({
                    scope: r.scope,
                    code: r.code,
                    displayName: r.displayName,
                    mailNickname: r.mailNickname,
                    graphGroupId: r.graphGroupId || '',
                    lastExportAt: now
                });
                if (r.scope === 'year' && /^\d{4}$/.test(String(r.code))) {
                    a.upsertCatalogLink({
                        kind: 'eltern',
                        code: String(r.code),
                        displayName: r.displayName,
                        mailNickname: r.mailNickname,
                        graphGroupId: r.graphGroupId || '',
                        mode: r.graphGroupId ? 'matched' : ''
                    });
                }
            } catch {
                /* ignore */
            }
        });
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'eltern-verteiler',
                action: 'export-script',
                summary: 'Exchange-Skript für ' + lists.length + ' Elternliste(n)'
            });
        }
    }

    function refresh(keepScript) {
        const prevScript = keepScript && getEl('evPsScript') ? getEl('evPsScript').value : '';
        Object.keys(PATTERN_KEYS).forEach(function (key) {
            const zone = document.querySelector('[data-ev-zone="' + key + '"]');
            if (zone && zone.querySelector('[data-token-type]')) syncPatternFromZone(key);
        });
        reloadSoll();
        renderList();
        renderDetail();
        if (keepScript && prevScript) setScript(prevScript);
    }

    function setTab(tab) {
        state.tab = tab === 'year' ? 'year' : 'class';
        state.selectedKey = '';
        state.selectedKeys.clear();
        document.querySelectorAll('[data-ev-tab]').forEach(function (btn) {
            btn.setAttribute('aria-selected', btn.getAttribute('data-ev-tab') === state.tab ? 'true' : 'false');
        });
        refresh();
    }

    function copyText(text) {
        const t = String(text || '');
        if (navigator.clipboard && navigator.clipboard.writeText) {
            return navigator.clipboard.writeText(t).then(
                function () {
                    toast('Kopiert', 'ok');
                },
                function () {
                    toast('Kopieren fehlgeschlagen', 'err');
                }
            );
        }
        toast('Zwischenablage nicht verfügbar', 'err');
        return Promise.resolve();
    }

    function downloadPs(text, base) {
        const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'eltern-verteiler-' + (base || 'sync') + '.ps1';
        document.body.appendChild(a);
        a.click();
        setTimeout(function () {
            URL.revokeObjectURL(a.href);
            a.remove();
        }, 500);
    }

    function bind() {
        wireNamingBuilders();
        wireSisImport();
        document.querySelectorAll('#evSetupSteps [data-ev-step]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                setSetupStep(btn.getAttribute('data-ev-step'));
            });
        });
        const btnDiag = getEl('evDiagnoseRefresh');
        if (btnDiag) btnDiag.addEventListener('click', function () {
            renderDiagnose();
            toast('Diagnose aktualisiert', 'ok');
        });
        const btnDiagPs = getEl('evDiagnoseScript');
        if (btnDiagPs) {
            btnDiagPs.addEventListener('click', function () {
                const eg = window.ms365ElternGuardians;
                const ta = getEl('evDiagnosePs');
                if (!eg || typeof eg.buildElternDiagnoseScript !== 'function') return;
                if (!ensureScriptPrerequisites()) return;
                const report = renderDiagnose();
                const script = eg.buildElternDiagnoseScript(
                    report && report.lists,
                    domainFromCore() || domain(),
                    schoolName()
                );
                if (ta) ta.value = script;
                toast('Diagnose-Skript erzeugt', 'ok');
            });
        }
        document.querySelectorAll('[data-ev-tab]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                setTab(btn.getAttribute('data-ev-tab'));
            });
        });
        const search = getEl('evSearch');
        if (search) search.addEventListener('input', renderList);
        const btnRefresh = getEl('evBtnRefresh');
        if (btnRefresh) btnRefresh.addEventListener('click', function () {
            refresh();
            toast('Aktualisiert', 'ok');
        });
        const btnOne = getEl('evBtnScriptOne');
        if (btnOne) {
            btnOne.addEventListener('click', function () {
                const row = findRow(state.selectedKey);
                if (!row) return toast('Keine Liste gewählt', 'err');
                if (!row.guardianCount) return toast('Keine Elternmails für diese Liste', 'err');
                showScriptForLists([row], 'Skript erzeugt');
            });
        }
        const btnSelAll = getEl('evBtnSelectAll');
        if (btnSelAll) btnSelAll.addEventListener('click', selectAllWithParents);
        const btnSelNone = getEl('evBtnSelectNone');
        if (btnSelNone) btnSelNone.addEventListener('click', clearSelection);
        const btnSel = getEl('evBtnScriptSelected');
        if (btnSel) {
            btnSel.addEventListener('click', function () {
                const lists = Array.from(state.selectedKeys)
                    .map(findRow)
                    .filter(function (r) {
                        return r && r.guardianCount > 0;
                    });
                if (!lists.length) return toast('Keine Listen ausgewählt (Checkboxen links oder „Alle auswählen“)', 'err');
                showScriptForLists(lists, 'Sammel-Skript erzeugt (' + lists.length + ')');
            });
        }
        const btnAll = getEl('evBtnScriptAll');
        if (btnAll) {
            btnAll.addEventListener('click', function () {
                const lists = rowsForTab().filter(function (r) {
                    return r.guardianCount > 0;
                });
                if (!lists.length) return toast('Keine Listen mit Elternmails', 'err');
                showScriptForLists(lists, 'Skript für ' + lists.length + ' Listen');
            });
        }
        const copy = getEl('evPsCopy');
        if (copy) {
            copy.addEventListener('click', function () {
                const ta = getEl('evPsScript');
                copyText(ta ? ta.value : '');
            });
        }
        const dl = getEl('evPsDownload');
        if (dl) {
            dl.addEventListener('click', function () {
                const ta = getEl('evPsScript');
                const row = findRow(state.selectedKey);
                const base =
                    state.lastScriptLabel ||
                    (row ? row.mailNickname || row.code : 'sync');
                downloadPs(ta ? ta.value : '', base);
            });
        }
    }

    function init() {
        if (!api()) {
            toast('Stammdaten-Modul nicht geladen', 'err');
            return;
        }
        loadPatternsFromSetup();
        bind();
        renderAllPatternBuilders();
        refresh();
        renderDiagnose();
        setSetupStep('1');
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
