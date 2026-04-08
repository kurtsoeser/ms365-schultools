(function () {
    'use strict';

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function normHeaderKey(k) {
        return String(k ?? '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '')
            .replace(/ä/g, 'ae')
            .replace(/ö/g, 'oe')
            .replace(/ü/g, 'ue')
            .replace(/ß/g, 'ss')
            .replace(/[^a-z0-9]/g, '');
    }

    function getField(row, candidates) {
        if (!row || typeof row !== 'object') return '';
        const map = new Map();
        Object.keys(row).forEach((k) => map.set(normHeaderKey(k), row[k]));
        for (const c of candidates) {
            const v = map.get(normHeaderKey(c));
            if (v != null && String(v).trim() !== '') return String(v).trim();
        }
        return '';
    }

    function ensureXlsxReady() {
        return typeof XLSX !== 'undefined' && XLSX.utils && typeof XLSX.read === 'function';
    }

    function sheetToJsonRows(workbook) {
        const sheetName = workbook.SheetNames && workbook.SheetNames[0];
        if (!sheetName) return [];
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return [];
        return XLSX.utils.sheet_to_json(sheet, { defval: '' });
    }

    function parseCsvTextToJsonRows(text) {
        if (!ensureXlsxReady()) return [];
        let s = String(text || '');
        if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
        let wb = XLSX.read(s, { type: 'string', FS: ';' });
        let rows = sheetToJsonRows(wb);
        if (!rows.length) {
            wb = XLSX.read(s, { type: 'string', FS: ',' });
            rows = sheetToJsonRows(wb);
        }
        return rows;
    }

    function downloadXlsxTemplate(filename, aoa, sheetName) {
        if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Daten');
        XLSX.writeFile(wb, filename);
        return true;
    }

    // UI binding (optional; nur wenn Elemente existieren)
    function bindUi() {
        const form = document.getElementById('tenantSettingsForm');
        if (!form) return;

        if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
            return;
        }

        const parseLinesToSubjects = window.ms365TenantSettingsParseSubjectsLines;
        const parseLinesToTeachers = window.ms365TenantSettingsParseTeachersLines;
        const parseLinesToStudents = window.ms365TenantSettingsParseStudentsLines;
        const parseLinesToClasses = window.ms365TenantSettingsParseClassesLines;
        const load = window.ms365TenantSettingsLoad;
        const save = window.ms365TenantSettingsSave;

        const taSubjects = document.getElementById('tenantSubjectsLines');
        const subjectsTbody = document.getElementById('tenantSubjectsTableBody');
        const btnAddSubjectRow = document.getElementById('tenantSubjectsAddRow');
        const taTeachers = document.getElementById('tenantTeachersLines');
        const teachersTbody = document.getElementById('tenantTeachersTableBody');
        const btnAddTeacherRow = document.getElementById('tenantTeachersAddRow');
        const taStudents = document.getElementById('tenantStudentsLines');
        const studentsTbody = document.getElementById('tenantStudentsTableBody');
        const btnAddStudentRow = document.getElementById('tenantStudentsAddRow');
        const taClasses = document.getElementById('tenantClassesLines');
        const classesTbody = document.getElementById('tenantClassesTableBody');
        const btnAddClassRow = document.getElementById('tenantClassesAddRow');
        const fileSubjects = document.getElementById('tenantSubjectsImportFile');
        const fileTeachers = document.getElementById('tenantTeachersImportFile');
        const fileStudents = document.getElementById('tenantStudentsImportFile');
        const fileClasses = document.getElementById('tenantClassesImportFile');
        const btnSubjectsTpl = document.getElementById('tenantSubjectsTemplateXlsx');
        const btnTeachersTpl = document.getElementById('tenantTeachersTemplateXlsx');
        const btnStudentsTpl = document.getElementById('tenantStudentsTemplateXlsx');
        const btnClassesTpl = document.getElementById('tenantClassesTemplateXlsx');
        const btnSave = document.getElementById('tenantSettingsSave');
        const btnReload = document.getElementById('tenantSettingsReload');
        const btnExport = document.getElementById('tenantSettingsExport');
        const btnExportHeader = document.getElementById('tenantSettingsExportHeader');
        const fileImport = document.getElementById('tenantSettingsImportFile');
        const btnClear = document.getElementById('tenantSettingsClear');
        const summary = document.getElementById('tenantSettingsSummary');
        const inpDefaultGradYear = document.getElementById('tenantDefaultGraduationYear');

        function seedDemoDataIfEmptyStorage() {
            // Demo-Daten nur beim allerersten Start (wenn noch nichts gespeichert ist)
            try {
                const raw = localStorage.getItem('ms365-tenant-settings-v1');
                if (raw) return false;
            } catch {
                // wenn localStorage nicht lesbar ist: nicht seeden
                return false;
            }

            const demo = {
                domain: 'ms365.schule',
                defaultGraduationYear: '2030',
                subjects: [
                    { code: 'M', name: 'Mathematik' },
                    { code: 'D', name: 'Deutsch' },
                    { code: 'E', name: 'Englisch' }
                ],
                teachers: [
                    { code: 'LEH', name: 'Vorname Lehrer', email: 'vorname.lehrer@ms365.schule' },
                    { code: 'MUS', name: 'Max Muster', email: 'max.muster@ms365.schule' }
                ],
                students: [
                    { klasse: '1A', name: 'Anna Beispiel', email: 'anna.beispiel@ms365.schule' },
                    { klasse: '1A', name: 'Ben Demo', email: 'ben.demo@ms365.schule' },
                    { klasse: '1B', name: 'Carla Test', email: 'carla.test@ms365.schule' },
                    { klasse: '2A', name: 'David Probe', email: 'david.probe@ms365.schule' },
                    { klasse: '2A', name: 'Eva Sample', email: 'eva.sample@ms365.schule' }
                ],
                classes: [
                    { code: '1A', year: '2030', name: 'Klasse 1A', headName: 'Vorname Lehrer', headEmail: 'vorname.lehrer@ms365.schule' },
                    { code: '1B', year: '2030', name: 'Klasse 1B', headName: 'Max Muster', headEmail: 'max.muster@ms365.schule' },
                    { code: '2A', year: '2030', name: 'Klasse 2A', headName: 'Vorname Lehrer', headEmail: 'vorname.lehrer@ms365.schule' }
                ]
            };

            const saved = save(demo);
            // Domain auch in der UI sichtbar machen
            try {
                if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                    window.ms365SetSchoolDomainNoAt(saved.domain);
                }
            } catch {
                // ignore
            }
            return true;
        }

        let autoSaveTimer = null;
        function autoSaveNow() {
            const subjects = typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
            const teachers = typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
            const students = typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
            const classes = typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
            const defaultGraduationYear = inpDefaultGradYear ? normStr(inpDefaultGradYear.value) : '';
            save({ domain, defaultGraduationYear, subjects, teachers, students, classes });
        }

        function scheduleAutoSave() {
            if (autoSaveTimer) clearTimeout(autoSaveTimer);
            autoSaveTimer = setTimeout(() => {
                autoSaveTimer = null;
                try {
                    autoSaveNow();
                } catch {
                    // ignore (z.B. während Import/Reset)
                }
            }, 450);
        }

        function setSummary(text, kind) {
            if (!summary) return;
            summary.style.display = 'block';
            summary.textContent = text;
            summary.dataset.kind = kind || 'info';
        }

        function subjectsToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code)};${normStr(x.name || '')}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getSubjectsFromTextarea() {
            return typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
        }

        function setSubjectsTextareaFromRows(rows) {
            if (!taSubjects) return;
            taSubjects.value = subjectsToLines(rows);
        }

        function renderSubjectsTableFromTextarea() {
            if (!subjectsTbody) return;
            const rows = getSubjectsFromTextarea();
            subjectsTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 3;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                subjectsTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getSubjectsFromTextarea();
                        if (!all[idx]) return renderSubjectsTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setSubjectsTextareaFromRows(all);
                        renderSubjectsTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getSubjectsFromTextarea();
                        if (!all[idx]) return renderSubjectsTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setSubjectsTextareaFromRows(all);
                        renderSubjectsTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getSubjectsFromTextarea();
                    all.splice(idx, 1);
                    setSubjectsTextareaFromRows(all);
                    renderSubjectsTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdAction);
                subjectsTbody.appendChild(tr);
            });
        }

        function teachersToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code)};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getTeachersFromTextarea() {
            return typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
        }

        function setTeachersTextareaFromRows(rows) {
            if (!taTeachers) return;
            taTeachers.value = teachersToLines(rows);
        }

        function startCellEdit(td, initialValue, onCommit) {
            const prevText = String(initialValue ?? '');
            const input = document.createElement('input');
            input.className = 'cell-editor';
            input.type = 'text';
            input.value = prevText;
            td.replaceChildren(input);
            input.focus();
            input.select();

            const commit = () => {
                const next = normStr(input.value);
                onCommit(next);
            };
            const cancel = () => {
                onCommit(prevText, { cancelled: true });
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
            input.addEventListener('blur', () => commit());
        }

        function renderTeachersTableFromTextarea() {
            if (!teachersTbody) return;
            const rows = getTeachersFromTextarea();
            teachersTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                teachersTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getTeachersFromTextarea();
                    all.splice(idx, 1);
                    setTeachersTextareaFromRows(all);
                    renderTeachersTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdEmail, tdAction);
                teachersTbody.appendChild(tr);
            });
        }

        function studentsToLines(rows) {
            return (rows || [])
                .map((x) => `${normStr(x.klasse || '')};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getStudentsFromTextarea() {
            return typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
        }

        function setStudentsTextareaFromRows(rows) {
            if (!taStudents) return;
            taStudents.value = studentsToLines(rows);
        }

        function renderStudentsTableFromTextarea() {
            if (!studentsTbody) return;
            const rows = getStudentsFromTextarea();
            studentsTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                studentsTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdClass = document.createElement('td');
                tdClass.innerHTML = `<code>${row.klasse || ''}</code>`;
                tdClass.title = 'Doppelklick zum Bearbeiten';
                tdClass.addEventListener('dblclick', () => {
                    startCellEdit(tdClass, row.klasse, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].klasse;
                        all[idx].klasse = meta && meta.cancelled ? prev : normStr(next);
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getStudentsFromTextarea();
                    all.splice(idx, 1);
                    setStudentsTextareaFromRows(all);
                    renderStudentsTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdClass, tdName, tdEmail, tdAction);
                studentsTbody.appendChild(tr);
            });
        }

        function classesToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const y = normStr(x.year || '');
                    const year = /^\d{4}$/.test(y) ? y : '';
                    return `${normCode(x.code)};${year};${normStr(x.name || '')};${normStr(x.headName || '')};${normStr(x.headEmail || '').toLowerCase()}`.trim();
                })
                .filter(Boolean)
                .join('\n');
        }

        function getClassesFromTextarea() {
            return typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
        }

        function setClassesTextareaFromRows(rows) {
            if (!taClasses) return;
            taClasses.value = classesToLines(rows);
        }

        function renderClassesTableFromTextarea() {
            if (!classesTbody) return;
            const rows = getClassesFromTextarea();
            classesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 6;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                classesTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdYear = document.createElement('td');
                tdYear.textContent = row.year || '';
                tdYear.title = 'Doppelklick zum Bearbeiten';
                tdYear.addEventListener('dblclick', () => {
                    startCellEdit(tdYear, row.year, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].year || '';
                        const n = normStr(next);
                        all[idx].year = meta && meta.cancelled ? prev : /^\d{4}$/.test(n) ? n : '';
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdHead = document.createElement('td');
                tdHead.textContent = row.headName || '';
                tdHead.title = 'Doppelklick zum Bearbeiten';
                tdHead.addEventListener('dblclick', () => {
                    startCellEdit(tdHead, row.headName, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].headName;
                        all[idx].headName = meta && meta.cancelled ? prev : normStr(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.headEmail || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.headEmail, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].headEmail;
                        all[idx].headEmail = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getClassesFromTextarea();
                    all.splice(idx, 1);
                    setClassesTextareaFromRows(all);
                    renderClassesTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdYear, tdName, tdHead, tdEmail, tdAction);
                classesTbody.appendChild(tr);
            });
        }

        function renderFromStorage() {
            const s = load();
            if (inpDefaultGradYear) {
                inpDefaultGradYear.value = /^\d{4}$/.test(normStr(s.defaultGraduationYear || ''))
                    ? normStr(s.defaultGraduationYear)
                    : '2030';
            }
            if (taSubjects) {
                taSubjects.value = (s.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            }
            if (taTeachers) {
                taTeachers.value = (s.teachers || [])
                    .map((x) => `${x.code};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            if (taStudents) {
                taStudents.value = (s.students || [])
                    .map((x) => `${x.klasse || ''};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            if (taClasses) {
                taClasses.value = (s.classes || [])
                    .map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim())
                    .join('\n');
            }
            renderSubjectsTableFromTextarea();
            renderTeachersTableFromTextarea();
            renderStudentsTableFromTextarea();
            renderClassesTableFromTextarea();
            setSummary(
                `Gespeichert: ${(s.subjects || []).length} Fächer, ${(s.teachers || []).length} Lehrkräfte, ${(s.students || []).length} Schüler, ${(s.classes || []).length} Klassen.`,
                'ok'
            );
        }

        function downloadJson(filename, obj) {
            const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 250);
        }

        function importFileToRows(file, onRows) {
            if (!file) return;
            if (!ensureXlsxReady()) {
                setSummary('Import: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
                return;
            }
            const name = String(file.name || '').toLowerCase();
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    let jsonRows = [];
                    if (name.endsWith('.csv')) {
                        const buf = new Uint8Array(e.target.result);
                        const tryDecoders = ['utf-8', 'windows-1252'];
                        for (const enc of tryDecoders) {
                            try {
                                const text = new TextDecoder(enc).decode(buf);
                                jsonRows = parseCsvTextToJsonRows(text);
                                if (jsonRows.length) break;
                            } catch {
                                // ignore
                            }
                        }
                    } else {
                        const data = new Uint8Array(e.target.result);
                        const wb = XLSX.read(data, { type: 'array' });
                        jsonRows = sheetToJsonRows(wb);
                    }
                    onRows(jsonRows || []);
                } catch (err) {
                    setSummary('Import fehlgeschlagen: ' + (err?.message || String(err)), 'warn');
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function importSubjectsRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'fach', 'abk', 'abkuerzung', 'abbreviation']);
                const name = getField(r, ['name', 'fachname', 'bezeichnung', 'subject', 'subjectname']);
                const c = normCode(code);
                if (!c) return;
                out.push({ code: c, name: normStr(name) });
            });
            if (taSubjects) taSubjects.value = out.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            setSummary(`Fächer importiert: ${out.length}`, 'ok');
        }

        function importTeachersRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'lehrer', 'abbrev', 'abbreviation']);
                let name = getField(r, ['name', 'lehrername', 'anzeigename', 'displayname']);
                let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                const c = normCode(code);
                if (!c) return;

                // Heuristik für Teillisten: wenn "Name" eigentlich eine E-Mail ist (enthält @),
                // dann korrekt zuordnen statt E-Mail als Name zu speichern.
                const nameNorm = normStr(name);
                const emailNorm = normStr(email).toLowerCase();
                const nameLooksLikeEmail = nameNorm.includes('@');
                const emailLooksLikeEmail = emailNorm.includes('@');

                if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                    email = nameNorm;
                    name = '';
                }

                out.push({ code: c, name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taTeachers) taTeachers.value = teachersToLines(out);
            renderTeachersTableFromTextarea();
            setSummary(`Lehrkräfte importiert: ${out.length}`, 'ok');
        }

        function importStudentsRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                let klasse = getField(r, ['klasse', 'class', 'gruppe', 'group']);
                let name = getField(r, ['name', 'schueler', 'schüler', 'anzeigename', 'displayname']);
                let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                if (!klasse && !name && !email) return;

                // Heuristik für Teillisten: wenn "Name" eigentlich eine E-Mail ist (enthält @),
                // dann korrekt zuordnen statt E-Mail als Name zu speichern.
                const nameNorm = normStr(name);
                const emailNorm = normStr(email).toLowerCase();
                const nameLooksLikeEmail = nameNorm.includes('@');
                const emailLooksLikeEmail = emailNorm.includes('@');

                if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                    email = nameNorm;
                    name = '';
                }

                out.push({ klasse: normStr(klasse), name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taStudents) taStudents.value = studentsToLines(out);
            renderStudentsTableFromTextarea();
            setSummary(`Schüler importiert: ${out.length}`, 'ok');
        }

        function importClassesRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                let code = getField(r, ['abkürzung', 'abkuerzung', 'abk', 'kuerzel', 'kürzel', 'code', 'klasseabk', 'classcode']);
                let year = getField(r, ['abschlussjahr', 'abschluss', 'year', 'graduationyear']);
                let name = getField(r, ['klasse', 'class', 'name', 'bezeichnung', 'classname']);
                let headName = getField(r, ['klassenvorstand', 'klassenvorstandname', 'kv', 'kvname', 'vorstand', 'head', 'headname']);
                let headEmail = getField(r, ['klassenvorstandemail', 'kvemail', 'e-mail', 'email', 'mail', 'upn', 'heademail']);
                if (!code && !year && !name && !headName && !headEmail) return;

                // Heuristik: falls "Klassenvorstand" eigentlich E-Mail ist
                const hn = normStr(headName);
                const he = normStr(headEmail).toLowerCase();
                if (hn.includes('@') && (!he || !he.includes('@'))) {
                    headEmail = hn;
                    headName = '';
                }

                out.push({
                    code: normCode(code),
                    year: /^\d{4}$/.test(normStr(year)) ? normStr(year) : '',
                    name: normStr(name),
                    headName: normStr(headName),
                    headEmail: normStr(headEmail).toLowerCase()
                });
            });
            if (taClasses) taClasses.value = classesToLines(out);
            renderClassesTableFromTextarea();
            setSummary(`Klassen importiert: ${out.length}`, 'ok');
        }

        if (btnSave) {
            btnSave.addEventListener('click', () => {
                const subjects = typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
                const teachers = typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
                const students = typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
                const classes = typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
                const domain =
                    typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
                const defaultGraduationYear = inpDefaultGradYear ? normStr(inpDefaultGradYear.value) : '';
                const saved = save({ domain, defaultGraduationYear, subjects, teachers, students, classes });
                setSummary(
                    `Gespeichert: ${(saved.subjects || []).length} Fächer, ${(saved.teachers || []).length} Lehrkräfte, ${(saved.students || []).length} Schüler, ${(saved.classes || []).length} Klassen.`,
                    'ok'
                );
                renderTeachersTableFromTextarea();
                renderStudentsTableFromTextarea();
                renderClassesTableFromTextarea();
            });
        }

        if (fileSubjects) {
            fileSubjects.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importSubjectsRows(rows));
                fileSubjects.value = '';
            });
        }
        if (fileTeachers) {
            fileTeachers.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importTeachersRows(rows));
                fileTeachers.value = '';
            });
        }
        if (fileStudents) {
            fileStudents.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importStudentsRows(rows));
                fileStudents.value = '';
            });
        }
        if (fileClasses) {
            fileClasses.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importClassesRows(rows));
                fileClasses.value = '';
            });
        }

        if (btnSubjectsTpl) {
            btnSubjectsTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Faecherliste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name'],
                        ['D', 'Deutsch'],
                        ['M', 'Mathematik'],
                        ['E', 'Englisch']
                    ],
                    'Faecher'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnTeachersTpl) {
            btnTeachersTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Lehrerliste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name', 'E-Mail'],
                        ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Lehrer'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnStudentsTpl) {
            btnStudentsTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Schuelerliste-Vorlage.xlsx',
                    [
                        ['Klasse', 'Name', 'E-Mail'],
                        ['1AK', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['1AK', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Schueler'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnClassesTpl) {
            btnClassesTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Klassenliste-Vorlage.xlsx',
                    [
                        ['Abkürzung', 'Abschlussjahr', 'Klasse', 'Klassenvorstand', 'E-Mail'],
                        ['1AK', '2030', '1A-Klasse', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['2BK', '2029', '2B-Klasse', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Klassen'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }

        if (btnReload) {
            btnReload.addEventListener('click', () => renderFromStorage());
        }

        if (btnExport) {
            btnExport.addEventListener('click', () => {
                const s = load();
                downloadJson('tenant-einstellungen.json', s);
            });
        }
        if (btnExportHeader) {
            btnExportHeader.addEventListener('click', () => {
                const s = load();
                downloadJson('tenant-einstellungen.json', s);
            });
        }

        if (btnClear) {
            btnClear.addEventListener('click', () => {
                try {
                    localStorage.removeItem('ms365-tenant-settings-v1');
                } catch {
                    // ignore
                }
                // UI/Domain wieder auf Standard zurücksetzen
                try {
                    const domainInput = document.getElementById('schoolEmailDomain');
                    if (domainInput) domainInput.value = 'ms365.schule';
                    if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                        window.ms365SetSchoolDomainNoAt('ms365.schule');
                    }
                } catch {
                    // ignore
                }
                if (inpDefaultGradYear) inpDefaultGradYear.value = '2030';
                if (taSubjects) taSubjects.value = '';
                if (taTeachers) taTeachers.value = '';
                if (taStudents) taStudents.value = '';
                if (taClasses) taClasses.value = '';
                renderSubjectsTableFromTextarea();
                renderTeachersTableFromTextarea();
                renderStudentsTableFromTextarea();
                renderClassesTableFromTextarea();
                setSummary('Tenant-Grundeinstellungen gelöscht (nur lokaler Browser-Speicher).', 'warn');
            });
        }

        if (fileImport) {
            fileImport.addEventListener('change', async (e) => {
                const f = e.target.files && e.target.files[0];
                if (!f) return;
                try {
                    const text = await f.text();
                    const obj = safeJsonParse(text);
                    if (!obj) {
                        setSummary('Import fehlgeschlagen: keine gültige JSON-Datei.', 'warn');
                        return;
                    }
                    const saved = save(obj);
                    if (inpDefaultGradYear) inpDefaultGradYear.value = normStr(saved.defaultGraduationYear || '2030');
                    if (taSubjects) taSubjects.value = (saved.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
                    if (taTeachers) taTeachers.value = (saved.teachers || []).map((x) => `${x.code};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    if (taStudents) taStudents.value = (saved.students || []).map((x) => `${x.klasse || ''};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    if (taClasses) taClasses.value = (saved.classes || []).map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim()).join('\n');
                    renderSubjectsTableFromTextarea();
                    renderTeachersTableFromTextarea();
                    renderStudentsTableFromTextarea();
                    renderClassesTableFromTextarea();
                    setSummary(
                        `Import OK: ${(saved.subjects || []).length} Fächer, ${(saved.teachers || []).length} Lehrkräfte, ${(saved.students || []).length} Schüler, ${(saved.classes || []).length} Klassen.`,
                        'ok'
                    );
                } catch (err) {
                    setSummary('Import fehlgeschlagen: ' + (err?.message || String(err)), 'warn');
                } finally {
                    fileImport.value = '';
                }
            });
        }

        const domainInput = document.getElementById('schoolEmailDomain');
        if (domainInput) {
            domainInput.addEventListener('input', () => scheduleAutoSave());
            domainInput.addEventListener('change', () => scheduleAutoSave());
        }
        if (inpDefaultGradYear) {
            inpDefaultGradYear.addEventListener('input', () => scheduleAutoSave());
            inpDefaultGradYear.addEventListener('change', () => scheduleAutoSave());
        }
        if (taSubjects) taSubjects.addEventListener('input', () => scheduleAutoSave());
        if (taSubjects) taSubjects.addEventListener('input', () => renderSubjectsTableFromTextarea());

        if (btnAddSubjectRow) {
            btnAddSubjectRow.addEventListener('click', () => {
                const all = getSubjectsFromTextarea();
                all.push({ code: '', name: '' });
                setSubjectsTextareaFromRows(all);
                renderSubjectsTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taTeachers) {
            taTeachers.addEventListener('input', () => renderTeachersTableFromTextarea());
            taTeachers.addEventListener('input', () => scheduleAutoSave());
        }
        if (btnAddTeacherRow) {
            btnAddTeacherRow.addEventListener('click', () => {
                const all = getTeachersFromTextarea();
                all.push({ code: '', name: '', email: '' });
                setTeachersTextareaFromRows(all);
                renderTeachersTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taStudents) {
            taStudents.addEventListener('input', () => renderStudentsTableFromTextarea());
            taStudents.addEventListener('input', () => scheduleAutoSave());
        }
        if (btnAddStudentRow) {
            btnAddStudentRow.addEventListener('click', () => {
                const all = getStudentsFromTextarea();
                all.push({ klasse: '', name: '', email: '' });
                setStudentsTextareaFromRows(all);
                renderStudentsTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taClasses) {
            taClasses.addEventListener('input', () => renderClassesTableFromTextarea());
            taClasses.addEventListener('input', () => scheduleAutoSave());
        }
        if (btnAddClassRow) {
            btnAddClassRow.addEventListener('click', () => {
                const all = getClassesFromTextarea();
                all.push({ code: '', name: '', headName: '', headEmail: '' });
                setClassesTextareaFromRows(all);
                renderClassesTableFromTextarea();
                scheduleAutoSave();
            });
        }

        const seeded = seedDemoDataIfEmptyStorage();
        renderFromStorage();
        if (seeded) {
            setSummary(
                'Demo-Daten wurden als Startvorlage geladen (Domain, Fächer, Lehrkräfte, Schüler, Klassen). Du kannst alles anpassen oder löschen.',
                'ok'
            );
        }
    }

    bindUi();
})();

