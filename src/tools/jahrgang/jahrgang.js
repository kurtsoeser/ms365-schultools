(function () {
    'use strict';

    let jgCurrentStep = 1;
    /** @type {{ klasse: string, jahr: string, displayName: string, suffix: string, mailNick: string, owner: string, memberLines: string }[]} */
    let jgRows = [];
    /** Bearbeitbare Vorschau (Schritt 1); Jahr leer = Standard-Abschlussjahr */
    /** @type {{ klasse: string, jahr: string, displayName: string, suffix: string, mailNick: string }[]} */
    let jgPreviewRows = [];
    let jgSuppressTextareaSync = false;

    const panelW = document.getElementById('panelWebuntis');
    const panelJ = document.getElementById('panelJahrgang');
    const btnModeW = document.getElementById('modeWebuntis');
    const btnModeJ = document.getElementById('modeJahrgang');
    const panelA = document.getElementById('panelArge');
    const btnModeA = document.getElementById('modeArge');
    const panelG = document.getElementById('panelGruppenPolicy');
    const btnModeG = document.getElementById('modeGruppenPolicy');

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    function showToast(msg) {
        const el = document.getElementById('toast');
        if (!el) return;
        el.textContent = msg;
        el.classList.add('show');
        clearTimeout(showToast._t);
        showToast._t = setTimeout(() => el.classList.remove('show'), 3500);
    }

    function startCellEdit(td, initialValue, onCommit) {
        const prevText = String(initialValue ?? '');
        const input = document.createElement('input');
        input.type = 'text';
        input.value = prevText;
        input.className = 'cell-editor';
        input.style.width = '100%';
        input.style.boxSizing = 'border-box';
        input.style.padding = '8px 10px';
        input.style.border = '1px solid #5e72e4';
        input.style.borderRadius = '10px';
        input.style.font = 'inherit';
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

    function setMode(which) {
        const w = which === 'webuntis';
        const j = which === 'jahrgang';
        const a = which === 'arge';
        const g = which === 'gruppenerstellung';
        if (panelW) panelW.style.display = w ? '' : 'none';
        if (panelJ) panelJ.style.display = j ? '' : 'none';
        if (panelA) panelA.style.display = a ? '' : 'none';
        if (panelG) panelG.style.display = g ? '' : 'none';
        if (btnModeW) btnModeW.classList.toggle('btn-success', w);
        if (btnModeJ) btnModeJ.classList.toggle('btn-success', j);
        if (btnModeA) btnModeA.classList.toggle('btn-success', a);
        if (btnModeG) btnModeG.classList.toggle('btn-success', g);
        const sdb = document.getElementById('schoolDomainBar');
        if (sdb) sdb.style.display = g ? 'none' : '';
    }

    if (btnModeW) btnModeW.addEventListener('click', () => setMode('webuntis'));
    if (btnModeJ)
        btnModeJ.addEventListener('click', () => {
            setMode('jahrgang');
            adoptJgClassesFromTenantSettingsIfEmpty();
            scheduleJgPreviewFromTextarea();
        });
    if (btnModeA) btnModeA.addEventListener('click', () => setMode('arge'));
    if (btnModeG) btnModeG.addEventListener('click', () => setMode('gruppenerstellung'));

    function applyInitialModeFromUrl() {
        try {
            const mode = new URLSearchParams(window.location.search).get('mode');
            if (!mode) return;
            if (mode.toLowerCase() === 'jahrgang') {
                setMode('jahrgang');
                adoptJgClassesFromTenantSettingsIfEmpty();
                scheduleJgPreviewFromTextarea();
            }
            else if (mode.toLowerCase() === 'arge') setMode('arge');
            else if (mode.toLowerCase() === 'gruppenerstellung' || mode.toLowerCase() === 'grouppolicy') setMode('gruppenerstellung');
            else if (mode.toLowerCase() === 'kursteams' || mode.toLowerCase() === 'kursteam' || mode.toLowerCase() === 'webuntis') setMode('webuntis');
        } catch {
            // ignore
        }
    }

    function adoptJgClassesFromTenantSettingsIfEmpty() {
        const ta = document.getElementById('jgClassLines');
        if (!ta) return;
        if (normStr(ta.value)) return;
        if (typeof window.ms365TenantSettingsLoad !== 'function') return;
        const s = window.ms365TenantSettingsLoad();
        const classes = Array.isArray(s?.classes) ? s.classes : [];
        if (!classes.length) return;
        const lines = classes
            .map((c) => {
                const code = normStr(c?.code) || normStr(c?.name);
                const year = normStr(c?.year || '');
                const name = normStr(c?.name || '');
                if (code && name && /^\d{4}$/.test(year)) return `${code};${year};${name}`;
                if (code && /^\d{4}$/.test(year)) return `${code};${year}`;
                if (code && name) return `${code};${name}`;
                return code;
            })
            .filter(Boolean);
        if (!lines.length) return;
        ta.value = lines.join('\n');
        scheduleJgPreviewFromTextarea();
        showToast('Jahrgang: Klassen aus Tenant-Einstellungen übernommen.');
    }

    let jgTenantClassesDebounce;
    function scheduleTenantClassesSyncFromJgTextarea() {
        clearTimeout(jgTenantClassesDebounce);
        jgTenantClassesDebounce = setTimeout(() => {
            try {
                const ta = document.getElementById('jgClassLines');
                if (!ta) return;
                if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function')
                    return;
                const current = window.ms365TenantSettingsLoad();
                const defaultY = getJgDefaultAbschlussjahr();
                const lines = String(ta.value || '').split(/\r\n|\n|\r/);
                const seen = new Set();
                const classes = [];
                lines.forEach((line) => {
                    const trimmed = normStr(line);
                    if (!trimmed || trimmed.startsWith('#')) return;
                    const r = parseClassLine(trimmed, defaultY);
                    if (r.skip || r.error) return;
                    const code = normCode(r.klasse);
                    const key = code.toLowerCase();
                    if (seen.has(key)) return;
                    seen.add(key);
                    classes.push({
                        code,
                        name: normStr(r.displayName || r.klasse),
                        year: normStr(r.jahr || ''),
                        headName: '',
                        headEmail: ''
                    });
                });

                const domain =
                    typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : normStr(current?.domain);
                window.ms365TenantSettingsSave({
                    domain,
                    defaultGraduationYear: defaultY,
                    subjects: current?.subjects || [],
                    teachers: current?.teachers || [],
                    students: current?.students || [],
                    classes
                });
            } catch {
                // ignore
            }
        }, 250);
    }

    function goToJgStep(step) {
        jgCurrentStep = step;
        document.querySelectorAll('.jg-step-content').forEach(el => {
            el.classList.toggle('active', parseFloat(el.dataset.jgStep) === step);
        });
        document.querySelectorAll('.jg-steps .step').forEach(el => {
            const s = parseFloat(el.dataset.jgStep);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
        if (step === 1) {
            if (jgRows.length) {
                jgPreviewRows = jgRows.map(r => ({
                    klasse: r.klasse,
                    jahr: r.jahr,
                    displayName: r.displayName || '',
                    suffix: r.suffix,
                    mailNick: r.mailNick
                }));
                renderJgPreviewTableBody();
            } else {
                scheduleJgPreviewFromTextarea();
            }
        }
        if (step === 3) {
            rebuildJgMembersTableFromRows();
        }
        if (typeof window.ms365ApplyStepProgress === 'function') {
            window.ms365ApplyStepProgress(document.querySelector('.jg-steps'), step, [1, 2, 3, 4, 5]);
        }
    }

    /** Alte Reihenfolge 1=Grundlagen, 2=Liste, 3=Besitzer → neue 1=Liste, 2=Besitzer, 3=Einstellungen */
    function migrateJgStepFromV1(step) {
        const m = { 1: 3, 2: 1, 3: 2, 4: 4 };
        const n = m[step];
        return n !== undefined ? n : step;
    }

    /** v2: 1–4 (Liste, Besitzer, Grundlagen, Ausführen) → v3: 1–5 mit Mitglieder als Schritt 3 */
    function migrateJgStepFromV2ToV3(step) {
        const m = { 1: 1, 2: 2, 3: 4, 4: 5 };
        const n = m[step];
        return n !== undefined ? n : step;
    }

    let jgPreviewDebounce;
    function scheduleJgPreviewFromTextarea() {
        clearTimeout(jgPreviewDebounce);
        jgPreviewDebounce = setTimeout(() => {
            syncJgPreviewRowsFromTextarea();
            renderJgPreviewTableBody();
        }, 120);
    }

    function scheduleJgPreviewRowsOnly() {
        clearTimeout(jgPreviewDebounce);
        jgPreviewDebounce = setTimeout(() => {
            if (jgPreviewRows.length) {
                recomputeJgPreviewMailNicks();
                updateJgPreviewMailCellsDom();
            } else {
                syncJgPreviewRowsFromTextarea();
                renderJgPreviewTableBody();
            }
        }, 120);
    }

    function getJgDefaultAbschlussjahr() {
        const el = document.getElementById('jgDefaultYear');
        const raw = (el && el.value ? el.value : '').trim();
        if (/^\d{4}$/.test(raw)) return raw;
        try {
            if (typeof window.ms365TenantSettingsLoad === 'function') {
                const s = window.ms365TenantSettingsLoad();
                const y = String(s?.defaultGraduationYear || '').trim();
                if (/^\d{4}$/.test(y)) return y;
            }
        } catch {
            // ignore
        }
        return '2030';
    }

    function jgEffectiveYear(jahr) {
        const y = String(jahr || '').trim();
        if (/^\d{4}$/.test(y)) return y;
        return getJgDefaultAbschlussjahr();
    }

    /** Entspricht dem Graph-displayName / Namen der Gruppe in Microsoft 365. */
    function jgM365DisplayName(row) {
        const dn = String(row.displayName || '').trim();
        return dn || String(row.klasse || '').trim();
    }

    function syncJgPreviewRowsFromTextarea() {
        const defaultY = getJgDefaultAbschlussjahr();
        const lines = document.getElementById('jgClassLines').value.split(/\r?\n/);
        const parsed = [];
        lines.forEach(line => {
            const r = parseClassLine(line, defaultY);
            if (r.skip || r.error) return;
            parsed.push(r);
        });
        const seenKlasse = new Set();
        jgPreviewRows = [];
        for (const p of parsed) {
            if (seenKlasse.has(p.klasse)) continue;
            seenKlasse.add(p.klasse);
            jgPreviewRows.push({
                klasse: p.klasse,
                jahr: p.jahr,
                displayName: p.displayName || '',
                suffix: p.suffix,
                mailNick: ''
            });
        }
        recomputeJgPreviewMailNicks();
    }

    function recomputeJgPreviewMailNicks() {
        const prefix = getPrefix();
        jgPreviewRows.forEach(r => {
            const m = String(r.klasse || '')
                .trim()
                .match(/^(\d+)([A-Za-z]+)$/);
            r.suffix = m ? m[2] : r.suffix || '';
            const year = jgEffectiveYear(r.jahr);
            r.mailNick = buildMailNickname(prefix, year, r.suffix);
        });
        resolveDuplicateNicks(jgPreviewRows);
    }

    function updateJgPreviewMailCellsDom() {
        const tbody = document.getElementById('jgPreviewBody');
        if (!tbody) return;
        const domain = getDomain() || '…';
        jgPreviewRows.forEach((r, i) => {
            const tr = tbody.querySelector(`tr[data-jg-index="${i}"]`);
            if (!tr) return;
            const tds = tr.querySelectorAll('td');
            if (tds.length < 6) return;
            tds[3].textContent = jgM365DisplayName(r);
            tds[4].textContent = r.mailNick;
            tds[4].style.fontFamily = 'Consolas,monospace';
            tds[4].style.fontSize = '0.9em';
            tds[5].textContent = r.mailNick + '@' + domain;
        });
    }

    function syncTextareaFromJgPreviewRows() {
        if (jgSuppressTextareaSync || !jgPreviewRows.length) return;
        const ta = document.getElementById('jgClassLines');
        if (!ta) return;
        const lines = jgPreviewRows.map(r => {
            const k = (r.klasse || '').trim();
            const y = String(r.jahr || '').trim();
            const dn = String(r.displayName || '').trim();
            if (/^\d{4}$/.test(y) && dn) return k + ';' + y + ';' + dn;
            if (/^\d{4}$/.test(y)) return k + ';' + y;
            if (dn) return k + ';' + dn;
            return k;
        });
        jgSuppressTextareaSync = true;
        ta.value = lines.join('\n');
        jgSuppressTextareaSync = false;
    }

    function renderJgPreviewTableBody() {
        const tbody = document.getElementById('jgPreviewBody');
        if (!tbody) return;
        try {
            const defaultY = getJgDefaultAbschlussjahr();
            const lines = document.getElementById('jgClassLines').value.split(/\r?\n/);
            let nonEmpty = 0;
            lines.forEach(line => {
                const t = line.trim();
                if (t && !t.startsWith('#')) nonEmpty++;
            });
            let hadError = false;
            if (!jgPreviewRows.length) {
                lines.forEach(line => {
                    const r = parseClassLine(line, defaultY);
                    if (r.skip) return;
                    if (r.error) hadError = true;
                });
                if (nonEmpty && hadError) {
                    tbody.innerHTML =
                        '<tr><td colspan="6" style="color:#6c757d;">Keine gültigen Zeilen. Erwartet z. B. <code>1AK</code>, <code>1AK;2030</code> oder <code>1AK;2030;Klasse 1A-HAK</code> (Klasse = Ziffern + Buchstaben).</td></tr>';
                } else {
                    tbody.innerHTML =
                        '<tr><td colspan="6" style="color:#6c757d;">Noch keine Zeilen – oben Klassen einfügen oder „+ Zeile hinzufügen“.</td></tr>';
                }
                return;
            }

            const domain = getDomain() || '…';
            tbody.replaceChildren();
            jgPreviewRows.forEach((r, i) => {
                const tr = document.createElement('tr');
                tr.dataset.jgIndex = String(i);

                const td1 = document.createElement('td');
                td1.innerHTML = `<code>${String(r.klasse || '')}</code>`;
                td1.title = 'Doppelklick zum Bearbeiten (z. B. 1AK)';
                td1.addEventListener('dblclick', () => {
                    startCellEdit(td1, r.klasse, (next, meta) => {
                        const prev = jgPreviewRows[i]?.klasse || '';
                        const v = meta && meta.cancelled ? prev : normStr(next);
                        jgPreviewRows[i].klasse = v;
                        recomputeJgPreviewMailNicks();
                        renderJgPreviewTableBody();
                    });
                });

                const td2 = document.createElement('td');
                td2.textContent = /^\d{4}$/.test(String(r.jahr || '').trim()) ? String(r.jahr).trim() : '';
                td2.title = 'Doppelklick zum Bearbeiten (vier Ziffern; leer = Standard)';
                td2.addEventListener('dblclick', () => {
                    startCellEdit(td2, r.jahr, (next, meta) => {
                        const prev = jgPreviewRows[i]?.jahr || '';
                        const raw = meta && meta.cancelled ? prev : normStr(next);
                        const v = /^\d{4}$/.test(raw) ? raw : '';
                        jgPreviewRows[i].jahr = v;
                        recomputeJgPreviewMailNicks();
                        updateJgPreviewMailCellsDom();
                        syncTextareaFromJgPreviewRows();
                        // year cell text
                        td2.textContent = v;
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = String(r.displayName || '');
                tdName.title = 'Doppelklick zum Bearbeiten (optional)';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, r.displayName, (next, meta) => {
                        const prev = jgPreviewRows[i]?.displayName || '';
                        const v = meta && meta.cancelled ? prev : normStr(next);
                        jgPreviewRows[i].displayName = v;
                        updateJgPreviewMailCellsDom();
                        syncTextareaFromJgPreviewRows();
                        // Klassenname + M365 DisplayName
                        tdName.textContent = v;
                        const dnCell = tr.querySelector('td.jg-preview-displayname');
                        if (dnCell) dnCell.textContent = jgM365DisplayName(jgPreviewRows[i]);
                    });
                });

                const tdDn = document.createElement('td');
                tdDn.className = 'jg-preview-displayname';
                tdDn.textContent = jgM365DisplayName(r);
                tdDn.title = 'Gruppenname in Microsoft 365 (DisplayName): Klassenname (falls vorhanden), sonst Klasse';
                tdDn.style.fontWeight = '600';
                tdDn.style.color = '#32325d';

                const td3 = document.createElement('td');
                td3.textContent = r.mailNick;
                td3.style.fontFamily = 'Consolas,monospace';
                td3.style.fontSize = '0.9em';

                const td4 = document.createElement('td');
                td4.textContent = r.mailNick + '@' + domain;

                tr.append(td1, td2, tdName, tdDn, td3, td4);
                tbody.appendChild(tr);
            });
        } catch (e) {
            console.error('Jahrgang-Vorschau:', e);
            tbody.innerHTML =
                '<tr><td colspan="6" style="color:#dc3545;">Vorschau konnte nicht berechnet werden. Seite neu laden oder Konsole prüfen.</td></tr>';
        }
    }

    /**
     * @param {string} line
     * @param {string} [defaultYearForSingleClass] Vierstellig, wenn die Zeile nur die Klasse enthält (z. B. „1AK“)
     */
    function parseClassLine(line, defaultYearForSingleClass) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) return { skip: true };
        const parts = trimmed.split(/[;,\t]/).map(s => s.trim()).filter(Boolean);
        const defYear =
            defaultYearForSingleClass && /^\d{4}$/.test(String(defaultYearForSingleClass).trim())
                ? String(defaultYearForSingleClass).trim()
                : '2030';

        if (parts.length === 1) {
            const klasse = parts[0];
            const m = klasse.match(/^(\d+)([A-Za-z]+)$/);
            if (!m) {
                return { error: 'Klasse erwartet z.B. 1AK (Ziffern + Buchstaben): ' + trimmed };
            }
            return { klasse, jahr: defYear, displayName: '', suffix: m[2] };
        }

        const klasse = parts[0];
        const m = klasse.match(/^(\d+)([A-Za-z]+)$/);
        if (!m) {
            return { error: 'Klasse erwartet z.B. 1AK (Ziffern + Buchstaben): ' + trimmed };
        }

        // 2..N Teile: tolerant lesen
        let jahr = defYear;
        let displayName = '';
        if (parts.length >= 2) {
            if (/^\d{4}$/.test(parts[1])) {
                jahr = parts[1];
                displayName = parts.length >= 3 ? parts.slice(2).join(' ') : '';
            } else {
                displayName = parts.slice(1).join(' ');
            }
        }
        // Sonderfall: klasse;name;year
        if (parts.length >= 3 && !/^\d{4}$/.test(parts[1]) && /^\d{4}$/.test(parts[2])) {
            displayName = parts[1];
            jahr = parts[2];
        }
        if (!/^\d{4}$/.test(jahr)) jahr = defYear;
        return { klasse, jahr, displayName, suffix: m[2] };
    }

    function getDomain() {
        if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
            return window.ms365GetSchoolDomainNoAt();
        }
        return '';
    }

    function getPrefix() {
        return (document.getElementById('jgPrefix').value || 'jg').trim().toLowerCase().replace(/[^a-z0-9]/g, '') || 'jg';
    }

    function suffixForNick(suffix) {
        const upper = document.getElementById('jgSuffixUpper').checked;
        const s = suffix.replace(/[^A-Za-z0-9]/g, '');
        return upper ? s.toUpperCase() : s.toLowerCase();
    }

    function buildMailNickname(prefix, year, suffix) {
        const suf = suffixForNick(suffix);
        return (prefix + year + '-' + suf).replace(/[^a-zA-Z0-9-]/g, '');
    }

    function resolveDuplicateNicks(rows) {
        const seen = new Map();
        rows.forEach(row => {
            let base = row.mailNick;
            let candidate = base;
            let n = 2;
            while (seen.has(candidate)) {
                candidate = base + n;
                n++;
            }
            row.mailNick = candidate;
            seen.set(candidate, true);
        });
    }

    const JG_STORAGE_KEY = 'ms365-jahrgang-state-v1';

    function getJgCreateTeams() {
        const el = document.getElementById('jgCreateTeams');
        return el ? !!el.checked : true;
    }

    function getJgExchangeSmtp() {
        const el = document.getElementById('jgExchangeSmtp');
        return el ? !!el.checked : true;
    }

    /** Mehrzeiliger Text → eindeutige UPNs (pro Klasse). */
    function parseJgMemberLinesText(raw) {
        const lines = String(raw || '').split(/\r\n|\n|\r/);
        const seen = new Set();
        const out = [];
        lines.forEach(line => {
            const t = String(line || '').trim();
            if (!t || t.startsWith('#')) return;
            const key = t.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push(t);
        });
        return out;
    }

    function rebuildJgMembersTableFromRows() {
        const domain = getDomain();
        const tbody = document.getElementById('jgMembersBody');
        if (!tbody) return;
        tbody.replaceChildren();
        jgRows.forEach((row, index) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = row.klasse;
            const tdName = document.createElement('td');
            tdName.textContent = row.displayName || '';
            const tdDn = document.createElement('td');
            tdDn.textContent = jgM365DisplayName(row);
            tdDn.title = 'DisplayName der Microsoft-365-Gruppe';
            tdDn.style.fontWeight = '600';
            tdDn.style.color = '#32325d';
            const td2 = document.createElement('td');
            td2.textContent = row.mailNick + '@' + domain;
            td2.style.fontFamily = 'Consolas,monospace';
            td2.style.fontSize = '0.9em';
            const td3 = document.createElement('td');
            const ta = document.createElement('textarea');
            ta.className = 'jg-member-lines';
            ta.rows = 4;
            ta.style.width = '100%';
            ta.style.minWidth = '220px';
            ta.style.padding = '8px';
            ta.style.fontFamily = 'Consolas,monospace';
            ta.style.fontSize = '0.9em';
            ta.style.boxSizing = 'border-box';
            ta.setAttribute('autocomplete', 'off');
            ta.placeholder = 'person@' + domain;
            ta.value = row.memberLines != null ? row.memberLines : '';
            ta.addEventListener('input', () => {
                jgRows[index].memberLines = ta.value;
                refreshJgScriptIfStep5();
            });
            ta.addEventListener('paste', () => setTimeout(refreshJgScriptIfStep5, 0));
            td3.appendChild(ta);
            tr.append(td1, tdName, tdDn, td2, td3);
            tbody.appendChild(tr);
        });
    }

    function refreshJgScriptIfStep5() {
        if (jgCurrentStep !== 5 || !jgRows.length) return;
        const missing = jgRows.filter(r => !r.owner);
        if (missing.length) return;
        const pre = document.getElementById('jgPowerShellScript');
        if (pre) pre.textContent = buildStandaloneJahrgangPs1(false, getJgCreateTeams(), getJgExchangeSmtp());
    }

    function rebuildJgOwnerTableFromRows() {
        const domain = getDomain();
        const tbody = document.getElementById('jgOwnerBody');
        if (!tbody) return;
        tbody.replaceChildren();
        jgRows.forEach((row, index) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = row.klasse;
            const tdName = document.createElement('td');
            tdName.textContent = row.displayName || '';
            const tdDn = document.createElement('td');
            tdDn.textContent = jgM365DisplayName(row);
            tdDn.title = 'DisplayName der Microsoft-365-Gruppe';
            tdDn.style.fontWeight = '600';
            tdDn.style.color = '#32325d';
            const td2 = document.createElement('td');
            td2.textContent = row.mailNick + '@' + domain;
            const td3 = document.createElement('td');
            td3.textContent = row.mailNick;
            const td4 = document.createElement('td');
            const inp = document.createElement('input');
            inp.type = 'email';
            inp.placeholder = 'lehrer@' + domain;
            inp.style.width = '100%';
            inp.style.padding = '8px';
            inp.style.boxSizing = 'border-box';
            inp.value = row.owner || '';
            inp.addEventListener('input', () => {
                jgRows[index].owner = inp.value.trim();
            });
            td4.appendChild(inp);
            tr.append(td1, tdName, tdDn, td2, td3, td4);
            tbody.appendChild(tr);
        });
    }

    function saveJahrgangState() {
        try {
            const state = {
                jgStepOrder: 'v3',
                jgCurrentStep,
                jgRows,
                jgPrefix: document.getElementById('jgPrefix').value,
                jgDefaultYear: document.getElementById('jgDefaultYear')
                    ? document.getElementById('jgDefaultYear').value
                    : '2030',
                jgSuffixUpper: document.getElementById('jgSuffixUpper').checked,
                jgCreateTeams: getJgCreateTeams(),
                jgExchangeSmtp: getJgExchangeSmtp(),
                jgClassLines: document.getElementById('jgClassLines').value,
                jgPowerShellScript: document.getElementById('jgPowerShellScript').textContent
            };
            localStorage.setItem(JG_STORAGE_KEY, JSON.stringify(state));
            showToast('Jahrgangsgruppen: Zwischenstand gespeichert.');
        } catch (e) {
            showToast('Speichern fehlgeschlagen: ' + e.message);
        }
    }

    function loadJahrgangState() {
        try {
            const raw = localStorage.getItem(JG_STORAGE_KEY);
            if (!raw) {
                showToast('Kein gespeicherter Stand für Jahrgangsgruppen.');
                return;
            }
            const state = JSON.parse(raw);
            let step = typeof state.jgCurrentStep === 'number' ? state.jgCurrentStep : 1;
            if (state.jgStepOrder === 'v3') {
                /* Schritte 1–5 */
            } else if (state.jgStepOrder === 'v2') {
                step = migrateJgStepFromV2ToV3(step);
            } else {
                step = migrateJgStepFromV1(step);
                step = migrateJgStepFromV2ToV3(step);
            }
            jgCurrentStep = Math.min(Math.max(1, step), 5);
            jgRows = Array.isArray(state.jgRows) ? state.jgRows : [];
            jgRows.forEach(function (row) {
                if (row.memberLines === undefined) {
                    row.memberLines = '';
                }
            });
            if (
                state.jgMemberEmails !== undefined &&
                String(state.jgMemberEmails || '').trim() !== ''
            ) {
                const legacy = String(state.jgMemberEmails);
                jgRows.forEach(function (row) {
                    if (!String(row.memberLines || '').trim()) {
                        row.memberLines = legacy;
                    }
                });
            }
            if (
                typeof window.ms365SetSchoolDomainNoAt === 'function' &&
                state.jgDomain !== undefined &&
                String(state.jgDomain).trim() !== ''
            ) {
                window.ms365SetSchoolDomainNoAt(state.jgDomain);
            }
            document.getElementById('jgPrefix').value = state.jgPrefix !== undefined ? state.jgPrefix : 'jg';
            const jgDefY = document.getElementById('jgDefaultYear');
            if (jgDefY) {
                jgDefY.value = state.jgDefaultYear !== undefined ? state.jgDefaultYear : '2030';
            }
            document.getElementById('jgSuffixUpper').checked = state.jgSuffixUpper !== false;
            const jgTeamsEl = document.getElementById('jgCreateTeams');
            if (jgTeamsEl) {
                jgTeamsEl.checked = state.jgCreateTeams !== undefined ? !!state.jgCreateTeams : true;
            }
            const jgExoEl = document.getElementById('jgExchangeSmtp');
            if (jgExoEl) {
                jgExoEl.checked = state.jgExchangeSmtp !== undefined ? !!state.jgExchangeSmtp : true;
            }
            document.getElementById('jgClassLines').value = state.jgClassLines || '';
            document.getElementById('jgParseError').style.display = 'none';
            const pre = document.getElementById('jgPowerShellScript');
            if (pre && state.jgPowerShellScript !== undefined) {
                pre.textContent = state.jgPowerShellScript;
            }
            updatePrefixExample();
            if (jgRows.length) {
                rebuildJgOwnerTableFromRows();
                rebuildJgMembersTableFromRows();
            } else {
                document.getElementById('jgOwnerBody').replaceChildren();
                const jmb = document.getElementById('jgMembersBody');
                if (jmb) jmb.replaceChildren();
            }
            goToJgStep(jgCurrentStep);
            showToast('Jahrgangsgruppen: Stand geladen.');
        } catch (e) {
            showToast('Laden fehlgeschlagen: ' + e.message);
        }
    }

    function clearJahrgangState() {
        if (!confirm('Gespeicherten Zwischenstand für Jahrgangsgruppen wirklich löschen?')) {
            return;
        }
        try {
            localStorage.removeItem(JG_STORAGE_KEY);
            jgCurrentStep = 1;
            jgRows = [];
            document.getElementById('jgPrefix').value = 'jg';
            const jgDefYClear = document.getElementById('jgDefaultYear');
            if (jgDefYClear) jgDefYClear.value = '2030';
            document.getElementById('jgSuffixUpper').checked = true;
            const jgTeamsClear = document.getElementById('jgCreateTeams');
            if (jgTeamsClear) jgTeamsClear.checked = true;
            const jgExoClear = document.getElementById('jgExchangeSmtp');
            if (jgExoClear) jgExoClear.checked = true;
            document.getElementById('jgClassLines').value = '';
            document.getElementById('jgParseError').style.display = 'none';
            document.getElementById('jgOwnerBody').replaceChildren();
            const jgMemBodyClear = document.getElementById('jgMembersBody');
            if (jgMemBodyClear) jgMemBodyClear.replaceChildren();
            document.getElementById('jgPowerShellScript').textContent = '';
            jgPreviewRows = [];
            updatePrefixExample();
            goToJgStep(1);
            showToast('Jahrgangsgruppen: Speicher geleert.');
        } catch (e) {
            showToast('Fehler: ' + e.message);
        }
    }

    window.ms365SaveJahrgang = saveJahrgangState;
    window.ms365LoadJahrgang = loadJahrgangState;
    window.ms365ClearJahrgang = clearJahrgangState;

    function updatePrefixExample() {
        const dom = getDomain();
        const pre = getPrefix();
        const ex = buildMailNickname(pre, '2030', 'AK');
        const el = document.getElementById('jgPrefixExample');
        if (el) {
            const fallback =
                typeof window.ms365DefaultSchoolDomainNoAt === 'function'
                    ? window.ms365DefaultSchoolDomainNoAt()
                    : 'ms365.schule';
            el.textContent = ex + '@' + (dom || fallback);
        }
    }

    ['schoolEmailDomain', 'jgPrefix', 'jgSuffixUpper'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', updatePrefixExample);
            el.addEventListener('change', updatePrefixExample);
            el.addEventListener('input', scheduleJgPreviewRowsOnly);
            el.addEventListener('change', scheduleJgPreviewRowsOnly);
            el.addEventListener('input', refreshJgScriptIfStep5);
            el.addEventListener('change', refreshJgScriptIfStep5);
        }
    });
    const jgDefaultYearEl = document.getElementById('jgDefaultYear');
    if (jgDefaultYearEl) {
        jgDefaultYearEl.addEventListener('input', scheduleJgPreviewRowsOnly);
        jgDefaultYearEl.addEventListener('change', scheduleJgPreviewRowsOnly);
        jgDefaultYearEl.addEventListener('input', refreshJgScriptIfStep5);
        jgDefaultYearEl.addEventListener('change', refreshJgScriptIfStep5);
        jgDefaultYearEl.addEventListener('input', () => scheduleTenantClassesSyncFromJgTextarea());
        jgDefaultYearEl.addEventListener('change', () => scheduleTenantClassesSyncFromJgTextarea());
    }
    const jgClassLinesEl = document.getElementById('jgClassLines');
    if (jgClassLinesEl) {
        jgClassLinesEl.addEventListener('input', () => {
            if (jgSuppressTextareaSync) return;
            scheduleJgPreviewFromTextarea();
            scheduleTenantClassesSyncFromJgTextarea();
        });
        jgClassLinesEl.addEventListener('change', () => {
            if (jgSuppressTextareaSync) return;
            scheduleJgPreviewFromTextarea();
            scheduleTenantClassesSyncFromJgTextarea();
        });
        jgClassLinesEl.addEventListener('paste', () =>
            setTimeout(() => {
                if (!jgSuppressTextareaSync) scheduleJgPreviewFromTextarea();
                scheduleTenantClassesSyncFromJgTextarea();
            }, 0)
        );
        jgClassLinesEl.addEventListener('input', refreshJgScriptIfStep5);
        jgClassLinesEl.addEventListener('change', refreshJgScriptIfStep5);
    }
    const jgTeamsEl = document.getElementById('jgCreateTeams');
    if (jgTeamsEl) jgTeamsEl.addEventListener('change', refreshJgScriptIfStep5);
    const jgExoEl = document.getElementById('jgExchangeSmtp');
    if (jgExoEl) jgExoEl.addEventListener('change', refreshJgScriptIfStep5);
    updatePrefixExample();

    document.getElementById('jgBack1').addEventListener('click', () => goToJgStep(1));

    const jgPreviewAddRow = document.getElementById('jgPreviewAddRow');
    if (jgPreviewAddRow) {
        jgPreviewAddRow.addEventListener('click', () => {
        jgPreviewRows.push({ klasse: '', jahr: '', displayName: '', suffix: '', mailNick: '' });
            recomputeJgPreviewMailNicks();
            renderJgPreviewTableBody();
        });
    }

    document.getElementById('jgParseAndGo2').addEventListener('click', () => {
        const errEl = document.getElementById('jgParseError');
        errEl.style.display = 'none';
        if (!jgPreviewRows.length) {
            syncJgPreviewRowsFromTextarea();
        }
        if (!jgPreviewRows.length) {
            errEl.textContent =
                'Bitte mindestens eine Klassenzeile oben eintragen oder in der Vorschau eine Zeile hinzufügen und ausfüllen.';
            errEl.style.display = 'block';
            return;
        }
        const rowErrors = [];
        jgPreviewRows.forEach((r, idx) => {
            const k = (r.klasse || '').trim();
            const m = k.match(/^(\d+)([A-Za-z]+)$/);
            if (!m) {
                rowErrors.push('Vorschau Zeile ' + (idx + 1) + ': Klasse ungültig (z. B. 1AK).');
            }
        });
        if (rowErrors.length) {
            errEl.textContent = rowErrors.join('\n');
            errEl.style.display = 'block';
            return;
        }

        if (
            typeof window.ms365IsTenantSchoolDomainConfigured !== 'function' ||
            !window.ms365IsTenantSchoolDomainConfigured()
        ) {
            errEl.textContent =
                'Bitte legen Sie die E-Mail-Domain der Schule in den Tenant-Einstellungen fest (Seite „Tenant-Einstellungen“).';
            errEl.style.display = 'block';
            if (typeof window.ms365ShowTenantDomainRequiredModal === 'function') {
                window.ms365ShowTenantDomainRequiredModal();
            }
            return;
        }

        recomputeJgPreviewMailNicks();
        const prefix = getPrefix();
        const ownerByKlasse = new Map(jgRows.map(r => [r.klasse, r.owner]));
        const memberLinesByKlasse = new Map(jgRows.map(r => [r.klasse, r.memberLines || '']));
        jgRows = jgPreviewRows.map(r => {
            const m = r.klasse.trim().match(/^(\d+)([A-Za-z]+)$/);
            const y = String(r.jahr || '').trim();
            const year = /^\d{4}$/.test(y) ? y : getJgDefaultAbschlussjahr();
            const klasseTrim = r.klasse.trim();
            const dn = String(r.displayName || '').trim();
            return {
                klasse: klasseTrim,
                jahr: year,
                displayName: dn,
                suffix: m[2],
                mailNick: buildMailNickname(prefix, year, m[2]),
                owner: ownerByKlasse.get(klasseTrim) || '',
                memberLines: memberLinesByKlasse.get(klasseTrim) || ''
            };
        });
        resolveDuplicateNicks(jgRows);

        rebuildJgOwnerTableFromRows();

        goToJgStep(2);
    });

    document.getElementById('jgGoTo3').addEventListener('click', () => goToJgStep(3));

    const jgMemberBack = document.getElementById('jgMemberBack');
    if (jgMemberBack) jgMemberBack.addEventListener('click', () => goToJgStep(2));
    const jgMemberNext = document.getElementById('jgMemberNext');
    if (jgMemberNext) jgMemberNext.addEventListener('click', () => goToJgStep(4));

    document.getElementById('jgBack2').addEventListener('click', () => goToJgStep(3));

    document.getElementById('jgGoTo4').addEventListener('click', () => {
        if (!jgRows.length) {
            showToast('Bitte zuerst die Klassenliste in Schritt 1 ausfüllen und zu Besitzern wechseln.');
            return;
        }
        const missing = jgRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle Klassen eine Besitzer-E-Mail (UPN) eintragen.');
            return;
        }
        if (
            getJgExchangeSmtp() &&
            (typeof window.ms365IsTenantSchoolDomainConfigured !== 'function' ||
                !window.ms365IsTenantSchoolDomainConfigured())
        ) {
            showToast('Für die Exchange-Option legen Sie die E-Mail-Domain der Schule in den Tenant-Einstellungen fest.');
            if (typeof window.ms365ShowTenantDomainRequiredModal === 'function') {
                window.ms365ShowTenantDomainRequiredModal();
            }
            return;
        }
        document.getElementById('jgPowerShellScript').textContent = buildPowerShellScript();
        goToJgStep(5);
    });

    document.getElementById('jgBack3').addEventListener('click', () => goToJgStep(4));

    document.getElementById('jgCopyScript').addEventListener('click', () => {
        const t = document.getElementById('jgPowerShellScript').textContent;
        navigator.clipboard.writeText(t).then(() => showToast('Script kopiert.'));
    });

    function psEscapeSingle(s) {
        return String(s).replace(/'/g, "''");
    }

    function buildPowerShellScript() {
        return buildStandaloneJahrgangPs1(false, getJgCreateTeams(), getJgExchangeSmtp());
    }

    function downloadBlob(filename, text, mime) {
        const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function downloadJson(filename, obj) {
        downloadBlob(filename, JSON.stringify(obj, null, 2), 'application/json;charset=utf-8');
    }

    function csvEscapeField(v, delimiter) {
        const d = delimiter || ';';
        const s = String(v ?? '');
        const mustQuote = s.includes('"') || s.includes('\n') || s.includes('\r') || s.includes(d);
        if (!mustQuote) return s;
        return '"' + s.replace(/"/g, '""') + '"';
    }

    function toCsv(rows, delimiter) {
        const d = delimiter || ';';
        return rows
            .map((r) => r.map((c) => csvEscapeField(c, d)).join(d))
            .join('\r\n');
    }

    function parseCsvLine(line, delimiter) {
        const d = delimiter || ';';
        const out = [];
        let cur = '';
        let inQ = false;
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (inQ) {
                if (ch === '"') {
                    const next = line[i + 1];
                    if (next === '"') {
                        cur += '"';
                        i++;
                    } else {
                        inQ = false;
                    }
                } else {
                    cur += ch;
                }
            } else {
                if (ch === '"') inQ = true;
                else if (ch === d) {
                    out.push(cur);
                    cur = '';
                } else {
                    cur += ch;
                }
            }
        }
        out.push(cur);
        return out.map((x) => normStr(x));
    }

    function detectCsvDelimiter(text) {
        const head = String(text || '').split(/\r\n|\n|\r/).slice(0, 5).join('\n');
        const semis = (head.match(/;/g) || []).length;
        const commas = (head.match(/,/g) || []).length;
        return semis >= commas ? ';' : ',';
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

    function jgBuildStateSnapshot() {
        return {
            kind: 'ms365-jahrgang-export-v1',
            exportedAt: new Date().toISOString(),
            jgCurrentStep,
            jgPrefix: document.getElementById('jgPrefix')?.value ?? 'jg',
            jgDefaultYear: document.getElementById('jgDefaultYear')?.value ?? '2030',
            jgSuffixUpper: !!document.getElementById('jgSuffixUpper')?.checked,
            jgCreateTeams: getJgCreateTeams(),
            jgExchangeSmtp: getJgExchangeSmtp(),
            jgClassLines: document.getElementById('jgClassLines')?.value ?? '',
            rows: (jgRows || []).map((r) => ({
                klasse: normStr(r.klasse),
                jahr: normStr(r.jahr),
                displayName: normStr(r.displayName || ''),
                suffix: normStr(r.suffix || ''),
                mailNick: normStr(r.mailNick || ''),
                owner: normStr(r.owner || ''),
                memberLines: String(r.memberLines ?? '')
            }))
        };
    }

    function applyImportedJgState(obj) {
        const o = obj && typeof obj === 'object' ? obj : {};
        const rows = Array.isArray(o.jgRows) ? o.jgRows : Array.isArray(o.rows) ? o.rows : [];
        const lines = o.jgClassLines !== undefined ? String(o.jgClassLines || '') : '';

        const prefixEl = document.getElementById('jgPrefix');
        const defYearEl = document.getElementById('jgDefaultYear');
        const upperEl = document.getElementById('jgSuffixUpper');
        if (prefixEl && o.jgPrefix !== undefined) prefixEl.value = String(o.jgPrefix || 'jg');
        if (defYearEl && o.jgDefaultYear !== undefined) defYearEl.value = String(o.jgDefaultYear || '2030');
        if (upperEl && o.jgSuffixUpper !== undefined) upperEl.checked = !!o.jgSuffixUpper;
        const teamsEl = document.getElementById('jgCreateTeams');
        if (teamsEl && o.jgCreateTeams !== undefined) teamsEl.checked = !!o.jgCreateTeams;
        const exoEl = document.getElementById('jgExchangeSmtp');
        if (exoEl && o.jgExchangeSmtp !== undefined) exoEl.checked = !!o.jgExchangeSmtp;

        const ta = document.getElementById('jgClassLines');
        if (ta) ta.value = lines;
        if (ta && !normStr(ta.value) && rows.length) {
            const reconstructed = rows
                .map((r) => {
                    const k = normStr(r.klasse || r.class || r.name);
                    const y = normStr(r.jahr || r.year);
                    if (!k) return '';
                    if (/^\d{4}$/.test(y)) return `${k};${y}`;
                    return k;
                })
                .filter(Boolean);
            ta.value = reconstructed.join('\n');
        }

        // owner/members aus importierten Zeilen übernehmen (matched by klasse)
        const ownerByKlasse = new Map();
        const memByKlasse = new Map();
        rows.forEach((r) => {
            const k = normStr(r.klasse || r.class || r.name);
            if (!k) return;
            ownerByKlasse.set(k, normStr(r.owner || r.ownerUpn || ''));
            memByKlasse.set(k, String(r.memberLines ?? r.members ?? ''));
            // optional: DisplayName / Klassenname
            // akzeptiere verschiedene Keys
            // - displayName: export v1
            // - className/name: CSV/Excel
            // - tenant classes use "name" for class display name
            // -> hier nur zwischenspeichern via Map an klasse-key
        });

        const nameByKlasse = new Map();
        rows.forEach((r) => {
            const k = normStr(r.klasse || r.class || r.code);
            if (!k) return;
            const dn = normStr(r.displayName || r.className || r.klassenname || r.name || '');
            if (dn) nameByKlasse.set(k, dn);
        });

        syncJgPreviewRowsFromTextarea();
        const prefix = getPrefix();
        jgRows = jgPreviewRows.map((r) => {
            const m = normStr(r.klasse).match(/^(\d+)([A-Za-z]+)$/);
            const y = normStr(r.jahr || '');
            const year = /^\d{4}$/.test(y) ? y : getJgDefaultAbschlussjahr();
            const suffix = m ? m[2] : normStr(r.suffix || '');
            const klasseTrim = normStr(r.klasse);
            return {
                klasse: klasseTrim,
                jahr: year,
                suffix,
                mailNick: buildMailNickname(prefix, year, suffix),
                displayName: normStr(r.displayName || nameByKlasse.get(klasseTrim) || ''),
                owner: ownerByKlasse.get(klasseTrim) || '',
                memberLines: memByKlasse.get(klasseTrim) || ''
            };
        });
        resolveDuplicateNicks(jgRows);

        rebuildJgOwnerTableFromRows();
        rebuildJgMembersTableFromRows();
        scheduleJgPreviewRowsOnly();
        refreshJgScriptIfStep5();
        updatePrefixExample();
        showToast('Jahrgang: Import übernommen.');
    }

    function exportJgCsv() {
        const domain = getDomain();
        if (!jgPreviewRows.length) syncJgPreviewRowsFromTextarea();
        const ownerByKlasse = new Map((jgRows || []).map((r) => [normStr(r.klasse), normStr(r.owner || '')]));
        const memByKlasse = new Map((jgRows || []).map((r) => [normStr(r.klasse), String(r.memberLines ?? '')]));
        const nameByKlasse = new Map((jgRows || []).map((r) => [normStr(r.klasse), normStr(r.displayName || '')]));

        const dataRows = jgPreviewRows.map((r) => {
            const klasse = normStr(r.klasse);
            const y = normStr(r.jahr || '');
            const year = /^\d{4}$/.test(y) ? y : '';
            const className = normStr(r.displayName || nameByKlasse.get(klasse) || '');
            const displayName = className || jgM365DisplayName(r);
            const mailNick = normStr(r.mailNick);
            const email = mailNick ? `${mailNick}@${domain}` : '';
            const owner = ownerByKlasse.get(klasse) || '';
            const members = (memByKlasse.get(klasse) || '').replace(/\r\n|\n|\r/g, '|');
            return [klasse, year, className, displayName, mailNick, email, owner, members];
        });

        const csv = toCsv(
            [
                ['klasse', 'abschlussjahr', 'className', 'displayName', 'mailNick', 'groupEmail', 'ownerUpn', 'membersUpnPipe'],
                ...dataRows
            ],
            ';'
        );
        downloadBlob(`jahrgang-export-${new Date().toISOString().slice(0, 10)}.csv`, csv, 'text/csv;charset=utf-8');
        showToast('Jahrgang: CSV exportiert.');
    }

    function importJgCsvText(text) {
        const raw = String(text || '');
        const delimiter = detectCsvDelimiter(raw);
        const lines = raw.split(/\r\n|\n|\r/).filter((l) => normStr(l));
        if (!lines.length) return;

        const first = parseCsvLine(lines[0], delimiter).map((x) => normHeaderKey(x));
        const hasHeader = first.some((h) =>
            ['klasse', 'abschlussjahr', 'classname', 'displayname', 'mailnick', 'ownerupn', 'membersupnpipe'].includes(h)
        );

        const idx = (key) => first.indexOf(key);
        const get = (row, key, altIdx) => {
            const i = hasHeader ? idx(key) : altIdx;
            if (i == null || i < 0) return '';
            return normStr(row[i] ?? '');
        };

        const outRows = [];
        const start = hasHeader ? 1 : 0;
        for (let i = start; i < lines.length; i++) {
            const row = parseCsvLine(lines[i], delimiter);
            const klasse = get(row, 'klasse', 0);
            const year = get(row, 'abschlussjahr', 1) || get(row, 'jahr', 1);
            const className = get(row, 'classname', 2);
            const displayName = get(row, 'displayname', 3);
            const owner = get(row, 'ownerupn', 6);
            const membersPipe = get(row, 'membersupnpipe', 7) || get(row, 'members', 7);
            const memberLines = membersPipe ? membersPipe.split('|').map((x) => normStr(x)).filter(Boolean).join('\n') : '';
            if (!klasse) continue;
            outRows.push({ klasse, jahr: year, displayName: className || displayName, owner, memberLines });
        }

        applyImportedJgState({ rows: outRows });
    }

    function buildStandaloneJahrgangPs1(standalone, createTeams, setExchangeSmtp) {
        if (createTeams === undefined) createTeams = true;
        if (setExchangeSmtp === undefined) setExchangeSmtp = true;
        const domain = getDomain();
        const domainTrim = (domain || '').trim();
        const setExoEffective = setExchangeSmtp && domainTrim.length > 0;
        const stamp = new Date().toISOString();
        const lines = [];
        const scopesLine = '$scopes = @("Group.ReadWrite.All","User.Read.All")';

        if (standalone) {
            lines.push('#Requires -Version 5.1');
            lines.push('# Jahrgangsgruppen (M365 Unified); optional Teams ($Ms365CreateTeams); optional Exchange-SMTP ($Ms365SetExchangeSmtp)');
            lines.push('# Erzeugt in der Browser-App am ' + stamp);
            lines.push('# Daten sind unten eingebettet.');
            lines.push('');
            lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
            lines.push('$ErrorActionPreference = "Continue"');
            lines.push('');
            lines.push('Write-Host ""');
            lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
            lines.push('Write-Host "  Jahrgangsgruppen (Microsoft Graph)"   -ForegroundColor Cyan');
            lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
            lines.push('Write-Host ""');
            lines.push(
                '# Meta-Modul Microsoft.Graph (einheitliche DLL-Versionen; PS-5.1 „4096 Funktionen“ per MaximumFunctionCount)'
            );
            lines.push('$MaximumFunctionCount = 32768');
            lines.push('Write-Host "Lade Microsoft.Graph ..." -ForegroundColor Gray');
            lines.push('try {');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('} catch {');
            lines.push('    Write-Host "Microsoft.Graph nicht gefunden – Installation (einmalig, kann einige Minuten dauern) ..." -ForegroundColor Yellow');
            lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('}');
            lines.push('');
            lines.push(scopesLine);
            lines.push('Write-Host "Starte Microsoft Graph-Anmeldung (Browser/Dialog oder Geraetecode) ..." -ForegroundColor Yellow');
            lines.push('Write-Host "Hinweis: Fenster ggf. im Hintergrund – Taskleiste pruefen." -ForegroundColor Gray');
            lines.push('$script:Ms365OldEap = $ErrorActionPreference');
            lines.push('$ErrorActionPreference = "Stop"');
            lines.push('try {');
            lines.push('    Connect-MgGraph -Scopes $scopes -NoWelcome');
            lines.push('} catch {');
            lines.push('    Write-Host ("Hinweis (interaktive Anmeldung): {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow');
            lines.push('}');
            lines.push('$ErrorActionPreference = $script:Ms365OldEap');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    Write-Host ""');
            lines.push('    Write-Host "Kein Graph-Kontext – Geraetecode-Anmeldung (Code erscheint unten, Browser: https://microsoft.com/devicelogin ) ..." -ForegroundColor Yellow');
            lines.push('    $ErrorActionPreference = "Stop"');
            lines.push('    try {');
            lines.push('        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome');
            lines.push('    } catch {');
            lines.push('        Write-Error ("Microsoft Graph: Anmeldung fehlgeschlagen: {0}" -f $_.Exception.Message)');
            lines.push('        exit 1');
            lines.push('    }');
            lines.push('    $ErrorActionPreference = $script:Ms365OldEap');
            lines.push('}');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    Write-Error "Microsoft Graph: Keine Sitzung – Anmeldung nicht erfolgreich. Skript wird beendet."');
            lines.push('    exit 1');
            lines.push('}');
            lines.push('$mgCtx = Get-MgContext');
            lines.push('Write-Host ("Angemeldet (Tenant: {0})" -f $mgCtx.TenantId) -ForegroundColor Green');
            lines.push('');
        } else {
            lines.push('# Microsoft Graph: Jahrgangsgruppen als Microsoft 365-Gruppen (Unified Group)');
            lines.push('# Voraussetzung: Install-Module Microsoft.Graph');
            lines.push('# https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup');
            lines.push('');
            lines.push('Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue');
            lines.push('$MaximumFunctionCount = 32768');
            lines.push('try {');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('} catch {');
            lines.push('    Write-Host "Microsoft.Graph nicht gefunden – Installation (einmalig, kann einige Minuten dauern) ..." -ForegroundColor Yellow');
            lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('}');
            lines.push('');
            lines.push(scopesLine);
            lines.push('Write-Host "Starte Microsoft Graph-Anmeldung (Browser/Dialog oder Geraetecode) ..." -ForegroundColor Yellow');
            lines.push('Write-Host "Hinweis: Fenster ggf. im Hintergrund – Taskleiste pruefen." -ForegroundColor Gray');
            lines.push('$script:Ms365OldEap = $ErrorActionPreference');
            lines.push('$ErrorActionPreference = "Stop"');
            lines.push('try {');
            lines.push('    Connect-MgGraph -Scopes $scopes -NoWelcome');
            lines.push('} catch {');
            lines.push('    Write-Host ("Hinweis (interaktive Anmeldung): {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow');
            lines.push('}');
            lines.push('$ErrorActionPreference = $script:Ms365OldEap');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    Write-Host ""');
            lines.push('    Write-Host "Kein Graph-Kontext – Geraetecode-Anmeldung (Code erscheint unten, Browser: https://microsoft.com/devicelogin ) ..." -ForegroundColor Yellow');
            lines.push('    $ErrorActionPreference = "Stop"');
            lines.push('    try {');
            lines.push('        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome');
            lines.push('    } catch {');
            lines.push('        throw ("Microsoft Graph: Anmeldung fehlgeschlagen: {0}" -f $_.Exception.Message)');
            lines.push('    }');
            lines.push('    $ErrorActionPreference = $script:Ms365OldEap');
            lines.push('}');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    throw "Microsoft Graph: Keine Sitzung – Anmeldung nicht erfolgreich."');
            lines.push('}');
            lines.push('$mgCtx = Get-MgContext');
            lines.push('Write-Host ("Angemeldet (Tenant: {0})" -f $mgCtx.TenantId) -ForegroundColor Green');
            lines.push('');
        }

        lines.push('$Ms365CreateTeams = $' + (createTeams ? 'true' : 'false'));
        lines.push('$Ms365SetExchangeSmtp = $' + (setExoEffective ? 'true' : 'false'));
        lines.push("$Ms365ExchangeDomain = '" + psEscapeSingle(domainTrim) + "'");
        lines.push('');
        if (setExoEffective) {
            lines.push('$script:Ms365ExoConnected = $false');
            lines.push('function Ensure-Ms365ExchangeOnline {');
            lines.push('    if ($script:Ms365ExoConnected) { return }');
            lines.push(
                '    Write-Host "Exchange Online: Modul laden und anmelden (zweiter Dialog) …" -ForegroundColor Yellow'
            );
            lines.push('    try {');
            lines.push('        Import-Module ExchangeOnlineManagement -ErrorAction Stop');
            lines.push('    } catch {');
            lines.push('        Write-Host "Installiere ExchangeOnlineManagement …" -ForegroundColor Yellow');
            lines.push('        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber');
            lines.push('        Import-Module ExchangeOnlineManagement -ErrorAction Stop');
            lines.push('    }');
            lines.push('    Connect-ExchangeOnline -ShowBanner:$false');
            lines.push('    $script:Ms365ExoConnected = $true');
            lines.push('    Write-Host "Exchange Online: angemeldet." -ForegroundColor Green');
            lines.push('}');
            lines.push('Ensure-Ms365ExchangeOnline');
            lines.push('');
        }
        lines.push('$rows = @(');
        jgRows.forEach((r, i) => {
            const last = i === jgRows.length - 1;
            const mems = parseJgMemberLinesText(r.memberLines || '');
            const memPart = mems.map(e => "'" + psEscapeSingle(e) + "'").join(',');
            const dn = String(r.displayName || '').trim();
            lines.push(
                "    [PSCustomObject]@{ Klasse = '" +
                    psEscapeSingle(r.klasse) +
                    "'; DisplayName = '" +
                    psEscapeSingle(dn || r.klasse) +
                    "'; MailNickname = '" +
                    psEscapeSingle(r.mailNick) +
                    "'; OwnerUpn = '" +
                    psEscapeSingle(r.owner) +
                    "'; Description = 'Jahrgangsgruppe " +
                    psEscapeSingle(dn || r.klasse) +
                    " (Abschluss " +
                    psEscapeSingle(r.jahr) +
                    ")'; MemberUpns = @(" +
                    memPart +
                    ') }' +
                    (last ? '' : ',')
            );
        });
        lines.push(')');
        lines.push('');
        lines.push('$i = 0');
        lines.push('foreach ($r in $rows) {');
        lines.push('    $i++');
        lines.push('    try {');
        lines.push("        $owner = Get-MgUser -UserId $r.OwnerUpn -ErrorAction Stop");
        lines.push(
            '        # M365 Unified Group: New-MgGroup -BodyParameter (Bulk-Muster, vgl. https://m365corner.com/m365-powershell/using-new-mggroup-in-graph-powershell.html )'
        );
        lines.push('        $groupBody = @{');
        lines.push('            DisplayName     = $r.DisplayName');
        lines.push('            Description     = $r.Description');
        lines.push('            MailNickname    = $r.MailNickname');
        lines.push('            MailEnabled     = $true');
        lines.push('            SecurityEnabled = $false');
        lines.push('            GroupTypes      = @("Unified")');
        lines.push('            Visibility      = "Private"');
        lines.push('        }');
        lines.push('        $group = New-MgGroup -BodyParameter $groupBody -ErrorAction Stop');
        lines.push('        Start-Sleep -Seconds 2  # Replikation vor Owner-Zuweisung');
        lines.push('        New-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id');
        lines.push('        try {');
        lines.push(
            '            $memberRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)" }'
        );
        lines.push(
            '            Invoke-MgGraphRequest -Method POST -Uri (' +
                "'" +
                'https://graph.microsoft.com/v1.0/groups/{0}/members/$ref' +
                "'" +
                ' -f $group.Id) -Body ($memberRef | ConvertTo-Json -Compress) -ErrorAction Stop'
        );
        lines.push('        } catch {');
        lines.push(
            '            Write-Host ("Hinweis (Besitzer als Mitglied): {0}" -f $_.Exception.Message) -ForegroundColor DarkGray'
        );
        lines.push('        }');
        lines.push('        if ($r.MemberUpns -and $r.MemberUpns.Count -gt 0) {');
        lines.push('            foreach ($mUpn in $r.MemberUpns) {');
        lines.push('                if ([string]::IsNullOrWhiteSpace($mUpn)) { continue }');
        lines.push('                try {');
        lines.push('                    $trimUpn = $mUpn.Trim()');
        lines.push('                    $memUser = Get-MgUser -UserId $trimUpn -ErrorAction Stop');
        lines.push('                    if ($memUser.Id -eq $owner.Id) { continue }');
        lines.push(
            '                    $memberRefExtra = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($memUser.Id)" }'
        );
        lines.push(
            '                    Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/$ref" -f $group.Id) -Body ($memberRefExtra | ConvertTo-Json -Compress) -ErrorAction Stop'
        );
        lines.push('                } catch {');
        lines.push(
            '                    if ($_.Exception.Message -match "already exist") {'
        );
        lines.push(
            '                        Write-Host ("  Hinweis (Mitglied {0}): bereits in der Gruppe." -f $mUpn.Trim()) -ForegroundColor DarkGray'
        );
        lines.push('                    } else {');
        lines.push(
            '                        Write-Host ("  Hinweis (Mitglied {0}): {1}" -f $mUpn.Trim(), $_.Exception.Message) -ForegroundColor DarkGray'
        );
        lines.push('                    }');
        lines.push('                }');
        lines.push('            }');
        lines.push('        }');
        lines.push('        if ($Ms365CreateTeams) {');
        lines.push('            $teamProps = @{');
        lines.push('                memberSettings = @{ allowCreatePrivateChannels = $true; allowCreateUpdateChannels = $true }');
        lines.push(
            '                messagingSettings = @{ allowUserEditMessages = $true; allowUserDeleteMessages = $true }'
        );
        lines.push('                funSettings = @{ allowGiphy = $true; giphyContentRating = "moderate" }');
        lines.push('                guestSettings = @{ allowCreateUpdateChannels = $false }');
        lines.push('            }');
        lines.push('            $teamUri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/team"');
        lines.push('            for ($ti = 0; $ti -lt 8; $ti++) {');
        lines.push('                try {');
        lines.push(
            '                    Invoke-MgGraphRequest -Method PUT -Uri $teamUri -Body $teamProps -ErrorAction Stop'
        );
        lines.push(
            '                    Write-Host ("Teams: {0} – Team bereitgestellt." -f $r.Klasse) -ForegroundColor Cyan'
        );
        lines.push('                    break');
        lines.push('                } catch {');
        lines.push('                    if ($ti -lt 7) {');
        lines.push(
            '                        Write-Host ("Teams: Warte auf Replikation ({0}/8) …" -f ($ti + 1)) -ForegroundColor DarkYellow'
        );
        lines.push('                        Start-Sleep -Seconds 10');
        lines.push('                    } else {');
        lines.push(
            '                        Write-Warning ("Teams: {0} – Team konnte nicht angelegt werden: {1}" -f $r.Klasse, $_.Exception.Message)'
        );
        lines.push('                    }');
        lines.push('                }');
        lines.push('            }');
        lines.push('        }');
        lines.push('        if ($Ms365SetExchangeSmtp -and $Ms365ExchangeDomain) {');
        lines.push('            $wantedSmtp = "$($r.MailNickname)@$Ms365ExchangeDomain"');
        lines.push('            for ($ei = 0; $ei -lt 6; $ei++) {');
        lines.push('                try {');
        lines.push(
            '                    Set-UnifiedGroup -Identity $group.Id -PrimarySmtpAddress $wantedSmtp -ErrorAction Stop'
        );
        lines.push(
            '                    Write-Host ("Exchange: {0} – PrimarySmtpAddress = {1}" -f $r.Klasse, $wantedSmtp) -ForegroundColor Green'
        );
        lines.push('                    break');
        lines.push('                } catch {');
        lines.push('                    if ($ei -lt 5) {');
        lines.push(
            '                        Write-Host ("Exchange: Warte auf Postfach ({0}/6) …" -f ($ei + 1)) -ForegroundColor DarkYellow'
        );
        lines.push('                        Start-Sleep -Seconds 15');
        lines.push('                    } else {');
        lines.push(
            '                        Write-Warning ("Exchange: {0} – PrimarySmtpAddress nicht gesetzt: {1}" -f $r.Klasse, $_.Exception.Message)'
        );
        lines.push('                    }');
        lines.push('                }');
        lines.push('            }');
        lines.push('        }');
        lines.push(
            '        Write-Host ("OK [{0}/{1}] {2} -> {3}" -f $i, $rows.Count, $r.Klasse, $r.MailNickname) -ForegroundColor Green'
        );
        lines.push('    }');
        lines.push('    catch {');
        lines.push('        $ex = $_.Exception');
        lines.push('        $detail = $ex.Message');
        lines.push('        if ($ex.InnerException) { $detail += " | " + $ex.InnerException.Message }');
        lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $r.Klasse, $detail)');
        lines.push('    }');
        lines.push('    Start-Sleep -Seconds 2');
        lines.push('}');
        lines.push('');
        lines.push(
            '# SMTP: Graph legt nur mailNickname an. Mit $Ms365SetExchangeSmtp wird die primäre Adresse per Exchange gesetzt.'
        );
        lines.push('# Zieldomain (App): ' + psEscapeSingle(domainTrim || domain));
        lines.push('# Set-UnifiedGroup: https://learn.microsoft.com/powershell/module/exchange/set-unifiedgroup');
        if (setExoEffective) {
            lines.push('if ($script:Ms365ExoConnected) {');
            lines.push(
                '    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}'
            );
            lines.push('}');
            lines.push('');
        }
        if (standalone) {
            lines.push('');
            lines.push('Write-Host ""');
            lines.push('Write-Host "Fertig." -ForegroundColor Cyan');
            lines.push('Read-Host "Enter druecken zum Beenden"');
        }
        return lines.join('\r\n');
    }

    function downloadJahrgangStandalonePackage() {
        if (!jgRows.length) {
            showToast('Keine Klassen – zuerst Klassenliste, Besitzer und Einstellungen abschließen.');
            return;
        }
        const missing = jgRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle Klassen einen Besitzer eintragen.');
            return;
        }
        if (typeof window.ms365BuildPolyglotCmd !== 'function') {
            showToast('polyglot-cmd.js fehlt – Seite neu laden.');
            return;
        }
        if (
            getJgExchangeSmtp() &&
            (typeof window.ms365IsTenantSchoolDomainConfigured !== 'function' ||
                !window.ms365IsTenantSchoolDomainConfigured())
        ) {
            showToast('Für die Exchange-Option legen Sie die E-Mail-Domain der Schule in den Tenant-Einstellungen fest.');
            if (typeof window.ms365ShowTenantDomainRequiredModal === 'function') {
                window.ms365ShowTenantDomainRequiredModal();
            }
            return;
        }
        const ps1 = buildStandaloneJahrgangPs1(true, getJgCreateTeams(), getJgExchangeSmtp());
        const cmd = window.ms365BuildPolyglotCmd({
            title: 'Jahrgangsgruppen-Anlage',
            echoLine: 'Starte Jahrgangsgruppen-Anlage Microsoft Graph ...',
            psBody: ps1
        });
        downloadBlob('Jahrgangsgruppen-Anlage.cmd', cmd);
        showToast('Jahrgangsgruppen-Anlage.cmd heruntergeladen – Doppelklick zum Start.');
    }

    window.downloadJahrgangStandalonePackage = downloadJahrgangStandalonePackage;

    applyInitialModeFromUrl();
    // jahrgang.html zeigt das Jahrgang-Panel direkt (ohne Mode-Buttons):
    // daher auch beim Seitenladen aus Tenant vorbefüllen.
    if (panelJ && panelJ.style.display !== 'none') {
        adoptJgClassesFromTenantSettingsIfEmpty();
        scheduleJgPreviewFromTextarea();
    }

    document.querySelectorAll('.jg-steps .step').forEach(el => {
        el.setAttribute('tabindex', '0');
        el.addEventListener('click', () => {
            const s = parseFloat(el.dataset.jgStep);
            if (s <= jgCurrentStep || el.classList.contains('completed')) {
                goToJgStep(s);
            }
        });
        el.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                el.click();
            }
        });
    });

    const elJgStepsInit = document.querySelector('.jg-steps');
    if (elJgStepsInit && typeof window.ms365ApplyStepProgress === 'function') {
        window.ms365ApplyStepProgress(elJgStepsInit, jgCurrentStep, [1, 2, 3, 4, 5]);
    }

    // bottom toolbar wiring (save/load + import/export)
    const btnSaveState = document.getElementById('btnSaveState');
    if (btnSaveState) btnSaveState.addEventListener('click', () => saveJahrgangState());
    const btnLoadState = document.getElementById('btnLoadState');
    if (btnLoadState) btnLoadState.addEventListener('click', () => loadJahrgangState());
    const btnClearStorage = document.getElementById('btnClearStorage');
    if (btnClearStorage) btnClearStorage.addEventListener('click', () => clearJahrgangState());

    const btnExportJson = document.getElementById('btnExportJgJson');
    if (btnExportJson) {
        btnExportJson.addEventListener('click', () => {
            const snap = jgBuildStateSnapshot();
            downloadJson(`jahrgang-export-${new Date().toISOString().slice(0, 10)}.json`, snap);
            showToast('Jahrgang: JSON exportiert.');
        });
    }

    const fileJson = document.getElementById('jgImportJsonFile');
    const btnImportJson = document.getElementById('btnImportJgJson');
    if (btnImportJson && fileJson) {
        btnImportJson.addEventListener('click', () => fileJson.click());
        fileJson.addEventListener('change', async (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            try {
                const text = await f.text();
                const obj = JSON.parse(text);
                applyImportedJgState(obj);
            } catch (err) {
                showToast('Jahrgang: JSON Import fehlgeschlagen: ' + (err?.message || String(err)));
            } finally {
                fileJson.value = '';
            }
        });
    }

    const btnExportCsv = document.getElementById('btnExportJgCsv');
    if (btnExportCsv) btnExportCsv.addEventListener('click', () => exportJgCsv());

    const fileCsv = document.getElementById('jgImportCsvFile');
    const btnImportCsv = document.getElementById('btnImportJgCsv');
    if (btnImportCsv && fileCsv) {
        btnImportCsv.addEventListener('click', () => fileCsv.click());
        fileCsv.addEventListener('change', async (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            try {
                const text = await f.text();
                importJgCsvText(text);
            } catch (err) {
                showToast('Jahrgang: CSV Import fehlgeschlagen: ' + (err?.message || String(err)));
            } finally {
                fileCsv.value = '';
            }
        });
    }
})();

