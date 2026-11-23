/**
 * tenant-settings-ui-import-flows.js
 * @param {Record<string, any>} ctx
 */
export function installTenantImportFlows(ctx) {
        ctx.sisApi = function sisApi() {
            return window.ms365SchoolSisImport || null;
        }

        ctx.hideSisDiff = function hideSisDiff() {
            ctx.pendingSisImport = null;
            const box = document.getElementById('tenantSisDiff');
            if (box) box.hidden = true;
        }

        ctx.rememberSisHistory = function rememberSisHistory(diff, source, mode) {
            try {
                const api = window.ms365AppDataV2;
                if (!api || typeof api.getSetup !== 'function' || typeof api.patchSetup !== 'function') return;
                const setup = api.getSetup() || {};
                const hist = Array.isArray(setup.sisImportHistory) ? setup.sisImportHistory.slice() : [];
                const c = (diff && diff.counts) || {};
                hist.push({
                    at: new Date().toISOString(),
                    source: source || '',
                    mode: mode,
                    added: c.added || 0,
                    updated: c.updated || 0,
                    removed: c.removed || 0,
                    conflicts: c.conflicts || 0
                });
                api.patchSetup({ sisImportHistory: hist.slice(-20) });
            } catch {
                /* ignore */
            }
            if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                window.ms365ActionLog.append({
                    tool: 'sis-import',
                    action: mode === 'replace' ? 'replace' : 'merge',
                    summary: (ctx.sisApi() && ctx.sisApi().summarizeSisDiff ? ctx.sisApi().summarizeSisDiff(diff) : '') + (source ? ' · ' + source : '')
                });
            }
        }

        ctx.applyPendingSis = function applyPendingSis(mode) {
            const pending = ctx.pendingSisImport;
            const sis = ctx.sisApi();
            if (!pending || !sis || typeof sis.applySisImport !== 'function') return;
            const merged = ctx.mergeSisImportViaAdapter(
                {
                    teachersLines: ctx.taTeachers ? ctx.taTeachers.value : '',
                    studentsLines: ctx.taStudents ? ctx.taStudents.value : '',
                    subjectsLines: ctx.taSubjects ? ctx.taSubjects.value : '',
                    classesLines: ctx.taClasses ? ctx.taClasses.value : ''
                },
                {
                    records: pending.records,
                    mode: mode,
                    source: pending.source
                },
                {
                    schoolSis: sis,
                    parseStudentsLines:
                        typeof window.ms365TenantSettingsParseStudentsLines === 'function'
                            ? window.ms365TenantSettingsParseStudentsLines
                            : function () {
                                  return [];
                              }
                }
            );
            const next = Array.isArray(merged.records) ? merged.records : [];
            if (ctx.taStudents && merged.lines && merged.lines.studentsLines != null) {
                ctx.taStudents.value = merged.lines.studentsLines;
            } else if (ctx.taStudents) {
                ctx.taStudents.value = sis.recordsToSemicolonLines(next);
            }
            ctx.renderStudentsTableFromTextarea();
            ctx.notifyTenantStudentsChanged('sis-import');
            const ySt = ctx.getDisplayedSchoolYearLabel() || ctx.currentSchoolYearLabel();
            ctx.setSummary(
                'Schülerimport übernommen (' +
                    (mode === 'replace' ? 'ersetzt' : 'zusammengeführt') +
                    '): ' +
                    next.length +
                    ' Zeilen — Schuljahr ' +
                    ySt,
                'ok'
            );
            ctx.rememberSisHistory(merged.studentDiff || pending.diff, pending.source, mode);
            ctx.renderStudentLifecyclePanel(pending.existing, ctx.getStudentsFromTextarea());
            ctx.renderStudentSammelLifecyclePanel(pending.existing, ctx.getStudentsFromTextarea());
            ctx.hideSisDiff();
        }

        ctx.showSisDiffPreview = function showSisDiffPreview(existing, result) {
            const sis = ctx.sisApi();
            const box = document.getElementById('tenantSisDiff');
            const meta = document.getElementById('tenantSisDiffMeta');
            const ul = document.getElementById('tenantSisDiffList');
            if (!sis || typeof sis.diffSisImport !== 'function' || !box) {
                if (ctx.taStudents) ctx.taStudents.value = (result && result.lines) || '';
                ctx.renderStudentsTableFromTextarea();
                return;
            }
            const records = result && Array.isArray(result.records) ? result.records : [];
            const diff = sis.diffSisImport(existing || [], records);
            ctx.pendingSisImport = {
                existing: existing || [],
                records: records,
                source: result && result.source ? result.source : '',
                diff: diff
            };
            if (meta) {
                const extra = [];
                if (result && result.meta && result.meta.emailsGenerated) {
                    extra.push(result.meta.emailsGenerated + ' Mails aus Namen vorgeschlagen');
                }
                if (result && result.meta && result.meta.unmatchedGuardians) {
                    extra.push(result.meta.unmatchedGuardians + ' Eltern ohne Schüler-Zuordnung');
                }
                meta.textContent =
                    (sis.summarizeSisDiff ? sis.summarizeSisDiff(diff) : '') +
                    (result && result.source ? ' · Quelle ' + result.source : '') +
                    (extra.length ? ' · ' + extra.join(' · ') : '') +
                    '. Zusammenführen behält lokale Zeilen, die in der Datei fehlen (geeignet für Aktualisierungen im Jahr).';
            }
            if (ul) {
                ul.innerHTML = '';
                (diff.conflicts || []).slice(0, 8).forEach(function (c) {
                    const li = document.createElement('li');
                    li.textContent = 'Konflikt: ' + (c.summary || c.type);
                    ul.appendChild(li);
                });
                (diff.updated || []).slice(0, 6).forEach(function (u) {
                    const li = document.createElement('li');
                    const prev = u.previous || {};
                    const rec = u.incoming || {};
                    li.textContent =
                        'Änderung: ' +
                        (prev.name || prev.email || '') +
                        (u.klasseChanged ? ' (' + (prev.klasse || '–') + ' → ' + (rec.klasse || '–') + ')' : '');
                    ul.appendChild(li);
                });
                if (!ul.childNodes.length) {
                    const li = document.createElement('li');
                    li.textContent = (diff.counts && diff.counts.added ? diff.counts.added + ' neue Zeilen. ' : '') + 'Keine Konflikte erkannt.';
                    ul.appendChild(li);
                }
            }
            box.hidden = false;
            try {
                box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            } catch {
                /* ignore */
            }
            ctx.setSummary('Schülerimport wartet auf Bestätigung – bitte Vorschau prüfen.', 'info');
        }

        ctx.hideTeachersWuDiff = function hideTeachersWuDiff(opts) {
            const restore = !opts || opts.restore !== false;
            if (
                restore &&
                ctx.pendingTeachersWuImport &&
                ctx.taTeachers &&
                typeof ctx.pendingTeachersWuImport.backupLines === 'string'
            ) {
                ctx.taTeachers.value = ctx.pendingTeachersWuImport.backupLines;
                ctx.teachersSortState.key = null;
                ctx.teachersSortState.dir = 1;
                ctx.renderTeachersTableFromTextarea();
            }
            ctx.pendingTeachersWuImport = null;
            const box = document.getElementById('tenantTeachersWuDiff');
            if (box) box.hidden = true;
        }

        /** Live-Vorschau in Textarea + Haupttabelle (ohne Autosave, bis bestätigt). */
        ctx.applyTeachersWuPreviewToMain = function applyTeachersWuPreviewToMain(mode) {
            const pending = ctx.pendingTeachersWuImport;
            if (!pending || !ctx.taTeachers) return;
            const merged =
                mode === 'replace'
                    ? ctx.replaceTeachersImportList(pending.incoming)
                    : ctx.mergeTeachersImportLists(pending.existing || [], pending.incoming);
            const wu = window.ms365WebuntisExportImport;
            ctx.taTeachers.value =
                wu && typeof wu.teachersToSemicolonLines === 'function'
                    ? wu.teachersToSemicolonLines(merged)
                    : ctx.teachersToLines(merged);
            ctx.teachersSortState.key = null;
            ctx.teachersSortState.dir = 1;
            ctx.renderTeachersTableFromTextarea();
        }

        ctx.showTeachersWuDiffPreview = function showTeachersWuDiffPreview(existing, result, routeSuffix) {
            const incoming = result && Array.isArray(result.teachers) ? result.teachers : [];
            const existingRows = existing || [];
            const diff = ctx.diffTeachersImport(existingRows, incoming);
            // Leere lokale Liste ohne Konflikte: direkt übernehmen (Tabelle sofort sichtbar).
            if (
                !existingRows.length &&
                incoming.length &&
                !(diff.conflicts && diff.conflicts.length)
            ) {
                ctx.applyTeachersWuImportDirect(incoming, result, routeSuffix);
                return;
            }
            ctx.pendingTeachersWuImport = {
                incoming: incoming,
                existing: existingRows,
                backupLines: ctx.taTeachers ? ctx.taTeachers.value : '',
                result: result || {},
                diff: diff,
                routeSuffix: routeSuffix || ''
            };
            const box = document.getElementById('tenantTeachersWuDiff');
            const meta = document.getElementById('tenantTeachersWuDiffMeta');
            const tbody = document.getElementById('tenantTeachersWuDiffBody');
            if (!box || !tbody) {
                ctx.applyTeachersWuImportDirect(incoming, result, routeSuffix);
                return;
            }
            const extra = [];
            if (result && result.meta && result.meta.emailsGenerated) {
                extra.push(result.meta.emailsGenerated + ' E-Mails aus Namen vorgeschlagen');
            }
            if (result && result.meta && result.meta.emailConflicts) {
                extra.push(result.meta.emailConflicts + ' E-Mail-Konflikte beim Erzeugen');
            }
            if (meta) {
                meta.textContent =
                    ctx.summarizeTeachersDiff(diff) +
                    ' · WebUntis-Lehrerexport' +
                    (routeSuffix || '') +
                    (extra.length ? ' · ' + extra.join(' · ') : '') +
                    '. Zusammenführen behält Lehrkräfte, die nur lokal stehen (Abgleich über Kürzel).';
            }
            tbody.replaceChildren();
            const maxRows = 120;
            let shown = 0;
            function appendRow(status, code, name, email, kind) {
                if (shown >= maxRows) return;
                shown++;
                const tr = document.createElement('tr');
                tr.className = 'ts-teachers-wu-diff__row--' + (kind || 'add');
                const tdS = document.createElement('td');
                tdS.textContent = status;
                const tdC = document.createElement('td');
                tdC.textContent = code;
                const tdN = document.createElement('td');
                tdN.textContent = name;
                const tdE = document.createElement('td');
                tdE.textContent = email;
                tr.append(tdS, tdC, tdN, tdE);
                tbody.appendChild(tr);
            }
            (diff.conflicts || []).forEach(function (c) {
                const inc = c.incoming || {};
                appendRow('Konflikt', c.code || inc.code, inc.name || '', inc.email || '', 'conflict');
            });
            (diff.added || []).forEach(function (r) {
                appendRow('Neu', r.code, r.name, r.email, 'add');
            });
            (diff.updated || []).forEach(function (u) {
                const inc = u.incoming || {};
                const bits = [];
                if (u.nameChanged) bits.push('Name');
                if (u.emailChanged) bits.push('E-Mail');
                appendRow('Änderung (' + bits.join(', ') + ')', inc.code, inc.name, inc.email, 'update');
            });
            if (!shown) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.textContent = 'Alle ' + (diff.counts && diff.counts.unchanged ? diff.counts.unchanged : incoming.length) + ' Kürzel bereits in der Liste – Sie können trotzdem zusammenführen oder verwerfen.';
                tr.appendChild(td);
                tbody.appendChild(tr);
            } else if (
                (diff.counts && diff.counts.unchanged) ||
                incoming.length > shown
            ) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = 'var(--muted)';
                const rest =
                    (diff.counts && diff.counts.unchanged ? diff.counts.unchanged + ' unverändert' : '') +
                    (incoming.length > shown ? ' · Auszug (' + shown + ' von ' + incoming.length + ')' : '');
                td.textContent = rest.trim();
                tr.appendChild(td);
                tbody.appendChild(tr);
            }
            box.hidden = false;
            // Haupttabelle sofort mit Merge-Vorschau füllen (Verwerfen stellt den Stand wieder her).
            ctx.applyTeachersWuPreviewToMain('merge');
            try {
                ctx.focusRegisterTabForImportTarget('teachers');
                box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            } catch {
                /* ignore */
            }
            if (diff.conflicts && diff.conflicts.length) {
                ctx.setSummary(
                    'WebUntis-Lehrer: bitte E-Mail-Konflikte prüfen – die Tabelle unten zeigt die Merge-Vorschau.',
                    'warn'
                );
            } else {
                ctx.setSummary(
                    'WebUntis-Lehrerimport: Tabelle unten zeigt die Vorschau – bitte Zusammenführen, Ersetzen oder Verwerfen.',
                    'info'
                );
            }
        }

        ctx.applyTeachersWuImportDirect = function applyTeachersWuImportDirect(incoming, result, routeSuffix) {
            if (ctx.taTeachers) {
                const wu = window.ms365WebuntisExportImport;
                const merged = ctx.mergeTeachersImportLists(ctx.getTeachersFromTextarea(), incoming);
                ctx.taTeachers.value =
                    wu && typeof wu.teachersToSemicolonLines === 'function'
                        ? wu.teachersToSemicolonLines(merged)
                        : ctx.teachersToLines(merged);
            }
            ctx.teachersSortState.key = null;
            ctx.teachersSortState.dir = 1;
            ctx.renderTeachersTableFromTextarea();
            const n = result && result.meta && result.meta.teacherCount ? result.meta.teacherCount : incoming.length;
            ctx.setSummary('Lehrkräfte importiert: ' + n + ' (WebUntis)' + (routeSuffix || ''), 'ok');
            ctx.scheduleAutoSave();
        }

        ctx.applyPendingTeachersWu = function applyPendingTeachersWu(mode) {
            const pending = ctx.pendingTeachersWuImport;
            if (!pending) return;
            const existing = pending.existing || [];
            // Merge-Vorschau liegt bereits in Textarea/Tabelle; bei Ersetzen neu schreiben.
            if (mode === 'replace') {
                ctx.applyTeachersWuPreviewToMain('replace');
            }
            const merged = ctx.getTeachersFromTextarea();
            const n = pending.incoming.length;
            const routeSuffix = pending.routeSuffix || '';
            ctx.hideTeachersWuDiff({ restore: false });
            ctx.setSummary(
                'Lehrkräfte übernommen (' +
                    (mode === 'replace' ? 'Liste ersetzt' : 'zusammengeführt') +
                    '): ' +
                    n +
                    ' aus WebUntis' +
                    routeSuffix,
                'ok'
            );
            ctx.renderTeacherLifecyclePanel(existing, merged);
            ctx.scheduleAutoSave();
        }

        ctx.lifecycleApi = function lifecycleApi() {
            return window.ms365StudentClassLifecycle || null;
        }

        ctx.getLifecycleContext = function getLifecycleContext() {
            const empty = { classTeams: [], schuelerGroupId: '' };
            try {
                const api = window.ms365AppDataV2;
                const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
                const classTeams = c && c.core && Array.isArray(c.core.classTeams) ? c.core.classTeams : [];
                const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : c && c.setup;
                const matched = setup && setup.matched ? setup.matched : {};
                return {
                    classTeams: classTeams,
                    schuelerGroupId: matched.schuelerGroupId ? String(matched.schuelerGroupId) : ''
                };
            } catch {
                return empty;
            }
        }

}
