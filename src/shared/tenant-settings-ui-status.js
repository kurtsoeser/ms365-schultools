/**
 * tenant-settings-ui-status.js
 * @param {Record<string, any>} ctx
 */
export function installTenantStatusPanel(ctx) {
        ctx.registerSaveDirty = false;
        ctx.registerDirtyCount = 0;
        ctx.lastSavedAt = null;
        try {
            const stored =
                localStorage.getItem('ms365-stammdaten-saved-at') ||
                sessionStorage.getItem('ms365-stammdaten-saved-at');
            const n = stored ? Number(stored) : 0;
            if (n > 0) ctx.lastSavedAt = new Date(n);
        } catch {
            /* Speicher nicht verfügbar */
        }

        ctx.formatSavedClock = function formatSavedClock(date) {
            if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
            const hh = String(date.getHours()).padStart(2, '0');
            const mm = String(date.getMinutes()).padStart(2, '0');
            const dd = String(date.getDate()).padStart(2, '0');
            const mo = String(date.getMonth() + 1).padStart(2, '0');
            return dd + '.' + mo + '.' + date.getFullYear() + ', ' + hh + ':' + mm;
        };

        ctx.refreshWorkbenchMeta = function refreshWorkbenchMeta() {
            const dirty = !!ctx.registerSaveDirty;
            document.querySelectorAll('[data-work-count]').forEach(function (el) {
                const body = document.getElementById(el.getAttribute('data-work-count') || '');
                if (!body) return;
                const rows = body.querySelectorAll('tr');
                let n = 0;
                rows.forEach(function (tr) {
                    if (tr.querySelector('td[colspan]')) return;
                    n += 1;
                });
                const one = el.getAttribute('data-work-one') || 'Datensatz';
                const many = el.getAttribute('data-work-many') || 'Datensätze';
                const noun = n === 1 ? one : many;
                el.textContent = n + ' ' + noun + ' · ' + (dirty ? 'ungespeichert' : 'gespeichert');
            });
        };

        ctx.observeWorkbenchCounts = function observeWorkbenchCounts() {
            if (ctx._workCountObs || typeof MutationObserver === 'undefined') return;
            const bodies = [];
            document.querySelectorAll('[data-work-count]').forEach(function (el) {
                const body = document.getElementById(el.getAttribute('data-work-count') || '');
                if (body) bodies.push(body);
            });
            if (!bodies.length) return;
            ctx._workCountObs = new MutationObserver(function () {
                ctx.refreshWorkbenchMeta();
            });
            bodies.forEach(function (body) {
                ctx._workCountObs.observe(body, { childList: true });
            });
            ctx.refreshWorkbenchMeta();
        };

        ctx.syncRegisterSaveBar = function syncRegisterSaveBar() {
            const dirty = !!ctx.registerSaveDirty;
            const dock = document.querySelector('.ts-admin-dock.ts-register-dock');
            const bar = document.getElementById('tenantSaveBar');
            if (dock) dock.classList.toggle('is-dirty', dirty);
            if (bar) {
                bar.classList.toggle('is-dirty', dirty);
                bar.setAttribute('data-dirty', dirty ? '1' : '0');
            }
            const hint = document.getElementById('tenantSaveDirtyHint');
            if (hint) {
                hint.hidden = false;
                hint.classList.toggle('is-dirty', dirty);
                hint.classList.toggle('is-clean', !dirty);
                if (dirty) {
                    const n = ctx.registerDirtyCount > 0 ? ctx.registerDirtyCount : 1;
                    hint.textContent =
                        '● ' + n + (n === 1 ? ' ungespeicherte Änderung' : ' ungespeicherte Änderungen');
                } else {
                    hint.textContent = '✓ Alle Änderungen gespeichert';
                }
            }
            const meta = document.getElementById('tenantSaveMeta');
            if (meta) {
                const clock = ctx.formatSavedClock(ctx.lastSavedAt);
                meta.textContent = clock ? 'Lokal · ' + clock : 'Lokal';
            }
            document.querySelectorAll('[data-tenant-action="save"]').forEach(function (btn) {
                const inHeader = !!(btn.closest && btn.closest('.ts-register-header-toolbar'));
                btn.hidden = !!inHeader;
                const label = btn.querySelector ? btn.querySelector('.ts-save-btn-label') : null;
                if (label) label.textContent = dirty ? 'Änderungen speichern' : 'Speichern';
                btn.classList.toggle('fb-btn--ghost', !dirty);
                btn.classList.toggle('ts-act-primary', dirty && !inHeader);
                btn.classList.toggle('is-dirty-primary', dirty && !inHeader);
            });
            if (typeof ctx.refreshWorkbenchMeta === 'function') ctx.refreshWorkbenchMeta();
        };

        ctx.markRegisterDirty = function markRegisterDirty() {
            ctx.registerSaveDirty = true;
            ctx.registerDirtyCount = (ctx.registerDirtyCount || 0) + 1;
            ctx.syncRegisterSaveBar();
        };

        ctx.markRegisterClean = function markRegisterClean(opts) {
            const wasDirty = !!ctx.registerSaveDirty;
            const stamp = !!(opts && opts.stamp);
            ctx.registerSaveDirty = false;
            ctx.registerDirtyCount = 0;
            if (wasDirty || stamp) {
                ctx.lastSavedAt = new Date();
                try {
                    const raw = String(ctx.lastSavedAt.getTime());
                    localStorage.setItem('ms365-stammdaten-saved-at', raw);
                    sessionStorage.setItem('ms365-stammdaten-saved-at', raw);
                } catch {
                    /* ignore */
                }
            }
            ctx.syncRegisterSaveBar();
        };

        ctx.flushPendingAutoSave = function flushPendingAutoSave() {
            if (ctx.autoSaveTimer) {
                clearTimeout(ctx.autoSaveTimer);
                ctx.autoSaveTimer = null;
            }
            try {
                ctx.autoSaveNow();
            } catch {
                // ignore (z.B. während Import/Reset)
            }
        }

        ctx.scheduleAutoSave = function scheduleAutoSave() {
            ctx.listEditSincePaint = true;
            ctx.markRegisterDirty();
            if (ctx.autoSaveTimer) clearTimeout(ctx.autoSaveTimer);
            ctx.autoSaveTimer = setTimeout(() => {
                ctx.autoSaveTimer = null;
                ctx.flushPendingAutoSave();
            }, 450);
        }

        window.ms365TenantSettingsFlushPendingSave = ctx.flushPendingAutoSave;
        document.querySelectorAll('details.ts-sga-board-raw, details.ts-tech-fold, details.ts-group-pop').forEach(function (el) {
            el.open = false;
        });
        const docRoot = document.documentElement;
        if (docRoot && docRoot.dataset && !docRoot.dataset.tenantGroupPopBound) {
            docRoot.dataset.tenantGroupPopBound = '1';
            document.addEventListener('click', function (ev) {
                const target = ev.target;
                document.querySelectorAll('details.ts-group-pop[open]').forEach(function (fold) {
                    if (target instanceof Node && fold.contains(target)) return;
                    fold.open = false;
                });
            });
            document.addEventListener('keydown', function (ev) {
                if (ev.key !== 'Escape') return;
                document.querySelectorAll('details.ts-group-pop[open]').forEach(function (fold) {
                    fold.open = false;
                });
            });
        }
        ctx.syncRegisterSaveBar();
        ctx.observeWorkbenchCounts();

        ctx.setSummary = function setSummary(text, kind) {
            if (!ctx.summary) return;
            if (ctx._summaryClearTimer) {
                clearTimeout(ctx._summaryClearTimer);
                ctx._summaryClearTimer = null;
            }
            ctx.summary.hidden = false;
            ctx.summary.style.display = 'block';
            ctx.summary.textContent = text;
            ctx.summary.dataset.kind = kind || 'info';
            if (ctx.summary.dataset.stand) ctx.summary.dataset.stand = '';
            const msg = String(text || '').trim();
            if (!msg) return;
            const ms = kind === 'warn' || kind === 'error' ? 16000 : 7000;
            ctx._summaryClearTimer = setTimeout(function () {
                ctx._summaryClearTimer = null;
                if (!ctx.summary) return;
                ctx.summary.textContent = '';
                ctx.summary.hidden = true;
                ctx.summary.style.display = 'none';
            }, ms);
        }

        ctx.buildStandSummaryBody = function buildStandSummaryBody(s, yLbl) {
            return (
                'Aktueller Stand: schulweit ' +
                (s.subjects || []).length +
                ' Fächer, ' +
                (s.arges || []).length +
                ' ARGEs, ' +
                (s.admin || []).length +
                ' Verwaltung, ' +
                (s.sga || []).length +
                ' SGA-Einträge, ' +
                (s.teachers || []).length +
                ' Lehrkräfte — für Schuljahr ' +
                yLbl +
                ': ' +
                (s.students || []).length +
                ' Schüler, ' +
                (s.studentCouncil || []).length +
                ' Schülervertretung, ' +
                (s.classes || []).length +
                ' Klassen.'
            );
        }

        ctx.setStandSummary = function setStandSummary(saveReason) {
            ctx.updateRegisterLastSavedMeta(saveReason);
            if (!ctx.summary) return;
            ctx.summary.textContent = '';
            ctx.summary.hidden = true;
            ctx.summary.style.display = 'none';
            ctx.summary.dataset.stand = '';
        }

        ctx.refreshStandSummaryIfVisible = function refreshStandSummaryIfVisible(saveReason) {
            if (!ctx.summary || ctx.summary.dataset.stand !== '1') return;
            ctx.setStandSummary(saveReason);
        }

        // ----------------------------------------------------------------
        // Status-Übersicht: „Was fehlt noch?"
        // ----------------------------------------------------------------
        ctx.renderStatusOverview = function renderStatusOverview() {
            const grid = document.getElementById('tenantStatusGrid');
            const footer = document.getElementById('tenantStatusFooter') || document.getElementById('tenantStatusLastCheck');
            if (!grid) return;

            const api = window.ms365AppDataV2;
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : {};
            const dirMatch = (setup && setup.directoryMatchByEmail) ? setup.directoryMatchByEmail : {};
            const classMatch = (setup && setup.classGroupMatchByKey) ? setup.classGroupMatchByKey : {};

            // Alle Personen-E-Mails aus allen Listen einsammeln
            const teacherRows = ctx.getTeachersFromTextarea ? ctx.getTeachersFromTextarea() : [];
            const adminRows = typeof ctx.getAdministrationGroups === 'function'
                ? (function () {
                    const gs = ctx.getAdministrationGroups();
                    const out = [];
                    (gs || []).forEach(function (g) {
                        (Array.isArray(g.people) ? g.people : []).forEach(function (p) {
                            if (p && p.email) out.push({ name: p.name || '', email: p.email });
                        });
                    });
                    return out;
                }())
                : [];
            const studentRows = ctx.getStudentsFromTextarea ? ctx.getStudentsFromTextarea() : [];
            const sgaRows = typeof ctx.getSgaFromTextarea === 'function' ? ctx.getSgaFromTextarea() : [];
            const svRows = typeof ctx.getStudentCouncilFromTextarea === 'function' ? ctx.getStudentCouncilFromTextarea() : [];

            function countEmailMatch(rows) {
                let total = 0; let matched = 0; let notFound = 0; let unchecked = 0;
                const seen = new Set();
                rows.forEach(function (r) {
                    const em = ctx.normStr(r && r.email || '').toLowerCase();
                    if (!em || em.indexOf('@') === -1) return;
                    if (seen.has(em)) return;
                    seen.add(em);
                    total++;
                    const m = dirMatch[em];
                    if (m && m.graphUserId) matched++;
                    else if (m && m.notFound) notFound++;
                    else unchecked++;
                });
                return { total, matched, notFound, unchecked };
            }

            const tStat = countEmailMatch(teacherRows);
            const aStat = countEmailMatch(adminRows);
            const sStat = countEmailMatch(studentRows);
            const sgaStat = countEmailMatch(sgaRows);
            const svStat = countEmailMatch(svRows);

            // Klassen-Gruppen
            const classRows = typeof ctx.getClassesFromTextarea === 'function' ? ctx.getClassesFromTextarea() : [];
            let clTotal = 0; let clMatched = 0; let clNotFound = 0; let clUnchecked = 0;
            const clSeen = new Set();
            classRows.forEach(function (row) {
                const key = (function () {
                    const code = ctx.normStr(row && row.code || '').toUpperCase();
                    const name = ctx.normStr(row && row.name || '').toUpperCase();
                    return code || name || '';
                }());
                if (!key || clSeen.has(key)) return;
                clSeen.add(key);
                clTotal++;
                const m = classMatch[key];
                if (m && m.groupId) clMatched++;
                else if (m && m.notFound) clNotFound++;
                else clUnchecked++;
            });

            // Fächer/ARGEs (ohne Gruppen-Match-Check: einfach zählen, wie viele vorhanden)
            const subjectRows = typeof ctx.getSubjectsFromTextarea === 'function' ? ctx.getSubjectsFromTextarea() : [];
            const argeRows = typeof ctx.getArgesFromTextarea === 'function' ? ctx.getArgesFromTextarea() : [];

            // Chips aufbauen
            const chips = [];

            function chip(icon, label, kind, title) {
                chips.push({ icon, label, kind, title: title || '' });
            }

            const tabStatus = {};

            function personChip(label, stat, tabId) {
                const { total, matched, notFound, unchecked } = stat;
                let title = '';
                if (total === 0) {
                    chip('bi-dash', label + ': keine Einträge', 'muted', 'Keine Einträge in der Liste.');
                    title = 'Keine Einträge in der Liste.';
                } else if (matched === total) {
                    chip('bi-check-circle-fill', label + ': alle ' + total + ' gematcht', 'ok', 'Alle ' + total + ' E-Mail-Adressen in Microsoft Entra gefunden.');
                    title = 'Alle ' + total + ' in Microsoft Entra gefunden.';
                } else if (unchecked > 0) {
                    chip('bi-question-circle', label + ': ' + unchecked + '/' + total + ' ungeprüft', 'warn',
                        unchecked + ' noch nicht geprüft, ' + notFound + ' nicht gefunden, ' + matched + ' gefunden.');
                    title = unchecked + ' ungeprüft, ' + notFound + ' nicht gefunden, ' + matched + ' gefunden.';
                } else if (notFound > 0) {
                    chip('bi-x-circle', label + ': ' + notFound + '/' + total + ' nicht gefunden', 'error',
                        notFound + ' in Microsoft Entra nicht gefunden, ' + matched + ' gefunden.');
                    title = notFound + ' in Entra nicht gefunden, ' + matched + ' gefunden.';
                } else {
                    chip('bi-check-circle-fill', label + ': alle ' + total + ' gematcht', 'ok', 'Alle gefunden.');
                    title = 'Alle in Microsoft Entra gefunden.';
                }
                if (tabId) {
                    const st = ctx.personMatchTabStatus(stat);
                    tabStatus[tabId] = { kind: st.kind, icon: st.icon, title };
                }
            }

            personChip('Lehrer', tStat, 'tabMainLehrer');
            const lehrerGid = ctx.getLehrerSammelGroupId();
            if (teacherRows.length > 0) {
                const lgTitle = lehrerGid
                    ? 'Lehrer:innen-Sammelgruppe ist mit Microsoft 365 verknüpft.'
                    : 'Lehrer:innen-Sammelgruppe noch nicht verknüpft – im Tab Lehrer prüfen oder anlegen.';
                chip(
                    lehrerGid ? 'bi-check-circle-fill' : 'bi-info-circle',
                    lehrerGid ? 'Lehrer:innen-Gruppe: verknüpft' : 'Lehrer:innen-Gruppe: offen',
                    lehrerGid ? 'ok' : 'warn',
                    lgTitle
                );
            }
            personChip('Verwaltung', aStat, 'tabMainVerwaltung');
            personChip('Schüler', sStat, 'tabMainSchueler');
            const schuelerGid = ctx.getSchuelerSammelGroupId();
            if (studentRows.length > 0) {
                const sgTitle = schuelerGid
                    ? 'Schüler:innen-Sammelgruppe ist mit Microsoft 365 verknüpft.'
                    : 'Schüler:innen-Sammelgruppe noch nicht verknüpft – im Tab Schüler prüfen oder anlegen.';
                chip(
                    schuelerGid ? 'bi-check-circle-fill' : 'bi-info-circle',
                    schuelerGid ? 'Schüler:innen-Gruppe: verknüpft' : 'Schüler:innen-Gruppe: offen',
                    schuelerGid ? 'ok' : 'warn',
                    sgTitle
                );
            }
            personChip('SGA', sgaStat, 'tabMainSga');
            personChip('Schülervertretung', svStat, 'tabMainSchuelervertretung');

            // Klassen-Gruppen
            let clTitle = '';
            if (clTotal === 0) {
                chip('bi-dash', 'Klassen: keine Einträge', 'muted', 'Keine Klassen eingetragen.');
                clTitle = 'Keine Klassen eingetragen.';
            } else if (clMatched === clTotal) {
                chip('bi-check-circle-fill', 'Klassen: ' + clTotal + ' gematcht', 'ok', 'Alle Klassen-Gruppen in Microsoft 365 gefunden.');
                clTitle = 'Alle ' + clTotal + ' Klassen-Gruppen in Microsoft 365 gefunden.';
            } else if (clUnchecked > 0) {
                chip('bi-question-circle', 'Klassen: ' + clUnchecked + '/' + clTotal + ' ungeprüft', 'warn',
                    clUnchecked + ' noch nicht geprüft, ' + clNotFound + ' nicht gefunden, ' + clMatched + ' gefunden.');
                clTitle = clUnchecked + ' ungeprüft, ' + clNotFound + ' nicht gefunden, ' + clMatched + ' gefunden.';
            } else {
                chip('bi-x-circle', 'Klassen: ' + clNotFound + '/' + clTotal + ' nicht gefunden', 'error',
                    clNotFound + ' Klassen-Gruppen fehlen in Microsoft 365, ' + clMatched + ' gefunden.');
                clTitle = clNotFound + ' Klassen-Gruppen fehlen, ' + clMatched + ' gefunden.';
            }
            const clSt = ctx.classMatchTabStatus(clTotal, clMatched, clNotFound, clUnchecked);
            tabStatus.tabMainKlassen = { kind: clSt.kind, icon: clSt.icon, title: clTitle };

            const belegungSnap =
                window.ms365AppDataV2 && typeof window.ms365AppDataV2.getUnterrichtsbelegung === 'function'
                    ? window.ms365AppDataV2.getUnterrichtsbelegung()
                    : null;
            const teachingSum = ctx.summarizeClassTeaching(
                classRows,
                belegungSnap && Array.isArray(belegungSnap.rows) ? belegungSnap.rows : []
            );
            if (teachingSum.classTotal > 0 || teachingSum.rowCount > 0) {
                const teachTitle =
                    teachingSum.withTeachers +
                    ' von ' +
                    teachingSum.classTotal +
                    ' Klassen mit unterrichtenden Lehrkräften (' +
                    teachingSum.teacherAssignments +
                    ' Zuordnungen) – Unterrichtsbelegung.';
                if (teachingSum.withoutTeachers === 0 && teachingSum.classTotal > 0) {
                    chip(
                        'bi-person-badge',
                        'Lehrer je Klasse: ' + teachingSum.withTeachers + '/' + teachingSum.classTotal,
                        'ok',
                        teachTitle
                    );
                    tabStatus.tabMainKlassenlehrer = {
                        kind: 'ok',
                        icon: 'bi-check-circle-fill',
                        title: teachTitle
                    };
                } else {
                    chip(
                        'bi-person-badge',
                        'Lehrer je Klasse: ' + teachingSum.withTeachers + '/' + teachingSum.classTotal,
                        teachingSum.withTeachers ? 'warn' : 'muted',
                        teachTitle +
                            (teachingSum.withoutTeachers
                                ? ' ' + teachingSum.withoutTeachers + ' ohne Belegung.'
                                : '')
                    );
                    tabStatus.tabMainKlassenlehrer = {
                        kind: teachingSum.withTeachers ? 'warn' : 'muted',
                        icon: teachingSum.withTeachers ? 'bi-exclamation-circle' : 'bi-dash',
                        title: teachTitle
                    };
                }
            } else {
                tabStatus.tabMainKlassenlehrer = {
                    kind: 'muted',
                    icon: 'bi-dash',
                    title: 'Noch keine Unterrichtsbelegung – Tab „Lehrer je Klasse“ oder Kursteams/Unterrichtsteams.'
                };
            }

            // Fächer / ARGEs – reine Zahl (kein Gruppen-Match vorhanden)
            let subLinked = 0;
            if (subjectRows.length > 0) {
                subjectRows.forEach(function (r) {
                    const l = ctx.getSubjectCatalogLink(r.code);
                    if (l && ctx.normStr(l.graphGroupId)) subLinked++;
                });
                const subTitle =
                    subLinked > 0
                        ? subLinked + ' von ' + subjectRows.length + ' Fachgruppen mit Microsoft-365-Gruppe verknüpft.'
                        : subjectRows.length +
                          ' Fachgruppen eingetragen – „Aus Microsoft 365 übernehmen“ oder ARGE & Fachgruppen.';
                chip(
                    subLinked === subjectRows.length && subjectRows.length > 0 ? 'bi-check-circle-fill' : 'bi-info-circle',
                    subjectRows.length + ' Fachgruppen' + (subLinked ? ' (' + subLinked + ' M365)' : ''),
                    subLinked === subjectRows.length && subjectRows.length > 0 ? 'ok' : 'muted',
                    subTitle
                );
            }
            const subSt = ctx.catalogLinkTabStatus(subjectRows.length, subLinked);
            tabStatus.tabMainSubjects = {
                kind: subjectRows.length ? subSt.kind : 'muted',
                icon: subjectRows.length ? subSt.icon : 'bi-dash',
                title:
                    subjectRows.length === 0
                        ? 'Noch keine Fächer eingetragen.'
                        : subLinked + ' von ' + subjectRows.length + ' mit M365-Gruppe verknüpft.'
            };

            let linkedN = 0;
            if (argeRows.length > 0) {
                argeRows.forEach(function (r) {
                    const l = ctx.getArgeCatalogLink(r.code);
                    if (l && ctx.normStr(l.graphGroupId)) linkedN++;
                });
                const argeTitle =
                    linkedN > 0
                        ? linkedN + ' von ' + argeRows.length + ' ARGEs mit Microsoft-365-Gruppe verknüpft.'
                        : argeRows.length + ' ARGEs eingetragen – „Aus Microsoft 365 übernehmen“ oder ARGE & Fachgruppen für Verknüpfungen.';
                chip(
                    linkedN === argeRows.length && argeRows.length > 0 ? 'bi-check-circle-fill' : 'bi-info-circle',
                    argeRows.length + ' ARGEs' + (linkedN ? ' (' + linkedN + ' M365)' : ''),
                    linkedN === argeRows.length && argeRows.length > 0 ? 'ok' : 'muted',
                    argeTitle
                );
            }
            const argeSt = ctx.catalogLinkTabStatus(argeRows.length, linkedN);
            tabStatus.tabMainArges = {
                kind: argeRows.length ? argeSt.kind : 'muted',
                icon: argeRows.length ? argeSt.icon : 'bi-dash',
                title:
                    argeRows.length === 0
                        ? 'Noch keine ARGEs eingetragen.'
                        : linkedN + ' von ' + argeRows.length + ' mit M365-Gruppe verknüpft.'
            };

            const schoolName = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? ctx.normStr(window.ms365GetSchoolDomainNoAt()) : '';
            if (schoolName && domain) {
                tabStatus.tabMainStammdaten = {
                    kind: 'ok',
                    icon: 'bi-check-circle-fill',
                    title: 'Schulname und Domain sind gesetzt.'
                };
            } else if (!schoolName && !domain) {
                tabStatus.tabMainStammdaten = {
                    kind: 'muted',
                    icon: 'bi-dash',
                    title: 'Schulname und Domain noch leer.'
                };
            } else {
                tabStatus.tabMainStammdaten = {
                    kind: 'warn',
                    icon: 'bi-exclamation-circle',
                    title: 'Schulname oder Domain fehlt noch.'
                };
            }

            // DOM rendern
            grid.replaceChildren();
            chips.forEach(function (c) {
                const span = document.createElement('span');
                span.className = 'ts-status-chip ts-status-chip--' + c.kind;
                if (c.title) { span.title = c.title; span.setAttribute('aria-label', c.label + ': ' + c.title); }
                span.innerHTML = '<i class="bi ' + ctx.escapeHtml(c.icon) + '" aria-hidden="true"></i> ' + ctx.escapeHtml(c.label);
                grid.appendChild(span);
            });

            if (footer) {
                const now = new Date();
                footer.textContent =
                    'Entra-Abgleich geprüft: ' + now.toLocaleTimeString('de-AT', { hour: '2-digit', minute: '2-digit' });
            }

            ctx.applyTabRegisterStatus(tabStatus);
        }

        window.ms365RenderTenantStatusOverview = ctx.renderStatusOverview;

}
