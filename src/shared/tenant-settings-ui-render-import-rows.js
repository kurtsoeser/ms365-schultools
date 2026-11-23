/**
 * tenant-settings-ui-render-import-rows.js
 * @param {Record<string, any>} ctx
 */
export function installTenantRenderImportRows(ctx) {
    ctx.importArgesRows = function importArgesRows(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = ctx.getField(r, ['kürzel', 'kuerzel', 'code', 'arge', 'abk', 'abkuerzung']);
            const name = ctx.getField(r, ['name', 'bezeichnung', 'titel', 'title']);
            const subs = ctx.getField(r, ['faecher', 'fächer', 'subjects', 'fach', 'subjectcodes']);
            const c = ctx.normCode(code);
            if (!c) return;
            const subjects = String(subs || '')
                .split(/[,\s|]+/)
                .map((x) => ctx.normCode(x))
                .filter(Boolean);
            out.push({ code: c, name: ctx.normStr(name), subjects });
        });
        if (ctx.taArges) ctx.taArges.value = ctx.argesToLines(out);
        ctx.renderArgesTableFromTextarea();
        ctx.setSummary(`ARGEs importiert: ${out.length} (schulweit)`, 'ok');
    }

    ctx.importTeachersRows = function importTeachersRows(jsonRows) {
        const prevTeachers = ctx.getTeachersFromTextarea();
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = ctx.getField(r, ['kürzel', 'kuerzel', 'code', 'lehrer', 'abbrev', 'abbreviation']);
            let name = ctx.getField(r, ['name', 'lehrername', 'anzeigename', 'displayname']);
            let email = ctx.getField(r, ['e-mail', 'email', 'mail', 'upn']);
            const c = ctx.normCode(code);
            if (!c) return;

            // Heuristik für Teillisten: wenn "Name" eigentlich eine E-Mail ist (enthält @),
            // dann korrekt zuordnen statt E-Mail als Name zu speichern.
            const nameNorm = ctx.normStr(name);
            const emailNorm = ctx.normStr(email).toLowerCase();
            const nameLooksLikeEmail = nameNorm.includes('@');
            const emailLooksLikeEmail = emailNorm.includes('@');

            if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                email = nameNorm;
                name = '';
            }

            out.push({ code: c, name: ctx.normStr(name), email: ctx.normStr(email).toLowerCase() });
        });
        if (ctx.taTeachers) ctx.taTeachers.value = ctx.teachersToLines(out);
        ctx.teachersSortState.key = null;
        ctx.teachersSortState.dir = 1;
        ctx.renderTeachersTableFromTextarea();
        ctx.setSummary(`Lehrkräfte importiert: ${out.length} (schulweit)`, 'ok');
        ctx.renderTeacherLifecyclePanel(prevTeachers, out);
        ctx.scheduleAutoSave();
    }

    ctx.importStudentsRows = function importStudentsRows(jsonRows) {
        const prev = ctx.getStudentsFromTextarea();
        const sis = window.ms365SchoolSisImport;
        if (sis && typeof sis.importStudentsAndGuardians === 'function') {
            const result = sis.importStudentsAndGuardians({ objectRows: jsonRows, source: 'auto' });
            ctx.showSisDiffPreview(prev, result);
            return;
        }
        const out = [];
        (jsonRows || []).forEach((r) => {
            let klasse = ctx.getField(r, ['klasse', 'class', 'gruppe', 'group']);
            let name = ctx.getField(r, ['name', 'schueler', 'schüler', 'anzeigename', 'displayname']);
            let email = ctx.getField(r, ['e-mail', 'email', 'mail', 'upn']);
            if (!klasse && !name && !email) return;

            const nameNorm = ctx.normStr(name);
            const emailNorm = ctx.normStr(email).toLowerCase();
            const nameLooksLikeEmail = nameNorm.includes('@');
            const emailLooksLikeEmail = emailNorm.includes('@');

            if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                email = nameNorm;
                name = '';
            }

            out.push({ klasse: ctx.normStr(klasse), name: ctx.normStr(name), email: ctx.normStr(email).toLowerCase() });
        });
        if (ctx.taStudents) ctx.taStudents.value = ctx.studentsToLines(out);
        ctx.studentsSortState.key = null;
        ctx.studentsSortState.dir = 1;
        ctx.renderStudentsTableFromTextarea();
        const ySt = ctx.getDisplayedSchoolYearLabel() || ctx.currentSchoolYearLabel();
        ctx.setSummary(`Schüler importiert: ${out.length} (Schuljahr ${ySt})`, 'ok');
        ctx.renderStudentLifecyclePanel(prev, out);
        ctx.renderStudentSammelLifecyclePanel(prev, out);
    }

    ctx.importClassesFromWebuntisPdfFile = function importClassesFromWebuntisPdfFile(file) {
        const wu = window.ms365WebuntisExportImport;
        const pdfApi = window.ms365PdfText;
        if (!wu || typeof wu.importClassesFromWebuntisPdf !== 'function') {
            ctx.setSummary('WebUntis-Klassenimport: Modul fehlt – Seite neu laden.', 'warn');
            return;
        }
        if (!pdfApi || typeof pdfApi.extractPdfFile !== 'function') {
            ctx.setSummary('PDF-Modul fehlt – Seite neu laden.', 'warn');
            return;
        }
        const skipEl = document.getElementById('tenantClassesSkipNoKv') || document.getElementById('swClassesSkipNoKv');
        const skipWithoutTeacher = skipEl != null ? !!skipEl.checked : true;
        const teachers = ctx.getTeachersFromTextarea();
        ctx.setSummary('Klassen-PDF wird gelesen …', 'ok');
        pdfApi
            .extractPdfFile(file)
            .then(function (content) {
                const result = wu.importClassesFromWebuntisPdf({
                    text: content.text,
                    words: content.words,
                    teachers: teachers,
                    skipWithoutTeacher: skipWithoutTeacher,
                    inferYear: true
                });
                if (!result.classes.length) {
                    ctx.setSummary('Keine Klassen im PDF erkannt (Kurzname / Klassenlehrkraft prüfen).', 'warn');
                    return;
                }
                if (ctx.taClasses) ctx.taClasses.value = result.lines;
                ctx.renderClassesTableFromTextarea();
                const bits = [
                    result.meta.classCount + ' Klassen',
                    result.meta.withTeacher + ' mit KV-Kürzel',
                    result.meta.matched + ' KV aus Lehrerliste'
                ];
                if (result.meta.missingTeacher) bits.push(result.meta.missingTeacher + ' Kürzel ohne Lehrerliste');
                if (result.schoolYear && result.schoolYear.label) bits.push('Schuljahr ' + result.schoolYear.label);
                ctx.setSummary('WebUntis-Klassen importiert: ' + bits.join(' · '), 'ok');
                ctx.scheduleAutoSave();
            })
            .catch(function (err) {
                ctx.setSummary('Klassen-PDF: ' + (err && err.message ? err.message : String(err)), 'warn');
            });
    }

    ctx.importClassesRows = function importClassesRows(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            let code = ctx.getField(r, ['abkürzung', 'abkuerzung', 'abk', 'kuerzel', 'kürzel', 'code', 'klasseabk', 'classcode', 'kurzname']);
            let year = ctx.getField(r, ['abschlussjahr', 'abschluss', 'year', 'graduationyear']);
            let name = ctx.getField(r, ['klasse', 'class', 'name', 'bezeichnung', 'classname', 'langname']);
            let headName = ctx.getField(r, ['klassenvorstand', 'klassenvorstandname', 'kv', 'kvname', 'vorstand', 'head', 'headname']);
            let headEmail = ctx.getField(r, ['klassenvorstandemail', 'kvemail', 'e-mail', 'email', 'mail', 'upn', 'heademail']);
            const headCode = ctx.getField(r, ['klassenlehrkraft', 'lehrkraft', 'kvkuerzel', 'headcode']);
            if (!code && !year && !name && !headName && !headEmail && !headCode) return;

            // Heuristik: falls "Klassenvorstand" eigentlich E-Mail ist
            const hn = ctx.normStr(headName);
            const he = ctx.normStr(headEmail).toLowerCase();
            if (hn.includes('@') && (!he || !he.includes('@'))) {
                headEmail = hn;
                headName = '';
            }

            let resolvedName = ctx.normStr(headName);
            let resolvedEmail = ctx.normStr(headEmail).toLowerCase();
            if (headCode && (!resolvedName || !resolvedEmail)) {
                const teachers = ctx.getTeachersFromTextarea();
                const hit = teachers.find((t) => ctx.normCode(t.code) === ctx.normCode(headCode));
                if (hit) {
                    if (!resolvedName) resolvedName = ctx.normStr(hit.name);
                    if (!resolvedEmail) resolvedEmail = ctx.normStr(hit.email).toLowerCase();
                } else if (!resolvedName) {
                    resolvedName = ctx.normCode(headCode);
                }
            }

            out.push({
                code: ctx.normCode(code),
                year: /^\d{4}$/.test(ctx.normStr(year)) ? ctx.normStr(year) : '',
                name: ctx.normStr(name),
                headName: resolvedName,
                headEmail: resolvedEmail
            });
        });
        if (ctx.taClasses) ctx.taClasses.value = ctx.classesToLines(out);
        ctx.renderClassesTableFromTextarea();
        const yCl = ctx.getDisplayedSchoolYearLabel() || ctx.currentSchoolYearLabel();
        ctx.setSummary(`Klassen importiert: ${out.length} (Schuljahr ${yCl})`, 'ok');
    }

    ctx.onTenantSettingsSaveClick = async () => {
            if (ctx.autoSaveTimer) {
                clearTimeout(ctx.autoSaveTimer);
                ctx.autoSaveTimer = null;
            }
            const prevLoaded = ctx.load();
            const prevStudents = (prevLoaded && prevLoaded.students) || [];
            const prevTeachers = (prevLoaded && prevLoaded.teachers) || [];
            const subjects = typeof ctx.parseLinesToSubjects === 'function' ? ctx.parseLinesToSubjects(ctx.taSubjects ? ctx.taSubjects.value : '') : [];
            const okSubjects = await ctx.confirmSubjectCatalogIfShrinks(
                subjects,
                'Speichern der Stammdaten'
            );
            if (!okSubjects) {
                ctx.setSummary('Speichern abgebrochen – Fächerliste unverändert.', 'warn');
                return;
            }
            const teachers = typeof ctx.parseLinesToTeachers === 'function' ? ctx.parseLinesToTeachers(ctx.taTeachers ? ctx.taTeachers.value : '') : [];
            const admin = ctx.getAdminFromTextarea();
            const adminRoles = ctx.getAdminRolesFromTextarea();
            const sga = typeof ctx.parseLinesToSga === 'function' ? ctx.parseLinesToSga(ctx.taSga ? ctx.taSga.value : '') : [];
            const sgaMode = ctx.selSgaMode ? ctx.normStr(ctx.selSgaMode.value || 'group').toLowerCase() : 'group';
            const students = typeof ctx.parseLinesToStudents === 'function' ? ctx.parseLinesToStudents(ctx.taStudents ? ctx.taStudents.value : '') : [];
            const studentCouncil =
                typeof ctx.parseLinesToStudentCouncil === 'function'
                    ? ctx.parseLinesToStudentCouncil(ctx.taStudentCouncil ? ctx.taStudentCouncil.value : '')
                    : [];
            const classes = typeof ctx.parseLinesToClasses === 'function' ? ctx.parseLinesToClasses(ctx.taClasses ? ctx.taClasses.value : '') : [];
            const schoolName = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
            const verifiedEmailDomains = ctx.verifiedDomainsInput ? ctx.normStr(ctx.verifiedDomainsInput.value || '') : '';
            const arges = typeof ctx.parseLinesToArges === 'function' ? ctx.parseLinesToArges(ctx.taArges ? ctx.taArges.value : '') : [];
            const administration = ctx.getAdministrationEntries();
            const prevSettings = typeof ctx.load === 'function' ? ctx.load() : {};
            const saved = ctx.save({
                schoolName,
                domain,
                verifiedEmailDomains,
                subjects,
                arges,
                teachers,
                administration,
                admin,
                adminRoles,
                verwaltungAudienceGroups: ctx.audienceGroupsSnapshot,
                adminAudienceMemberships: ctx.audienceMembershipsSnapshot,
                verwaltungOrganigramm:
                    prevSettings && prevSettings.verwaltungOrganigramm
                        ? prevSettings.verwaltungOrganigramm
                        : { version: 1, nodes: [] },
                schulOrganigramm:
                    prevSettings && prevSettings.schulOrganigramm
                        ? prevSettings.schulOrganigramm
                        : { version: 1, nodes: [] },
                sgaMode,
                sga,
                students,
                studentCouncil,
                classes
            });
            ctx.renderSubjectsTableFromTextarea();
            ctx.renderArgesTableFromTextarea();
            ctx.renderTeachersTableFromTextarea();
            ctx.renderAdminRolesTableFromTextarea();
            ctx.renderAdminTableFromTextarea();
            ctx.renderSgaTableFromTextarea();
            ctx.renderStudentsTableFromTextarea();
            ctx.renderStudentCouncilTableFromTextarea();
            ctx.renderClassesTableFromTextarea();
            ctx.renderStatusOverview();
            ctx.renderStudentLifecyclePanel(prevStudents, students);
            ctx.renderStudentSammelLifecyclePanel(prevStudents, students);
            ctx.renderTeacherLifecyclePanel(prevTeachers, teachers);
            ctx.dispatchTenantSettingsChanged(saved, 'manual-save');
            if (typeof ctx.markRegisterClean === 'function') ctx.markRegisterClean({ stamp: true });
            if (typeof ctx.setSummary === 'function') {
                ctx.setSummary('Speichern erfolgreich – Stammdaten sind gesichert.', 'ok');
            }
            try {
                if (typeof window.ms365ShowToast === 'function') {
                    window.ms365ShowToast('Speichern erfolgreich – Stammdaten sind gesichert.', {
                        kind: 'ok',
                        title: 'Stammdaten'
                    });
                }
            } catch {
                /* ignore */
            }
            try {
                const sync = window.ms365SchoolProfileLicenseSync;
                if (sync && typeof sync.push === 'function') {
                    sync.push({ toast: false, force: true }).catch(function () {
                        /* Betreiber-Meldung optional */
                    });
                }
            } catch {
                /* ignore */
            }
            try {
                const spo = window.ms365StammdatenSpoAutoSync;
                if (spo && typeof spo.schedulePush === 'function') {
                    spo.schedulePush(0);
                }
            } catch {
                // ignore
            }
    };
    ctx.btnSaves.forEach((btn) => btn.addEventListener('click', ctx.onTenantSettingsSaveClick));

    window.addEventListener('pagehide', function () {
        if (!ctx.listEditSincePaint && !ctx.autoSaveTimer) return;
        if (ctx.autoSaveTimer) {
            clearTimeout(ctx.autoSaveTimer);
            ctx.autoSaveTimer = null;
        }
        try {
            ctx.autoSaveNow();
        } catch {
            // ignore
        }
    });
    document.addEventListener('visibilitychange', function () {
        if (document.visibilityState !== 'hidden') return;
        if (!ctx.listEditSincePaint && !ctx.autoSaveTimer) return;
        if (ctx.autoSaveTimer) {
            clearTimeout(ctx.autoSaveTimer);
            ctx.autoSaveTimer = null;
        }
        try {
            ctx.autoSaveNow();
        } catch {
            // ignore
        }
    });

    // Struktur -> Listen (Writeback)
}
