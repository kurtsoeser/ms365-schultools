
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

/** Ab Schema 2: data-step entspricht der angezeigten Schrittnummer (0–8). */
const KURSTEAM_STEP_SCHEMA = 2;
const STATE_KIND = 'ms365-kursteams-state';
const STATE_FILE_VERSION = 1;

function migrateKursteamStepFromStorage(step, storedSchema) {
    if (storedSchema >= KURSTEAM_STEP_SCHEMA) return step;
    const legacy = {
        2.5: 3,
        3: 4,
        4: 5,
        5: 6,
        6: 7,
        5.5: 8
    };
    return Object.prototype.hasOwnProperty.call(legacy, step) ? legacy[step] : step;
}

function safeEl(id) {
    return document.getElementById(id);
}

function safeInputValue(id, fallback) {
    const el = safeEl(id);
    if (!el) return fallback;
    return el.value;
}

function safeCheckbox(id, fallback) {
    const el = safeEl(id);
    if (!el) return fallback;
    return !!el.checked;
}

/**
 * Aktuellen Kursteams-Stand als plain object (für localStorage + JSON-Datei).
 */
ns.buildKursteamStateSnapshot = function buildKursteamStateSnapshot() {
    if (typeof ns.getPatternFromBuilder === 'function') {
        try {
            ns.teamNamePattern = ns.getPatternFromBuilder();
        } catch {
            /* ignore */
        }
    }

    return {
        kind: STATE_KIND,
        version: STATE_FILE_VERSION,
        stepSchema: KURSTEAM_STEP_SCHEMA,
        exportedAt: new Date().toISOString(),
        rawData: Array.isArray(ns.rawData) ? ns.rawData : [],
        filteredData: Array.isArray(ns.filteredData) ? ns.filteredData : [],
        teamsData: Array.isArray(ns.teamsData) ? ns.teamsData : [],
        teacherEmailMapping:
            ns.teacherEmailMapping && typeof ns.teacherEmailMapping === 'object'
                ? ns.teacherEmailMapping
                : {},
        teamsGenerated: !!ns.teamsGenerated,
        currentStep: Number.isFinite(ns.currentStep) ? ns.currentStep : 0,
        yearPrefix: safeInputValue('yearPrefix', 'SJ26'),
        schoolDomain:
            typeof window.ms365GetSchoolDomainNoAt === 'function'
                ? window.ms365GetSchoolDomainNoAt()
                : '',
        teamSeparator: safeInputValue('teamSeparator', ' | '),
        teamNamePattern: ns.teamNamePattern || null,
        excludeSubjects: safeInputValue('excludeSubjects', 'ORD,DIR,KV'),
        removeDuplicates: safeCheckbox('removeDuplicates', true),
        kursteamEntryMode: ns.kursteamEntryMode,
        studentRosterRaw: ns.studentRosterRaw || '',
        studentRosterPreferGroup: safeCheckbox('studentRosterPreferGroup', true),
        studentRosterSkipCombinedClasses: safeCheckbox('studentRosterSkipCombinedClasses', true),
        studentRosterHideNoMatch: safeCheckbox('studentRosterHideNoMatch', true),
        studentRosterTeamSelection: ns.studentRosterTeamSelection || {},
        webuntisPaste: safeInputValue('webuntisPasteInput', '')
    };
};

function looksLikeKursteamState(obj) {
    if (!obj || typeof obj !== 'object') return false;
    if (obj.kind === STATE_KIND) return true;
    return (
        Array.isArray(obj.rawData) ||
        Array.isArray(obj.filteredData) ||
        Array.isArray(obj.teamsData) ||
        (obj.teacherEmailMapping && typeof obj.teacherEmailMapping === 'object')
    );
}

/**
 * Stellt einen Snapshot wieder her (UI + Namespace).
 * @returns {{ rows: number, teams: number, step: number }}
 */
ns.applyKursteamStateSnapshot = function applyKursteamStateSnapshot(state) {
    if (!looksLikeKursteamState(state)) {
        throw new Error(
            'Keine gültige Kursteams-Stand-Datei (erwarte rawData/teamsData oder kind=ms365-kursteams-state).'
        );
    }

    ns.rawData = Array.isArray(state.rawData) ? state.rawData : [];
    ns.filteredData = Array.isArray(state.filteredData) ? state.filteredData : [];
    ns.teamsData = Array.isArray(state.teamsData) ? state.teamsData : [];
    ns.teacherEmailMapping =
        state.teacherEmailMapping && typeof state.teacherEmailMapping === 'object'
            ? state.teacherEmailMapping
            : {};
    ns.teamsGenerated = !!state.teamsGenerated;
    ns.kursteamEntryMode =
        state.kursteamEntryMode === 'manual' || state.kursteamEntryMode === 'webuntis'
            ? state.kursteamEntryMode
            : ns.rawData.length
              ? 'webuntis'
              : 'unset';

    const yp = safeEl('yearPrefix');
    if (yp) yp.value = state.yearPrefix || 'SJ26';

    if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
        const sd = state.schoolDomain;
        const legacy = state.emailDomain;
        if (sd !== undefined && sd !== null && String(sd).trim() !== '') {
            window.ms365SetSchoolDomainNoAt(sd);
        } else if (legacy !== undefined && legacy !== null && String(legacy).trim() !== '') {
            window.ms365SetSchoolDomainNoAt(
                String(legacy)
                    .trim()
                    .replace(/^@+/, '')
            );
        }
    }

    const sep = safeEl('teamSeparator');
    if (sep) sep.value = state.teamSeparator !== undefined ? state.teamSeparator : ' | ';

    ns.teamNamePattern = state.teamNamePattern || null;
    if (typeof ns.renderTeamNameBuilder === 'function') ns.renderTeamNameBuilder();

    const ex = safeEl('excludeSubjects');
    if (ex) ex.value = state.excludeSubjects !== undefined ? state.excludeSubjects : 'ORD,DIR,KV';
    const rd = safeEl('removeDuplicates');
    if (rd) rd.checked = state.removeDuplicates !== false;
    if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();

    const paste = safeEl('webuntisPasteInput');
    if (paste && typeof state.webuntisPaste === 'string') paste.value = state.webuntisPaste;

    ns.studentRosterRaw = state.studentRosterRaw || '';
    const pref = safeEl('studentRosterPreferGroup');
    const skip = safeEl('studentRosterSkipCombinedClasses');
    const hide = safeEl('studentRosterHideNoMatch');
    if (pref) pref.checked = state.studentRosterPreferGroup !== false;
    if (skip) skip.checked = state.studentRosterSkipCombinedClasses !== false;
    if (hide) hide.checked = state.studentRosterHideNoMatch !== false;
    if (ns.studentRosterRaw && typeof ns.parseStudentRosterFromText === 'function') {
        ns.parseStudentRosterFromText(ns.studentRosterRaw);
    }
    ns.studentRosterTeamSelection = state.studentRosterTeamSelection || {};
    if (typeof ns.refreshStudentRosterUI === 'function') ns.refreshStudentRosterUI();

    if (ns.rawData.length) {
        const total = safeEl('totalRecords');
        const us = safeEl('uniqueSubjects');
        const ut = safeEl('uniqueTeachers');
        const stats = safeEl('importStats');
        if (total) total.textContent = String(ns.rawData.length);
        if (us) us.textContent = String(new Set(ns.rawData.map((r) => r.fach).filter((f) => f)).size);
        if (ut) ut.textContent = String(new Set(ns.rawData.map((r) => r.lehrer).filter((l) => l)).size);
        if (stats) stats.style.display = 'block';
        if (typeof ns.setContinueButton === 'function') {
            ns.setContinueButton('continueBtn1', true, '');
        }
    }

    if (ns.filteredData.length) {
        const fr = safeEl('filteredRecords');
        const fs = safeEl('filterStats');
        if (fr) fr.textContent = String(ns.filteredData.length);
        if (fs) fs.style.display = 'block';
        if (typeof ns.displayFilteredData === 'function') ns.displayFilteredData();
        if (typeof ns.setContinueButton === 'function') {
            ns.setContinueButton('continueBtn2', true, '');
        }
    }

    const mapCount = Object.keys(ns.teacherEmailMapping).length;
    if (mapCount) {
        const tc = safeEl('teacherCount');
        const info = safeEl('teacherMappingInfo');
        if (tc) tc.textContent = String(mapCount);
        if (info) info.style.display = 'block';
        if (typeof ns.displayTeacherMappingTable === 'function') ns.displayTeacherMappingTable();
    }

    if (ns.teamsData.length && ns.teamsGenerated) {
        if (typeof ns.displayTeamsData === 'function') ns.displayTeamsData();
        if (typeof ns.setContinueButton === 'function') {
            ns.setContinueButton('continueBtn4', true, '');
        }
    }

    if (typeof ns.updateStep5Checklist === 'function') ns.updateStep5Checklist();
    if (typeof ns.updateStep4Checklist === 'function') ns.updateStep4Checklist();

    const hasRows = ns.rawData.length > 0;
    const stepRaw = state.currentStep !== undefined ? state.currentStep : hasRows ? 1 : 0;
    const step = migrateKursteamStepFromStorage(stepRaw, state.stepSchema || 0);
    ns.currentStep = step;
    if (typeof ns.goToStep === 'function') ns.goToStep(step);

    ns.autoSaveDirty = false;
    return {
        rows: ns.rawData.length,
        teams: ns.teamsData.length,
        step
    };
};

ns.saveStateToStorage = function saveStateToStorage(options) {
    const quiet = !!(options && options.quiet);
    try {
        const state = ns.buildKursteamStateSnapshot();
        localStorage.setItem(ns.STORAGE_KEY, JSON.stringify(state));
        ns.autoSaveDirty = false;
        if (!quiet) {
            ns.showToast(
                'Kursteams: Zwischenstand gespeichert (' +
                    state.rawData.length +
                    ' Zeilen, ' +
                    state.teamsData.length +
                    ' Teams).'
            );
        }
        return true;
    } catch (e) {
        const msg = String(e && e.message ? e.message : e);
        const quota = /quota|speicher|exceeded/i.test(msg) || e.name === 'QuotaExceededError';
        if (!quiet) {
            ns.showToast(
                'Speichern fehlgeschlagen: ' +
                    msg +
                    (quota ? ' – bitte „JSON exportieren“ nutzen (localStorage voll).' : '')
            );
        }
        return false;
    }
};

ns.loadStateFromStorage = function loadStateFromStorage() {
    try {
        const raw = localStorage.getItem(ns.STORAGE_KEY);
        if (!raw) {
            ns.showToast('Kein gespeicherter Stand gefunden.');
            return false;
        }
        const state = JSON.parse(raw);
        const info = ns.applyKursteamStateSnapshot(state);
        ns.showToast(
            'Kursteams: Stand geladen (' +
                info.rows +
                ' Zeilen, ' +
                info.teams +
                ' Teams, Schritt ' +
                info.step +
                ').'
        );
        return true;
    } catch (e) {
        ns.showToast('Laden fehlgeschlagen: ' + (e.message || e));
        return false;
    }
};

ns.exportKursteamStateJson = function exportKursteamStateJson() {
    try {
        const state = ns.buildKursteamStateSnapshot();
        if (
            !state.rawData.length &&
            !state.teamsData.length &&
            !Object.keys(state.teacherEmailMapping).length
        ) {
            ns.showToast('Nichts zu exportieren – zuerst Daten importieren oder Teams erzeugen.');
            return;
        }
        const stamp = new Date().toISOString().slice(0, 10);
        const filename = 'kursteams-stand-' + stamp + '.json';
        const body = JSON.stringify(state, null, 2);
        if (typeof ns.downloadBlob === 'function') {
            ns.downloadBlob(filename, body, 'application/json;charset=utf-8');
        } else {
            const blob = new Blob([body], { type: 'application/json;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = filename;
            a.click();
            URL.revokeObjectURL(a.href);
        }
        try {
            localStorage.setItem(ns.STORAGE_KEY, JSON.stringify(state));
            ns.autoSaveDirty = false;
        } catch {
            /* JSON-Download ist die Backup-Priorität */
        }
        ns.showToast(
            'Kursteams: JSON exportiert (' +
                state.rawData.length +
                ' Zeilen, ' +
                state.teamsData.length +
                ' Teams).'
        );
    } catch (e) {
        ns.showToast('JSON-Export fehlgeschlagen: ' + (e.message || e));
    }
};

ns.importKursteamStateJsonText = function importKursteamStateJsonText(text) {
    const state = JSON.parse(text);
    const info = ns.applyKursteamStateSnapshot(state);
    try {
        localStorage.setItem(ns.STORAGE_KEY, JSON.stringify(ns.buildKursteamStateSnapshot()));
        ns.autoSaveDirty = false;
    } catch {
        /* Import in UI trotzdem ok */
    }
    ns.showToast(
        'Kursteams: JSON importiert (' +
            info.rows +
            ' Zeilen, ' +
            info.teams +
            ' Teams, Schritt ' +
            info.step +
            ').'
    );
    return info;
};

ns.clearStorage = function clearStorage() {
    ns.confirmModal(
        'Lokalen Speicher löschen',
        'Den gespeicherten Zwischenstand für Kursteams in diesem Browser wirklich löschen?',
        () => {
            try {
                localStorage.removeItem(ns.STORAGE_KEY);
                ns.autoSaveDirty = false;
                ns.showToast('Kursteams: Lokaler Speicher wurde geleert.');
            } catch (e) {
                ns.showToast('Fehler: ' + e.message);
            }
        }
    );
};

function wireJsonImportExport() {
    const btnExport = document.getElementById('btnExportKursteamJson');
    if (btnExport) {
        btnExport.addEventListener('click', () => ns.exportKursteamStateJson());
    }
    const fileInput = document.getElementById('kursteamImportJsonFile');
    const btnImport = document.getElementById('btnImportKursteamJson');
    if (btnImport && fileInput) {
        btnImport.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', async (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            try {
                const text = await f.text();
                const hasWork =
                    (Array.isArray(ns.rawData) && ns.rawData.length) ||
                    (Array.isArray(ns.teamsData) && ns.teamsData.length);
                if (hasWork) {
                    ns.confirmModal(
                        'JSON importieren',
                        'Vorhandene Kursteams-Daten in dieser Sitzung werden überschrieben. Fortfahren?',
                        () => {
                            try {
                                ns.importKursteamStateJsonText(text);
                            } catch (err) {
                                ns.showToast('JSON-Import fehlgeschlagen: ' + (err.message || err));
                            }
                        }
                    );
                } else {
                    ns.importKursteamStateJsonText(text);
                }
            } catch (err) {
                ns.showToast('JSON-Import fehlgeschlagen: ' + (err.message || err));
            } finally {
                fileInput.value = '';
            }
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireJsonImportExport);
} else {
    wireJsonImportExport();
}

ns.autoSaveDirty = false;

ns.markAutoSaveDirty = function markAutoSaveDirty() {
    ns.autoSaveDirty = true;
};

setInterval(function () {
    if (ns.autoSaveDirty && (ns.rawData.length || ns.teamsData.length)) {
        ns.saveStateToStorage({ quiet: true });
    }
}, 60000);

window.addEventListener('beforeunload', function (e) {
    if (ns.autoSaveDirty && (ns.rawData.length || ns.teamsData.length)) {
        const msg = 'Es gibt ungespeicherte Kursteams-Daten. Seite wirklich verlassen?';
        e.preventDefault();
        e.returnValue = msg;
        return msg;
    }
});

window.exportKursteamStateJson = ns.exportKursteamStateJson;
