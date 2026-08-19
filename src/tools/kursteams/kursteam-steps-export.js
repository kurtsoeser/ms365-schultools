import {
    buildStandaloneKursteamPs1,
    buildStandaloneKursteamPs1V2,
    buildKursteamCsvPreviewPs1
} from './kursteam-ps-export.js';

const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

ns.updateTeacherStats = function updateTeacherStats() {
    // Lehrer die in den importierten Unterrichtsdaten vorkommen
    const uniqueTeachers = new Set((ns.filteredData || []).map(row => (row.lehrer || '').toUpperCase().trim()).filter(Boolean));
    const teachersArray = Array.from(uniqueTeachers);
    const mappedCount = teachersArray.filter(t => ns.teacherEmailMapping[t]).length;
    const unmappedCount = teachersArray.length - mappedCount;

    document.getElementById('uniqueTeachersNeeded').textContent = teachersArray.length;
    document.getElementById('mappedTeachers').textContent = mappedCount;
    document.getElementById('unmappedTeachers').textContent = unmappedCount;
    document.getElementById('teacherRequiredStats').style.display = 'grid';

    // Fehlende-Lehrer-Sektion
    if (unmappedCount > 0) ns.displayMissingTeachers(teachersArray);
    else document.getElementById('missingTeachersSection').style.display = 'none';

    // Zuordnungstabelle + Schul-Einstellungen-Sync-Panel
    ns.displayTeacherMappingTableWithUsage(teachersArray);
    ns.displayTenantTeacherSyncPanel(teachersArray);
    if (typeof ns.updateStep4Checklist === 'function') ns.updateStep4Checklist();
};

/**
 * Zeigt fehlende Lehrer (kommen in Unterrichtsdaten vor, haben aber keine E-Mail-Zuordnung).
 * Schlägt E-Mail aus Schul-Einstellungen vor falls dort ein namensgleicher Eintrag ohne E-Mail existiert.
 */
ns.displayMissingTeachers = function displayMissingTeachers(allTeachers) {
    const unmappedTeachers = allTeachers.filter(t => !ns.teacherEmailMapping[t]);
    if (unmappedTeachers.length === 0) {
        document.getElementById('missingTeachersSection').style.display = 'none';
        return;
    }
    const emailDomain =
        typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
            ? window.ms365GetTeacherEmailDomainSuffix()
            : '@';

    // Schul-Einstellungen: Lehrer ohne E-Mail aber mit Name als Vorschlag-Quelle
    const tenantTeachers = typeof window.ms365TenantSettingsLoad === 'function'
        ? (window.ms365TenantSettingsLoad().teachers || [])
        : [];
    const tenantByCode = new Map(tenantTeachers.map(t => [String(t.code || '').toUpperCase(), t]));

    const tbody = document.getElementById('missingTeachersBody');
    tbody.replaceChildren();
    unmappedTeachers.forEach(kuerzel => {
        const tenantEntry = tenantByCode.get(kuerzel);
        // Vorschlag: aus Schul-Einstellungen (Name→E-Mail ableiten) oder Domain-Fallback
        const suggestedEmail = (tenantEntry && tenantEntry.email)
            ? tenantEntry.email
            : kuerzel.toLowerCase() + emailDomain;
        const tenantName = tenantEntry && tenantEntry.name ? tenantEntry.name : '';
        const inTenant = !!tenantEntry;

        const tr = document.createElement('tr');

        // Kürzel + Schul-Einstellungen-Badge
        const td1 = document.createElement('td');
        td1.style.whiteSpace = 'nowrap';
        const strong = document.createElement('strong');
        strong.textContent = kuerzel;
        td1.appendChild(strong);
        if (tenantName) {
            const small = document.createElement('div');
            small.style.cssText = 'font-size:0.8em;color:var(--text-secondary);';
            small.textContent = tenantName;
            td1.appendChild(small);
        }
        if (inTenant) {
            const badge = document.createElement('span');
            badge.style.cssText = 'font-size:0.72em;background:var(--brand1);color:#fff;border-radius:4px;padding:1px 5px;margin-top:2px;display:inline-block;';
            badge.title = 'In Schul-Einstellungen vorhanden';
            badge.textContent = 'Schule ✓';
            td1.appendChild(badge);
        }

        // E-Mail Eingabefeld (sofort editierbar)
        const td2 = document.createElement('td');
        const input = document.createElement('input');
        input.type = 'email';
        input.className = 'kt-team-draft-input';
        input.placeholder = suggestedEmail;
        input.value = suggestedEmail;
        input.style.minWidth = '220px';
        td2.appendChild(input);

        // Aktion
        const td3 = document.createElement('td');
        td3.style.whiteSpace = 'nowrap';
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-small btn-brand';
        btn.innerHTML = '<i class="bi bi-check2"></i> Übernehmen';
        btn.addEventListener('click', () => {
            const email = input.value.trim().toLowerCase();
            if (!email || !email.includes('@')) {
                ns.showToast('Bitte eine gültige E-Mail eingeben.');
                input.focus();
                return;
            }
            ns.quickAddTeacher(kuerzel, email, true);
        });
        td3.appendChild(btn);

        tr.append(td1, td2, td3);
        tbody.appendChild(tr);
    });
    document.getElementById('missingTeachersSection').style.display = 'block';
};

/**
 * Schnelles Hinzufügen aus der Fehlenden-Tabelle (ohne Modal wenn direkt=true).
 */
ns.quickAddTeacher = function quickAddTeacher(kuerzel, suggestedEmail, direct) {
    function doAdd(email) {
        ns.teacherEmailMapping[kuerzel] = email;
        document.getElementById('teacherCount').textContent = Object.keys(ns.teacherEmailMapping).length;
        document.getElementById('teacherMappingInfo').style.display = 'block';
        if (typeof window.ms365TenantSettingsSave === 'function') {
            // In Schul-Einstellungen zurückschreiben
            const current = window.ms365TenantSettingsLoad();
            const teachers = Array.isArray(current && current.teachers) ? [...current.teachers] : [];
            const idx = teachers.findIndex(t => String(t.code || '').toUpperCase() === kuerzel);
            if (idx >= 0) {
                teachers[idx] = { ...teachers[idx], email };
            } else {
                teachers.push({ code: kuerzel, name: '', email });
            }
            window.ms365TenantSettingsSave({ ...current, teachers });
        }
        if (typeof ns.markAutoSaveDirty === 'function') ns.markAutoSaveDirty();
        ns.showToast(kuerzel + ' → ' + email + ' gespeichert (auch in Schul-Einstellungen).');
        ns.updateTeacherStats();
    }

    if (direct) {
        doAdd(suggestedEmail);
        return;
    }
    ns.openModal(
        'E-Mail für ' + kuerzel,
        '<label for="quickEmail">E-Mail-Adresse</label><input type="email" id="quickEmail" value="' +
            ns.attrEscape(suggestedEmail) + '" style="width:100%;margin-top:6px;">',
        () => {
            const email = document.getElementById('quickEmail').value.trim().toLowerCase();
            if (!email || !email.includes('@')) {
                ns.showToast('Bitte eine gültige E-Mail eingeben.');
                return;
            }
            ns.closeModal();
            doAdd(email);
        }
    );
};

/**
 * Zeigt ein Panel mit Lehrern aus den Schul-Einstellungen die noch KEINE E-Mail haben –
 * damit der User weiß was er noch in den Einstellungen ergänzen sollte.
 * Außerdem: Lehrer aus Schul-Einstellungen die in den Unterrichtsdaten vorkommen aber
 * noch nicht im Mapping sind → werden als Vorschlag angeboten.
 */
ns.displayTenantTeacherSyncPanel = function displayTenantTeacherSyncPanel(requiredTeachers) {
    const panel = document.getElementById('tenantTeacherSyncPanel');
    if (!panel) return;

    if (typeof window.ms365TenantSettingsLoad !== 'function') {
        panel.style.display = 'none';
        return;
    }

    const tenantSettings = window.ms365TenantSettingsLoad();
    const tenantTeachers = Array.isArray(tenantSettings.teachers) ? tenantSettings.teachers : [];

    if (!tenantTeachers.length) {
        panel.innerHTML = '<p style="color:var(--text-secondary);font-size:0.88em;margin:0;">' +
            '<i class="bi bi-info-circle"></i> Keine Lehrer in den <strong>Schul-Grundeinstellungen</strong> hinterlegt. ' +
            'Lehrerliste dort eintragen → hier automatisch verfügbar.</p>';
        panel.style.display = 'block';
        return;
    }

    // Lehrer aus Schul-Einstellungen die ein Kürzel haben und in Unterrichtsdaten vorkommen
    const requiredSet = new Set(requiredTeachers);
    const tenantMapped = tenantTeachers.filter(t => t.code && t.email && requiredSet.has(t.code.toUpperCase()));
    const tenantUnmappedRequired = tenantTeachers.filter(t => t.code && !t.email && requiredSet.has(t.code.toUpperCase()));
    const tenantNotRequired = tenantTeachers.filter(t => t.code && t.email && !requiredSet.has(t.code.toUpperCase()));
    const tenantNoEmail = tenantTeachers.filter(t => t.code && !t.email && !requiredSet.has(t.code.toUpperCase()));

    let html = '<p style="font-size:0.88em;color:var(--text-secondary);margin:0 0 10px;">' +
        '<strong>' + tenantTeachers.length + '</strong> Lehrer in Schul-Einstellungen gespeichert – ' +
        '<strong style="color:var(--ok1);">' + tenantMapped.length + '</strong> davon mit E-Mail und in diesen Unterrichtsdaten aktiv.';

    if (tenantUnmappedRequired.length > 0) {
        html += ' <strong style="color:#e6a817;">' + tenantUnmappedRequired.length + ' Lehrer</strong> ' +
            'aus diesen Daten sind in den Schul-Einstellungen ohne E-Mail – dort ergänzen!';
    }
    html += '</p>';

    // Fehlende E-Mails in Schul-Einstellungen (die in Unterrichtsdaten vorkommen)
    if (tenantUnmappedRequired.length > 0) {
        html += '<details style="margin-top:8px;"><summary style="cursor:pointer;font-size:0.88em;font-weight:600;color:var(--heading);">' +
            '<i class="bi bi-exclamation-triangle-fill" style="color:#e6a817;"></i> ' +
            tenantUnmappedRequired.length + ' Lehrer aus Unterrichtsdaten ohne E-Mail in Schul-Einstellungen</summary>' +
            '<div style="margin-top:8px;overflow-x:auto;"><table style="width:100%;font-size:0.86em;"><thead><tr>' +
            '<th>Kürzel</th><th>Name</th><th>Hinweis</th></tr></thead><tbody>';
        tenantUnmappedRequired.forEach(t => {
            const kuerzel = String(t.code || '').toUpperCase();
            html += '<tr><td><strong>' + ns.escapeHtml(kuerzel) + '</strong></td>' +
                '<td>' + ns.escapeHtml(t.name || '–') + '</td>' +
                '<td style="color:#e6a817;font-size:0.85em;">Bitte in <em>Schul-Grundeinstellungen</em> → Lehrerliste ergänzen, ' +
                'oder oben direkt ➕ Hinzufügen</td></tr>';
        });
        html += '</tbody></table></div></details>';
    }

    // Lehrer aus Schul-Einstellungen die in diesen Unterrichtsdaten NICHT vorkommen
    if (tenantNoEmail.length > 0) {
        html += '<details style="margin-top:6px;"><summary style="cursor:pointer;font-size:0.85em;color:var(--muted);">' +
            tenantNoEmail.length + ' weitere Lehrer in Schul-Einstellungen ohne E-Mail (nicht in diesen Daten)</summary>' +
            '<p style="font-size:0.82em;color:var(--muted);margin:6px 0 0;">Diese Lehrer kommen in den importierten Unterrichtsdaten nicht vor – sie müssen hier nicht zugeordnet werden.</p></details>';
    }

    if (tenantNotRequired.length > 0) {
        html += '<p style="font-size:0.82em;color:var(--muted);margin:6px 0 0;">' +
            tenantNotRequired.length + ' weitere Lehrer mit E-Mail aus Schul-Einstellungen sind in diesen Unterrichtsdaten nicht aktiv.</p>';
    }

    panel.innerHTML = html;
    panel.style.display = 'block';
};

function getKursteamContentRoot() {
    const panel = document.getElementById('panelWebuntis');
    return panel ? panel.querySelector('.content') : null;
}

/** Verschachtelte .step-content-Blöcke (HTML-Fehler) als direkte Kinder von .content auslagern. */
function repairKursteamStepDom() {
    const content = getKursteamContentRoot();
    if (!content) return false;

    const steps = Array.from(content.querySelectorAll('.step-content'));
    const needsRepair = steps.some((el) => el.parentElement !== content);
    if (!needsRepair) return false;

    steps.sort((a, b) => {
        const sa = parseInt(String(a.getAttribute('data-step') || '0'), 10);
        const sb = parseInt(String(b.getAttribute('data-step') || '0'), 10);
        return (Number.isFinite(sa) ? sa : 0) - (Number.isFinite(sb) ? sb : 0);
    });
    steps.forEach((el) => content.appendChild(el));
    return true;
}

function getKursteamStepContentEl(step) {
    const content = getKursteamContentRoot();
    if (!content) return null;
    return content.querySelector(':scope > .step-content[data-step="' + step + '"]');
}

ns.goToStep = function goToStep(rawStep) {
    const panel = document.getElementById('panelWebuntis');
    if (!panel) return;

    let step = parseInt(String(rawStep).trim(), 10);
    if (!Number.isFinite(step) || step < 0 || step > 8) return;

    // Nur ab Schritt 6 (Graph/CSV/Schüler): generierte Teams nötig.
    // Schritt 5 („Teams konfigurieren“) ist bewusst ausgenommen — dort wird erst generiert.
    if (step === 7 || step === 8) {
        const validTeams = ns.teamsData.filter(t => t.isValid);
        if (!ns.teamsGenerated || validTeams.length === 0) {
            ns.showToast('Bitte zuerst unter „Teams konfigurieren“ auf „Team-Namen generieren“ klicken (mindestens ein gültiges Team).');
            step = 5;
        }
    }

    let contentEl = getKursteamStepContentEl(step);
    if (!contentEl && repairKursteamStepDom()) {
        contentEl = getKursteamStepContentEl(step);
    }
    const tabEl =
        panel.querySelector('.steps > .step[data-step="' + step + '"]') ||
        document.querySelector('#panelWebuntis .steps > .step[data-step="' + step + '"]');
    if (!contentEl || !tabEl) return;

    const contentRoot = getKursteamContentRoot();
    (contentRoot ? contentRoot.querySelectorAll(':scope > .step-content') : panel.querySelectorAll('.step-content')).forEach((el) =>
        el.classList.remove('active')
    );
    panel.querySelectorAll('.steps > .step').forEach(el => {
        el.classList.remove('active');
        el.classList.remove('completed');
    });

    contentEl.classList.add('active');
    tabEl.classList.add('active');
    try {
        tabEl.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
    } catch (e) {
        /* ignore */
    }

    const stepOrder = [0, 1, 2, 3, 4, 5, 7, 8];
    const currentIndex = stepOrder.indexOf(step);
    if (currentIndex >= 0) {
        for (let i = 0; i < currentIndex; i++) {
            const prev = panel.querySelector('.steps > .step[data-step="' + stepOrder[i] + '"]');
            if (prev) prev.classList.add('completed');
        }
    }

    ns.currentStep = step;

    if (step === 1) {
        const seeded =
            typeof ns.seedWebuntisPasteIfEmpty === 'function' ? ns.seedWebuntisPasteIfEmpty() : false;
        if (seeded) {
            ns.showToast('Demo: 6 Beispielzeilen aus Schul‑Standards vorbelegt.');
        }
    }
    if (step === 2 && typeof ns.refreshSubjectFilterUI === 'function') {
        ns.refreshSubjectFilterUI();
    }

    const hint = document.getElementById('manualKursteamHint');
    if (hint) hint.style.display = step === 2 && ns.kursteamEntryMode === 'manual' ? 'block' : 'none';

    if (step === 3) {
        if (typeof ns.displayEditableData === 'function') ns.displayEditableData();
        if (typeof ns.displayManualTeamsPreview === 'function') ns.displayManualTeamsPreview();
    }
    if (step === 4) {
        ns.updateTeacherStats();
        if (typeof ns.updateStep4Checklist === 'function') ns.updateStep4Checklist();
    }
    if (step === 5) {
        const manRow = document.getElementById('kursteamManualAddRow');
        if (manRow) manRow.style.display = ns.teamsGenerated ? '' : 'none';
        if (typeof ns.updateStep5Checklist === 'function') ns.updateStep5Checklist();
    }
    if (step === 8) {
        if (typeof ns.seedStudentRosterFromTenantIfEmpty === 'function') {
            const seeded = ns.seedStudentRosterFromTenantIfEmpty();
            if (seeded === 'demo') ns.showToast('Demo: Schülerliste aus Schul‑Standards vorbelegt.');
            else if (seeded === 'tenant') ns.showToast('Schülerliste aus Schul‑Einstellungen übernommen.');
        }
        if (typeof ns.refreshStudentRosterUI === 'function') ns.refreshStudentRosterUI();
    }
    if (step === 7) ns.prepareCSVExport();

    const stepsBar = panel.querySelector('.steps');
    if (typeof window.ms365ApplyStepProgress === 'function') {
        window.ms365ApplyStepProgress(stepsBar, step, stepOrder);
    }

    if (typeof ns.focusStepHeading === 'function') {
        requestAnimationFrame(() => ns.focusStepHeading(step));
    }
};

ns.prepareCSVExport = function prepareCSVExport() {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    document.getElementById('exportCount').textContent = validTeams.length;

    const warn = document.getElementById('step4NoTeamsWarning');
    const ready = document.getElementById('step4ReadyHint');
    const dl = document.getElementById('btnDownloadCsv');
    if (validTeams.length === 0) {
        warn.style.display = 'block';
        ready.style.display = 'none';
        dl.disabled = true;
    } else {
        warn.style.display = 'none';
        ready.style.display = 'block';
        dl.disabled = false;
    }

    let csvPreview = ns.buildCsvRow(['TeamName', 'Gruppenmail', 'Besitzer']);
    validTeams.slice(0, 5).forEach(team => {
        csvPreview += ns.buildCsvRow([team.teamName, team.gruppenmail, team.besitzer]);
    });
    if (validTeams.length > 5) {
        csvPreview += '... (' + (validTeams.length - 5) + ' weitere Teams)\n';
    }
    document.getElementById('csvPreview').textContent = csvPreview;

    const psPreview = document.getElementById('powershellScript');
    if (psPreview) {
        psPreview.textContent = buildStandaloneKursteamPs1V2(validTeams, ns.psEscapeSingle);
    }

    const psCsvPreview = document.getElementById('powershellScriptCsv');
    if (psCsvPreview) psCsvPreview.textContent = buildKursteamCsvPreviewPs1();

    const btnMain = document.getElementById('btnDownloadKursteam');
    if (btnMain) btnMain.disabled = validTeams.length === 0;

    const btnClassic = document.getElementById('btnDownloadKursteamClassic');
    if (btnClassic) btnClassic.disabled = validTeams.length === 0;

    if (typeof ns.refreshKursteamBackendUi === 'function') ns.refreshKursteamBackendUi();
};

ns.downloadCSV = function downloadCSV() {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    if (validTeams.length === 0) {
        ns.showToast('Keine gültigen Teams zum Exportieren.');
        return;
    }
    let csv = ns.buildCsvRow(['TeamName', 'Gruppenmail', 'Besitzer']);
    validTeams.forEach(team => {
        csv += ns.buildCsvRow([team.teamName, team.gruppenmail, team.besitzer]);
    });
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'neueteams.csv';
    link.click();
    URL.revokeObjectURL(link.href);
};

ns.copyPowerShell = function copyPowerShell() {
    const script = document.getElementById('powershellScript')?.textContent || '';
    if (!script.trim()) {
        ns.showToast('Kein Script – zuerst Teams generieren.');
        return;
    }
    navigator.clipboard.writeText(script).then(() => {
        ns.showToast('PowerShell-Script in die Zwischenablage kopiert.');
    });
};

ns.copyPowerShellCsv = function copyPowerShellCsv() {
    const script = document.getElementById('powershellScriptCsv')?.textContent || '';
    navigator.clipboard.writeText(script).then(() => {
        ns.showToast('CSV-PowerShell in die Zwischenablage kopiert.');
    });
};

function downloadKursteamCmdPackage(filename, title, echoLine, buildPs1) {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    if (!validTeams.length) {
        ns.showToast('Keine gültigen Teams – zuerst Team-Namen generieren.');
        return;
    }
    if (typeof window.ms365BuildPolyglotCmd !== 'function') {
        ns.showToast('polyglot-cmd.js fehlt – Seite neu laden.');
        return;
    }
    const ps1 = buildPs1(validTeams, ns.psEscapeSingle);
    const cmd = window.ms365BuildPolyglotCmd({ title, echoLine, psBody: ps1 });
    ns.downloadBlob(filename, cmd);
}

ns.downloadKursteamStandalonePackage = function downloadKursteamStandalonePackage() {
    downloadKursteamCmdPackage(
        'Kursteam-Anlage.cmd',
        'Kursteam-Anlage',
        'Starte Kursteam-Anlage mit PowerShell ...',
        buildStandaloneKursteamPs1V2
    );
    ns.showToast('Kursteam-Anlage.cmd heruntergeladen – bei Abbruch fortsetzbar.');
};

ns.downloadKursteamStandalonePackageClassic = function downloadKursteamStandalonePackageClassic() {
    downloadKursteamCmdPackage(
        'Kursteam-Anlage-einfach.cmd',
        'Kursteam-Anlage (einfach)',
        'Starte Kursteam-Anlage (einfache Variante) ...',
        buildStandaloneKursteamPs1
    );
    ns.showToast('Kursteam-Anlage-einfach.cmd heruntergeladen.');
};

/** @deprecated Alias – nutzt jetzt die empfohlene Variante */
ns.downloadKursteamStandalonePackageV2 = ns.downloadKursteamStandalonePackage;

ns.resetApp = function resetApp() {
    ns.confirmModal('App zurücksetzen', 'Alle Daten in dieser Sitzung wirklich verwerfen? (Lokaler Zwischenstand bleibt, bis Sie ihn löschen.)', () => {
        location.reload();
    });
};

// Step-Header klickbar + keyboard
document.querySelectorAll('#panelWebuntis .steps > .step').forEach(step => {
    step.setAttribute('tabindex', '0');
    step.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            step.click();
        }
    });
    step.addEventListener('click', function () {
        const stepNum = parseInt(String(this.dataset.step).trim(), 10);
        if (!Number.isFinite(stepNum) || stepNum < 0 || stepNum > 8) return;
        ns.goToStep(stepNum);
    });
});

// Global exports für HTML onclick
window.goToStep = function goToStepExport(step) {
    return ns.goToStep(step);
};
window.downloadCSV = ns.downloadCSV;
window.copyPowerShell = ns.copyPowerShell;
window.resetApp = ns.resetApp;
window.downloadKursteamStandalonePackage = ns.downloadKursteamStandalonePackage;
window.downloadKursteamStandalonePackageClassic = ns.downloadKursteamStandalonePackageClassic;
window.downloadKursteamStandalonePackageV2 = ns.downloadKursteamStandalonePackageV2;
window.copyPowerShellCsv = ns.copyPowerShellCsv;

repairKursteamStepDom();
// Defensive Initialisierung: falls ein Browser- oder HMR-Zwischenzustand
// alle active-Klassen verloren hat, den aktuellen Schritt erneut aktivieren.
ns.goToStep(Number.isFinite(ns.currentStep) ? ns.currentStep : 0);

// Snapshot für Microsoft Graph im Browser (kursteam-graph.js).
window.ms365GetKursteamSnapshotForGraph = function () {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    if (!validTeams.length) return null;
    return {
        teams: validTeams.map(t => ({
            teamName: t.teamName,
            gruppenmail: t.gruppenmail,
            besitzer: String(t.besitzer || '').trim()
        }))
    };
};

document.addEventListener('DOMContentLoaded', () => {
    const panel = document.getElementById('panelWebuntis');
    if (!panel || typeof window.ms365ApplyStepProgress !== 'function') return;
    const order = [0, 1, 2, 3, 4, 5, 7, 8];
    const parsed = parseInt(String(ns.currentStep).trim(), 10);
    const step = Number.isFinite(parsed) ? parsed : 0;
    window.ms365ApplyStepProgress(panel.querySelector('.steps'), step, order);
});



