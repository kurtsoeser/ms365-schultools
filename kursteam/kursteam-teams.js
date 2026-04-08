(function () {
    'use strict';

    const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

    ns.applyFilters = function applyFilters() {
        const excludeSubjects = document
            .getElementById('excludeSubjects')
            .value.split(',')
            .map(s => s.trim().toUpperCase())
            .filter(s => s.length > 0);
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
            const td0 = document.createElement('td');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger';
            btn.textContent = '❌';
            btn.addEventListener('click', () => ns.removeRow(index));
            td0.appendChild(btn);
            const td1 = document.createElement('td');
            td1.textContent = row.klasse;
            const td2 = document.createElement('td');
            td2.textContent = row.fach;
            const td3 = document.createElement('td');
            td3.textContent = row.lehrer;
            const td4 = document.createElement('td');
            td4.textContent = row.gruppe || '-';
            tr.append(td0, td1, td2, td3, td4);
            tbody.appendChild(tr);
        });
        const hasRows = ns.filteredData.length > 0;
        document.getElementById('dataTableContainer').style.display = hasRows ? 'block' : 'none';
        document.getElementById('continueBtn2').style.display = hasRows ? 'inline-block' : 'none';
    };

    ns.removeRow = function removeRow(index) {
        const row = ns.filteredData[index];
        ns.filteredData.splice(index, 1);
        if (row && row.id !== undefined && row.id !== null) {
            const ri = ns.rawData.findIndex(r => r.id === row.id);
            if (ri >= 0) ns.rawData.splice(ri, 1);
        }
        ns.invalidateTeams();
        ns.displayFilteredData();
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
        document.getElementById('excludeSubjects').value = 'ORD,DIR,KV';
        document.getElementById('removeDuplicates').checked = true;
        ns.applyFilters();
    };

    ns.generateTeamNames = function generateTeamNames() {
        const yearPrefix = document.getElementById('yearPrefix').value;
        const emailDomain =
            typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
                ? window.ms365GetTeacherEmailDomainSuffix()
                : '@';
        const separator = document.getElementById('teamSeparator').value;

        ns.teamsData = ns.filteredData.map(row => {
            let klasseForName = row.klasse;
            if (row.klasse.includes(',')) klasseForName = ns.combineClassNames(row.klasse);

            const teamName = `${yearPrefix}${separator}${klasseForName}${separator}${row.fach}`;
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

        ns.teamsData.forEach((team, index) => {
            const tr = document.createElement('tr');
            if (!team.isValid) tr.classList.add('error-row');

            const td1 = document.createElement('td');
            td1.appendChild(document.createTextNode(team.teamName));
            if (team.originalClass && team.originalClass.includes(',')) {
                td1.appendChild(document.createElement('br'));
                const small = document.createElement('small');
                small.style.color = '#6c757d';
                small.textContent = '(Original: ' + team.originalClass + ')';
                td1.appendChild(small);
            }

            const td2 = document.createElement('td');
            td2.appendChild(document.createTextNode(team.gruppenmail));
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
            td3.appendChild(document.createElement('br'));
            const smallM = document.createElement('small');
            smallM.style.color = team.mappingUsed ? '#28a745' : '#ffc107';
            smallM.textContent = team.mappingUsed ? '✓ Mapping' : '⚠ Generiert (' + team.lehrerCode + ')';
            td3.appendChild(smallM);

            const td4 = document.createElement('td');
            td4.textContent = team.isValid ? '✅' : '❌ ' + (team.error || 'Fehler');

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

        document.getElementById('validationResults').style.display = 'block';
        document.getElementById('teamsTableContainer').style.display = 'block';
        document.getElementById('continueBtn3').style.display = 'inline-block';
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
    window.applyFilters = ns.applyFilters;
    window.resetFilters = ns.resetFilters;
    window.generateTeamNames = ns.generateTeamNames;
    window.addManualKursteamTeam = ns.addManualKursteamTeam;
})();

