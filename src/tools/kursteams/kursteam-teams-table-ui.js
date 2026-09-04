
const KTS = window.ms365KursteamTeamsSortLogic;
const KTF = window.ms365KursteamTeamsFilterLogic;
window.ms365AssertModules({ KTS, KTF }, 'kursteam-teams-table-ui.js');

function readTeamsFilterFromDom() {
    return {
        klasse: document.getElementById('teamsFilterKlasse')?.value || '',
        fach: document.getElementById('teamsFilterFach')?.value || '',
        lehrer: document.getElementById('teamsFilterLehrer')?.value || '',
        status: document.getElementById('teamsFilterStatus')?.value || '',
        q: document.getElementById('teamsFilterQ')?.value || ''
    };
}

function fillDatalist(id, values) {
    const list = document.getElementById(id);
    if (!list) return;
    list.replaceChildren();
    (values || []).forEach((v) => {
        const opt = document.createElement('option');
        opt.value = v;
        list.appendChild(opt);
    });
}

function refreshTeamsFilterDatalists(ns) {
    fillDatalist('teamsFilterKlasseList', KTF.collectUniqueTeamFilterValues(ns.teamsData, 'klasse'));
    fillDatalist('teamsFilterFachList', KTF.collectUniqueTeamFilterValues(ns.teamsData, 'fach'));
    fillDatalist('teamsFilterLehrerList', KTF.collectUniqueTeamFilterValues(ns.teamsData, 'lehrer'));
}

function updateTeamsFilterSummary(shown, total) {
    const el = document.getElementById('teamsFilterSummary');
    if (!el) return;
    if (!total) {
        el.textContent = '';
        return;
    }
    if (shown === total) {
        el.textContent = total + ' Team(s)';
    } else {
        el.textContent = shown + ' von ' + total + ' Team(s) angezeigt';
    }
}

/**
 * Rendert die Team-Tabelle (Schritt mit Validierung / manuelle Zeilen).
 * @param {object} ns window.ms365Kursteam
 */
function render(ns) {
    const tbody = document.getElementById('teamsTableBody');
    tbody.replaceChildren();

    const validCount = ns.teamsData.filter((t) => t.isValid).length;
    const invalidCount = ns.teamsData.length - validCount;
    const mappedCount = ns.teamsData.filter((t) => t.mappingUsed).length;
    const dupAdj = ns.teamsData.filter((t) => t.mailNicknameAdjusted).length;
    document.getElementById('duplicateMailAdjustments').textContent = dupAdj;

    if (!ns.teamsSort) ns.teamsSort = { key: 'teamName', dir: 1 };
    ns.teamsFilter = readTeamsFilterFromDom();

    refreshTeamsFilterDatalists(ns);

    const filtered = KTF.filterTeamsWithIndices(ns.teamsData, ns.teamsFilter);
    const sortedTeams = KTS.sortTeamsWithIndices(
        filtered.map((x) => x.team),
        ns.teamsSort
    );
    const indexByTeam = new Map(filtered.map((f) => [f.team, f.index]));
    const viewFinal = sortedTeams.map(({ team }) => ({
        team,
        index: indexByTeam.has(team) ? indexByTeam.get(team) : -1
    }));

    updateTeamsFilterSummary(viewFinal.length, ns.teamsData.length);

    const table = document.getElementById('teamsTableContainer');
    const ths = table ? table.querySelectorAll('th[data-teams-sort-key]') : [];
    ths.forEach((th) => {
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

    viewFinal.forEach(({ team, index }) => {
        if (index < 0) return;
        const tr = document.createElement('tr');
        if (!team.isValid) tr.classList.add('error-row');

        if (team.ktManualDraft) {
            tr.classList.add('kt-team-draft-row');
            tr.setAttribute('data-team-draft-index', String(index));

            const td1 = document.createElement('td');
            const inpName = document.createElement('input');
            inpName.type = 'text';
            inpName.className = 'kt-team-draft-input';
            inpName.setAttribute('data-team-draft-field', 'teamName');
            inpName.value = team.teamName || '';
            inpName.placeholder = 'z. B. SJ26 | 1A | D';
            inpName.autocomplete = 'off';
            td1.appendChild(inpName);

            const td2 = document.createElement('td');
            const inpGm = document.createElement('input');
            inpGm.type = 'text';
            inpGm.className = 'kt-team-draft-input';
            inpGm.setAttribute('data-team-draft-field', 'gruppenmail');
            inpGm.value = team.gruppenmail || '';
            inpGm.placeholder = 'z. B. SJ26-1A-D';
            inpGm.autocomplete = 'off';
            td2.appendChild(inpGm);

            const td3 = document.createElement('td');
            const inpOwn = document.createElement('input');
            inpOwn.type = 'email';
            inpOwn.className = 'kt-team-draft-input';
            inpOwn.setAttribute('data-team-draft-field', 'besitzer');
            inpOwn.value = team.besitzer || '';
            inpOwn.placeholder = 'besitzer@schule.de';
            inpOwn.autocomplete = 'off';
            td3.appendChild(inpOwn);

            const td4 = document.createElement('td');
            td4.textContent = '…';

            const td5 = document.createElement('td');
            const bDel = document.createElement('button');
            bDel.type = 'button';
            bDel.className = 'btn btn-small btn-danger kt-delete-btn';
            bDel.textContent = '🗑️';
            bDel.setAttribute('aria-label', 'Zeile entfernen');
            bDel.addEventListener('click', () => ns.deleteTeam(index));
            td5.appendChild(bDel);

            tr.append(td1, td2, td3, td4, td5);

            tr.addEventListener('focusout', (e) => {
                const r = e.relatedTarget;
                if (r && tr.contains(r)) return;
                setTimeout(() => {
                    if (document.activeElement && tr.contains(document.activeElement)) return;
                    ns.commitManualTeamDraftRow(index);
                }, 0);
            });

            tbody.appendChild(tr);
            return;
        }

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
        smallM.textContent = team.mappingUsed ? '✓ Mapping' : '⚠ Generiert (' + (team.lehrerCode || '') + ')';
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
    const filterBar = document.getElementById('teamsFilterBar');
    if (filterBar) filterBar.style.display = ns.teamsData.length ? 'block' : 'none';
    if (typeof ns.updateStep5Checklist === 'function') ns.updateStep5Checklist();

    const manRow = document.getElementById('kursteamManualAddRow');
    if (manRow) manRow.style.display = ns.teamsGenerated ? '' : 'none';
}

let teamsFilterWired = false;
function wireTeamsFilterOnce(ns) {
    if (teamsFilterWired) return;
    teamsFilterWired = true;
    const ids = ['teamsFilterKlasse', 'teamsFilterFach', 'teamsFilterLehrer', 'teamsFilterStatus', 'teamsFilterQ'];
    ids.forEach((id) => {
        const el = document.getElementById(id);
        if (!el) return;
        const evt = el.tagName === 'SELECT' ? 'change' : 'input';
        el.addEventListener(evt, () => {
            if (typeof ns.displayTeamsData === 'function') ns.displayTeamsData();
        });
    });
    const reset = document.getElementById('teamsFilterReset');
    if (reset) {
        reset.addEventListener('click', () => {
            ids.forEach((id) => {
                const el = document.getElementById(id);
                if (!el) return;
                el.value = '';
            });
            if (typeof ns.displayTeamsData === 'function') ns.displayTeamsData();
        });
    }
}

window.ms365KursteamTeamsTableUI = {
    render,
    wireTeamsFilterOnce
};
