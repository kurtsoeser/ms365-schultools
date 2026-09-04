
function normTeamsFilterToken(s) {
    return String(s || '').trim().toUpperCase();
}

/**
 * @param {object} team
 * @returns {{ klasse: string, fach: string, lehrer: string }}
 */
function teamFilterFields(team) {
    const t = team || {};
    return {
        klasse: normTeamsFilterToken(t.originalClass || t.klasse || ''),
        fach: normTeamsFilterToken(t.fach || ''),
        lehrer: normTeamsFilterToken(t.lehrerCode || t.lehrer || '')
    };
}

/**
 * @param {Array} teamsData
 * @param {{ klasse?: string, fach?: string, lehrer?: string, status?: string, q?: string }} filters
 *   status: '' | 'valid' | 'invalid'
 *   q: Freitext über Team-Name, Gruppenmail, Besitzer
 * @returns {Array<{ team: object, index: number }>}
 */
function filterTeamsWithIndices(teamsData, filters) {
    const f = filters || {};
    const klasse = normTeamsFilterToken(f.klasse);
    const fach = normTeamsFilterToken(f.fach);
    const lehrer = normTeamsFilterToken(f.lehrer);
    const status = String(f.status || '').trim().toLowerCase();
    const q = normTeamsFilterToken(f.q);
    const out = [];

    (teamsData || []).forEach((team, index) => {
        if (!team) return;
        if (team.ktManualDraft) {
            // Entwürfe immer zeigen, damit man sie nicht „verliert“
            out.push({ team, index });
            return;
        }
        const fields = teamFilterFields(team);
        if (klasse && !fields.klasse.includes(klasse)) return;
        if (fach && !fields.fach.includes(fach)) return;
        if (lehrer && !fields.lehrer.includes(lehrer)) return;
        if (status === 'valid' && !team.isValid) return;
        if (status === 'invalid' && team.isValid) return;
        if (q) {
            const hay = [
                team.teamName,
                team.gruppenmail,
                team.besitzer,
                fields.klasse,
                fields.fach,
                fields.lehrer,
                team.gruppe
            ]
                .map((x) => normTeamsFilterToken(x))
                .join(' ');
            if (!hay.includes(q)) return;
        }
        out.push({ team, index });
    });
    return out;
}

/**
 * Unique Werte für Datalists (sortiert).
 * @param {Array} teamsData
 * @param {'klasse'|'fach'|'lehrer'} field
 */
function collectUniqueTeamFilterValues(teamsData, field) {
    const set = new Set();
    (teamsData || []).forEach((team) => {
        if (!team || team.ktManualDraft) return;
        const f = teamFilterFields(team);
        const v = f[field];
        if (v) set.add(v);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
}

window.ms365KursteamTeamsFilterLogic = {
    normTeamsFilterToken,
    teamFilterFields,
    filterTeamsWithIndices,
    collectUniqueTeamFilterValues
};
