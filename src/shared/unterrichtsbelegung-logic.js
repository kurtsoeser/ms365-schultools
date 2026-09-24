/**
 * Unterrichtsbelegung: bereinigte Kursteam-Liste → kanonische Belegung
 * (Klasse × Lehrkraft × Fach) für App-Persistenz und Folge-Module.
 */

/**
 * @typedef {{
 *   klasse: string,
 *   lehrerCode: string,
 *   lehrerEmail: string,
 *   fach: string,
 *   gruppe: string,
 *   teamName: string,
 *   gruppenmail: string
 * }} UnterrichtsbelegungRow
 */

/**
 * @typedef {{
 *   updatedAt: string,
 *   yearPrefix: string,
 *   source: string,
 *   teamsCount: number,
 *   classCount: number,
 *   teacherCount: number,
 *   rows: UnterrichtsbelegungRow[]
 * }} UnterrichtsbelegungSnapshot
 */

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

/**
 * Eine gültige Belegungszeile aus einem Kursteam-Eintrag.
 * @param {object} team
 * @returns {UnterrichtsbelegungRow|null}
 */
export function rowFromTeamEntry(team) {
    if (!team || typeof team !== 'object') return null;
    if (team.ktManualDraft) return null;
    if (team.isValid === false) return null;

    const klasse = normStr(team.originalClass || team.klasse);
    const lehrerEmail = normEmail(team.besitzer || team.lehrerEmail);
    const fach = normStr(team.fach);
    const teamName = normStr(team.teamName);
    const gruppenmail = normStr(team.gruppenmail).toLowerCase();

    if (!klasse && !fach && !lehrerEmail && !teamName) return null;

    let lehrerCode = normCode(team.lehrerCode || team.lehrer);
    if (!lehrerCode && lehrerEmail) {
        const local = lehrerEmail.split('@')[0] || '';
        lehrerCode = normCode(local);
    }

    return {
        klasse,
        lehrerCode,
        lehrerEmail,
        fach,
        gruppe: normStr(team.gruppe),
        teamName,
        gruppenmail
    };
}

/**
 * @param {object[]} teamsData
 * @returns {UnterrichtsbelegungRow[]}
 */
export function buildRowsFromTeamsData(teamsData) {
    const rows = [];
    const seen = new Set();
    (Array.isArray(teamsData) ? teamsData : []).forEach((t) => {
        const row = rowFromTeamEntry(t);
        if (!row) return;
        const key = [row.klasse, row.lehrerCode, row.fach, row.gruppe, row.gruppenmail].join('|');
        if (seen.has(key)) return;
        seen.add(key);
        rows.push(row);
    });
    rows.sort((a, b) => {
        const c = a.klasse.localeCompare(b.klasse, 'de');
        if (c !== 0) return c;
        const f = a.fach.localeCompare(b.fach, 'de');
        if (f !== 0) return f;
        return a.lehrerCode.localeCompare(b.lehrerCode, 'de');
    });
    return rows;
}

/**
 * @param {UnterrichtsbelegungRow[]} rows
 * @returns {Map<string, UnterrichtsbelegungRow[]>}
 */
export function groupRowsByKlasse(rows) {
    const map = new Map();
    (Array.isArray(rows) ? rows : []).forEach((r) => {
        const k = normStr(r && r.klasse) || '(ohne Klasse)';
        if (!map.has(k)) map.set(k, []);
        map.get(k).push(r);
    });
    return map;
}

/**
 * Eindeutige Lehrkraft-E-Mails (oder Codes) pro Klasse.
 * @param {UnterrichtsbelegungRow[]} rows
 * @returns {Map<string, { email: string, code: string }[]>}
 */
export function teachersByKlasse(rows) {
    const by = new Map();
    groupRowsByKlasse(rows).forEach((list, klasse) => {
        const seen = new Set();
        const teachers = [];
        list.forEach((r) => {
            const key = r.lehrerEmail || r.lehrerCode;
            if (!key || seen.has(key)) return;
            seen.add(key);
            teachers.push({ email: r.lehrerEmail, code: r.lehrerCode });
        });
        teachers.sort((a, b) => (a.code || a.email).localeCompare(b.code || b.email, 'de'));
        by.set(klasse, teachers);
    });
    return by;
}

/**
 * @param {object|null|undefined} raw
 * @returns {UnterrichtsbelegungSnapshot|null}
 */
export function normalizeBelegungSnapshot(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const rows = buildRowsFromTeamsData(
        Array.isArray(raw.rows)
            ? raw.rows.map((r) => ({
                  originalClass: r.klasse,
                  lehrerCode: r.lehrerCode,
                  besitzer: r.lehrerEmail,
                  fach: r.fach,
                  gruppe: r.gruppe,
                  teamName: r.teamName,
                  gruppenmail: r.gruppenmail,
                  isValid: true
              }))
            : []
    );
    if (!rows.length && !normStr(raw.updatedAt) && !normStr(raw.yearPrefix)) return null;

    const classSet = new Set();
    const teacherSet = new Set();
    rows.forEach((r) => {
        if (r.klasse) classSet.add(r.klasse);
        if (r.lehrerEmail) teacherSet.add(r.lehrerEmail);
        else if (r.lehrerCode) teacherSet.add(r.lehrerCode);
    });

    return {
        updatedAt: normStr(raw.updatedAt) || new Date().toISOString(),
        yearPrefix: normStr(raw.yearPrefix),
        source: normStr(raw.source) || 'kursteams',
        teamsCount: Number.isFinite(Number(raw.teamsCount)) ? Number(raw.teamsCount) : rows.length,
        classCount: classSet.size,
        teacherCount: teacherSet.size,
        rows
    };
}

/**
 * Snapshot aus der Kursteam-Endliste (vor/nach Anlage).
 * @param {object[]} teamsData
 * @param {{ yearPrefix?: string, source?: string, updatedAt?: string }} [meta]
 * @returns {UnterrichtsbelegungSnapshot|null}
 */
export function buildSnapshotFromTeamsData(teamsData, meta) {
    const rows = buildRowsFromTeamsData(teamsData);
    if (!rows.length) return null;
    const m = meta && typeof meta === 'object' ? meta : {};
    return normalizeBelegungSnapshot({
        updatedAt: m.updatedAt || new Date().toISOString(),
        yearPrefix: m.yearPrefix || '',
        source: m.source || 'kursteams',
        teamsCount: rows.length,
        rows
    });
}

export function summarizeBelegung(snapshot) {
    const s = normalizeBelegungSnapshot(snapshot);
    if (!s) return { rows: 0, classes: 0, teachers: 0, yearPrefix: '' };
    return {
        rows: s.rows.length,
        classes: s.classCount,
        teachers: s.teacherCount,
        yearPrefix: s.yearPrefix
    };
}
