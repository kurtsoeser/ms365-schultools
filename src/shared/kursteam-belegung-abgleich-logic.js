/**
 * Abgleich: geplanter Unterricht (WebUntis / Kursteam-Wizard) ↔ Teams in M365.
 */
import { buildRowsFromTeamsData } from './unterrichtsbelegung-logic.js';
import { KURSTEAM_STORAGE_KEY } from './webuntis-sync-monitor-logic.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function normMail(v) {
    return normStr(v).toLowerCase();
}

function normClass(v) {
    return normStr(v).toUpperCase();
}

/**
 * @param {{ klasse?: string, lehrerCode?: string, fach?: string, gruppe?: string }} row
 */
export function teachingSlotKey(row) {
    const r = row && typeof row === 'object' ? row : {};
    return [
        normClass(r.klasse),
        normCode(r.lehrerCode || r.lehrer),
        normStr(r.fach).toUpperCase(),
        normStr(r.gruppe).toUpperCase()
    ].join('|');
}

/**
 * @param {object[]} filtered
 */
export function plannedRowsFromWebuntisLines(filtered) {
    const seen = new Set();
    const out = [];
    (Array.isArray(filtered) ? filtered : []).forEach((r) => {
        if (!r || typeof r !== 'object') return;
        const klasse = normStr(r.klasse);
        const fach = normStr(r.fach);
        const lehrerCode = normCode(r.lehrer);
        const gruppe = normStr(r.gruppe);
        if (!klasse && !fach && !lehrerCode) return;
        const key = teachingSlotKey({ klasse, lehrerCode, fach, gruppe });
        if (seen.has(key)) return;
        seen.add(key);
        out.push({
            klasse,
            fach,
            lehrerCode,
            gruppe,
            gruppenmail: '',
            teamName: '',
            lehrerEmail: ''
        });
    });
    out.sort((a, b) => teachingSlotKey(a).localeCompare(teachingSlotKey(b), 'de'));
    return out;
}

/**
 * @param {object|null|undefined} state Kursteam localStorage snapshot
 */
export function collectPlannedRowsFromKursteamState(state) {
    const st = state && typeof state === 'object' ? state : {};
    if (Array.isArray(st.teamsData) && st.teamsData.length && st.teamsGenerated) {
        const rows = buildRowsFromTeamsData(st.teamsData);
        if (rows.length) {
            return { source: 'teamsData', label: 'Kursteam-Wizard (generierte Teamliste)', rows };
        }
    }
    if (Array.isArray(st.filteredData) && st.filteredData.length) {
        return {
            source: 'filteredData',
            label: 'WebUntis / bereinigte Stundenplan-Zeilen',
            rows: plannedRowsFromWebuntisLines(st.filteredData)
        };
    }
    if (Array.isArray(st.rawData) && st.rawData.length) {
        return {
            source: 'rawData',
            label: 'WebUntis Rohdaten (ungefiltert)',
            rows: plannedRowsFromWebuntisLines(st.rawData)
        };
    }
    return { source: 'none', label: 'Kein Kursteam-Stand im Browser', rows: [] };
}

/**
 * @param {object|null} belegungSnapshot
 */
export function collectPlannedRowsFromBelegung(belegungSnapshot) {
    const snap = belegungSnapshot && typeof belegungSnapshot === 'object' ? belegungSnapshot : null;
    const rows = snap && Array.isArray(snap.rows) ? snap.rows.map((r) => Object.assign({}, r)) : [];
    return {
        source: rows.length ? 'belegung' : 'none',
        label: rows.length ? 'Unterrichtsbelegung (App-Daten)' : 'Keine Unterrichtsbelegung',
        rows
    };
}

/**
 * Plan-Quellen zusammenführen (Wizard/WebUntis hat Vorrang vor reiner Belegung ohne Wizard).
 * @param {{ kursteamState?: object|null, belegungSnapshot?: object|null, preferBelegung?: boolean }} input
 */
export function resolvePlannedRowsForAbgleich(input) {
    const inp = input && typeof input === 'object' ? input : {};
    const fromState = collectPlannedRowsFromKursteamState(inp.kursteamState);
    const fromBeleg = collectPlannedRowsFromBelegung(inp.belegungSnapshot);

    if (inp.preferBelegung && fromBeleg.rows.length) {
        return fromBeleg;
    }
    if (fromState.rows.length) return fromState;
    if (fromBeleg.rows.length) return fromBeleg;
    return fromState;
}

/**
 * @param {object[]} m365Rows Zeilen mit graphGroupId (Graph-Import oder gespeicherte Belegung)
 */
export function normalizeM365RowsForAbgleich(m365Rows) {
    const out = [];
    const seen = new Set();
    (Array.isArray(m365Rows) ? m365Rows : []).forEach((r) => {
        if (!r || typeof r !== 'object') return;
        const graphGroupId = normStr(r.graphGroupId);
        const gruppenmail = normMail(r.gruppenmail);
        if (!graphGroupId && !gruppenmail) return;
        const key = gruppenmail || graphGroupId;
        if (seen.has(key)) return;
        seen.add(key);
        out.push({
            klasse: normStr(r.klasse),
            lehrerCode: normCode(r.lehrerCode),
            lehrerEmail: normMail(r.lehrerEmail),
            fach: normStr(r.fach),
            gruppe: normStr(r.gruppe),
            teamName: normStr(r.teamName),
            gruppenmail,
            graphGroupId
        });
    });
    return out;
}

/**
 * @param {object[]} plannedRows
 * @param {object[]} m365Rows
 */
export function buildKursteamAbgleichReport(plannedRows, m365Rows) {
    const plan = (Array.isArray(plannedRows) ? plannedRows : []).map((r) => Object.assign({}, r));
    const m365 = normalizeM365RowsForAbgleich(m365Rows);

    const m365ByMail = new Map();
    const m365BySlot = new Map();
    m365.forEach((row, idx) => {
        if (row.gruppenmail) m365ByMail.set(row.gruppenmail, { row, idx });
        const sk = teachingSlotKey(row);
        if (sk && sk !== '|||') m365BySlot.set(sk, { row, idx });
    });

    const matched = [];
    const missingInM365 = [];
    const mailConflict = [];
    const usedM365 = new Set();

    plan.forEach((p) => {
        const mail = normMail(p.gruppenmail);
        const slot = teachingSlotKey(p);
        let hit = null;
        let matchBy = '';

        if (mail && m365ByMail.has(mail)) {
            hit = m365ByMail.get(mail);
            matchBy = 'gruppenmail';
        } else if (slot && slot !== '|||' && m365BySlot.has(slot)) {
            hit = m365BySlot.get(slot);
            matchBy = 'unterricht';
        }

        if (!hit) {
            missingInM365.push({
                klasse: p.klasse,
                fach: p.fach,
                lehrerCode: p.lehrerCode,
                gruppe: p.gruppe,
                gruppenmail: p.gruppenmail || '',
                teamName: p.teamName || ''
            });
            return;
        }

        usedM365.add(hit.idx);
        const m = hit.row;
        if (mail && m.gruppenmail && mail !== m.gruppenmail) {
            mailConflict.push({
                klasse: p.klasse || m.klasse,
                fach: p.fach || m.fach,
                lehrerCode: p.lehrerCode || m.lehrerCode,
                planMail: mail,
                m365Mail: m.gruppenmail,
                teamName: m.teamName || p.teamName || ''
            });
        }
        matched.push({
            matchBy,
            klasse: p.klasse || m.klasse,
            fach: p.fach || m.fach,
            lehrerCode: p.lehrerCode || m.lehrerCode,
            gruppenmail: m.gruppenmail || mail,
            teamName: m.teamName || p.teamName || '',
            graphGroupId: m.graphGroupId || ''
        });
    });

    const onlyInM365 = m365
        .filter((_, idx) => !usedM365.has(idx))
        .map((m) => ({
            klasse: m.klasse,
            fach: m.fach,
            lehrerCode: m.lehrerCode,
            gruppenmail: m.gruppenmail,
            teamName: m.teamName,
            graphGroupId: m.graphGroupId
        }));

    return {
        counts: {
            planned: plan.length,
            m365: m365.length,
            matched: matched.length,
            missingInM365: missingInM365.length,
            onlyInM365: onlyInM365.length,
            mailConflict: mailConflict.length
        },
        matched,
        missingInM365,
        onlyInM365,
        mailConflict
    };
}

/**
 * @param {Storage|null} [storage]
 */
export function readKursteamStateFromBrowserStorage(storage) {
    const store = storage || (typeof localStorage !== 'undefined' ? localStorage : null);
    if (!store || typeof store.getItem !== 'function') return null;
    try {
        const raw = store.getItem(KURSTEAM_STORAGE_KEY);
        if (!raw) return null;
        return JSON.parse(raw);
    } catch {
        return null;
    }
}
