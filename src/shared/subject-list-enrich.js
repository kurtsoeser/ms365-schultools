/**
 * Fächerkatalog ergänzen (additiv) – z. B. nach versehentlichem Überschreiben
 * durch Demo-Import oder schmale Schularbeiten-Stammdaten.
 */
import { normStr, normCode } from './utils/strings.js';
import {
    collectPlannedRowsFromKursteamState,
    readKursteamStateFromBrowserStorage
} from './kursteam-belegung-abgleich-logic.js';

/**
 * Fachgruppen-Zeile aus einer Entra-Gruppe ableiten (Mail-Nickname / Anzeigename).
 * @param {{ mailNickname?: string, displayName?: string }} g
 * @param {{ mailPrefix?: string }} [options] z. B. subjectGroupMailPrefix aus den Stammdaten
 * @returns {{ code: string, name: string }}
 */
export function deriveSubjectEntryFromGroup(g, options) {
    const opts = options && typeof options === 'object' ? options : {};
    const nick = normStr(g && g.mailNickname).toLowerCase();
    const dn = normStr(g && g.displayName);
    const custom = normStr(opts.mailPrefix)
        .toLowerCase()
        .replace(/[^a-z0-9-]/g, '');
    const customBare = custom.replace(/-+$/g, '');

    const prefixes = ['fachgruppe-', 'fachgruppe.', 'fach-', 'fach.', 'fach'];
    if (customBare) {
        prefixes.unshift(customBare + '-', customBare + '.', customBare);
    }

    let rest = nick;
    for (let i = 0; i < prefixes.length; i++) {
        const p = prefixes[i];
        if (rest.startsWith(p)) {
            rest = rest.slice(p.length);
            break;
        }
    }
    let code = rest ? normCode(rest.replace(/[^a-z0-9]/gi, '')) : '';

    if (!code && dn) {
        const m = dn.match(/^(?:Fachgruppe|Fachschaft|Fach)\s*[-–:]?\s*(.+)$/i);
        if (m) {
            const token = normStr(m[1]).split(/\s+/)[0] || '';
            code = normCode(token);
        }
    }
    if (!code && nick) code = normCode(nick.replace(/[^a-z0-9]/gi, '').slice(0, 20));
    const name = dn || code || normStr(g && g.mailNickname);
    return { code, name };
}

/**
 * @param {object[]} existing
 * @param {object[]} additions
 * @returns {{ subjects: object[], changed: boolean, stats: { added: number, namesFilled: number } }}
 */
export function mergeSubjectCatalogRows(existing, additions) {
    const byCode = new Map();
    (existing || []).forEach(function (s) {
        const code = normCode(s && s.code);
        if (!code) return;
        byCode.set(code, { code: code, name: normStr(s && s.name) });
    });
    let added = 0;
    let namesFilled = 0;
    (additions || []).forEach(function (s) {
        const code = normCode(s && s.code);
        if (!code) return;
        const name = normStr(s && s.name);
        if (!byCode.has(code)) {
            byCode.set(code, { code: code, name: name });
            added++;
            return;
        }
        const cur = byCode.get(code);
        if (!normStr(cur.name) && name) {
            cur.name = name;
            namesFilled++;
        }
    });
    const subjects = Array.from(byCode.values()).sort(function (a, b) {
        return a.code.localeCompare(b.code, 'de');
    });
    return {
        subjects: subjects,
        changed: added > 0 || namesFilled > 0,
        stats: { added: added, namesFilled: namesFilled, total: subjects.length }
    };
}

/**
 * @param {object[]} rows Schulstruktur-Zeilen
 */
export function subjectHintsFromStructureRows(rows) {
    const out = [];
    const seen = new Set();
    (rows || []).forEach(function (r) {
        if (!r || typeof r !== 'object') return;
        const typ = normStr(r.typ).toLowerCase();
        let code = normCode(r.ktFach);
        if (!code && (typ === 'fach' || typ === 'fachschaft' || typ.indexOf('fach') !== -1)) {
            code = normCode(r.bezeichnung);
        }
        if (!code || seen.has(code)) return;
        seen.add(code);
        const name =
            normStr(r.bezeichnung) && normCode(r.bezeichnung) !== code ? normStr(r.bezeichnung) : '';
        out.push({ code: code, name: name });
    });
    return out;
}

/**
 * Fachsegment aus typischem Teamnamen „SJ26 | 1A | D“ (letztes Pipe-Segment).
 * @param {string} teamName
 */
export function subjectCodeFromTeamNamePipe(teamName) {
    const tn = normStr(teamName);
    if (!tn || tn.indexOf('|') === -1) return '';
    const parts = tn.split('|').map(function (p) {
        return normStr(p);
    });
    const candidate = parts.length ? parts[parts.length - 1] : '';
    if (!candidate || /\s/.test(candidate) || candidate.length > 12) return '';
    return normCode(candidate);
}

/**
 * @param {string} fach
 */
function looksLikeSubjectShortCode(fach) {
    const s = normStr(fach);
    if (!s || /\s/.test(s)) return false;
    if (s.length > 6) return false;
    if (/[a-zäöüß]/.test(s)) return false;
    return true;
}

/**
 * @param {{ fach?: string, teamName?: string }} row Unterrichtsbelegung / Kursteam-Zeile
 * @returns {{ code: string, name: string }}
 */
export function subjectHintFromBelegungRow(row) {
    const r = row && typeof row === 'object' ? row : {};
    const fach = normStr(r.fach);
    const teamName = normStr(r.teamName);
    let code = '';
    let name = '';

    if (fach && looksLikeSubjectShortCode(fach)) {
        code = normCode(fach);
    } else if (fach) {
        name = fach;
        code = subjectCodeFromTeamNamePipe(teamName);
    }
    if (!code) code = subjectCodeFromTeamNamePipe(teamName);
    if (!code && fach) code = normCode(fach.replace(/\s+/g, ''));

    return { code: code, name: name };
}

/**
 * @param {object[]} rows Unterrichtsbelegung (App-Daten, Kursteam-Wizard, WebUntis)
 */
export function subjectHintsFromUnterrichtsbelegungRows(rows) {
    const out = [];
    const seen = new Set();
    (rows || []).forEach(function (r) {
        const hint = subjectHintFromBelegungRow(r);
        const code = hint.code;
        if (!code || seen.has(code)) return;
        seen.add(code);
        out.push({ code: code, name: normStr(hint.name) });
    });
    return out;
}

/**
 * Alle gespeicherten Belegungszeilen aus App-Daten v2 (alle Schuljahre).
 * @param {object|null|undefined} container ms365-schooltool-data-v2
 */
export function belegungRowsFromAppDataContainer(container) {
    const merged = [];
    const seen = new Set();
    const c = container && typeof container === 'object' ? container : null;
    const by = c && c.years && c.years.byLabel && typeof c.years.byLabel === 'object' ? c.years.byLabel : null;
    if (!by) return merged;

    Object.keys(by).forEach(function (label) {
        const snap = by[label] && by[label].unterrichtsbelegung;
        const rows = snap && Array.isArray(snap.rows) ? snap.rows : [];
        rows.forEach(function (r) {
            if (!r || typeof r !== 'object') return;
            const key =
                normStr(r.gruppenmail).toLowerCase() ||
                [normStr(r.klasse), normStr(r.fach), normStr(r.lehrerCode), normStr(r.teamName)].join('|');
            if (seen.has(key)) return;
            seen.add(key);
            merged.push(r);
        });
    });
    return merged;
}

/**
 * @param {{
 *   appDataContainer?: object|null,
 *   kursteamState?: object|null,
 *   storage?: Storage|null,
 *   extraRows?: object[]
 * }} [options]
 */
export function gatherUnterrichtsbelegungRowsForSubjectEnrich(options) {
    const o = options && typeof options === 'object' ? options : {};
    const merged = [];
    const seen = new Set();

    function pushRow(r) {
        if (!r || typeof r !== 'object') return;
        const key =
            normStr(r.gruppenmail).toLowerCase() ||
            [normStr(r.klasse), normStr(r.fach), normStr(r.lehrerCode), normStr(r.teamName)].join('|');
        if (seen.has(key)) return;
        seen.add(key);
        merged.push(r);
    }

    belegungRowsFromAppDataContainer(o.appDataContainer).forEach(pushRow);

    let ktState = o.kursteamState;
    if (ktState == null && o.storage !== false) {
        ktState = readKursteamStateFromBrowserStorage(o.storage);
    }
    if (ktState && typeof ktState === 'object') {
        collectPlannedRowsFromKursteamState(ktState).rows.forEach(pushRow);
    }

    (o.extraRows || []).forEach(pushRow);
    return merged;
}

export function subjectHintsFromArges(arges) {
    const out = [];
    const seen = new Set();
    (arges || []).forEach(function (a) {
        const list = Array.isArray(a && a.subjects) ? a.subjects : [];
        list.forEach(function (raw) {
            const code = normCode(raw);
            if (!code || seen.has(code)) return;
            seen.add(code);
            out.push({ code: code, name: '' });
        });
    });
    return out;
}

/**
 * @param {object[]} subjects
 * @param {{
 *   legacySubjects?: object[],
 *   structureRows?: object[],
 *   arges?: object[],
 *   unterrichtsbelegungRows?: object[],
 *   extraSubjects?: object[]
 * }} ctx
 */
export function enrichSubjectsFromCatalogSources(subjects, ctx) {
    const o = ctx && typeof ctx === 'object' ? ctx : {};
    const belegRows =
        Array.isArray(o.unterrichtsbelegungRows) && o.unterrichtsbelegungRows.length
            ? o.unterrichtsbelegungRows
            : gatherUnterrichtsbelegungRowsForSubjectEnrich({
                  appDataContainer: o.appDataContainer,
                  kursteamState: o.kursteamState,
                  storage: o.storage,
                  extraRows: o.extraBelegungRows
              });
    const hints = []
        .concat(o.legacySubjects || [])
        .concat(subjectHintsFromStructureRows(o.structureRows || []))
        .concat(subjectHintsFromArges(o.arges || []))
        .concat(subjectHintsFromUnterrichtsbelegungRows(belegRows))
        .concat(o.extraSubjects || []);
    const merged = mergeSubjectCatalogRows(subjects, hints);
    return merged;
}

/**
 * @param {object[]} current
 * @param {object[]} next
 * @returns {{ wouldShrink: boolean, removed: string[], currentCount: number, nextCount: number }}
 */
export function subjectCatalogShrinkReport(current, next) {
    const curCodes = [];
    const curSet = new Set();
    (current || []).forEach(function (s) {
        const c = normCode(s && s.code);
        if (!c || curSet.has(c)) return;
        curSet.add(c);
        curCodes.push(c);
    });
    const nextSet = new Set();
    (next || []).forEach(function (s) {
        const c = normCode(s && s.code);
        if (c) nextSet.add(c);
    });
    const removed = curCodes.filter(function (c) {
        return !nextSet.has(c);
    });
    return {
        wouldShrink: removed.length > 0,
        removed: removed,
        currentCount: curSet.size,
        nextCount: nextSet.size
    };
}

/**
 * @param {object|null|undefined} obj Import-JSON (Tenant v1, v2-Container oder Browser-Backup)
 * @returns {object[]}
 */
export function subjectsFromImportPayload(obj) {
    if (!obj || typeof obj !== 'object') return [];
    if (obj.core && Array.isArray(obj.core.subjects)) return obj.core.subjects.slice();
    if (Array.isArray(obj.subjects)) return obj.subjects.slice();
    const loc = obj.localStorage;
    if (loc && typeof loc === 'object') {
        let raw = loc['ms365-schooltool-data-v2'];
        if (raw != null && raw !== '') {
            try {
                const v2 = typeof raw === 'string' ? JSON.parse(raw) : raw;
                if (v2 && v2.core && Array.isArray(v2.core.subjects)) return v2.core.subjects.slice();
            } catch {
                /* ignore */
            }
        }
        raw = loc['ms365-tenant-settings-v1'];
        if (raw != null && raw !== '') {
            try {
                const v1 = typeof raw === 'string' ? JSON.parse(raw) : raw;
                if (v1 && Array.isArray(v1.subjects)) return v1.subjects.slice();
            } catch {
                /* ignore */
            }
        }
    }
    return [];
}

/**
 * @param {{ wouldShrink?: boolean, removed?: string[], currentCount?: number, nextCount?: number }} report
 * @param {string} [sourceLabel]
 */
export function formatSubjectCatalogShrinkConfirmDe(report, sourceLabel) {
    const r = report || {};
    if (!r.wouldShrink) return '';
    const removed = Array.isArray(r.removed) ? r.removed : [];
    const sample = removed.slice(0, 18).join(', ');
    const more =
        removed.length > 18 ? '\n… und ' + String(removed.length - 18) + ' weitere Kürzel' : '';
    const head = sourceLabel ? String(sourceLabel).trim() + '\n\n' : '';
    return (
        head +
        'Es würden ' +
        String(removed.length) +
        ' Fach-Kürzel aus der Stammdaten-Fächerliste entfernt (' +
        String(r.currentCount != null ? r.currentCount : '?') +
        ' → ' +
        String(r.nextCount != null ? r.nextCount : '?') +
        ').\n\n' +
        'Die Schularbeiten-FachMeta-Liste ist nur ein Teilkatalog – sie darf die volle Fächerliste nicht ersetzen.\n\n' +
        (sample ? 'Betroffene Kürzel (Auszug): ' + sample + more + '\n\n' : '') +
        'Trotzdem übernehmen?'
    );
}
