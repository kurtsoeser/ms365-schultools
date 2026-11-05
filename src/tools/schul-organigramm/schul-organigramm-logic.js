/**
 * Schul-Organigramm – Datenmodell & Seed aus Stammdaten.
 *
 * Spec: docs/schul-organigramm.md (Phase 0)
 * Kanonisch: ms365-schooltool-data-v2.core.schulOrganigramm
 * Knoten referenzieren bestehende Listen (refKind + refKey), keine parallele Personenliste.
 *
 * @typedef {'area'|'section'|'role'|'arge'|'class'|'kursteam'|'sga'|'council'|'custom'|'person'} SchulOrgNodeType
 * @typedef {'adminRole'|'arge'|'argeLeiter'|'klassenvorstand'|'classFunction'|'argeFunction'|'class'|'kursteam'|'sga'|'studentCouncil'|'audience'|'adminPerson'|'sammelgruppe'|'jahrgang'|'elternJahrgang'|'elternClass'|'klassenchatJahrgang'|'klassenchatClass'|'none'} SchulOrgRefKind
 */

import { normalizeVerwaltungAudienceGroups } from '../../shared/administration-audience-groups.js';
import { classCodeFromDisplayName } from '../../shared/class-list-enrich.js';
import { resolveClassGraphGroupId } from '../../shared/membership-hygiene.js';
import {
    buildKnownClassCodeIndex,
    expandKlasseToKnownCodes
} from '../klassenchats/klassenchats-logic.js';

/** Schema-Version 2: Areas + Canvas-Koordinaten + view (siehe docs/schul-organigramm-landkarte.md). */
export const SCHUL_ORG_VERSION = 2;

/** Aggregat-Knoten für Schülervertretung / SGA (keine Personen-Kacheln). */
export const SCHUL_ORG_BOARD_REF = '_board';

/** Standard-Landkarten-Bereiche für Lehrer-/Schüler-/Eltern-Sammelgruppe. */
export const SCHUL_ORG_SEC_LEHRER = 'sec_lehrer';
export const SCHUL_ORG_SEC_SCHUELER = 'sec_schueler';
export const SCHUL_ORG_SEC_ELTERN = 'sec_eltern';
/** KlassenChats: Teams-Gruppenchats der Lehrkräfte einer Klasse (Jahrgang → Klasse). */
export const SCHUL_ORG_SEC_KLASSENCHATS = 'sec_klassenchats';
/** Unter ARGEs: Sammelbereich aller ARGE-Leitungen (Stammdaten + M365-Besitzer). */
export const SCHUL_ORG_SEC_ARGE = 'sec_arge';
export const SCHUL_ORG_SEC_ARGE_LEITER = 'sec_arge_leiter';
/** Root: Sammelbereich aller Klassenvorstände (aus Klassen-Stammdaten / Gruppen-Besitzer). */
export const SCHUL_ORG_SEC_KLASSENVORSTAND = 'sec_klassenvorstand';

/** Prefix Root-Sammelbereiche je Klassenfunktion (z. B. Klassensprecher). */
export const SCHUL_ORG_CLASS_FN_SEC_PREFIX = 'sec_classfn_';

/** Prefix Sammelbereiche je ARGE-Funktion (unter ARGEs). */
export const SCHUL_ORG_ARGE_FN_SEC_PREFIX = 'sec_argefn_';

/**
 * @typedef {'class'|'arge'} SchulOrgFunctionSource
 * @typedef {'head'|'functionsMap'} SchulOrgFunctionSeat
 * @typedef {{
 *   id: string,
 *   label: string,
 *   labelPlural: string,
 *   source: SchulOrgFunctionSource,
 *   seat: SchulOrgFunctionSeat,
 *   seatKey?: string,
 *   personKind: 'teacher'|'student',
 *   boardId: string,
 *   boardParent: 'root'|'sec_arge',
 *   refKind: string,
 *   tone: string,
 *   seatLabel: string,
 *   aggregate: boolean
 * }} SchulOrgFunctionAggregateDef
 */

/**
 * Builtin-Klassenfunktionen (Sitze in der Klasse + Sammelbereich).
 * @type {{ id: string, label: string, labelPlural: string, personKind: 'student' }[]}
 */
export const SCHUL_ORG_DEFAULT_CLASS_FUNCTIONS = [
    {
        id: 'klassensprecher',
        label: 'Klassensprecher',
        labelPlural: 'Klassensprecher',
        personKind: 'student'
    }
];

/**
 * Optionale Builtin-ARGE-Funktionen (leer; Erweiterung über settings.argeFunctions).
 * @type {{ id: string, label: string, labelPlural: string, personKind: 'teacher' }[]}
 */
export const SCHUL_ORG_DEFAULT_ARGE_FUNCTIONS = [];

/** Fallback-Jahrgang für Klassen ohne führende Ziffer im Kürzel. */
export const SCHUL_ORG_JAHRGANG_SONST = 'sonst';

/** SGA-Container und die drei Pflicht-Gruppen (Schülervertretung / Lehrer / Eltern). */
export const SCHUL_ORG_SGA_BOARD_ID = 'sga_board';
export const SCHUL_ORG_SGA_SV_AREA_ID = 'sga_gruppe_schueler';
/** @type {{ id: string, label: string, scope: 'student'|'teacher'|'external', tone: string, x: number, y: number }[]} */
export const SCHUL_ORG_SGA_SCOPE_AREAS = [
    {
        id: SCHUL_ORG_SGA_SV_AREA_ID,
        label: 'Schülervertretung',
        scope: 'student',
        tone: 'schueler',
        x: 12,
        y: 8
    },
    { id: 'sga_gruppe_lehrer', label: 'Lehrer', scope: 'teacher', tone: 'lehrer', x: 12, y: 180 },
    { id: 'sga_gruppe_eltern', label: 'Eltern', scope: 'external', tone: 'council', x: 12, y: 352 }
];

/**
 * Die drei SGA-Schülerplätze = Schülervertretung (Schulsprecher + STVs).
 * @type {{ id: string, refKey: string, label: string, seat: number }[]}
 */
export const SCHUL_ORG_SV_SEATS = [
    { id: 'sv_schulsprecher', refKey: 'schulsprecher', label: 'Schulsprecher', seat: 0 },
    {
        id: 'sv_schulsprecher_stv1',
        refKey: 'schulsprecher_stv1',
        label: 'Schulsprecher-STV1',
        seat: 1
    },
    {
        id: 'sv_schulsprecher_stv2',
        refKey: 'schulsprecher_stv2',
        label: 'Schulsprecher-STV2',
        seat: 2
    }
];

/** @type {Record<string, SchulOrgNodeType>} */
export const NODE = {
    /** Verschachtelbarer Bereich (Landkarte). Alias historisch: section */
    AREA: 'area',
    SECTION: 'area',
    ROLE: 'role',
    ARGE: 'arge',
    CLASS: 'class',
    /** Unterrichtsteam / Kursteam (unter Klasse; Mehrklassen → mehrere Knoten) */
    KURSTEAM: 'kursteam',
    SGA: 'sga',
    COUNCIL: 'council',
    PERSON: 'person',
    CUSTOM: 'custom'
};

/**
 * @param {object|null|undefined} n
 */
export function isAreaNode(n) {
    const t = normStr(n && n.type).toLowerCase();
    return t === NODE.AREA || t === 'section';
}

/**
 * @param {string} type
 */
function canonicalNodeType(type) {
    const t = normStr(type).toLowerCase() || NODE.CUSTOM;
    if (t === 'section') return NODE.AREA;
    return t;
}

/**
 * Default-Canvas-Positionen für Top-Level-Areas (flach).
 * Weite Abstände: hohe ARGE-/Rollenlisten dürfen sich nicht überdecken.
 * @param {number} index
 */
export function defaultCanvasSlot(index) {
    const col = index % 3;
    const row = Math.floor(index / 3);
    return { x: 48 + col * 920, y: 48 + row * 720 };
}

/**
 * Canvas-Koordinate parsen – null/''/false nicht als 0 werten (sonst alles bei 0,0).
 * @param {unknown} v
 * @returns {number|null}
 */
export function parseSchulOrgCanvasCoord(v) {
    if (typeof v === 'number' && Number.isFinite(v)) return Math.round(v);
    if (typeof v === 'string') {
        const t = v.trim();
        if (!t) return null;
        const n = Number(t);
        if (Number.isFinite(n)) return Math.round(n);
    }
    return null;
}

function normStr(v) {
    return String(v ?? '').trim();
}

/**
 * Stabiler Schlüssel für ein Kursteam (gruppenübergreifend, ohne Klassen-Platzierung).
 * @param {{ gruppenmail?: string, graphGroupId?: string, teamName?: string, fach?: string, lehrerEmail?: string, lehrerCode?: string, klasse?: string }} row
 */
export function kursteamStableKey(row) {
    const r = row && typeof row === 'object' ? row : {};
    const mail = normStr(r.gruppenmail).toLowerCase().replace(/[^a-z0-9._-]/g, '');
    if (mail) return mail.slice(0, 72);
    const gid = normStr(r.graphGroupId).replace(/[^a-zA-Z0-9_-]/g, '');
    if (gid) return ('g_' + gid).slice(0, 72);
    const bits = [r.teamName, r.fach, r.lehrerEmail || r.lehrerCode, r.klasse]
        .map(function (x) {
            return normStr(x)
                .toLowerCase()
                .replace(/[^a-z0-9]+/g, '-');
        })
        .filter(Boolean);
    return (bits.join('_') || 'kurs').slice(0, 72);
}

/**
 * Anzeige-Label für Kursteam-Kachel.
 * @param {object} row
 */
export function kursteamLabelFromRow(row) {
    const r = row && typeof row === 'object' ? row : {};
    const name = normStr(r.teamName);
    if (name) return name;
    const fach = normStr(r.fach);
    const lehrer = normStr(r.lehrerName) || normCode(r.lehrerCode) || normStr(r.lehrerEmail);
    if (fach && lehrer) return fach + ' · ' + lehrer;
    return fach || lehrer || 'Kursteam';
}

/**
 * Unterrichtsbelegung → Kursteam-Platzierungen je Stammdaten-Klasse.
 * Mehrklassen-Zellen (1AK1BK / 1A~1B) erzeugen je Klasse einen Eintrag.
 *
 * @param {object|null|undefined} belegung Snapshot `{ rows: [...] }`
 * @param {Array<{ code?: string, name?: string }>} classes
 * @returns {{ stableKey: string, classCode: string, row: object, sharedClassCodes: string[] }[]}
 */
export function kursteamPlacementsFromBelegung(belegung, classes) {
    const rows =
        belegung && typeof belegung === 'object' && Array.isArray(belegung.rows)
            ? belegung.rows
            : Array.isArray(belegung)
              ? belegung
              : [];
    // Index mit kanonischen Codes (auch aus Anzeigenamen), damit „Klasse 2HMB“ matcht
    const enrichedClasses = [];
    (Array.isArray(classes) ? classes : []).forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        enrichedClasses.push({ code: code, name: normStr(c && c.name) || code });
    });
    const classIndex = buildKnownClassCodeIndex(enrichedClasses);
    /** @type {{ stableKey: string, classCode: string, row: object, sharedClassCodes: string[] }[]} */
    const out = [];
    const seen = new Set();
    rows.forEach(function (row) {
        if (!row || typeof row !== 'object') return;
        const klasseRaw = normStr(row.klasse);
        if (!klasseRaw) return;
        const klasseCompact = klasseRaw.replace(/\./g, '').replace(/\s+/g, '');
        const canonicalHint =
            canonicalSchulOrgClassCode(klasseRaw) || canonicalSchulOrgClassCode(klasseCompact);
        let codes = expandKlasseToKnownCodes(klasseRaw, classIndex);
        if (!codes.length && klasseCompact !== klasseRaw) {
            codes = expandKlasseToKnownCodes(klasseCompact, classIndex);
        }
        if (!codes.length && canonicalHint) {
            codes = expandKlasseToKnownCodes(canonicalHint, classIndex);
        }
        // Fallback: kanonischer Code, wenn Index ihn kennt
        if (!codes.length && canonicalHint && classIndex.canonByUpper.has(canonicalHint)) {
            codes = [classIndex.canonByUpper.get(canonicalHint)];
        }
        const targets = codes.length
            ? codes
            : canonicalHint
              ? [canonicalHint]
              : klasseRaw
                ? [klasseRaw]
                : [];
        if (!targets.length) return;
        const stable = kursteamStableKey(row);
        if (!stable) return;
        const shared = targets
            .map(function (c) {
                return canonicalSchulOrgClassCode(c) || normCode(c) || normStr(c);
            })
            .filter(Boolean);
        targets.forEach(function (code) {
            const cc = canonicalSchulOrgClassCode(code) || normCode(code) || normStr(code);
            if (!cc) return;
            const placeKey = stable + '@' + cc;
            if (seen.has(placeKey)) return;
            seen.add(placeKey);
            out.push({
                stableKey: stable,
                classCode: cc,
                row: row,
                sharedClassCodes: shared
            });
        });
    });
    out.sort(function (a, b) {
        const c = a.classCode.localeCompare(b.classCode, 'de');
        if (c !== 0) return c;
        return kursteamLabelFromRow(a.row).localeCompare(kursteamLabelFromRow(b.row), 'de');
    });
    return out;
}

/**
 * @param {{ stableKey: string, classCode: string, row: object, sharedClassCodes: string[] }} place
 * @param {number} order
 */
function kursteamNodeFromPlacement(place, order) {
    const row = place.row || {};
    const classCode = place.classCode;
    const stable = place.stableKey;
    const refKey = stable + '@' + classCode;
    const gid = normStr(row.graphGroupId);
    const mail = normStr(row.gruppenmail).toLowerCase();
    /** @type {object} */
    const node = {
        id: 'kursteam_' + refKey.replace(/[^a-zA-Z0-9@_.-]/g, '_').slice(0, 96),
        type: NODE.KURSTEAM,
        label: kursteamLabelFromRow(row),
        parentId: 'class_' + classCode,
        order: order,
        tone: 'kursteam',
        refKind: 'kursteam',
        refKey: refKey,
        meta: {
            stableKey: stable,
            fach: normStr(row.fach),
            gruppe: normStr(row.gruppe),
            lehrerCode: normCode(row.lehrerCode),
            lehrerEmail: normStr(row.lehrerEmail).toLowerCase(),
            lehrerName: normStr(row.lehrerName),
            teamName: normStr(row.teamName),
            gruppenmail: mail,
            graphGroupId: gid,
            klasseRaw: normStr(row.klasse),
            placementClass: classCode,
            sharedClassCodes: Array.isArray(place.sharedClassCodes)
                ? place.sharedClassCodes.slice()
                : [classCode]
        }
    };
    if (gid || mail) {
        node.m365 = {
            groupMode: gid ? 'match' : 'none',
            groupKind: 'unified',
            groupId: gid || null,
            groupLabel: node.label,
            groupMail: mail || null,
            mailboxMode: 'none',
            mailboxUpn: null,
            mailboxLabel: null,
            mailboxId: null
        };
    }
    return node;
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function newId(prefix) {
    return (
        String(prefix || 'n') +
        '_' +
        Math.random().toString(36).slice(2, 8) +
        Date.now().toString(36).slice(-4)
    );
}

/**
 * Section-ID für eine Verwaltungs-Zielgruppe (builtin oder va-*).
 * @param {string} audienceGroupId
 */
export function audienceSectionId(audienceGroupId) {
    const id = normStr(audienceGroupId);
    if (!id) return '';
    if (id === 'schulleitung') return 'sec_schulleitung';
    if (id === 'verwaltung') return 'sec_verwaltung';
    if (id.startsWith('va-')) return 'sec_aud_' + id;
    return '';
}

/**
 * @param {string} sectionId
 * @returns {string} audience group id oder ''
 */
export function audienceIdFromSectionId(sectionId) {
    const sid = normStr(sectionId);
    if (sid === 'sec_schulleitung') return 'schulleitung';
    if (sid === 'sec_verwaltung') return 'verwaltung';
    if (sid.indexOf('sec_aud_') === 0) return sid.slice('sec_aud_'.length);
    return '';
}

/**
 * Ob eine Section Rollen aufnehmen darf (Builtin + Custom-Verwaltungsgruppen).
 * @param {string} sectionId
 * @param {object} [sectionNode]
 */
export function isRoleCapableSection(sectionId, sectionNode) {
    const sid = normStr(sectionId);
    if (sid === 'sec_schulleitung' || sid === 'sec_verwaltung') return true;
    if (sid.indexOf('sec_aud_') === 0) return true;
    if (sectionNode && normStr(sectionNode.refKind) === 'audience' && normStr(sectionNode.refKey)) {
        return true;
    }
    return false;
}

/**
 * Audience-Gruppen-ID einer Area (builtin / va-*), sonst ''.
 * @param {object|null|undefined} areaNode
 */
export function audienceIdForArea(areaNode) {
    if (!areaNode) return '';
    if (normStr(areaNode.refKind) === 'audience' && normStr(areaNode.refKey)) {
        return normStr(areaNode.refKey);
    }
    return audienceIdFromSectionId(areaNode.id);
}

/**
 * Standard-Bereich Lehrkräfte-, Schüler-, Eltern- oder KlassenChats-Sammelgruppe.
 * @param {object|null|undefined} areaNode
 * @returns {'lehrer'|'schueler'|'eltern'|'klassenchats'|''}
 */
/**
 * Bereich „ARGE-Leiter“ (Sammel der ARGE-Leitungen / M365-Owner).
 * @param {object|null|undefined} areaNode
 */
export function isArgeLeiterArea(areaNode) {
    if (!areaNode) return false;
    if (normStr(areaNode.id) === SCHUL_ORG_SEC_ARGE_LEITER) return true;
    if (normStr(areaNode.refKind).toLowerCase() === 'argeleiter') return true;
    return false;
}

/**
 * Root-Bereich „Klassenvorstände“ (Sammel aller Klassen-KVs).
 * @param {object|null|undefined} areaNode
 */
export function isKlassenvorstandArea(areaNode) {
    if (!areaNode) return false;
    if (normStr(areaNode.id) === SCHUL_ORG_SEC_KLASSENVORSTAND) return true;
    if (normStr(areaNode.refKind).toLowerCase() === 'klassenvorstand') return true;
    return false;
}

/**
 * @param {unknown} raw
 * @returns {string}
 */
export function cleanSchulOrgClassFunctionId(raw) {
    return normStr(raw)
        .toLowerCase()
        .replace(/[^a-z0-9_-]+/g, '')
        .slice(0, 40);
}

/**
 * @param {string} functionId
 * @returns {string}
 */
export function classFunctionBoardId(functionId) {
    const id = cleanSchulOrgClassFunctionId(functionId) || 'rolle';
    return SCHUL_ORG_CLASS_FN_SEC_PREFIX + id;
}

/**
 * @param {unknown} raw
 * @returns {Record<string, { name: string, email: string }>}
 */
export function normalizeClassFunctionsMap(raw) {
    /** @type {Record<string, { name: string, email: string }>} */
    const out = {};
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return out;
    Object.keys(/** @type {object} */ (raw)).forEach(function (key) {
        const id = cleanSchulOrgClassFunctionId(key);
        if (!id) return;
        const seat = /** @type {object} */ (raw)[key];
        if (!seat || typeof seat !== 'object') return;
        const name = normStr(seat.name);
        const email = normStr(seat.email).toLowerCase();
        if (!name && !email) return;
        out[id] = { name: name, email: email };
    });
    return out;
}

/**
 * Aktive Klassenfunktionen (Builtin + optionale settings.classFunctions).
 * @param {object} [settings]
 * @returns {{ id: string, label: string, labelPlural: string, personKind: 'student' }[]}
 */
export function listSchulOrgClassFunctions(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    /** @type {Map<string, { id: string, label: string, labelPlural: string, personKind: 'student' }>} */
    const byId = new Map();
    SCHUL_ORG_DEFAULT_CLASS_FUNCTIONS.forEach(function (def) {
        byId.set(def.id, Object.assign({}, def));
    });
    (Array.isArray(s.classFunctions) ? s.classFunctions : []).forEach(function (row) {
        const id = cleanSchulOrgClassFunctionId(row && row.id);
        if (!id) return;
        const label = normStr(row && row.label) || id;
        const labelPlural = normStr(row && row.labelPlural) || label;
        byId.set(id, {
            id: id,
            label: label,
            labelPlural: labelPlural,
            personKind: 'student'
        });
    });
    return Array.from(byId.values());
}

/**
 * Aktive ARGE-Funktionen (Builtin + settings.argeFunctions).
 * @param {object} [settings]
 * @returns {{ id: string, label: string, labelPlural: string, personKind: 'teacher' }[]}
 */
export function listSchulOrgArgeFunctions(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    /** @type {Map<string, { id: string, label: string, labelPlural: string, personKind: 'teacher' }>} */
    const byId = new Map();
    SCHUL_ORG_DEFAULT_ARGE_FUNCTIONS.forEach(function (def) {
        byId.set(def.id, Object.assign({}, def));
    });
    (Array.isArray(s.argeFunctions) ? s.argeFunctions : []).forEach(function (row) {
        const id = cleanSchulOrgClassFunctionId(row && row.id);
        if (!id) return;
        const label = normStr(row && row.label) || id;
        const labelPlural = normStr(row && row.labelPlural) || label;
        byId.set(id, {
            id: id,
            label: label,
            labelPlural: labelPlural,
            personKind: 'teacher'
        });
    });
    return Array.from(byId.values());
}

/**
 * @param {string} functionId
 * @returns {string}
 */
export function argeFunctionBoardId(functionId) {
    const id = cleanSchulOrgClassFunctionId(functionId) || 'rolle';
    return SCHUL_ORG_ARGE_FN_SEC_PREFIX + id;
}

/**
 * Alle Funktions-Sammelbereiche: Leitungssitze + Klassen-/ARGE-Funktionen.
 * Regel: gleiche Funktion über Gruppen → ein aggregierender Landkarten-Bereich.
 * @param {object} [settings]
 * @returns {SchulOrgFunctionAggregateDef[]}
 */
export function listSchulOrgFunctionAggregates(settings) {
    /** @type {SchulOrgFunctionAggregateDef[]} */
    const out = [
        {
            id: 'klassenvorstand',
            label: 'Klassenvorstand',
            labelPlural: 'Klassenvorstände',
            source: 'class',
            seat: 'head',
            personKind: 'teacher',
            boardId: SCHUL_ORG_SEC_KLASSENVORSTAND,
            boardParent: 'root',
            refKind: 'klassenvorstand',
            tone: 'lehrer',
            seatLabel: 'Klassenvorstand',
            aggregate: true
        },
        {
            id: 'argeLeiter',
            label: 'ARGE-Leitung',
            labelPlural: 'ARGE-Leiter',
            source: 'arge',
            seat: 'head',
            personKind: 'teacher',
            boardId: SCHUL_ORG_SEC_ARGE_LEITER,
            boardParent: 'sec_arge',
            refKind: 'argeLeiter',
            tone: 'arge',
            seatLabel: 'ARGE-Leitung',
            aggregate: true
        }
    ];
    listSchulOrgClassFunctions(settings).forEach(function (def) {
        out.push({
            id: def.id,
            label: def.label,
            labelPlural: def.labelPlural,
            source: 'class',
            seat: 'functionsMap',
            seatKey: def.id,
            personKind: 'student',
            boardId: classFunctionBoardId(def.id),
            boardParent: 'root',
            refKind: 'classFunction',
            tone: 'schueler',
            seatLabel: def.label,
            aggregate: true
        });
    });
    listSchulOrgArgeFunctions(settings).forEach(function (def) {
        out.push({
            id: 'argefn_' + def.id,
            label: def.label,
            labelPlural: def.labelPlural,
            source: 'arge',
            seat: 'functionsMap',
            seatKey: def.id,
            personKind: 'teacher',
            boardId: argeFunctionBoardId(def.id),
            boardParent: 'sec_arge',
            refKind: 'argeFunction',
            tone: 'arge',
            seatLabel: def.label,
            aggregate: true
        });
    });
    return out;
}

/**
 * @param {object|null|undefined} areaNode
 * @returns {string}
 */
export function classFunctionIdForArea(areaNode) {
    if (!areaNode) return '';
    if (normStr(areaNode.refKind).toLowerCase() === 'classfunction') {
        return cleanSchulOrgClassFunctionId(areaNode.refKey);
    }
    const id = normStr(areaNode.id);
    if (id.indexOf(SCHUL_ORG_CLASS_FN_SEC_PREFIX) === 0) {
        return cleanSchulOrgClassFunctionId(id.slice(SCHUL_ORG_CLASS_FN_SEC_PREFIX.length));
    }
    return '';
}

/**
 * @param {object|null|undefined} areaNode
 * @returns {string}
 */
export function argeFunctionIdForArea(areaNode) {
    if (!areaNode) return '';
    if (normStr(areaNode.refKind).toLowerCase() === 'argefunction') {
        return cleanSchulOrgClassFunctionId(areaNode.refKey);
    }
    const id = normStr(areaNode.id);
    if (id.indexOf(SCHUL_ORG_ARGE_FN_SEC_PREFIX) === 0) {
        return cleanSchulOrgClassFunctionId(id.slice(SCHUL_ORG_ARGE_FN_SEC_PREFIX.length));
    }
    return '';
}

/**
 * Root-Sammelbereich einer Klassenfunktion (z. B. alle Klassensprecher).
 * @param {object|null|undefined} areaNode
 */
export function isClassFunctionArea(areaNode) {
    return !!classFunctionIdForArea(areaNode);
}

/**
 * @param {object|null|undefined} areaNode
 */
export function isArgeFunctionArea(areaNode) {
    return !!argeFunctionIdForArea(areaNode);
}

/**
 * Sammelbereich einer schulweiten Funktion (KV, ARGE-Leiter, Klassen-/ARGE-Funktion).
 * @param {object|null|undefined} areaNode
 */
export function isFunctionAggregateArea(areaNode) {
    return (
        isKlassenvorstandArea(areaNode) ||
        isArgeLeiterArea(areaNode) ||
        isClassFunctionArea(areaNode) ||
        isArgeFunctionArea(areaNode)
    );
}

/**
 * @param {object|null|undefined} areaNode
 * @param {object} [settings]
 * @returns {SchulOrgFunctionAggregateDef|null}
 */
export function functionAggregateDefForArea(areaNode, settings) {
    if (!areaNode) return null;
    const defs = listSchulOrgFunctionAggregates(settings);
    const id = normStr(areaNode.id);
    const refKind = normStr(areaNode.refKind).toLowerCase();
    const refKey = normStr(areaNode.refKey);
    return (
        defs.find(function (d) {
            if (d.boardId === id) return true;
            if (normStr(d.refKind).toLowerCase() === refKind) {
                if (d.seat === 'functionsMap') {
                    return cleanSchulOrgClassFunctionId(d.seatKey) === cleanSchulOrgClassFunctionId(refKey);
                }
                return cleanSchulOrgClassFunctionId(d.id) === cleanSchulOrgClassFunctionId(refKey) || !refKey;
            }
            return false;
        }) || null
    );
}

/**
 * Besetzungen einer Klassenfunktion schulweit (eindeutig nach E-Mail).
 * @param {object} settings
 * @param {string} functionId
 * @returns {Array<{ name: string, email: string, label: string, classCode: string, className: string, functionId: string, functionLabel: string }>}
 */
export function collectClassFunctionPeopleFromSettings(settings, functionId) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const fnId = cleanSchulOrgClassFunctionId(functionId);
    if (!fnId) return [];
    const def = listSchulOrgClassFunctions(s).find(function (d) {
        return d.id === fnId;
    });
    const functionLabel = (def && def.label) || fnId;
    /** @type {Map<string, { name: string, email: string, label: string, classCode: string, className: string, functionId: string, functionLabel: string, classLabels: string[] }>} */
    const byMail = new Map();
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        const seats = normalizeClassFunctionsMap(c && c.functions);
        const seat = seats[fnId];
        if (!seat) return;
        const email = normStr(seat.email).toLowerCase();
        const name = normStr(seat.name);
        if (!email || email.indexOf('@') === -1) return;
        const classCode =
            canonicalSchulOrgClassCode(c && c.code) ||
            canonicalSchulOrgClassCode(c && c.name) ||
            normCode(c && c.code);
        const className = normStr(c && c.name) || classCode;
        const existing = byMail.get(email);
        if (existing) {
            if (className && existing.classLabels.indexOf(className) < 0) {
                existing.classLabels.push(className);
            }
            return;
        }
        byMail.set(email, {
            name: name,
            email: email,
            label: name || email,
            classCode: classCode,
            className: className,
            functionId: fnId,
            functionLabel: functionLabel,
            classLabels: className ? [className] : []
        });
    });
    return Array.from(byMail.values())
        .map(function (p) {
            const hint =
                p.classLabels.length > 1
                    ? p.classLabels.join(', ')
                    : p.classLabels[0] || p.className || '';
            return {
                name: p.name,
                email: p.email,
                label: p.label,
                classCode: p.classCode,
                className: hint || p.className,
                functionId: p.functionId,
                functionLabel: p.functionLabel
            };
        })
        .sort(function (a, b) {
            return String(a.name || a.email).localeCompare(String(b.name || b.email), 'de');
        });
}

/**
 * Besetzungen einer ARGE-Funktion schulweit.
 * @param {object} settings
 * @param {string} functionId
 */
export function collectArgeFunctionPeopleFromSettings(settings, functionId) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const fnId = cleanSchulOrgClassFunctionId(functionId);
    if (!fnId) return [];
    const def = listSchulOrgArgeFunctions(s).find(function (d) {
        return d.id === fnId;
    });
    const functionLabel = (def && def.label) || fnId;
    /** @type {Map<string, { name: string, email: string, label: string, argeCode: string, argeName: string, functionId: string, functionLabel: string, groupLabels: string[] }>} */
    const byMail = new Map();
    (Array.isArray(s.arges) ? s.arges : []).forEach(function (a) {
        const seats = normalizeClassFunctionsMap(a && a.functions);
        const seat = seats[fnId];
        if (!seat) return;
        const email = normStr(seat.email).toLowerCase();
        const name = normStr(seat.name);
        if (!email || email.indexOf('@') === -1) return;
        const argeCode = normCode(a && a.code);
        const argeName = normStr(a && a.name) || argeCode;
        const existing = byMail.get(email);
        if (existing) {
            if (argeName && existing.groupLabels.indexOf(argeName) < 0) {
                existing.groupLabels.push(argeName);
            }
            return;
        }
        byMail.set(email, {
            name: name,
            email: email,
            label: name || email,
            argeCode: argeCode,
            argeName: argeName,
            functionId: fnId,
            functionLabel: functionLabel,
            groupLabels: argeName ? [argeName] : []
        });
    });
    return Array.from(byMail.values())
        .map(function (p) {
            const hint =
                p.groupLabels.length > 1
                    ? p.groupLabels.join(', ')
                    : p.groupLabels[0] || p.argeName || '';
            return {
                name: p.name,
                email: p.email,
                label: p.label,
                argeCode: p.argeCode,
                argeName: hint || p.argeName,
                functionId: p.functionId,
                functionLabel: p.functionLabel
            };
        })
        .sort(function (a, b) {
            return String(a.name || a.email).localeCompare(String(b.name || b.email), 'de');
        });
}

/**
 * Personen eines Funktions-Sammelbereichs (Registry).
 * @param {object} settings
 * @param {string|SchulOrgFunctionAggregateDef} aggregateOrId
 */
export function collectFunctionAggregatePeople(settings, aggregateOrId) {
    const def =
        aggregateOrId && typeof aggregateOrId === 'object'
            ? aggregateOrId
            : listSchulOrgFunctionAggregates(settings).find(function (d) {
                  return d.id === aggregateOrId || d.boardId === aggregateOrId;
              });
    if (!def || !def.aggregate) return [];
    if (def.source === 'class' && def.seat === 'head') {
        return collectKlassenvorstandPeopleFromSettings(settings).map(function (p) {
            return Object.assign({}, p, {
                functionId: def.id,
                functionLabel: def.label,
                seatLabel: def.seatLabel
            });
        });
    }
    if (def.source === 'arge' && def.seat === 'head') {
        return collectArgeLeiterPeopleFromSettings(settings).map(function (p) {
            return Object.assign({}, p, {
                functionId: def.id,
                functionLabel: def.label,
                seatLabel: def.seatLabel,
                className: p.argeName,
                classCode: p.argeCode
            });
        });
    }
    if (def.source === 'class' && def.seat === 'functionsMap') {
        return collectClassFunctionPeopleFromSettings(settings, def.seatKey || def.id).map(function (p) {
            return Object.assign({}, p, { seatLabel: def.seatLabel });
        });
    }
    if (def.source === 'arge' && def.seat === 'functionsMap') {
        return collectArgeFunctionPeopleFromSettings(settings, def.seatKey || '').map(function (p) {
            return Object.assign({}, p, {
                seatLabel: def.seatLabel,
                className: p.argeName,
                classCode: p.argeCode
            });
        });
    }
    return [];
}

/**
 * ARGE-Leitungen aus Stammdaten (eindeutig nach E-Mail).
 * @param {object} settings
 * @returns {Array<{ name: string, email: string, label: string, argeCode: string, argeName: string }>}
 */
export function collectArgeLeiterPeopleFromSettings(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    /** @type {Map<string, { name: string, email: string, label: string, argeCode: string, argeName: string }>} */
    const byMail = new Map();
    (Array.isArray(s.arges) ? s.arges : []).forEach(function (a) {
        const email = normStr(a && a.headEmail).toLowerCase();
        const name = normStr(a && a.headName);
        if (!email || email.indexOf('@') === -1) return;
        if (byMail.has(email)) return;
        const argeName = normStr(a && a.name) || normCode(a && a.code);
        byMail.set(email, {
            name: name,
            email: email,
            label: name || email,
            argeCode: normCode(a && a.code),
            argeName: argeName
        });
    });
    return Array.from(byMail.values()).sort(function (a, b) {
        return String(a.name || a.email).localeCompare(String(b.name || b.email), 'de');
    });
}

/**
 * Klassenvorstände aus Stammdaten (eindeutig nach E-Mail; Klassenkürzel als Hinweis).
 * @param {object} settings
 * @returns {Array<{ name: string, email: string, label: string, classCode: string, className: string }>}
 */
export function collectKlassenvorstandPeopleFromSettings(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    /** @type {Map<string, { name: string, email: string, label: string, classCode: string, className: string, classLabels: string[] }>} */
    const byMail = new Map();
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        const email = normStr(c && c.headEmail).toLowerCase();
        const name = normStr(c && c.headName);
        if (!email || email.indexOf('@') === -1) return;
        const classCode =
            canonicalSchulOrgClassCode(c && c.code) ||
            canonicalSchulOrgClassCode(c && c.name) ||
            normCode(c && c.code);
        const className = normStr(c && c.name) || classCode;
        const existing = byMail.get(email);
        if (existing) {
            if (className && existing.classLabels.indexOf(className) < 0) {
                existing.classLabels.push(className);
            }
            return;
        }
        byMail.set(email, {
            name: name,
            email: email,
            label: name || email,
            classCode: classCode,
            className: className,
            classLabels: className ? [className] : []
        });
    });
    return Array.from(byMail.values())
        .map(function (p) {
            const hint =
                p.classLabels.length > 1
                    ? p.classLabels.join(', ')
                    : p.classLabels[0] || p.className;
            return {
                name: p.name,
                email: p.email,
                label: p.label,
                classCode: p.classCode,
                className: hint
            };
        })
        .sort(function (a, b) {
            return String(a.name || a.email).localeCompare(String(b.name || b.email), 'de');
        });
}

/**
 * Graph-Gruppen-IDs der ARGEs (Catalog-Links).
 * @param {object|null|undefined} setup
 * @returns {Array<{ groupId: string, code: string, label: string }>}
 */
export function listArgeCatalogGroupRefs(setup) {
    const links =
        setup && Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
    /** @type {Array<{ groupId: string, code: string, label: string }>} */
    const out = [];
    const seen = new Set();
    links.forEach(function (l) {
        if (!l || String(l.kind || '').toLowerCase() !== 'arge') return;
        const gid = normStr(l.graphGroupId);
        if (!gid || seen.has(gid)) return;
        seen.add(gid);
        out.push({
            groupId: gid,
            code: normCode(l.code),
            label: normStr(l.displayName) || normStr(l.code) || gid
        });
    });
    return out;
}

/**
 * Stellt sicher: unter „ARGEs / Fachschaften“ existiert der Bereich ARGE-Leiter.
 * Übernimmt ggf. einen manuell angelegten Bereich gleichen Namens.
 * @param {unknown} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammArgeLeiterStructure(layout, settings) {
    const base = normalizeSchulOrganigramm(layout);
    void settings;
    let secArge = base.nodes.find(function (n) {
        return n.id === SCHUL_ORG_SEC_ARGE || n.id === 'sec_arge';
    });
    if (!secArge) {
        base.nodes.push({
            id: SCHUL_ORG_SEC_ARGE,
            type: NODE.SECTION,
            label: 'ARGEs / Fachschaften',
            parentId: null,
            order: 40,
            tone: 'arge',
            refKind: 'none',
            refKey: ''
        });
        secArge = base.nodes[base.nodes.length - 1];
    }

    const labelMatch = function (n) {
        return /^arge[\s_-]*leiter/i.test(normStr(n && n.label));
    };
    let leiter = base.nodes.find(function (n) {
        return isArgeLeiterArea(n);
    });
    if (!leiter) {
        leiter = base.nodes.find(function (n) {
            return isAreaNode(n) && labelMatch(n) && String(n.parentId || '') === String(secArge.id);
        });
    }
    if (!leiter) {
        leiter = base.nodes.find(function (n) {
            return isAreaNode(n) && labelMatch(n);
        });
    }

    if (!leiter) {
        base.nodes.push({
            id: SCHUL_ORG_SEC_ARGE_LEITER,
            type: NODE.SECTION,
            label: 'ARGE-Leiter',
            parentId: secArge.id,
            order: 0,
            tone: 'arge',
            refKind: 'argeLeiter',
            refKey: 'argeLeiter',
            meta: { argeLeiterBoard: true, nestPeople: true }
        });
    } else {
        const keepM365 = leiter.m365;
        leiter.id = SCHUL_ORG_SEC_ARGE_LEITER;
        leiter.type = NODE.SECTION;
        leiter.parentId = secArge.id;
        leiter.tone = leiter.tone || 'arge';
        leiter.refKind = 'argeLeiter';
        leiter.refKey = 'argeLeiter';
        if (!normStr(leiter.label) || labelMatch(leiter)) leiter.label = 'ARGE-Leiter';
        leiter.meta = Object.assign({}, leiter.meta || {}, {
            argeLeiterBoard: true,
            nestPeople: true
        });
        if (keepM365) leiter.m365 = keepM365;
        // Doppelte manuelle „ARGE-Leiter“-Areas entfernen
        base.nodes = base.nodes.filter(function (n) {
            if (n === leiter || n.id === SCHUL_ORG_SEC_ARGE_LEITER) return true;
            if (isAreaNode(n) && labelMatch(n) && String(n.parentId || '') === String(secArge.id)) {
                return false;
            }
            return true;
        });
    }

    return normalizeSchulOrganigramm(base);
}

/**
 * Stellt sicher: Root-Bereich „Klassenvorstände“ (außerhalb Schüler/ARGE).
 * Übernimmt ggf. einen manuell angelegten Bereich gleichen Namens.
 * @param {unknown} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammKlassenvorstandStructure(layout, settings) {
    const base = normalizeSchulOrganigramm(layout);
    void settings;
    const labelMatch = function (n) {
        return /^klassen\s*vorst[aä]nde?/i.test(normStr(n && n.label));
    };
    let board = base.nodes.find(function (n) {
        return isKlassenvorstandArea(n);
    });
    if (!board) {
        board = base.nodes.find(function (n) {
            return isAreaNode(n) && labelMatch(n) && (n.parentId == null || n.parentId === '');
        });
    }
    if (!board) {
        board = base.nodes.find(function (n) {
            return isAreaNode(n) && labelMatch(n);
        });
    }

    if (!board) {
        let maxOrder = base.nodes.reduce(function (m, n) {
            return Math.max(m, typeof n.order === 'number' ? n.order : 0);
        }, 0);
        base.nodes.push({
            id: SCHUL_ORG_SEC_KLASSENVORSTAND,
            type: NODE.SECTION,
            label: 'Klassenvorstände',
            parentId: null,
            order: maxOrder + 1,
            tone: 'lehrer',
            refKind: 'klassenvorstand',
            refKey: 'klassenvorstand',
            meta: { klassenvorstandBoard: true, nestPeople: true }
        });
    } else {
        const keepM365 = board.m365;
        board.id = SCHUL_ORG_SEC_KLASSENVORSTAND;
        board.type = NODE.SECTION;
        board.parentId = null;
        board.tone = board.tone || 'lehrer';
        board.refKind = 'klassenvorstand';
        board.refKey = 'klassenvorstand';
        if (!normStr(board.label) || labelMatch(board)) board.label = 'Klassenvorstände';
        board.meta = Object.assign({}, board.meta || {}, {
            klassenvorstandBoard: true,
            nestPeople: true
        });
        if (keepM365) board.m365 = keepM365;
        base.nodes = base.nodes.filter(function (n) {
            if (n === board || n.id === SCHUL_ORG_SEC_KLASSENVORSTAND) return true;
            if (isAreaNode(n) && labelMatch(n)) return false;
            return true;
        });
    }

    return normalizeSchulOrganigramm(base);
}

/**
 * @param {unknown} layout
 * @param {SchulOrgFunctionAggregateDef} def
 * @param {string|null} parentId
 */
function ensureOneFunctionAggregateBoard(layout, def, parentId) {
    let base = normalizeSchulOrganigramm(layout);
    const boardId = def.boardId;
    const wantParent = parentId == null || parentId === '' ? null : String(parentId);
    const labelMatch = function (n) {
        const lab = normStr(n && n.label).toLowerCase();
        return (
            lab === normStr(def.labelPlural).toLowerCase() ||
            lab === normStr(def.label).toLowerCase() ||
            lab === normStr(def.labelPlural).toLowerCase() + 'innen'
        );
    };
    const matchesDef = function (n) {
        if (!isAreaNode(n)) return false;
        if (normStr(n.id) === boardId) return true;
        if (normStr(n.refKind).toLowerCase() === normStr(def.refKind).toLowerCase()) {
            if (def.seat === 'functionsMap') {
                return (
                    cleanSchulOrgClassFunctionId(n.refKey) ===
                    cleanSchulOrgClassFunctionId(def.seatKey || def.id)
                );
            }
            return true;
        }
        return false;
    };
    let board = base.nodes.find(matchesDef);
    if (!board) {
        board = base.nodes.find(function (n) {
            return (
                isAreaNode(n) &&
                labelMatch(n) &&
                (wantParent == null
                    ? n.parentId == null || n.parentId === ''
                    : String(n.parentId || '') === wantParent)
            );
        });
    }
    if (!board) {
        board = base.nodes.find(function (n) {
            return isAreaNode(n) && labelMatch(n);
        });
    }
    const refKey =
        def.seat === 'functionsMap'
            ? cleanSchulOrgClassFunctionId(def.seatKey || def.id)
            : cleanSchulOrgClassFunctionId(def.id) || def.id;
    if (!board) {
        let maxOrder = base.nodes.reduce(function (m, n) {
            return Math.max(m, typeof n.order === 'number' ? n.order : 0);
        }, 0);
        base.nodes.push({
            id: boardId,
            type: NODE.SECTION,
            label: def.labelPlural,
            parentId: wantParent,
            order: maxOrder + 1,
            tone: def.tone,
            refKind: def.refKind,
            refKey: refKey,
            meta: {
                functionAggregate: true,
                functionAggregateId: def.id,
                classFunctionBoard: def.refKind === 'classFunction',
                argeFunctionBoard: def.refKind === 'argeFunction',
                klassenvorstandBoard: def.refKind === 'klassenvorstand',
                argeLeiterBoard: def.refKind === 'argeLeiter',
                classFunctionId: def.refKind === 'classFunction' ? refKey : undefined,
                argeFunctionId: def.refKind === 'argeFunction' ? refKey : undefined,
                nestPeople: true
            }
        });
    } else {
        const keepM365 = board.m365;
        board.id = boardId;
        board.type = NODE.SECTION;
        board.parentId = wantParent;
        board.tone = board.tone || def.tone;
        board.refKind = def.refKind;
        board.refKey = refKey;
        if (!normStr(board.label) || labelMatch(board)) board.label = def.labelPlural;
        board.meta = Object.assign({}, board.meta || {}, {
            functionAggregate: true,
            functionAggregateId: def.id,
            classFunctionBoard: def.refKind === 'classFunction',
            argeFunctionBoard: def.refKind === 'argeFunction',
            klassenvorstandBoard: def.refKind === 'klassenvorstand',
            argeLeiterBoard: def.refKind === 'argeLeiter',
            classFunctionId: def.refKind === 'classFunction' ? refKey : undefined,
            argeFunctionId: def.refKind === 'argeFunction' ? refKey : undefined,
            nestPeople: true
        });
        if (keepM365) board.m365 = keepM365;
        base.nodes = base.nodes.filter(function (n) {
            if (n === board || n.id === boardId) return true;
            if (matchesDef(n) && n !== board) return false;
            if (isAreaNode(n) && labelMatch(n) && n !== board) {
                const sameParent =
                    wantParent == null
                        ? n.parentId == null || n.parentId === ''
                        : String(n.parentId || '') === wantParent;
                if (sameParent) return false;
            }
            return true;
        });
    }
    return normalizeSchulOrganigramm(base);
}

/**
 * Stellt sicher: Root-Sammelbereiche je Klassenfunktion (z. B. Klassensprecher).
 * @param {unknown} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammClassFunctionStructure(layout, settings) {
    let base = normalizeSchulOrganigramm(layout);
    listSchulOrgFunctionAggregates(settings)
        .filter(function (d) {
            return d.refKind === 'classFunction' && d.aggregate;
        })
        .forEach(function (def) {
            base = ensureOneFunctionAggregateBoard(base, def, null);
        });
    return base;
}

/**
 * Stellt sicher: Sammelbereiche je ARGE-Funktion unter ARGEs.
 * @param {unknown} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammArgeFunctionStructure(layout, settings) {
    let base = normalizeSchulOrganigramm(layout);
    let secArge = base.nodes.find(function (n) {
        return n.id === SCHUL_ORG_SEC_ARGE || n.id === 'sec_arge';
    });
    if (!secArge && listSchulOrgArgeFunctions(settings).length) {
        base.nodes.push({
            id: SCHUL_ORG_SEC_ARGE,
            type: NODE.SECTION,
            label: 'ARGEs / Fachschaften',
            parentId: null,
            order: 40,
            tone: 'arge',
            refKind: 'none',
            refKey: ''
        });
        secArge = base.nodes[base.nodes.length - 1];
    }
    const parentId = secArge ? secArge.id : SCHUL_ORG_SEC_ARGE;
    listSchulOrgFunctionAggregates(settings)
        .filter(function (d) {
            return d.refKind === 'argeFunction' && d.aggregate;
        })
        .forEach(function (def) {
            base = ensureOneFunctionAggregateBoard(base, def, parentId);
        });
    return base;
}

/**
 * Alle Funktions-Sammelbereiche (KV, ARGE-Leiter, Klassen-/ARGE-Funktionen).
 * @param {unknown} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammFunctionAggregateStructure(layout, settings) {
    let base = ensureSchulOrganigrammArgeLeiterStructure(layout, settings);
    base = ensureSchulOrganigrammKlassenvorstandStructure(base, settings);
    base = ensureSchulOrganigrammClassFunctionStructure(base, settings);
    base = ensureSchulOrganigrammArgeFunctionStructure(base, settings);
    return base;
}

/**
 * ARGEs als Bereichskacheln unter sec_arge (Legacy-Leaves → Areas, Labels/Meta aus Stammdaten).
 * @param {unknown} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammArgeStructure(layout, settings) {
    const base = ensureSchulOrganigrammArgeLeiterStructure(layout, settings);
    const s = settings && typeof settings === 'object' ? settings : {};
    const teachers = Array.isArray(s.teachers) ? s.teachers : [];
    let secArge = base.nodes.find(function (n) {
        return n.id === SCHUL_ORG_SEC_ARGE || n.id === 'sec_arge';
    });
    let maxOrder = base.nodes.reduce(function (m, n) {
        return Math.max(m, typeof n.order === 'number' ? n.order : 0);
    }, 0);
    if (!secArge) {
        maxOrder += 1;
        secArge = {
            id: SCHUL_ORG_SEC_ARGE,
            type: NODE.SECTION,
            label: 'ARGEs / Fachschaften',
            parentId: null,
            order: maxOrder,
            tone: 'arge',
            refKind: 'none',
            refKey: ''
        };
        base.nodes.push(secArge);
    }

    /** @type {Set<string>} */
    const wantCodes = new Set();
    (Array.isArray(s.arges) ? s.arges : []).forEach(function (a) {
        const code = normCode(a && a.code);
        const name = normStr(a && a.name) || code;
        if (!code && !name) return;
        const key = code || name;
        wantCodes.add(String(key).toUpperCase());
        const headEmail = normStr(a && a.headEmail).toLowerCase();
        const teacher =
            headEmail.indexOf('@') !== -1
                ? teachers.find(function (t) {
                      return normStr(t && t.email).toLowerCase() === headEmail;
                  })
                : null;
        const subjects = Array.isArray(a && a.subjects) ? a.subjects.slice() : [];
        const wantId = argeNodeId(code) || 'arge_' + normCode(name) || 'arge_' + key;
        let node = base.nodes.find(function (n) {
            if (!isArgeNode(n)) return false;
            return (
                n.id === wantId ||
                normCode(n.refKey) === code ||
                (code && normCode(n.refKey) === code)
            );
        });
        if (!node) {
            maxOrder += 1;
            base.nodes.push({
                id: wantId,
                type: NODE.SECTION,
                label: name,
                parentId: secArge.id,
                order: maxOrder,
                tone: 'arge',
                refKind: 'arge',
                refKey: key,
                meta: {
                    nestPeople: true,
                    audience: 'arge',
                    headName: normStr(a && a.headName),
                    headEmail: headEmail,
                    teacherLinked: !!teacher,
                    teacherCode: teacher ? normCode(teacher.code) : '',
                    subjects: subjects
                }
            });
            return;
        }
        const keepM365 = node.m365;
        const keepParent = node.parentId;
        node.id = wantId;
        node.type = NODE.SECTION;
        node.refKind = 'arge';
        node.refKey = key;
        node.tone = node.tone || 'arge';
        node.label = name;
        // Ungültiger Parent → zurück unter ARGEs; sonst manuelle Unterbereiche behalten
        const pid = keepParent == null || keepParent === '' ? null : String(keepParent);
        if (!pid || pid === SCHUL_ORG_SEC_ARGE_LEITER) {
            node.parentId = secArge.id;
        } else if (pid !== String(secArge.id)) {
            const parent = base.nodes.find(function (n) {
                return n.id === pid;
            });
            if (!parent || !isAreaNode(parent) || !canPlaceSchulOrgNodeInSection(node, pid, base)) {
                node.parentId = secArge.id;
            }
        } else {
            node.parentId = secArge.id;
        }
        node.meta = Object.assign({}, node.meta || {}, {
            nestPeople: true,
            audience: 'arge',
            headName: normStr(a && a.headName),
            headEmail: headEmail,
            teacherLinked: !!teacher,
            teacherCode: teacher ? normCode(teacher.code) : '',
            subjects: subjects
        });
        if (keepM365) node.m365 = keepM365;
    });

    base.nodes = base.nodes.filter(function (n) {
        if (!isArgeNode(n)) return true;
        const key = String(normCode(n.refKey) || '').toUpperCase();
        if (!key) return false;
        return wantCodes.has(key);
    });

    return normalizeSchulOrganigramm(base);
}

export function sammelgruppeKeyForArea(areaNode) {
    if (!isAreaNode(areaNode)) return '';
    const rk = normStr(areaNode.refKind).toLowerCase();
    const key = normStr(areaNode.refKey).toLowerCase();
    if (
        rk === 'sammelgruppe' &&
        (key === 'lehrer' || key === 'schueler' || key === 'eltern' || key === 'klassenchats')
    ) {
        return key;
    }
    const id = normStr(areaNode.id);
    if (id === SCHUL_ORG_SEC_LEHRER) return 'lehrer';
    if (id === SCHUL_ORG_SEC_SCHUELER) return 'schueler';
    if (id === SCHUL_ORG_SEC_ELTERN) return 'eltern';
    if (id === SCHUL_ORG_SEC_KLASSENCHATS) return 'klassenchats';
    return '';
}

/**
 * @param {object|null|undefined} areaNode
 */
export function isSammelgruppeArea(areaNode) {
    return !!sammelgruppeKeyForArea(areaNode);
}

/**
 * Kanonisches Klassenkürzel (z. B. „Klasse 5HM - SJ2026/27“ → „5HM“).
 * Verhindert Doppelknoten unter Jahrgang und „Sonstige Klassen“.
 * @param {string} codeOrName
 * @returns {string}
 */
export function canonicalSchulOrgClassCode(codeOrName) {
    const raw = normStr(codeOrName);
    if (!raw) return '';
    const fromDisplay = classCodeFromDisplayName(raw);
    if (fromDisplay) return fromDisplay;

    // „Klasse 5HM …“ / „Klasse2HMB“ (ohne Leerzeichen)
    let s = raw.replace(/^Klasse\s*/i, '').trim();
    s = s.replace(/\s*[-–]\s*SJ\s*\d{2,4}.*$/i, '').trim();
    s = s.replace(/\s*[-–]\s*Schuljahr.*$/i, '').trim();
    const again = classCodeFromDisplayName(s) || classCodeFromDisplayName('Klasse ' + s);
    if (again) return again;

    const upper = normCode(s || raw);
    // Reines Kürzel (5HM, 1AK, VK, …)
    if (/^\d+[A-ZÄÖÜß]*$/i.test(upper) || /^[A-ZÄÖÜß]{1,16}$/i.test(upper)) {
        return upper;
    }
    const compact = upper.replace(/[^0-9A-ZÄÖÜß]/gi, '');
    if (compact && /^\d+[A-ZÄÖÜß]*$/i.test(compact)) return compact;
    if (compact && compact.length <= 16 && !/^KLASSE/i.test(compact)) return compact;
    const first = (s || raw).split(/\s*[-–]\s*/)[0];
    return normCode(first).replace(/[^0-9A-ZÄÖÜß]/gi, '') || upper;
}

/**
 * @param {{ code?: string, name?: string }|null|undefined} row
 */
export function classCodeFromSettingsRow(row) {
    return (
        canonicalSchulOrgClassCode(row && row.code) ||
        canonicalSchulOrgClassCode(row && row.name) ||
        ''
    );
}

/**
 * Schulstufe / Jahrgang aus Klassenkürzel (z. B. 5A → "5", 10B → "10").
 * @param {string} classCode
 * @returns {string} Stufe oder SCHUL_ORG_JAHRGANG_SONST
 */
export function schulstufeFromClassCode(classCode) {
    const code = canonicalSchulOrgClassCode(classCode) || normCode(classCode);
    if (!code) return SCHUL_ORG_JAHRGANG_SONST;
    const m = /^(\d+)/.exec(code);
    return m ? m[1] : SCHUL_ORG_JAHRGANG_SONST;
}

/**
 * @param {string} stufe
 */
export function jahrgangAreaId(stufe) {
    const key = normStr(stufe).toLowerCase() || SCHUL_ORG_JAHRGANG_SONST;
    return 'sec_jahrgang_' + key;
}

/**
 * @param {string} stufe
 */
export function jahrgangAreaLabel(stufe) {
    const key = normStr(stufe);
    if (!key || key.toLowerCase() === SCHUL_ORG_JAHRGANG_SONST) return 'Sonstige Klassen';
    return key + '. Jahrgang';
}

/**
 * @param {object|null|undefined} areaNode
 * @returns {string} Jahrgangs-Key („1“, „2“, „sonst“) oder ''
 */
export function jahrgangKeyForArea(areaNode) {
    if (!isAreaNode(areaNode)) return '';
    if (normStr(areaNode.refKind).toLowerCase() === 'jahrgang' && normStr(areaNode.refKey)) {
        return normStr(areaNode.refKey).toLowerCase();
    }
    const id = normStr(areaNode.id);
    const m = /^sec_jahrgang_(.+)$/i.exec(id);
    return m ? String(m[1]).toLowerCase() : '';
}

/**
 * @param {object|null|undefined} areaNode
 */
export function isJahrgangArea(areaNode) {
    return !!jahrgangKeyForArea(areaNode);
}

/**
 * Eltern-Jahrgang unter sec_eltern (parallel zu Schüler-Jahrgängen).
 * @param {string} stufe
 */
export function elternJahrgangAreaId(stufe) {
    const key = normStr(stufe).toLowerCase() || SCHUL_ORG_JAHRGANG_SONST;
    return 'sec_eltern_jahrgang_' + key;
}

/**
 * @param {object|null|undefined} areaNode
 * @returns {string} Jahrgangs-Key oder ''
 */
export function elternJahrgangKeyForArea(areaNode) {
    if (!isAreaNode(areaNode)) return '';
    if (normStr(areaNode.refKind).toLowerCase() === 'elternJahrgang' && normStr(areaNode.refKey)) {
        return normStr(areaNode.refKey).toLowerCase();
    }
    const id = normStr(areaNode.id);
    const m = /^sec_eltern_jahrgang_(.+)$/i.exec(id);
    return m ? String(m[1]).toLowerCase() : '';
}

/**
 * @param {object|null|undefined} areaNode
 */
export function isElternJahrgangArea(areaNode) {
    return !!elternJahrgangKeyForArea(areaNode);
}

/**
 * @param {string} classCode
 */
export function elternClassNodeId(classCode) {
    const code = canonicalSchulOrgClassCode(classCode) || normCode(classCode);
    return code ? 'eltern_class_' + code : '';
}

/**
 * Klassen-Kachel unter Eltern (Erziehungsberechtigte der Klasse).
 * @param {object|null|undefined} node
 */
export function isElternClassNode(node) {
    if (!node) return false;
    if (normStr(node.refKind).toLowerCase() === 'elternclass') return true;
    return /^eltern_class_/i.test(normStr(node.id));
}

/**
 * Anzahl eindeutiger Elternmails aus Schüler-parentPairs (Stammdaten).
 * @param {object} settings
 */
export function countUniqueGuardiansInSettings(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const seen = new Set();
    (Array.isArray(s.students) ? s.students : []).forEach(function (st) {
        (Array.isArray(st && st.parentPairs) ? st.parentPairs : []).forEach(function (p) {
            const em = normStr(p && p.email).toLowerCase();
            if (em && em.indexOf('@') !== -1) seen.add(em);
        });
    });
    (Array.isArray(s.guardians) ? s.guardians : []).forEach(function (g) {
        const em = normStr(g && g.email).toLowerCase();
        if (em && em.indexOf('@') !== -1) seen.add(em);
    });
    return seen.size;
}

/**
 * KlassenChat-Jahrgang unter sec_klassenchats.
 * @param {string} stufe
 */
export function klassenchatJahrgangAreaId(stufe) {
    const key = normStr(stufe).toLowerCase() || SCHUL_ORG_JAHRGANG_SONST;
    return 'sec_klassenchat_jahrgang_' + key;
}

/**
 * @param {object|null|undefined} areaNode
 * @returns {string}
 */
export function klassenchatJahrgangKeyForArea(areaNode) {
    if (!isAreaNode(areaNode)) return '';
    if (normStr(areaNode.refKind).toLowerCase() === 'klassenchatjahrgang' && normStr(areaNode.refKey)) {
        return normStr(areaNode.refKey).toLowerCase();
    }
    const id = normStr(areaNode.id);
    const m = /^sec_klassenchat_jahrgang_(.+)$/i.exec(id);
    return m ? String(m[1]).toLowerCase() : '';
}

/**
 * @param {object|null|undefined} areaNode
 */
export function isKlassenchatJahrgangArea(areaNode) {
    return !!klassenchatJahrgangKeyForArea(areaNode);
}

/**
 * @param {string} classCode
 */
export function klassenchatClassNodeId(classCode) {
    const code = canonicalSchulOrgClassCode(classCode) || normCode(classCode);
    return code ? 'klassenchat_class_' + code : '';
}

/**
 * Klassen-Kachel unter KlassenChats (Lehrkräfte-Gruppenchat der Klasse).
 * @param {object|null|undefined} node
 */
export function isKlassenchatClassNode(node) {
    if (!node) return false;
    if (normStr(node.refKind).toLowerCase() === 'klassenchatclass') return true;
    return /^klassenchat_class_/i.test(normStr(node.id));
}

/**
 * Schüler-Klasse als Bereichskachel (Jahrgang → Klasse, KV + Schüler + M365-Gruppe).
 * @param {object|null|undefined} node
 */
export function isClassAreaNode(node) {
    if (!node || !isAreaNode(node)) return false;
    if (isElternClassNode(node) || isKlassenchatClassNode(node)) return false;
    if (normStr(node.refKind).toLowerCase() === 'class') return true;
    return /^class_/i.test(normStr(node.id));
}

/**
 * Klassen-Knoten (Bereich oder Legacy-Leaf).
 * @param {object|null|undefined} node
 */
export function isClassNode(node) {
    if (!node) return false;
    if (normStr(node.type).toLowerCase() === NODE.CLASS) return true;
    return isClassAreaNode(node);
}

/**
 * @param {string} code
 */
export function argeNodeId(code) {
    const c = normCode(code);
    return c ? 'arge_' + c : '';
}

/**
 * ARGE als Bereichskachel (wie KlassenChat): Leitung + Mitglieder + optional M365-Gruppe.
 * @param {object|null|undefined} node
 */
export function isArgeAreaNode(node) {
    if (!node || !isAreaNode(node)) return false;
    if (isArgeLeiterArea(node)) return false;
    if (normStr(node.refKind).toLowerCase() === 'arge') return true;
    return /^arge_/i.test(normStr(node.id));
}

/**
 * ARGE-Knoten (Bereich oder Legacy-Leaf).
 * @param {object|null|undefined} node
 */
export function isArgeNode(node) {
    if (!node) return false;
    if (isArgeLeiterArea(node)) return false;
    if (normStr(node.type).toLowerCase() === NODE.ARGE) return true;
    return isArgeAreaNode(node);
}

/**
 * Vorhandene Teams-Chat-Metadaten pro Klasse aus years.*.classChats.
 * @param {object} settings
 * @returns {Map<string, { chatId: string, topic: string, memberEmails: string[], webUrl: string }>}
 */
export function classChatMetaByKlasse(settings) {
    /** @type {Map<string, { chatId: string, topic: string, memberEmails: string[], webUrl: string }>} */
    const map = new Map();
    const s = settings && typeof settings === 'object' ? settings : {};
    const snap = s.classChats && typeof s.classChats === 'object' ? s.classChats : null;
    (snap && Array.isArray(snap.items) ? snap.items : []).forEach(function (row) {
        const klasse = canonicalSchulOrgClassCode(row && row.klasse) || normCode(row && row.klasse);
        if (!klasse) return;
        const kind = normStr(row && row.kind).toLowerCase();
        const chatId = normStr(row && row.chatId);
        const channelId = normStr(row && row.channelId);
        // Gruppenchat oder privater Kanal (channelId) zählen als angelegte Ressource
        const resourceId = chatId || channelId;
        if (!resourceId) return;
        map.set(klasse, {
            chatId: resourceId,
            topic: normStr(row && row.topic),
            memberEmails: Array.isArray(row.memberEmails)
                ? row.memberEmails.map(function (e) {
                      return normStr(e).toLowerCase();
                  }).filter(Boolean)
                : [],
            webUrl: normStr(row && row.webUrl),
            kind: kind === 'privatechannel' || (!chatId && channelId) ? 'privateChannel' : 'groupChat',
            channelId: channelId
        });
    });
    return map;
}

const BUILTIN_AREA_IDS = {
    sec_schulleitung: true,
    sec_verwaltung: true,
    sec_lehrer: true,
    sec_schueler: true,
    sec_eltern: true,
    sec_klassenchats: true,
    sec_arge: true,
    sec_arge_leiter: true,
    sec_klassenvorstand: true,
    sec_classfn_klassensprecher: true,
    sec_klassen: true,
    // Legacy-Wrapper „Schülervertretung & SGA“ – Migration entfernt ihn zugunsten von sga_board
    sga_board: true,
    sga_gruppe_schueler: true,
    sga_gruppe_lehrer: true,
    sga_gruppe_eltern: true
};

/** Schulleitung/Verwaltung: von der Landkarte entfernbar (Sync legt sie nicht erneut an). */
const OPTIONAL_DELETABLE_BUILTIN_AREA_IDS = {
    sec_schulleitung: true,
    sec_verwaltung: true
};

/**
 * @param {object|null|undefined} areaNode
 * @returns {'student'|'teacher'|'external'|''}
 */
export function sgaScopeForArea(areaNode) {
    if (!areaNode) return '';
    if (normStr(areaNode.refKind) === 'sgaScope') {
        const k = normStr(areaNode.refKey).toLowerCase();
        if (k === 'student' || k === 'teacher' || k === 'external') return k;
    }
    const id = normStr(areaNode.id);
    const hit = SCHUL_ORG_SGA_SCOPE_AREAS.find(function (s) {
        return s.id === id;
    });
    return hit ? hit.scope : '';
}

/**
 * @param {object|null|undefined} areaNode
 */
export function isSgaScopeArea(areaNode) {
    return !!sgaScopeForArea(areaNode);
}

/**
 * @param {object|null|undefined} areaNode
 */
export function isSgaBoardArea(areaNode) {
    if (!areaNode) return false;
    return (
        normStr(areaNode.id) === SCHUL_ORG_SGA_BOARD_ID ||
        (isAreaNode(areaNode) &&
            normStr(areaNode.refKind) === 'sga' &&
            normStr(areaNode.refKey) === SCHUL_ORG_BOARD_REF)
    );
}

/**
 * @param {object|null|undefined} node
 * @returns {{ id: string, refKey: string, label: string, seat: number }|null}
 */
export function svSeatForNode(node) {
    if (!node) return null;
    const id = normStr(node.id);
    const key = normStr(node.refKey).toLowerCase();
    return (
        SCHUL_ORG_SV_SEATS.find(function (s) {
            return s.id === id || s.refKey === key;
        }) || null
    );
}

/**
 * @param {object|null|undefined} node
 */
export function isSvSeatNode(node) {
    return !!svSeatForNode(node);
}

/**
 * Person für einen SV-Sitz: zuerst Schülervertretung-Liste, sonst SGA-Schülerplatz.
 * @param {object} settings
 * @param {number} seatIndex
 * @returns {{ name: string, email: string, klasse?: string }|null}
 */
export function personForSvSeat(settings, seatIndex) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const idx = Number(seatIndex);
    if (!Number.isFinite(idx) || idx < 0) return null;
    const council = (Array.isArray(s.studentCouncil) ? s.studentCouncil : []).filter(function (row) {
        return normStr(row && row.name) || normStr(row && row.email);
    });
    if (council[idx]) {
        return {
            name: normStr(council[idx].name),
            email: normStr(council[idx].email).toLowerCase(),
            klasse: normStr(council[idx].klasse)
        };
    }
    const sgaStudents = (Array.isArray(s.sga) ? s.sga : []).filter(function (row) {
        if (!(normStr(row && row.name) || normStr(row && row.email))) return false;
        const sc = normStr(row && row.scope).toLowerCase();
        return !sc || sc === 'student' || sc === 'schueler' || sc === 'schüler';
    });
    if (sgaStudents[idx]) {
        return {
            name: normStr(sgaStudents[idx].name),
            email: normStr(sgaStudents[idx].email).toLowerCase(),
            klasse: ''
        };
    }
    return null;
}

/**
 * Standard-Seed-Bereiche (nicht löschbar).
 * @param {object|null|undefined} areaNode
 */
export function isBuiltinSchulOrganigrammArea(areaNode) {
    if (!isAreaNode(areaNode)) return false;
    const id = normStr(areaNode.id);
    if (BUILTIN_AREA_IDS[id]) return true;
    if (isSammelgruppeArea(areaNode)) return true;
    if (isJahrgangArea(areaNode)) return true;
    if (isSgaBoardArea(areaNode) || isSgaScopeArea(areaNode)) return true;
    if (isFunctionAggregateArea(areaNode)) return true;
    const aud = audienceIdForArea(areaNode);
    if (aud === 'schulleitung' || aud === 'verwaltung') return true;
    return false;
}

/**
 * Schulleitung / Verwaltung (Personal) – absichtlich von der Landkarte entfernbar.
 * @param {object|null|undefined} areaNode
 */
export function isOptionalDeletableBuiltinArea(areaNode) {
    if (!isAreaNode(areaNode)) return false;
    return !!OPTIONAL_DELETABLE_BUILTIN_AREA_IDS[normStr(areaNode.id)];
}

/**
 * Selbst angelegte / verschachtelte Bereiche – sowie Schulleitung/Verwaltung.
 * Lehrer/Schüler/ARGE/Jahrgang/SGA bleiben geschützt.
 * @param {object|null|undefined} areaNode
 */
export function canDeleteSchulOrganigrammArea(areaNode) {
    if (!isAreaNode(areaNode)) return false;
    if (isOptionalDeletableBuiltinArea(areaNode)) return true;
    return !isBuiltinSchulOrganigrammArea(areaNode);
}

/**
 * @param {unknown} layout
 * @returns {string[]}
 */
export function schulOrganigrammRemovedAreaIds(layout) {
    const o = layout && typeof layout === 'object' ? layout : {};
    const list = Array.isArray(o.removedAreaIds) ? o.removedAreaIds : [];
    /** @type {string[]} */
    const out = [];
    const seen = new Set();
    list.forEach(function (id) {
        const s = normStr(id);
        if (!s || seen.has(s)) return;
        seen.add(s);
        out.push(s);
    });
    return out;
}

/**
 * @param {unknown} layout
 * @param {string} areaId
 */
export function withRemovedSchulOrganigrammArea(layout, areaId) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(areaId);
    if (!id) return base;
    const removed = schulOrganigrammRemovedAreaIds(base);
    if (removed.indexOf(id) < 0) removed.push(id);
    return Object.assign({}, base, { removedAreaIds: removed });
}

/**
 * Area kann direkte Personen aufnehmen (Slots oder Audience-Membership).
 * @param {object|null|undefined} areaNode
 */
export function areaAcceptsDirectPeople(areaNode) {
    return isAreaNode(areaNode);
}

/**
 * @param {string} email
 */
export function personNodeIdFromEmail(email) {
    const em = normStr(email).toLowerCase();
    if (!em) return newId('person');
    const slug = em.replace(/[^a-z0-9]+/g, '_').slice(0, 48);
    return 'person_' + slug;
}

/** @typedef {'unified'|'security'|'mailSecurity'|'distribution'|'groupChat'} SchulOrgGroupKind */
/** @typedef {'unified'|'security'|'mailSecurity'|'distribution'|'groupChat'|'mailbox'} SchulOrgM365ObjectKind */

/** Alle Microsoft-365-Objekttypen, die einem Bereich zugeordnet werden können. */
export const SCHUL_ORG_M365_OBJECT_KINDS = [
    'unified',
    'security',
    'mailSecurity',
    'distribution',
    'groupChat',
    'mailbox'
];

/**
 * @param {unknown} kind
 * @returns {SchulOrgM365ObjectKind|''}
 */
export function normalizeSchulOrgM365ObjectKind(kind) {
    const k = normStr(kind).toLowerCase();
    if (!k) return '';
    if (k === 'mailbox' || k === 'sharedmailbox' || k === 'shared_mailbox' || k === 'postfach') {
        return 'mailbox';
    }
    if (k === 'security' || k === 'sg') return 'security';
    if (k === 'mailsecurity' || k === 'mail_security' || k === 'mesg') return 'mailSecurity';
    if (k === 'distribution' || k === 'dl' || k === 'verteiler') return 'distribution';
    if (
        k === 'groupchat' ||
        k === 'group_chat' ||
        k === 'chat' ||
        k === 'gruppenchat' ||
        k === 'teamschat' ||
        k === 'teams_chat'
    ) {
        return 'groupChat';
    }
    if (
        k === 'unified' ||
        k === 'm365' ||
        k === 'microsoft365' ||
        k === 'gruppe' ||
        k === 'group'
    ) {
        return 'unified';
    }
    return 'unified';
}

/**
 * @param {unknown} kind
 */
export function schulOrgM365ObjectKindLabel(kind) {
    const k = normalizeSchulOrgM365ObjectKind(kind);
    if (k === 'mailbox') return 'Freigegebenes Postfach';
    if (!k) return 'Microsoft 365';
    return schulOrgGroupKindLabel(k);
}

/**
 * @param {unknown} kind
 */
export function schulOrgM365ObjectKindIcon(kind) {
    const k = normalizeSchulOrgM365ObjectKind(kind);
    if (k === 'mailbox') return schulOrgMailboxIcon();
    if (!k) return 'bi-microsoft';
    return schulOrgGroupKindIcon(k);
}

/**
 * @returns {Record<SchulOrgM365ObjectKind, { mode: string, id: string|null, label: string|null, mail: string|null, upn: string|null, webUrl: string|null }>}
 */
export function emptySchulOrgM365Bindings() {
    /** @type {Record<string, { mode: string, id: string|null, label: string|null, mail: string|null, upn: string|null, webUrl: string|null }>} */
    const out = {};
    SCHUL_ORG_M365_OBJECT_KINDS.forEach(function (kind) {
        out[kind] = {
            mode: 'none',
            id: null,
            label: null,
            mail: null,
            upn: null,
            webUrl: null
        };
    });
    return /** @type {any} */ (out);
}

/**
 * Eine Binding-Zeile normalisieren.
 * @param {unknown} raw
 * @param {SchulOrgM365ObjectKind} kind
 */
export function normalizeSchulOrgM365Binding(raw, kind) {
    const k = normalizeSchulOrgM365ObjectKind(kind) || 'unified';
    const b = raw && typeof raw === 'object' ? raw : {};
    const modeRaw = normStr(b.mode || b.groupMode || b.mailboxMode).toLowerCase();
    const mode = modeRaw === 'match' || modeRaw === 'create' ? modeRaw : 'none';
    const id = mode === 'none' ? null : normStr(b.id || b.groupId || b.chatId || b.mailboxId) || null;
    const label =
        mode === 'none' ? null : normStr(b.label || b.groupLabel || b.mailboxLabel) || null;
    const mail =
        mode === 'none'
            ? null
            : normStr(b.mail || b.groupMail).toLowerCase() || null;
    const upn =
        k === 'mailbox' && mode !== 'none'
            ? normStr(b.upn || b.mailboxUpn).toLowerCase() || null
            : null;
    const webUrl = k === 'groupChat' && mode !== 'none' ? normStr(b.webUrl || b.chatWebUrl) || null : null;
    return {
        mode: mode,
        id: id,
        label: label,
        mail: mail,
        upn: upn,
        webUrl: webUrl
    };
}

/**
 * Legacy groupMode/groupKind/mailbox* → bindings und zurück spiegeln.
 * @param {unknown} raw
 */
export function normalizeSchulOrgM365(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const bindings = emptySchulOrgM365Bindings();
    const srcBindings = raw.bindings && typeof raw.bindings === 'object' ? raw.bindings : null;
    if (srcBindings) {
        SCHUL_ORG_M365_OBJECT_KINDS.forEach(function (kind) {
            if (srcBindings[kind] != null) {
                bindings[kind] = normalizeSchulOrgM365Binding(srcBindings[kind], kind);
            }
        });
    } else {
        // Altformat: ein groupKind + optionales Postfach
        // groupMode=none / mailboxMode=none gewinnt – verwaiste IDs werden verworfen
        const gk = normalizeSchulOrgGroupKind(raw.groupKind || 'unified');
        const gm = normStr(raw.groupMode).toLowerCase();
        const groupMode = gm === 'match' || gm === 'create' ? gm : 'none';
        if (groupMode !== 'none') {
            const slot = gk === 'groupChat' ? 'groupChat' : gk;
            bindings[slot] = normalizeSchulOrgM365Binding(
                {
                    mode: groupMode,
                    id: raw.groupId,
                    label: raw.groupLabel,
                    mail: raw.groupMail,
                    webUrl: raw.chatWebUrl || raw.webUrl
                },
                slot
            );
        }
        const mm = normStr(raw.mailboxMode).toLowerCase();
        const mailboxMode = mm === 'match' || mm === 'create' ? mm : 'none';
        if (mailboxMode !== 'none') {
            bindings.mailbox = normalizeSchulOrgM365Binding(
                {
                    mode: mailboxMode,
                    id: raw.mailboxId,
                    label: raw.mailboxLabel,
                    upn: raw.mailboxUpn
                },
                'mailbox'
            );
        }
    }

    // Legacy-Spiegel: nur Entra-Gruppe + Postfach (Abgleich/Provision).
    // Teams-Gruppenchat gehört NICHT in groupMode/groupId – sonst landet die Chat-ID
    // in Graph-Gruppen-Pfaden und UI-Sync (isChat ↔ groupMode) gerät in Widerspruch.
    const entraOrder = ['unified', 'security', 'mailSecurity', 'distribution'];
    let primaryKind = 'unified';
    let primary = bindings.unified;
    for (let i = 0; i < entraOrder.length; i++) {
        const kind = entraOrder[i];
        const b = bindings[kind];
        if (b && (b.mode === 'match' || b.mode === 'create' || b.id)) {
            primaryKind = kind;
            primary = b;
            break;
        }
    }
    const mb = bindings.mailbox;
    return {
        bindings: bindings,
        groupMode: primary && primary.mode !== 'none' ? primary.mode : 'none',
        groupKind: primaryKind,
        groupId: primary && primary.mode !== 'none' ? primary.id : null,
        groupLabel: primary && primary.mode !== 'none' ? primary.label : null,
        groupMail: primary && primary.mode !== 'none' ? primary.mail : null,
        mailboxMode: mb && mb.mode !== 'none' ? mb.mode : 'none',
        mailboxUpn: mb && mb.mode !== 'none' ? mb.upn : null,
        mailboxLabel: mb && mb.mode !== 'none' ? mb.label : null,
        mailboxId: mb && mb.mode !== 'none' ? mb.id : null
    };
}

/**
 * @param {object|null|undefined} m365
 * @param {SchulOrgM365ObjectKind} kind
 */
export function getSchulOrgM365Binding(m365, kind) {
    const norm = normalizeSchulOrgM365(m365 || {});
    const k = normalizeSchulOrgM365ObjectKind(kind);
    if (!norm || !k) return normalizeSchulOrgM365Binding(null, 'unified');
    return norm.bindings[k] || normalizeSchulOrgM365Binding(null, k);
}

/**
 * @param {object|null|undefined} m365
 * @param {SchulOrgM365ObjectKind} kind
 * @param {object|null} patch
 */
export function setSchulOrgM365Binding(m365, kind, patch) {
    const norm = normalizeSchulOrgM365(m365 || {}) || {
        bindings: emptySchulOrgM365Bindings()
    };
    const k = normalizeSchulOrgM365ObjectKind(kind);
    if (!k) return normalizeSchulOrgM365(norm);
    if (patch == null) {
        norm.bindings[k] = normalizeSchulOrgM365Binding(null, k);
    } else {
        norm.bindings[k] = normalizeSchulOrgM365Binding(
            Object.assign({}, norm.bindings[k] || {}, patch),
            k
        );
    }
    return normalizeSchulOrgM365({ bindings: norm.bindings });
}

/**
 * @param {unknown} kind
 * @returns {SchulOrgGroupKind}
 */
export function normalizeSchulOrgGroupKind(kind) {
    const k = normStr(kind).toLowerCase();
    if (k === 'security' || k === 'sg') return 'security';
    if (k === 'mailsecurity' || k === 'mail_security' || k === 'mesg') return 'mailSecurity';
    if (k === 'distribution' || k === 'dl' || k === 'verteiler') return 'distribution';
    if (
        k === 'groupchat' ||
        k === 'group_chat' ||
        k === 'chat' ||
        k === 'gruppenchat' ||
        k === 'teamschat' ||
        k === 'teams_chat'
    ) {
        return 'groupChat';
    }
    return 'unified';
}

/**
 * @param {unknown} kind
 */
export function schulOrgGroupKindLabel(kind) {
    const k = normalizeSchulOrgGroupKind(kind);
    if (k === 'security') return 'Sicherheitsgruppe';
    if (k === 'mailSecurity') return 'E-Mail-Sicherheitsgruppe';
    if (k === 'distribution') return 'Verteilerliste';
    if (k === 'groupChat') return 'Teams-Gruppenchat';
    return 'Microsoft 365-Gruppe';
}

/**
 * Bootstrap-Icon-Klasse für Entra-/Exchange-/Teams-Objekttypen.
 * @param {unknown} kind
 * @returns {string}
 */
export function schulOrgGroupKindIcon(kind) {
    const k = normalizeSchulOrgGroupKind(kind);
    if (k === 'security') return 'bi-shield-lock';
    if (k === 'mailSecurity') return 'bi-envelope-paper';
    if (k === 'distribution') return 'bi-mailbox';
    if (k === 'groupChat') return 'bi-chat-dots-fill';
    return 'bi-people-fill';
}

/** Icon für freigegebenes Postfach (parallel zu groupKind). */
export function schulOrgMailboxIcon() {
    return 'bi-envelope-at';
}

/**
 * Ob der Objekttyp eine Entra-/Exchange-Gruppe ist (anlegbar via Graph-Gruppen-API).
 * @param {unknown} kind
 */
export function schulOrgGroupKindIsEntraGroup(kind) {
    return normalizeSchulOrgGroupKind(kind) !== 'groupChat';
}

/**
 * @param {object|null|undefined} g Graph-Gruppe
 * @returns {SchulOrgGroupKind}
 */
export function inferSchulOrgGroupKind(g) {
    if (!g || typeof g !== 'object') return 'unified';
    const types = Array.isArray(g.groupTypes) ? g.groupTypes : [];
    if (types.indexOf('Unified') !== -1) return 'unified';
    if (g.securityEnabled && g.mailEnabled) return 'mailSecurity';
    if (g.securityEnabled && !g.mailEnabled) return 'security';
    if (g.mailEnabled) return 'distribution';
    return 'unified';
}

/**
 * Vorschlag für Graph-mailNickname aus Bereichsname.
 * @param {string} displayName
 */
export function suggestSchulOrgMailNickname(displayName) {
    let s = normStr(displayName)
        .toLowerCase()
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss');
    try {
        s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    } catch {
        /* ignore */
    }
    s = s
        .replace(/[^a-z0-9._-]+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^[-._]+|[-._]+$/g, '');
    if (!s) s = 'bereich';
    return s.slice(0, 60);
}

/**
 * Lokale Platzhalter-Mail (manuell ohne M365) – kein Graph-User-Match.
 * @param {string} email
 */
export function isSchulOrgLocalPlaceholderEmail(email) {
    const em = normStr(email).toLowerCase();
    return !em || em.indexOf('@') === -1 || em.endsWith('@local.schul-org');
}

/**
 * Person hat eine echte M365-fähige E-Mail (kein lokaler Platzhalter).
 * @param {{ email?: string }|null|undefined} person
 */
export function personHasM365UserLink(person) {
    const em = normStr(person && person.email).toLowerCase();
    if (!em || em.indexOf('@') === -1) return false;
    return !isSchulOrgLocalPlaceholderEmail(em);
}

/**
 * Chat-ID für KlassenChat-Kachel aus live settings.classChats (oder node.meta).
 * @param {object|null|undefined} node
 * @param {object|null|undefined} settings
 */
export function resolveClassChatIdForNode(node, settings) {
    if (!isKlassenchatClassNode(node)) return '';
    const metaId = normStr(node && node.meta && node.meta.chatId);
    if (metaId) return metaId;
    const code =
        canonicalSchulOrgClassCode(node && node.refKey) ||
        canonicalSchulOrgClassCode(node && node.label) ||
        '';
    if (!code) return '';
    const hit = classChatMetaByKlasse(settings).get(code);
    return hit ? normStr(hit.chatId) : '';
}

/**
 * Ob ein Knoten mit M365-Gruppe, Postfach und/oder Teams-Gruppenchat verknüpft ist.
 * @param {object|null|undefined} node
 * @param {string} [resolvedGroupId] – nur nutzen, wenn am Knoten nicht explizit groupMode=none
 * @param {string} [resolvedChatId] – live aus classChats
 * @returns {{
 *   hasGroup: boolean,
 *   hasMailbox: boolean,
 *   hasChat: boolean,
 *   showChatMark: boolean,
 *   groupId: string,
 *   chatId: string,
 *   groupKind: SchulOrgGroupKind,
 *   displayKind: SchulOrgGroupKind|null,
 *   linked: boolean
 * }}
 */
export function nodeM365LinkFlags(node, resolvedGroupId, resolvedChatId) {
    const m365 = node && node.m365 && typeof node.m365 === 'object' ? node.m365 : null;
    const norm = m365 ? normalizeSchulOrgM365(m365) : null;
    const bindings = norm ? norm.bindings : emptySchulOrgM365Bindings();
    const isChatArea = isKlassenchatClassNode(node);
    const resolved = normStr(resolvedGroupId);
    const entraKinds = ['unified', 'security', 'mailSecurity', 'distribution'];

    /** @type {Array<{ kind: string, icon: string, title: string }>} */
    const marks = [];
    let gid = '';
    let primaryGroupKind = /** @type {SchulOrgGroupKind} */ ('unified');

    entraKinds.forEach(function (kind) {
        const b = bindings[kind] || normalizeSchulOrgM365Binding(null, kind);
        let id = '';
        if (b.mode === 'none') {
            id = '';
        } else if (b.mode === 'create') {
            id = normStr(b.id);
        } else if (b.mode === 'match') {
            // Eigene ID, sonst geerbte Auflösung nur für unified (Klassengruppen-Match)
            id = normStr(b.id) || (kind === 'unified' ? resolved : '');
        }
        if (id) {
            marks.push({
                kind: kind,
                icon: schulOrgM365ObjectKindIcon(kind),
                title: schulOrgM365ObjectKindLabel(kind) + ' verknüpft'
            });
            if (!gid) {
                gid = id;
                primaryGroupKind = /** @type {SchulOrgGroupKind} */ (kind);
            }
        }
    });

    // Ohne m365-Block: Tenant-/Klassen-Auflösung als unified
    if (!gid && !m365 && resolved) {
        gid = resolved;
        primaryGroupKind = 'unified';
        marks.push({
            kind: 'unified',
            icon: schulOrgM365ObjectKindIcon('unified'),
            title: schulOrgM365ObjectKindLabel('unified') + ' verknüpft'
        });
    }

    const hasGroup = !!gid;
    const mb = bindings.mailbox || normalizeSchulOrgM365Binding(null, 'mailbox');
    const hasMailbox =
        (mb.mode === 'match' || mb.mode === 'create') && !!(normStr(mb.upn) || normStr(mb.id));
    if (hasMailbox) {
        marks.push({
            kind: 'mailbox',
            icon: schulOrgM365ObjectKindIcon('mailbox'),
            title: 'Freigegebenes Postfach verknüpft'
        });
    }

    const chatBind = bindings.groupChat || normalizeSchulOrgM365Binding(null, 'groupChat');
    const metaChatId = normStr(node && node.meta && node.meta.chatId);
    const chatId =
        metaChatId ||
        normStr(chatBind.id) ||
        normStr(resolvedChatId);
    const hasChat = !!chatId;
    const showChatMark =
        hasChat ||
        isChatArea ||
        chatBind.mode === 'match' ||
        chatBind.mode === 'create';
    if (showChatMark) {
        marks.push({
            kind: 'groupChat',
            icon: schulOrgM365ObjectKindIcon('groupChat'),
            title: hasChat
                ? 'Teams-Gruppenchat verknüpft'
                : 'Teams-Gruppenchat (KlassenChat)'
        });
    }

    let displayKind = /** @type {SchulOrgGroupKind|null} */ (null);
    if (showChatMark) displayKind = 'groupChat';
    else if (hasGroup) displayKind = primaryGroupKind;

    return {
        hasGroup: hasGroup,
        hasMailbox: hasMailbox,
        hasChat: hasChat,
        showChatMark: showChatMark,
        groupId: gid,
        chatId: chatId,
        groupKind: primaryGroupKind,
        displayKind: displayKind,
        marks: marks,
        bindings: bindings,
        linked: hasGroup || hasMailbox || hasChat || showChatMark
    };
}

/**
 * Bereiche/Knoten, die dieselbe M365-Gruppe bereits nutzen (Landkarte).
 * @param {unknown} layout
 * @param {string} groupId
 * @param {{ excludeNodeId?: string }} [opts]
 * @returns {Array<{ id: string, label: string, refKind: string, type: string }>}
 */
export function findSchulOrgNodesUsingM365Group(layout, groupId, opts) {
    const gid = normStr(groupId);
    if (!gid) return [];
    const excludeId = normStr(opts && opts.excludeNodeId);
    const base = normalizeSchulOrganigramm(layout);
    /** @type {Array<{ id: string, label: string, refKind: string, type: string }>} */
    const out = [];
    base.nodes.forEach(function (n) {
        if (!n || !n.id) return;
        if (excludeId && String(n.id) === excludeId) return;
        const m365 = n.m365 && typeof n.m365 === 'object' ? n.m365 : null;
        let used = '';
        if (m365) {
            const norm = normalizeSchulOrgM365(m365);
            const bindings = norm && norm.bindings ? norm.bindings : {};
            ['unified', 'security', 'mailSecurity', 'distribution'].forEach(function (kind) {
                if (used) return;
                const b = bindings[kind];
                if (b && b.mode !== 'none' && normStr(b.id) === gid) used = gid;
            });
            if (!used) {
                const mode = normStr(m365.groupMode).toLowerCase() || 'none';
                if (mode !== 'none') used = normStr(m365.groupId);
            }
        }
        if (!used && n.meta && typeof n.meta === 'object') {
            used = normStr(n.meta.graphGroupId);
        }
        if (!used || used !== gid) return;
        out.push({
            id: String(n.id),
            label: normStr(n.label) || String(n.id),
            refKind: normStr(n.refKind),
            type: normStr(n.type)
        });
    });
    return out;
}

/**
 * Bereiche, die dasselbe freigegebene Postfach bereits nutzen.
 * @param {unknown} layout
 * @param {string} mailboxUpnOrId
 * @param {{ excludeNodeId?: string }} [opts]
 * @returns {Array<{ id: string, label: string, refKind: string, type: string }>}
 */
export function findSchulOrgNodesUsingM365Mailbox(layout, mailboxUpnOrId, opts) {
    const key = normStr(mailboxUpnOrId).toLowerCase();
    if (!key) return [];
    const excludeId = normStr(opts && opts.excludeNodeId);
    const base = normalizeSchulOrganigramm(layout);
    /** @type {Array<{ id: string, label: string, refKind: string, type: string }>} */
    const out = [];
    base.nodes.forEach(function (n) {
        if (!n || !n.id) return;
        if (excludeId && String(n.id) === excludeId) return;
        const m365 = n.m365 && typeof n.m365 === 'object' ? n.m365 : null;
        if (!m365) return;
        const mode = normStr(m365.mailboxMode).toLowerCase();
        if (mode === 'none' || !mode) return;
        const upn = normStr(m365.mailboxUpn).toLowerCase();
        const mid = normStr(m365.mailboxId);
        if (upn !== key && mid !== key && mid.toLowerCase() !== key) return;
        out.push({
            id: String(n.id),
            label: normStr(n.label) || String(n.id),
            refKind: normStr(n.refKind),
            type: normStr(n.type)
        });
    });
    return out;
}

/**
 * Kurzer Anzeigetext für Konflikt-Hinweise (z. B. „KlassenChat ‚1A‘“).
 * @param {{ label?: string, refKind?: string }} node
 */
export function schulOrgM365ConflictLabel(node) {
    const label = normStr(node && node.label) || 'Bereich';
    const rk = normStr(node && node.refKind).toLowerCase();
    if (rk === 'klassenchatclass' || rk === 'klassenchatjahrgang') return 'KlassenChat „' + label + '“';
    if (rk === 'class' || rk === 'jahrgang') return 'Klasse/Jahrgang „' + label + '“';
    if (rk === 'arge' || rk === 'argeleiter') return 'ARGE „' + label + '“';
    if (rk === 'elternclass' || rk === 'elternjahrgang') return 'Eltern „' + label + '“';
    if (rk === 'sammelgruppe') return 'Sammelbereich „' + label + '“';
    return 'Bereich „' + label + '“';
}

/**
 * M365-Bindung einer Area setzen/entfernen.
 * @param {unknown} layout
 * @param {string} areaId
 * @param {object|null} m365Patch – null = entfernen
 */
/**
 * Meta nur als flache, JSON-sichere Felder übernehmen (keine DOM-/Zyklus-Refs).
 * @param {object} raw
 * @returns {object}
 */
function sanitizeSchulOrgNodeMeta(raw) {
    const src = raw && typeof raw === 'object' ? raw : {};
    /** @type {Record<string, unknown>} */
    const out = {};
    Object.keys(src).forEach(function (key) {
        const k = String(key || '');
        if (!k || k.charAt(0) === '_' || k === '__proto__' || k === 'constructor') return;
        const v = src[k];
        if (v == null) return;
        const t = typeof v;
        if (t === 'string' || t === 'number' || t === 'boolean') {
            out[k] = v;
            return;
        }
        if (Array.isArray(v)) {
            const arr = [];
            for (let i = 0; i < Math.min(v.length, 200); i += 1) {
                const item = v[i];
                const it = typeof item;
                if (it === 'string' || it === 'number' || it === 'boolean') arr.push(item);
                else if (item && it === 'object') {
                    try {
                        arr.push(JSON.parse(JSON.stringify(item)));
                    } catch {
                        /* skip non-serializable */
                    }
                }
            }
            if (arr.length) out[k] = arr;
            return;
        }
        if (t === 'object') {
            try {
                out[k] = JSON.parse(JSON.stringify(v));
            } catch {
                /* skip */
            }
        }
    });
    return out;
}

export function setSchulOrganigrammAreaM365(layout, areaId, m365Patch) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(areaId);
    const node = base.nodes.find(function (n) {
        return n.id === id && isAreaNode(n);
    });
    if (!node) return base;
    if (m365Patch == null) {
        delete node.m365;
        return normalizeSchulOrganigramm(base);
    }
    const p = m365Patch && typeof m365Patch === 'object' ? m365Patch : {};
    if (p.bindings && typeof p.bindings === 'object') {
        node.m365 = normalizeSchulOrgM365({ bindings: p.bindings });
    } else {
        // Legacy-Patch: ein groupKind + Postfach.
        // Bei explizitem groupKind andere Entra-Slots leeren (sonst bleibt z. B. unified create
        // aktiv und wird als primaryKind bevorzugt).
        const merged = normalizeSchulOrgM365(node.m365 || {}) || {
            bindings: emptySchulOrgM365Bindings()
        };
        const groupKind = normalizeSchulOrgGroupKind(p.groupKind || merged.groupKind || 'unified');
        const groupMode = normStr(p.groupMode) || 'none';
        const mailboxMode = normStr(p.mailboxMode) || 'none';
        if (p.groupMode !== undefined || p.groupId !== undefined || p.groupKind !== undefined) {
            if (p.groupKind !== undefined) {
                ['unified', 'security', 'mailSecurity', 'distribution', 'groupChat'].forEach(
                    function (kind) {
                        if (kind === groupKind) return;
                        if (!merged.bindings[kind]) return;
                        merged.bindings[kind] = normalizeSchulOrgM365Binding(null, kind);
                    }
                );
            }
            merged.bindings[groupKind] = normalizeSchulOrgM365Binding(
                {
                    mode: groupMode === 'match' || groupMode === 'create' ? groupMode : 'none',
                    id: p.groupId,
                    label: p.groupLabel,
                    mail: p.groupMail
                },
                groupKind
            );
        }
        if (p.mailboxMode !== undefined || p.mailboxUpn !== undefined || p.mailboxId !== undefined) {
            merged.bindings.mailbox = normalizeSchulOrgM365Binding(
                {
                    mode: mailboxMode === 'match' || mailboxMode === 'create' ? mailboxMode : 'none',
                    id: p.mailboxId,
                    label: p.mailboxLabel,
                    upn: p.mailboxUpn
                },
                'mailbox'
            );
        }
        node.m365 = normalizeSchulOrgM365({ bindings: merged.bindings });
    }
    return normalizeSchulOrganigramm(base);
}

/**
 * Klassengruppen-Matches (Werkzeug Klassengruppen / classTeams + Register-Abgleich)
 * auf Klassen-Bereiche der Landkarte legen. Explizit gelöste Verknüpfungen
 * (groupMode=none) bleiben unangetastet.
 *
 * @param {unknown} layout
 * @param {object} settings – Stammdaten mit classes[]
 * @param {{
 *   classTeams?: object[],
 *   classGroupMatchByKey?: Record<string, { groupId?: string, notFound?: boolean, displayName?: string }>
 * }} [linkCtx]
 * @returns {{ layout: object, linked: number, updated: number }}
 */
export function applyClassGroupLinksToSchulOrganigramm(layout, settings, linkCtx) {
    const base = normalizeSchulOrganigramm(layout);
    const s = settings && typeof settings === 'object' ? settings : {};
    const ctx = linkCtx && typeof linkCtx === 'object' ? linkCtx : {};
    const classTeams = Array.isArray(ctx.classTeams) ? ctx.classTeams : [];
    const matchMap =
        ctx.classGroupMatchByKey && typeof ctx.classGroupMatchByKey === 'object'
            ? ctx.classGroupMatchByKey
            : {};
    const classes = Array.isArray(s.classes) ? s.classes : [];
    let linked = 0;
    let updated = 0;

    base.nodes.forEach(function (n) {
        if (!isClassAreaNode(n) || !cleanRefToken(n.refKey)) return;
        const mode = n.m365 ? normStr(n.m365.groupMode).toLowerCase() : '';
        // Nutzer hat die Gruppe bewusst gelöst
        if (mode === 'none') return;

        const code =
            canonicalSchulOrgClassCode(n.refKey) ||
            canonicalSchulOrgClassCode(n.label) ||
            normCode(n.refKey);
        if (!code) return;
        const cls =
            classes.find(function (c) {
                const cc =
                    canonicalSchulOrgClassCode(c && c.code) ||
                    canonicalSchulOrgClassCode(c && c.name) ||
                    normCode(c && c.code);
                return cc === code;
            }) || {
                code: code,
                name: normStr(n.label) || code,
                year: normStr(n.meta && n.meta.year)
            };
        const gid = resolveClassGraphGroupId(cls, classTeams, matchMap);
        if (!gid) return;
        linked += 1;
        const prev = n.m365 && typeof n.m365 === 'object' ? n.m365 : {};
        const prevGid = normStr(prev.groupId);
        const team = classTeams.find(function (t) {
            return normStr(t && t.graphGroupId) === gid;
        });
        const matchEntry = matchMap[code] || matchMap[normCode(cls.code)] || matchMap[normCode(cls.name)];
        const label =
            normStr(prev.groupLabel) ||
            normStr(team && team.displayName) ||
            normStr(matchEntry && matchEntry.displayName) ||
            'Klassengruppe';
        if (prevGid === gid && mode === 'match') {
            if (!normStr(prev.groupLabel) && label) {
                n.m365 = Object.assign({}, prev, { groupLabel: label });
                updated += 1;
            }
            return;
        }
        n.m365 = {
            groupMode: 'match',
            groupKind: normalizeSchulOrgGroupKind(prev.groupKind || 'unified'),
            groupId: gid,
            groupLabel: label,
            groupMail: prev.groupMail ? normStr(prev.groupMail).toLowerCase() : null,
            mailboxMode:
                normStr(prev.mailboxMode).toLowerCase() === 'match' ||
                normStr(prev.mailboxMode).toLowerCase() === 'create'
                    ? normStr(prev.mailboxMode).toLowerCase()
                    : 'none',
            mailboxUpn: prev.mailboxUpn ? normStr(prev.mailboxUpn).toLowerCase() : null,
            mailboxLabel: prev.mailboxLabel ? normStr(prev.mailboxLabel) : null,
            mailboxId: prev.mailboxId ? normStr(prev.mailboxId) : null
        };
        updated += 1;
    });

    return {
        layout: normalizeSchulOrganigramm(base),
        linked: linked,
        updated: updated
    };
}

/**
 * Direkten Person-Slot unter einer Area anlegen/aktualisieren.
 * @param {unknown} layout
 * @param {string} areaId
 * @param {{ email: string, name?: string }} person
 */
export function upsertSchulOrganigrammAreaPerson(layout, areaId, person) {
    const base = normalizeSchulOrganigramm(layout);
    const parentId = normStr(areaId);
    const parent = base.nodes.find(function (n) {
        return n.id === parentId && isAreaNode(n);
    });
    if (!parent) return base;
    const email = normStr(person && person.email).toLowerCase();
    if (!email || email.indexOf('@') === -1) return base;
    const name = normStr(person && person.name) || email;
    const existing = base.nodes.find(function (n) {
        return (
            n.type === NODE.PERSON &&
            normStr(n.refKind) === 'adminPerson' &&
            normStr(n.refKey).toLowerCase() === email
        );
    });
    if (existing) {
        existing.parentId = parentId;
        existing.label = name;
        existing.refKind = 'adminPerson';
        existing.refKey = email;
        return normalizeSchulOrganigramm(base);
    }
    const maxOrder = base.nodes.reduce(function (m, n) {
        return String(n.parentId || '') === parentId
            ? Math.max(m, typeof n.order === 'number' ? n.order : 0)
            : m;
    }, 0);
    base.nodes.push({
        id: personNodeIdFromEmail(email),
        type: NODE.PERSON,
        label: name,
        parentId: parentId,
        order: maxOrder + 1,
        tone: 'staff',
        refKind: 'adminPerson',
        refKey: email
    });
    return normalizeSchulOrganigramm(base);
}

/**
 * Person-Slot entfernen (nur Layout-Knoten, nicht Stammdaten-Rolle).
 * @param {unknown} layout
 * @param {string} email
 * @param {string} [areaId] – wenn gesetzt, nur unter dieser Area
 */
export function removeSchulOrganigrammAreaPerson(layout, email, areaId) {
    const base = normalizeSchulOrganigramm(layout);
    const em = normStr(email).toLowerCase();
    const aid = normStr(areaId);
    if (!em) return base;
    base.nodes = base.nodes.filter(function (n) {
        if (n.type !== NODE.PERSON || normStr(n.refKind) !== 'adminPerson') return true;
        if (normStr(n.refKey).toLowerCase() !== em) return true;
        if (aid && String(n.parentId || '') !== aid) return true;
        return false;
    });
    return normalizeSchulOrganigramm(base);
}

function listCustomAudienceGroups(settings) {
    const groups = normalizeVerwaltungAudienceGroups(
        settings && settings.verwaltungAudienceGroups
    );
    return groups.filter(function (g) {
        return g && !g.builtin && normStr(g.id).indexOf('va-') === 0;
    });
}

/**
 * @param {unknown} raw
 */
export function normalizeSchulOrganigramm(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const nodesIn = Array.isArray(o.nodes) ? o.nodes : [];
    const seen = new Set();
    const nodes = [];
    let rootAreaIndex = 0;
    nodesIn.forEach(function (n, idx) {
        if (!n || typeof n !== 'object') return;
        let id = normStr(n.id);
        if (!id || seen.has(id)) id = newId('og');
        seen.add(id);
        let type = canonicalNodeType(n.type);
        if (type === 'group' && (n.parentId == null || n.parentId === '')) {
            type = NODE.AREA;
        }
        const parentId = n.parentId == null || n.parentId === '' ? null : normStr(n.parentId);
        const node = {
            id: id,
            type: type,
            label: normStr(n.label || n.name) || 'Knoten',
            parentId: parentId,
            order: typeof n.order === 'number' ? n.order : idx,
            tone: normStr(n.tone) || '',
            refKind: normStr(n.refKind) || 'none',
            refKey: normStr(n.refKey || n.roleCode || n.code)
        };
        const cx = parseSchulOrgCanvasCoord(n.x);
        const cy = parseSchulOrgCanvasCoord(n.y);
        const rkLower = normStr(n.refKind).toLowerCase();
        // Schüler-/Eltern-/KlassenChat-Jahrgänge: immer Auto-Pack (keine freien Orte/Größen)
        const isJahrgang =
            rkLower === 'jahrgang' ||
            rkLower === 'elternjahrgang' ||
            rkLower === 'klassenchatjahrgang' ||
            /^sec_jahrgang_/i.test(id) ||
            /^sec_eltern_jahrgang_/i.test(id) ||
            /^sec_klassenchat_jahrgang_/i.test(id);
        // Freie Canvas-Orte: nur Root-Areas. Jahrgänge/Klassen immer Auto-Pack (sonst Überlappung).
        if (isAreaNode(node) && parentId == null) {
            if (cx != null && cy != null) {
                node.x = cx;
                node.y = cy;
            } else {
                const slot = defaultCanvasSlot(rootAreaIndex++);
                node.x = slot.x;
                node.y = slot.y;
            }
        } else if (
            !isJahrgang &&
            cx != null &&
            cy != null &&
            !(parentId != null && (cx > 400 || cy > 300))
        ) {
            // Optionale Relativ-Orte für verschachtelte Custom-Areas / Leaves
            node.x = cx;
            node.y = cy;
        }
        const cw = parseSchulOrgCanvasCoord(n.w);
        const ch = parseSchulOrgCanvasCoord(n.h);
        // Fenstergröße: vor allem Root-Areas; Jahrgänge nie (Auto-Pack)
        if (!isJahrgang && cw != null && ch != null) {
            const clamped = clampSchulOrganigrammAreaSize(cw, ch);
            if (clamped.w != null) node.w = clamped.w;
            if (clamped.h != null) node.h = clamped.h;
        } else if (!isJahrgang && isAreaNode(node) && parentId == null) {
            if (cw != null) {
                const clamped = clampSchulOrganigrammAreaSize(cw, SCHUL_ORG_AREA_MIN_H);
                if (clamped.w != null) node.w = clamped.w;
            }
            if (ch != null) {
                const clamped = clampSchulOrganigrammAreaSize(SCHUL_ORG_AREA_MIN_W, ch);
                if (clamped.h != null) node.h = clamped.h;
            }
        }
        if (n.collapsed === true) node.collapsed = true;
        if (n.m365 && typeof n.m365 === 'object') {
            node.m365 = normalizeSchulOrgM365(n.m365);
        }
        if (n.externalIds && typeof n.externalIds === 'object') {
            const ext = {};
            ['bip', 'sokrates', 'webuntis', 'graphGroupId'].forEach(function (key) {
                const v = normStr(n.externalIds[key]);
                if (v) ext[key] = v;
            });
            if (Object.keys(ext).length) node.externalIds = ext;
        }
        if (n.meta && typeof n.meta === 'object') {
            node.meta = sanitizeSchulOrgNodeMeta(n.meta);
        }
        nodes.push(node);
    });
    nodes.sort(function (a, b) {
        return (a.order || 0) - (b.order || 0);
    });
    const viewIn = o.view && typeof o.view === 'object' ? o.view : {};
    const k = Number(viewIn.k);
    const view = {
        x: Number.isFinite(Number(viewIn.x)) ? Number(viewIn.x) : 0,
        y: Number.isFinite(Number(viewIn.y)) ? Number(viewIn.y) : 0,
        k: Number.isFinite(k) && k > 0.15 && k < 4 ? k : 1
    };
    const removedAreaIds = schulOrganigrammRemovedAreaIds(o);
    return {
        version: SCHUL_ORG_VERSION,
        view: view,
        nodes: nodes,
        removedAreaIds: removedAreaIds
    };
}

/**
 * Area unter Parent legen / Reihenfolge setzen.
 * @param {unknown} layout
 * @param {string} areaId
 * @param {string|null} newParentId
 * @param {string} [beforeId]
 */
export function moveSchulOrganigrammArea(layout, areaId, newParentId, beforeId) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(areaId);
    const node = base.nodes.find(function (n) {
        return n.id === id;
    });
    if (!node || !isAreaNode(node)) return base;
    const parentId = newParentId == null || newParentId === '' ? null : String(newParentId);
    if (parentId === id) return base;
    // Zyklus verhindern
    if (parentId) {
        let walk = parentId;
        const guard = new Set();
        while (walk) {
            if (walk === id) return base;
            if (guard.has(walk)) break;
            guard.add(walk);
            const p = base.nodes.find(function (n) {
                return n.id === walk;
            });
            walk = p && p.parentId != null ? String(p.parentId) : '';
        }
        const parent = base.nodes.find(function (n) {
            return n.id === parentId;
        });
        if (!parent || !isAreaNode(parent)) return base;
    }
    node.parentId = parentId;
    // Unterbereich: keine freien x/y – Canvas packt unter dem Elternkopf
    if (parentId) {
        delete node.x;
        delete node.y;
    }
    const before = normStr(beforeId);
    const siblings = base.nodes
        .filter(function (n) {
            const np = n.parentId == null || n.parentId === '' ? null : String(n.parentId);
            return np === parentId && n.id !== id;
        })
        .sort(function (a, b) {
            return (a.order || 0) - (b.order || 0);
        });
    const ordered = [];
    let inserted = false;
    siblings.forEach(function (s) {
        if (before && s.id === before) {
            ordered.push(node);
            inserted = true;
        }
        ordered.push(s);
    });
    if (!inserted) ordered.push(node);
    ordered.forEach(function (n, i) {
        n.order = i;
    });
    return normalizeSchulOrganigramm(base);
}

/**
 * @param {{ nodes?: object[] }} layout
 * @param {string|null} parentId
 */
export function childrenOf(layout, parentId) {
    const pid = parentId == null || parentId === '' ? null : String(parentId);
    return normalizeSchulOrganigramm(layout)
        .nodes.filter(function (n) {
            const np = n.parentId == null || n.parentId === '' ? null : String(n.parentId);
            return np === pid;
        })
        .sort(function (a, b) {
            return (a.order || 0) - (b.order || 0);
        });
}

/**
 * Layout aus Stammdaten ableiten (Phase 1: Seed / Sync fehlender Bereiche).
 * @param {object} settings – normalizeSettings / load()-Ergebnis
 */
export function seedSchulOrganigrammFromStammdaten(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const nodes = [];
    let order = 0;

    /**
     * @param {string} id
     * @param {string} label
     * @param {string} tone
     * @param {{ refKind?: string, refKey?: string }} [ref]
     */
    function pushSection(id, label, tone, ref) {
        nodes.push({
            id: id,
            type: NODE.SECTION,
            label: label,
            parentId: null,
            order: order++,
            tone: tone || 'custom',
            refKind: (ref && ref.refKind) || 'none',
            refKey: (ref && ref.refKey) || ''
        });
        return id;
    }

    const secLead = pushSection('sec_schulleitung', 'Schulleitung', 'lead', {
        refKind: 'audience',
        refKey: 'schulleitung'
    });
    const secStaff = pushSection('sec_verwaltung', 'Verwaltung (Personal)', 'staff', {
        refKind: 'audience',
        refKey: 'verwaltung'
    });
    listCustomAudienceGroups(s).forEach(function (g) {
        const sid = audienceSectionId(g.id);
        if (!sid) return;
        pushSection(sid, g.label, 'staff', { refKind: 'audience', refKey: g.id });
    });
    const teachers = Array.isArray(s.teachers) ? s.teachers : [];
    const students = Array.isArray(s.students) ? s.students : [];
    nodes.push({
        id: SCHUL_ORG_SEC_LEHRER,
        type: NODE.SECTION,
        label: 'Lehrkräfte',
        parentId: null,
        order: order++,
        tone: 'lehrer',
        refKind: 'sammelgruppe',
        refKey: 'lehrer',
        meta: { peopleCount: teachers.length, board: true, sammelgruppe: 'lehrer' }
    });
    const classRowsRaw = Array.isArray(s.classes) ? s.classes : [];
    /** @type {Map<string, object>} kanonischer Code → Stammdatenzeile (Duplikate zusammenführen) */
    const classByCanon = new Map();
    classRowsRaw.forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        if (!classByCanon.has(code)) {
            classByCanon.set(code, Object.assign({}, c, { code: code }));
            return;
        }
        const prev = classByCanon.get(code);
        // Beste Meta behalten (KV, Jahr, Anzeigename)
        classByCanon.set(
            code,
            Object.assign({}, prev, {
                code: code,
                name: normStr(prev.name) || normStr(c && c.name) || code,
                year: normStr(prev.year) || normStr(c && c.year),
                headName: normStr(prev.headName) || normStr(c && c.headName),
                headEmail: normStr(prev.headEmail) || normStr(c && c.headEmail)
            })
        );
    });
    const classRows = Array.from(classByCanon.values());
    const classCount = classRows.length;
    nodes.push({
        id: SCHUL_ORG_SEC_SCHUELER,
        type: NODE.SECTION,
        label: 'Schüler',
        parentId: null,
        order: order++,
        tone: 'schueler',
        refKind: 'sammelgruppe',
        refKey: 'schueler',
        meta: {
            classCount: classCount,
            studentCount: students.length,
            board: true,
            sammelgruppe: 'schueler',
            nestClasses: true
        }
    });
    const secArge = pushSection('sec_arge', 'ARGEs / Fachschaften', 'arge');
    nodes.push({
        id: SCHUL_ORG_SEC_ARGE_LEITER,
        type: NODE.SECTION,
        label: 'ARGE-Leiter',
        parentId: secArge,
        order: order++,
        tone: 'arge',
        refKind: 'argeLeiter',
        refKey: 'argeLeiter',
        meta: { argeLeiterBoard: true, nestPeople: true }
    });
    nodes.push({
        id: SCHUL_ORG_SEC_KLASSENVORSTAND,
        type: NODE.SECTION,
        label: 'Klassenvorstände',
        parentId: null,
        order: order++,
        tone: 'lehrer',
        refKind: 'klassenvorstand',
        refKey: 'klassenvorstand',
        meta: { klassenvorstandBoard: true, nestPeople: true }
    });
    listSchulOrgClassFunctions(s).forEach(function (def) {
        nodes.push({
            id: classFunctionBoardId(def.id),
            type: NODE.SECTION,
            label: def.labelPlural,
            parentId: null,
            order: order++,
            tone: 'schueler',
            refKind: 'classFunction',
            refKey: def.id,
            meta: {
                functionAggregate: true,
                classFunctionBoard: true,
                classFunctionId: def.id,
                nestPeople: true
            }
        });
    });
    listSchulOrgArgeFunctions(s).forEach(function (def) {
        nodes.push({
            id: argeFunctionBoardId(def.id),
            type: NODE.SECTION,
            label: def.labelPlural,
            parentId: secArge,
            order: order++,
            tone: 'arge',
            refKind: 'argeFunction',
            refKey: def.id,
            meta: {
                functionAggregate: true,
                argeFunctionBoard: true,
                argeFunctionId: def.id,
                nestPeople: true
            }
        });
    });
    // Schüler → Jahrgänge → Klassen
    /** @type {Map<string, string>} */
    const jahrgangParentByStufe = new Map();
    const stufen = [];
    const stufeSeen = new Set();
    classRows.forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        const stufe = schulstufeFromClassCode(code);
        if (stufeSeen.has(stufe)) return;
        stufeSeen.add(stufe);
        stufen.push(stufe);
    });
    stufen.sort(function (a, b) {
        if (a === SCHUL_ORG_JAHRGANG_SONST) return 1;
        if (b === SCHUL_ORG_JAHRGANG_SONST) return -1;
        return Number(a) - Number(b) || String(a).localeCompare(String(b), 'de');
    });
    stufen.forEach(function (stufe, idx) {
        const jid = jahrgangAreaId(stufe);
        jahrgangParentByStufe.set(stufe, jid);
        nodes.push({
            id: jid,
            type: NODE.SECTION,
            label: jahrgangAreaLabel(stufe),
            parentId: SCHUL_ORG_SEC_SCHUELER,
            order: order++,
            tone: 'class',
            refKind: 'jahrgang',
            refKey: stufe,
            // Nebeneinander (Canvas packt Jahrgänge unter Schüler horizontal)
            x: 12 + idx * 312,
            y: 8,
            meta: { jahrgang: stufe, nestClasses: true }
        });
    });

    // Eltern → Jahrgänge → Klassen (Erziehungsberechtigte aus Schüler-Stammdaten)
    const guardianCount = countUniqueGuardiansInSettings(s);
    nodes.push({
        id: SCHUL_ORG_SEC_ELTERN,
        type: NODE.SECTION,
        label: 'Eltern',
        parentId: null,
        order: order++,
        tone: 'council',
        refKind: 'sammelgruppe',
        refKey: 'eltern',
        meta: {
            classCount: classCount,
            guardianCount: guardianCount,
            board: true,
            sammelgruppe: 'eltern',
            nestClasses: true
        }
    });
    /** @type {Map<string, string>} */
    const elternJahrgangParentByStufe = new Map();
    stufen.forEach(function (stufe) {
        const jid = elternJahrgangAreaId(stufe);
        elternJahrgangParentByStufe.set(stufe, jid);
        nodes.push({
            id: jid,
            type: NODE.SECTION,
            label: jahrgangAreaLabel(stufe),
            parentId: SCHUL_ORG_SEC_ELTERN,
            order: order++,
            tone: 'council',
            refKind: 'elternJahrgang',
            refKey: stufe,
            // Keine festen Orte – Canvas packt Jahrgänge bei Fensterbreite um
            meta: { jahrgang: stufe, nestClasses: true, audience: 'eltern' }
        });
    });
    classRows.forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        const stufe = schulstufeFromClassCode(code);
        const parentId =
            elternJahrgangParentByStufe.get(stufe) || elternJahrgangAreaId(stufe);
        const eid = elternClassNodeId(code);
        if (!eid) return;
        nodes.push({
            id: eid,
            type: NODE.SECTION,
            label: normStr(c && c.name) || 'Klasse ' + code,
            parentId: parentId,
            order: order++,
            tone: 'council',
            refKind: 'elternClass',
            refKey: code,
            meta: {
                year: normStr(c && c.year),
                jahrgang: stufe,
                audience: 'eltern',
                nestPeople: true
            }
        });
    });

    // KlassenChats → Jahrgänge → Klassen (Lehrkräfte-Gruppenchats aus Belegung/KV)
    const chatMetaByKlasse = classChatMetaByKlasse(s);
    nodes.push({
        id: SCHUL_ORG_SEC_KLASSENCHATS,
        type: NODE.SECTION,
        label: 'KlassenChats',
        parentId: null,
        order: order++,
        tone: 'lehrer',
        refKind: 'sammelgruppe',
        refKey: 'klassenchats',
        meta: {
            classCount: classCount,
            chatCount: chatMetaByKlasse.size,
            board: true,
            sammelgruppe: 'klassenchats',
            nestClasses: true
        }
    });
    /** @type {Map<string, string>} */
    const klassenchatJahrgangParentByStufe = new Map();
    stufen.forEach(function (stufe) {
        const jid = klassenchatJahrgangAreaId(stufe);
        klassenchatJahrgangParentByStufe.set(stufe, jid);
        nodes.push({
            id: jid,
            type: NODE.SECTION,
            label: jahrgangAreaLabel(stufe),
            parentId: SCHUL_ORG_SEC_KLASSENCHATS,
            order: order++,
            tone: 'lehrer',
            refKind: 'klassenchatJahrgang',
            refKey: stufe,
            // Keine festen Orte – Canvas packt Jahrgänge bei Fensterbreite um
            meta: { jahrgang: stufe, nestClasses: true, audience: 'klassenchats' }
        });
    });
    classRows.forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        const stufe = schulstufeFromClassCode(code);
        const parentId =
            klassenchatJahrgangParentByStufe.get(stufe) || klassenchatJahrgangAreaId(stufe);
        const kid = klassenchatClassNodeId(code);
        if (!kid) return;
        const chatMeta = chatMetaByKlasse.get(code) || null;
        nodes.push({
            id: kid,
            type: NODE.SECTION,
            label: normStr(c && c.name) || 'Klasse ' + code,
            parentId: parentId,
            order: order++,
            tone: 'lehrer',
            refKind: 'klassenchatClass',
            refKey: code,
            meta: {
                year: normStr(c && c.year),
                jahrgang: stufe,
                audience: 'klassenchats',
                nestPeople: true,
                chatId: chatMeta ? chatMeta.chatId : '',
                chatTopic: chatMeta ? chatMeta.topic : '',
                chatWebUrl: chatMeta ? chatMeta.webUrl : ''
            },
            // Intent: Klassenchat = Teams-Gruppenchat (Chat-ID liegt in meta)
            m365: {
                groupMode: 'none',
                groupKind: 'groupChat',
                groupId: null,
                groupLabel: null,
                groupMail: null,
                mailboxMode: 'none',
                mailboxUpn: null,
                mailboxLabel: null,
                mailboxId: null
            }
        });
    });

    // SGA ist eigener Root-Bereich (kein Wrapper „Schülervertretung & SGA“ mehr)

    const roles = Array.isArray(s.adminRoles) ? s.adminRoles : [];
    const admin = Array.isArray(s.admin) ? s.admin : [];
    roles.forEach(function (role) {
        const code = normCode(role && role.code);
        const name = normStr(role && role.name) || code;
        if (!code && !name) return;
        const tier = normStr(role && role.tier).toLowerCase();
        const parentId =
            tier === 'schulleitung' || /direktion|direktor/i.test(name + code) ? secLead : secStaff;
        const people = admin.filter(function (p) {
            const r = normStr(p && p.role).toLowerCase();
            return r === name.toLowerCase() || r === code.toLowerCase();
        });
        nodes.push({
            id: 'role_' + (code || newId('r')),
            type: NODE.ROLE,
            label: name,
            parentId: parentId,
            order: order++,
            tone: parentId === secLead ? 'lead' : 'staff',
            refKind: 'adminRole',
            refKey: code || name,
            meta: { peopleCount: people.length, seatMode: normStr(role && role.seatMode) }
        });
    });

    (Array.isArray(s.arges) ? s.arges : []).forEach(function (a) {
        const code = normCode(a && a.code);
        const name = normStr(a && a.name) || code;
        if (!code && !name) return;
        const headEmail = normStr(a && a.headEmail).toLowerCase();
        const teacher =
            headEmail.indexOf('@') !== -1
                ? teachers.find(function (t) {
                      return normStr(t && t.email).toLowerCase() === headEmail;
                  })
                : null;
        const subjects = Array.isArray(a && a.subjects) ? a.subjects.slice() : [];
        nodes.push({
            id: argeNodeId(code) || 'arge_' + newId('a'),
            type: NODE.SECTION,
            label: name,
            parentId: secArge,
            order: order++,
            tone: 'arge',
            refKind: 'arge',
            refKey: code || name,
            meta: {
                nestPeople: true,
                audience: 'arge',
                headName: normStr(a && a.headName),
                headEmail: headEmail,
                teacherLinked: !!teacher,
                teacherCode: teacher ? normCode(teacher.code) : '',
                subjects: subjects
            }
        });
    });

    /** @type {Set<string>} */
    const knownClassCodes = new Set();
    classRows.forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        knownClassCodes.add(code);
        const stufe = schulstufeFromClassCode(code);
        const parentId = jahrgangParentByStufe.get(stufe) || jahrgangAreaId(stufe);
        const headEmail = normStr(c && c.headEmail).toLowerCase();
        const teacher =
            headEmail.indexOf('@') !== -1
                ? teachers.find(function (t) {
                      return normStr(t && t.email).toLowerCase() === headEmail;
                  })
                : null;
        const fnSeats = normalizeClassFunctionsMap(c && c.functions);
        const classRep = fnSeats.klassensprecher || null;
        nodes.push({
            id: 'class_' + code,
            type: NODE.SECTION,
            label: normStr(c && c.name) || code,
            parentId: parentId,
            order: order++,
            tone: 'class',
            refKind: 'class',
            refKey: code,
            meta: {
                nestPeople: true,
                audience: 'schueler',
                year: normStr(c && c.year),
                jahrgang: stufe,
                headName: normStr(c && c.headName),
                headEmail: headEmail,
                teacherLinked: !!teacher,
                teacherCode: teacher ? normCode(teacher.code) : '',
                classRepName: classRep ? normStr(classRep.name) : '',
                classRepEmail: classRep ? normStr(classRep.email).toLowerCase() : ''
            }
        });
    });

    // Kursteams aus Unterrichtsbelegung → unter jede verknüpfte Klasse (Mehrklassen: mehrfach)
    const belegung = s.unterrichtsbelegung;
    kursteamPlacementsFromBelegung(belegung, classRows).forEach(function (place) {
        const cc = canonicalSchulOrgClassCode(place.classCode) || place.classCode;
        if (!knownClassCodes.has(cc)) return;
        nodes.push(kursteamNodeFromPlacement(Object.assign({}, place, { classCode: cc }), order++));
    });

    const sgaRows = (Array.isArray(s.sga) ? s.sga : []).filter(function (row) {
        return normStr(row && row.name) || normStr(row && row.email);
    });
    const sgaCounts = { student: 0, teacher: 0, external: 0 };
    sgaRows.forEach(function (row) {
        const sc = normStr(row && row.scope).toLowerCase();
        if (sc === 'teacher' || sc === 'lehrer') sgaCounts.teacher += 1;
        else if (sc === 'external' || sc === 'extern' || sc === 'eltern') sgaCounts.external += 1;
        else sgaCounts.student += 1;
    });
    const councilRows = (Array.isArray(s.studentCouncil) ? s.studentCouncil : []).filter(function (row) {
        return normStr(row && row.name) || normStr(row && row.email);
    });
    const svPeopleCount = Math.max(councilRows.length, sgaCounts.student);
    nodes.push({
        id: SCHUL_ORG_SGA_BOARD_ID,
        type: NODE.AREA,
        label: 'SGA',
        parentId: null,
        order: order++,
        tone: 'sga',
        refKind: 'sga',
        refKey: SCHUL_ORG_BOARD_REF,
        x: 12,
        y: 8,
        meta: {
            board: true,
            peopleCount: sgaRows.length,
            studentCount: sgaCounts.student,
            teacherCount: sgaCounts.teacher,
            externalCount: sgaCounts.external
        }
    });
    SCHUL_ORG_SGA_SCOPE_AREAS.forEach(function (scopeDef) {
        const count =
            scopeDef.scope === 'teacher'
                ? sgaCounts.teacher
                : scopeDef.scope === 'external'
                  ? sgaCounts.external
                  : Math.max(sgaCounts.student, svPeopleCount);
        nodes.push({
            id: scopeDef.id,
            type: NODE.AREA,
            label: scopeDef.label,
            parentId: SCHUL_ORG_SGA_BOARD_ID,
            order: order++,
            tone: scopeDef.tone,
            refKind: 'sgaScope',
            refKey: scopeDef.scope,
            x: scopeDef.x,
            y: scopeDef.y,
            meta: {
                board: true,
                sgaScope: scopeDef.scope,
                peopleCount: count,
                svSeats: scopeDef.scope === 'student' ? true : undefined
            }
        });
    });
    // Schülervertretung = die drei SGA-Schülerplätze mit Funktionen
    SCHUL_ORG_SV_SEATS.forEach(function (seat) {
        const person = personForSvSeat(s, seat.seat);
        nodes.push({
            id: seat.id,
            type: NODE.COUNCIL,
            label: seat.label,
            parentId: SCHUL_ORG_SGA_SV_AREA_ID,
            order: order++,
            tone: 'council',
            refKind: 'studentCouncil',
            refKey: seat.refKey,
            meta: {
                svSeat: seat.seat,
                peopleCount: person ? 1 : 0,
                personName: person ? person.name : '',
                personEmail: person ? person.email : ''
            }
        });
    });

    return normalizeSchulOrganigramm({ version: SCHUL_ORG_VERSION, nodes: nodes });
}

/**
 * SGA-Board + drei Scope-Bereiche sicherstellen (Migration / Sync).
 * @param {object} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammSgaStructure(layout, settings) {
    const base = normalizeSchulOrganigramm(layout);
    const s = settings && typeof settings === 'object' ? settings : {};
    const sgaRows = (Array.isArray(s.sga) ? s.sga : []).filter(function (row) {
        return normStr(row && row.name) || normStr(row && row.email);
    });
    const sgaCounts = { student: 0, teacher: 0, external: 0 };
    sgaRows.forEach(function (row) {
        const sc = normStr(row && row.scope).toLowerCase();
        if (sc === 'teacher' || sc === 'lehrer') sgaCounts.teacher += 1;
        else if (sc === 'external' || sc === 'extern' || sc === 'eltern') sgaCounts.external += 1;
        else sgaCounts.student += 1;
    });

    let board = base.nodes.find(function (n) {
        return n.id === SCHUL_ORG_SGA_BOARD_ID;
    });
    if (!board) {
        board = {
            id: SCHUL_ORG_SGA_BOARD_ID,
            type: NODE.AREA,
            label: 'SGA',
            parentId: null,
            order: 99,
            tone: 'sga',
            refKind: 'sga',
            refKey: SCHUL_ORG_BOARD_REF,
            x: 12,
            y: 8
        };
        base.nodes.push(board);
    } else {
        const keepM365 = board.m365;
        board.type = NODE.AREA;
        board.refKind = 'sga';
        board.refKey = SCHUL_ORG_BOARD_REF;
        board.tone = board.tone || 'sga';
        // Ein Root-Bereich „SGA“ (Stammdaten), kein Nest unter sec_vertretung
        board.parentId = null;
        board.label = 'SGA';
        if (keepM365) board.m365 = keepM365;
    }
    board.meta = Object.assign({}, board.meta || {}, {
        board: true,
        peopleCount: sgaRows.length,
        studentCount: sgaCounts.student,
        teacherCount: sgaCounts.teacher,
        externalCount: sgaCounts.external
    });

    SCHUL_ORG_SGA_SCOPE_AREAS.forEach(function (scopeDef) {
        let node = base.nodes.find(function (n) {
            return n.id === scopeDef.id;
        });
        const count =
            scopeDef.scope === 'teacher'
                ? sgaCounts.teacher
                : scopeDef.scope === 'external'
                  ? sgaCounts.external
                  : sgaCounts.student;
        if (!node) {
            node = {
                id: scopeDef.id,
                type: NODE.AREA,
                label: scopeDef.label,
                parentId: SCHUL_ORG_SGA_BOARD_ID,
                order: 100,
                tone: scopeDef.tone,
                refKind: 'sgaScope',
                refKey: scopeDef.scope,
                x: scopeDef.x,
                y: scopeDef.y
            };
            base.nodes.push(node);
        } else {
            node.type = NODE.AREA;
            node.parentId = SCHUL_ORG_SGA_BOARD_ID;
            node.refKind = 'sgaScope';
            node.refKey = scopeDef.scope;
            node.tone = scopeDef.tone;
            node.label = node.label || scopeDef.label;
            // Relativ-Koordinaten zurücksetzen – Canvas stapelt ohne Überlappung
            node.x = undefined;
            node.y = undefined;
        }
        node.meta = Object.assign({}, node.meta || {}, {
            board: true,
            sgaScope: scopeDef.scope,
            peopleCount: count,
            svSeats: scopeDef.scope === 'student' ? true : undefined
        });
        if (scopeDef.id === SCHUL_ORG_SGA_SV_AREA_ID) {
            node.label = 'Schülervertretung';
        }
    });

    // Legacy-Wrapper „Schülervertretung & SGA“ + alte Aggregat-Kacheln entfernen
    base.nodes.forEach(function (n) {
        if (normStr(n.parentId) !== 'sec_vertretung') return;
        if (isSgaScopeArea(n) || isSvSeatNode(n) || n.id === SCHUL_ORG_SGA_BOARD_ID) {
            n.parentId = SCHUL_ORG_SGA_BOARD_ID;
            return;
        }
        n.parentId = null;
    });
    base.nodes = base.nodes.filter(function (n) {
        if (n.id === 'sec_vertretung') return false;
        if (n.id === 'council_board') return false;
        if (
            normStr(n.type).toLowerCase() === NODE.COUNCIL &&
            normStr(n.refKey) === SCHUL_ORG_BOARD_REF &&
            !isSvSeatNode(n)
        ) {
            return false;
        }
        return true;
    });

    let maxOrder = base.nodes.reduce(function (m, n) {
        return Math.max(m, typeof n.order === 'number' ? n.order : 0);
    }, 0);
    SCHUL_ORG_SV_SEATS.forEach(function (seat) {
        let node = base.nodes.find(function (n) {
            return n.id === seat.id || (normStr(n.refKind) === 'studentCouncil' && normStr(n.refKey) === seat.refKey);
        });
        const person = personForSvSeat(s, seat.seat);
        if (!node) {
            maxOrder += 1;
            node = {
                id: seat.id,
                type: NODE.COUNCIL,
                label: seat.label,
                parentId: SCHUL_ORG_SGA_SV_AREA_ID,
                order: maxOrder,
                tone: 'council',
                refKind: 'studentCouncil',
                refKey: seat.refKey
            };
            base.nodes.push(node);
        } else {
            node.id = seat.id;
            node.type = NODE.COUNCIL;
            node.parentId = SCHUL_ORG_SGA_SV_AREA_ID;
            node.refKind = 'studentCouncil';
            node.refKey = seat.refKey;
            node.label = seat.label;
            node.tone = 'council';
        }
        node.meta = Object.assign({}, node.meta || {}, {
            svSeat: seat.seat,
            peopleCount: person ? 1 : 0,
            personName: person ? person.name : '',
            personEmail: person ? person.email : ''
        });
    });

    return normalizeSchulOrganigramm(base);
}

/**
 * Jahrgangs-Bereiche unter „Schüler“ sicherstellen und Klassen dort einhängen.
 * Migriert Legacy `sec_klassen` → Jahrgänge. M365-Gruppen pro Jahrgang später möglich (refKind: jahrgang).
 * @param {object} layout
 * @param {object} [settings]
 */
/**
 * Area-/Custom-Knoten, die wie eine Klasse aussehen (Label „Klasse 4HMA …“),
 * aber kein echter class-Knoten sind – oft Geister unter „Sonstige Klassen“.
 * @param {object|null|undefined} n
 * @param {Set<string>} [knownCodes]
 */
export function isPhantomSchulOrgClassNode(n, knownCodes) {
    if (!n || typeof n !== 'object') return false;
    const t = normStr(n.type).toLowerCase();
    // Echte Schüler-Klassen (Leaf oder Bereich) sind nie Geister
    if (t === NODE.CLASS || isClassAreaNode(n) || isClassNode(n)) return false;
    if (t !== NODE.AREA && t !== 'section' && t !== NODE.CUSTOM && t !== 'group') return false;
    if (
        isJahrgangArea(n) ||
        isElternJahrgangArea(n) ||
        isElternClassNode(n) ||
        isKlassenchatJahrgangArea(n) ||
        isKlassenchatClassNode(n) ||
        isArgeAreaNode(n) ||
        isArgeLeiterArea(n) ||
        isKlassenvorstandArea(n) ||
        isFunctionAggregateArea(n)
    ) {
        return false;
    }
    if (isSammelgruppeArea(n) || isBuiltinSchulOrganigrammArea(n)) return false;
    if (isSgaBoardArea(n) || isSgaScopeArea(n)) return false;
    const refKind = normStr(n.refKind).toLowerCase();
    if (
        refKind === 'jahrgang' ||
        refKind === 'elternjahrgang' ||
        refKind === 'elternclass' ||
        refKind === 'klassenchatjahrgang' ||
        refKind === 'klassenchatclass' ||
        refKind === 'arge' ||
        refKind === 'argeleiter' ||
        refKind === 'klassenvorstand' ||
        refKind === 'classfunction' ||
        refKind === 'argefunction' ||
        refKind === 'sammelgruppe' ||
        refKind === 'audience'
    ) {
        return false;
    }
    // Ohne refKind class: nur Label/Id-ähnlich „Klasse …“ zählt als Phantom
    if (refKind === 'class') return false;
    const code =
        canonicalSchulOrgClassCode(n.refKey) ||
        canonicalSchulOrgClassCode(n.label) ||
        canonicalSchulOrgClassCode(String(n.id || '').replace(/^class_/i, ''));
    if (!code) return false;
    if (knownCodes && knownCodes.size && knownCodes.has(code)) return true;
    // Label sieht aus wie Stammdaten-Klasse
    if (/^Klasse\s+/i.test(normStr(n.label)) && /^\d+/.test(code)) return true;
    if (normStr(n.tone) === 'class' && /^\d+/.test(code)) return true;
    return false;
}

/**
 * Geister-Klassen-Areas entfernen bzw. in echte class-Knoten überführen und deduplizieren.
 * @param {object} layout
 * @param {object} [settings]
 * @returns {{ layout: object, stripped: number }}
 */
export function stripPhantomSchulOrgClassAreas(layout, settings) {
    const base = normalizeSchulOrganigramm(layout);
    const s = settings && typeof settings === 'object' ? settings : {};
    /** @type {Set<string>} */
    const knownCodes = new Set();
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (code) knownCodes.add(code);
    });
    base.nodes.forEach(function (n) {
        if (!isClassNode(n)) return;
        const code =
            canonicalSchulOrgClassCode(n.refKey) ||
            canonicalSchulOrgClassCode(String(n.id).replace(/^class_/i, ''));
        if (code) knownCodes.add(code);
    });

    let stripped = 0;
    /** @type {object[]} */
    const converted = [];
    base.nodes = base.nodes.filter(function (n) {
        if (!isPhantomSchulOrgClassNode(n, knownCodes)) return true;
        const code =
            canonicalSchulOrgClassCode(n.refKey) ||
            canonicalSchulOrgClassCode(n.label) ||
            canonicalSchulOrgClassCode(String(n.id).replace(/^class_/i, ''));
        if (!code) {
            stripped += 1;
            return false;
        }
        // Wenn es die Klasse schon gibt → Phantom verwerfen
        const already = base.nodes.some(function (x) {
            return isClassNode(x) && (canonicalSchulOrgClassCode(x.refKey) || '') === code;
        });
        if (already || converted.some(function (x) {
            return x.refKey === code;
        })) {
            stripped += 1;
            return false;
        }
        // Sonst in echten Klassen-Bereich wandeln
        const stufe = schulstufeFromClassCode(code);
        converted.push({
            id: 'class_' + code,
            type: NODE.SECTION,
            label: normStr(n.label) || code,
            parentId: jahrgangAreaId(stufe),
            order: typeof n.order === 'number' ? n.order : 0,
            tone: 'class',
            refKind: 'class',
            refKey: code,
            meta: Object.assign({}, n.meta || {}, {
                jahrgang: stufe,
                nestPeople: true,
                audience: 'schueler'
            }),
            m365: n.m365
        });
        stripped += 1;
        return false;
    });
    converted.forEach(function (n) {
        base.nodes.push(n);
    });
    return { layout: normalizeSchulOrganigramm(base), stripped: stripped };
}

/**
 * Klassen-Knoten auf kanonische Codes zusammenführen (Duplikate Jahrgang + Sonstige).
 * Kinder (Unterrichtsteams) folgen dem behaltenen Knoten.
 * @param {object} layout
 * @returns {{ layout: object, merged: number }}
 */
export function mergeDuplicateSchulOrganigrammClasses(layout) {
    const base = normalizeSchulOrganigramm(layout);
    /** @type {Map<string, object[]>} */
    const byCanon = new Map();
    base.nodes.forEach(function (n) {
        if (!isClassNode(n) && normStr(n.refKind).toLowerCase() !== 'class') return;
        if (isElternClassNode(n) || isKlassenchatClassNode(n)) return;
        const code =
            canonicalSchulOrgClassCode(n.refKey) ||
            canonicalSchulOrgClassCode(String(n.id).replace(/^class_/i, '')) ||
            canonicalSchulOrgClassCode(n.label);
        if (!code) return;
        // Legacy-Leaf / Area → Klassen-Bereich
        n.type = NODE.SECTION;
        n.refKind = 'class';
        n.tone = n.tone || 'class';
        n.meta = Object.assign({}, n.meta || {}, { nestPeople: true, audience: 'schueler' });
        if (!byCanon.has(code)) byCanon.set(code, []);
        byCanon.get(code).push(n);
    });

    /** @type {Set<string>} */
    const dropIds = new Set();
    /** @type {Map<string, string>} alter parentId → neuer */
    const reparent = new Map();
    let merged = 0;

    byCanon.forEach(function (group, code) {
        if (group.length < 2) {
            const only = group[0];
            if (!only) return;
            const wantId = 'class_' + code;
            const stufe = schulstufeFromClassCode(code);
            only.type = NODE.SECTION;
            only.refKind = 'class';
            only.refKey = code;
            only.tone = only.tone || 'class';
            only.meta = Object.assign({}, only.meta || {}, {
                jahrgang: stufe,
                nestPeople: true,
                audience: 'schueler'
            });
            if (only.id !== wantId) {
                reparent.set(String(only.id), wantId);
                only.id = wantId;
            }
            only.parentId = jahrgangAreaId(stufe);
            return;
        }
        // Bevorzugen: korrekte id class_CODE, dann mehr Kursteam-Kinder, dann Jahrgang ≠ sonst
        group.sort(function (a, b) {
            const aIdOk = a.id === 'class_' + code ? 1 : 0;
            const bIdOk = b.id === 'class_' + code ? 1 : 0;
            if (bIdOk !== aIdOk) return bIdOk - aIdOk;
            const aKids = base.nodes.filter(function (n) {
                return n.parentId === a.id;
            }).length;
            const bKids = base.nodes.filter(function (n) {
                return n.parentId === b.id;
            }).length;
            if (bKids !== aKids) return bKids - aKids;
            const aSonst = String(a.parentId || '').indexOf('sonst') >= 0 ? 1 : 0;
            const bSonst = String(b.parentId || '').indexOf('sonst') >= 0 ? 1 : 0;
            return aSonst - bSonst;
        });
        const keep = group[0];
        const wantId = 'class_' + code;
        const stufe = schulstufeFromClassCode(code);
        const oldKeepId = String(keep.id);
        keep.type = NODE.SECTION;
        keep.refKind = 'class';
        keep.refKey = code;
        keep.tone = keep.tone || 'class';
        keep.id = wantId;
        keep.parentId = jahrgangAreaId(stufe);
        keep.meta = Object.assign({}, keep.meta || {}, {
            jahrgang: stufe,
            nestPeople: true,
            audience: 'schueler'
        });
        if (oldKeepId !== wantId) reparent.set(oldKeepId, wantId);
        for (let i = 1; i < group.length; i++) {
            const loser = group[i];
            dropIds.add(String(loser.id));
            reparent.set(String(loser.id), wantId);
            // Meta vom Duplikat nachziehen, wenn Keeper leer
            if (!normStr(keep.meta && keep.meta.headEmail) && normStr(loser.meta && loser.meta.headEmail)) {
                keep.meta = Object.assign({}, keep.meta || {}, {
                    headName: loser.meta.headName,
                    headEmail: loser.meta.headEmail,
                    teacherLinked: loser.meta.teacherLinked,
                    teacherCode: loser.meta.teacherCode
                });
            }
            if (!normStr(keep.label) && normStr(loser.label)) keep.label = loser.label;
            merged += 1;
        }
    });

    if (!dropIds.size && !reparent.size) {
        return { layout: base, merged: 0 };
    }

    base.nodes = base.nodes.filter(function (n) {
        return !dropIds.has(String(n.id));
    });
    base.nodes.forEach(function (n) {
        const pid = n.parentId == null ? '' : String(n.parentId);
        if (pid && reparent.has(pid)) {
            n.parentId = reparent.get(pid);
        }
        // Kursteam-Platzierung an kanonische Klasse anpassen
        if (normStr(n.type).toLowerCase() === NODE.KURSTEAM || n.refKind === 'kursteam') {
            const rk = normStr(n.refKey);
            const at = rk.indexOf('@');
            if (at > 0) {
                const stable = rk.slice(0, at);
                const oldClass = rk.slice(at + 1);
                const newClass = canonicalSchulOrgClassCode(oldClass) || oldClass;
                if (newClass && newClass !== oldClass) {
                    n.refKey = stable + '@' + newClass;
                    n.parentId = 'class_' + newClass;
                    n.id = ('kursteam_' + n.refKey.replace(/[^a-zA-Z0-9@_.-]/g, '_')).slice(0, 96);
                    n.meta = Object.assign({}, n.meta || {}, { placementClass: newClass });
                }
            }
        }
    });

    return { layout: normalizeSchulOrganigramm(base), merged: merged };
}

export function ensureSchulOrganigrammJahrgangStructure(layout, settings) {
    const merged = mergeDuplicateSchulOrganigrammClasses(layout);
    const base = merged.layout;
    const s = settings && typeof settings === 'object' ? settings : {};

    /** @type {Set<string>} */
    const stufen = new Set();
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        stufen.add(schulstufeFromClassCode(code));
    });
    base.nodes.forEach(function (n) {
        if (!isClassNode(n)) return;
        const code =
            canonicalSchulOrgClassCode(n.refKey) ||
            canonicalSchulOrgClassCode(String(n.id).replace(/^class_/i, ''));
        if (code) stufen.add(schulstufeFromClassCode(code));
    });
    if (!stufen.size && base.nodes.some(function (n) {
        return isJahrgangArea(n);
    })) {
        // Bestehende Jahrgangs-Areas behalten, auch ohne aktuelle Klassenliste
        base.nodes.forEach(function (n) {
            const k = jahrgangKeyForArea(n);
            if (k) stufen.add(k);
        });
    }

    const ordered = Array.from(stufen);
    ordered.sort(function (a, b) {
        if (a === SCHUL_ORG_JAHRGANG_SONST) return 1;
        if (b === SCHUL_ORG_JAHRGANG_SONST) return -1;
        return Number(a) - Number(b) || String(a).localeCompare(String(b), 'de');
    });

    let maxOrder = base.nodes.reduce(function (m, n) {
        return Math.max(m, typeof n.order === 'number' ? n.order : 0);
    }, 0);

    ordered.forEach(function (stufe) {
        const jid = jahrgangAreaId(stufe);
        let node = base.nodes.find(function (n) {
            return n.id === jid || jahrgangKeyForArea(n) === stufe;
        });
        if (!node) {
            maxOrder += 1;
            node = {
                id: jid,
                type: NODE.SECTION,
                label: jahrgangAreaLabel(stufe),
                parentId: SCHUL_ORG_SEC_SCHUELER,
                order: maxOrder,
                tone: 'class',
                refKind: 'jahrgang',
                refKey: stufe,
                meta: { jahrgang: stufe, nestClasses: true }
            };
            base.nodes.push(node);
        } else {
            node.id = jid;
            node.type = NODE.SECTION;
            node.parentId = SCHUL_ORG_SEC_SCHUELER;
            node.refKind = 'jahrgang';
            node.refKey = stufe;
            node.tone = node.tone || 'class';
            node.label = jahrgangAreaLabel(stufe);
            // Keine festen Orte – Canvas packt Jahrgänge nebeneinander
            delete node.x;
            delete node.y;
            node.meta = Object.assign({}, node.meta || {}, { jahrgang: stufe, nestClasses: true });
        }
    });

    base.nodes.forEach(function (n) {
        if (!isClassNode(n)) return;
        const code =
            canonicalSchulOrgClassCode(n.refKey) ||
            canonicalSchulOrgClassCode(String(n.id).replace(/^class_/i, ''));
        if (!code) return;
        n.type = NODE.SECTION;
        n.refKind = 'class';
        n.refKey = code;
        n.id = 'class_' + code;
        n.tone = n.tone || 'class';
        const stufe = schulstufeFromClassCode(code);
        const wantParent = jahrgangAreaId(stufe);
        if (n.parentId !== wantParent) {
            n.parentId = wantParent;
            n.x = undefined;
            n.y = undefined;
        }
        n.meta = Object.assign({}, n.meta || {}, {
            jahrgang: stufe,
            nestPeople: true,
            audience: 'schueler'
        });
    });

    // Leere Jahrgangs-Areas ohne Klassen entfernen (außer noch Kinder anderer Art)
    base.nodes = base.nodes.filter(function (n) {
        if (!isJahrgangArea(n)) return true;
        const hasKids = base.nodes.some(function (ch) {
            return ch.parentId === n.id;
        });
        return hasKids;
    });

    // Legacy-Container „Klassen“ entfernen, wenn leer / nur noch Hülle
    const klassen = base.nodes.find(function (n) {
        return n.id === 'sec_klassen';
    });
    if (klassen) {
        const stillKids = base.nodes.some(function (n) {
            return n.parentId === 'sec_klassen';
        });
        if (!stillKids) {
            base.nodes = base.nodes.filter(function (n) {
                return n.id !== 'sec_klassen';
            });
        }
    }

    return normalizeSchulOrganigramm(base);
}

/**
 * Eltern-Bereich: Jahrgänge + Klassen-Kacheln (Erziehungsberechtigte) sicherstellen.
 * @param {object} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammElternStructure(layout, settings) {
    const base = normalizeSchulOrganigramm(layout);
    const s = settings && typeof settings === 'object' ? settings : {};

    let elternRoot = base.nodes.find(function (n) {
        return n.id === SCHUL_ORG_SEC_ELTERN || sammelgruppeKeyForArea(n) === 'eltern';
    });
    let maxOrder = base.nodes.reduce(function (m, n) {
        return Math.max(m, typeof n.order === 'number' ? n.order : 0);
    }, 0);
    if (!elternRoot) {
        maxOrder += 1;
        elternRoot = {
            id: SCHUL_ORG_SEC_ELTERN,
            type: NODE.SECTION,
            label: 'Eltern',
            parentId: null,
            order: maxOrder,
            tone: 'council',
            refKind: 'sammelgruppe',
            refKey: 'eltern',
            meta: {
                board: true,
                sammelgruppe: 'eltern',
                nestClasses: true,
                guardianCount: countUniqueGuardiansInSettings(s)
            }
        };
        base.nodes.push(elternRoot);
    } else {
        elternRoot.id = SCHUL_ORG_SEC_ELTERN;
        elternRoot.refKind = 'sammelgruppe';
        elternRoot.refKey = 'eltern';
        elternRoot.tone = elternRoot.tone || 'council';
        elternRoot.meta = Object.assign({}, elternRoot.meta || {}, {
            board: true,
            sammelgruppe: 'eltern',
            nestClasses: true,
            guardianCount: countUniqueGuardiansInSettings(s)
        });
    }

    /** @type {Set<string>} */
    const stufen = new Set();
    /** @type {Map<string, object>} */
    const classByCanon = new Map();
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        classByCanon.set(code, c);
        stufen.add(schulstufeFromClassCode(code));
    });
    base.nodes.forEach(function (n) {
        if (!isElternClassNode(n)) return;
        const code = canonicalSchulOrgClassCode(n.refKey) || canonicalSchulOrgClassCode(n.label);
        if (code) stufen.add(schulstufeFromClassCode(code));
    });

    const ordered = Array.from(stufen);
    ordered.sort(function (a, b) {
        if (a === SCHUL_ORG_JAHRGANG_SONST) return 1;
        if (b === SCHUL_ORG_JAHRGANG_SONST) return -1;
        return Number(a) - Number(b) || String(a).localeCompare(String(b), 'de');
    });

    ordered.forEach(function (stufe) {
        const jid = elternJahrgangAreaId(stufe);
        let node = base.nodes.find(function (n) {
            return n.id === jid || elternJahrgangKeyForArea(n) === stufe;
        });
        if (!node) {
            maxOrder += 1;
            node = {
                id: jid,
                type: NODE.SECTION,
                label: jahrgangAreaLabel(stufe),
                parentId: SCHUL_ORG_SEC_ELTERN,
                order: maxOrder,
                tone: 'council',
                refKind: 'elternJahrgang',
                refKey: stufe,
                meta: { jahrgang: stufe, nestClasses: true, audience: 'eltern' }
            };
            base.nodes.push(node);
        } else {
            node.id = jid;
            node.type = NODE.SECTION;
            node.parentId = SCHUL_ORG_SEC_ELTERN;
            node.refKind = 'elternJahrgang';
            node.refKey = stufe;
            node.tone = node.tone || 'council';
            node.label = jahrgangAreaLabel(stufe);
            delete node.x;
            delete node.y;
            node.meta = Object.assign({}, node.meta || {}, {
                jahrgang: stufe,
                nestClasses: true,
                audience: 'eltern'
            });
        }
    });

    /** @type {Set<string>} */
    const wantClassCodes = new Set(classByCanon.keys());
    classByCanon.forEach(function (c, code) {
        const eid = elternClassNodeId(code);
        const stufe = schulstufeFromClassCode(code);
        const wantParent = elternJahrgangAreaId(stufe);
        let node = base.nodes.find(function (n) {
            return n.id === eid || (isElternClassNode(n) && canonicalSchulOrgClassCode(n.refKey) === code);
        });
        if (!node) {
            maxOrder += 1;
            base.nodes.push({
                id: eid,
                type: NODE.SECTION,
                label: normStr(c && c.name) || 'Klasse ' + code,
                parentId: wantParent,
                order: maxOrder,
                tone: 'council',
                refKind: 'elternClass',
                refKey: code,
                meta: {
                    year: normStr(c && c.year),
                    jahrgang: stufe,
                    audience: 'eltern',
                    nestPeople: true
                }
            });
            return;
        }
        node.id = eid;
        node.type = NODE.SECTION;
        node.refKind = 'elternClass';
        node.refKey = code;
        node.parentId = wantParent;
        node.tone = node.tone || 'council';
        if (!normStr(node.label)) node.label = normStr(c && c.name) || 'Klasse ' + code;
        node.meta = Object.assign({}, node.meta || {}, {
            year: normStr(c && c.year) || (node.meta && node.meta.year) || '',
            jahrgang: stufe,
            audience: 'eltern',
            nestPeople: true
        });
    });

    // Verwaiste Eltern-Klassen entfernen
    base.nodes = base.nodes.filter(function (n) {
        if (!isElternClassNode(n)) return true;
        const code = canonicalSchulOrgClassCode(n.refKey);
        return !!(code && wantClassCodes.has(code));
    });

    // Leere Eltern-Jahrgänge entfernen
    base.nodes = base.nodes.filter(function (n) {
        if (!isElternJahrgangArea(n)) return true;
        return base.nodes.some(function (ch) {
            return ch.parentId === n.id;
        });
    });

    return normalizeSchulOrganigramm(base);
}

/**
 * KlassenChats-Bereich: Jahrgänge + Klassen-Kacheln (Lehrkräfte-Gruppenchats) sicherstellen.
 * @param {object} layout
 * @param {object} [settings]
 */
export function ensureSchulOrganigrammKlassenchatsStructure(layout, settings) {
    const base = normalizeSchulOrganigramm(layout);
    const s = settings && typeof settings === 'object' ? settings : {};
    const chatMetaByKlasse = classChatMetaByKlasse(s);

    let root = base.nodes.find(function (n) {
        return n.id === SCHUL_ORG_SEC_KLASSENCHATS || sammelgruppeKeyForArea(n) === 'klassenchats';
    });
    let maxOrder = base.nodes.reduce(function (m, n) {
        return Math.max(m, typeof n.order === 'number' ? n.order : 0);
    }, 0);
    if (!root) {
        maxOrder += 1;
        root = {
            id: SCHUL_ORG_SEC_KLASSENCHATS,
            type: NODE.SECTION,
            label: 'KlassenChats',
            parentId: null,
            order: maxOrder,
            tone: 'lehrer',
            refKind: 'sammelgruppe',
            refKey: 'klassenchats',
            meta: {
                board: true,
                sammelgruppe: 'klassenchats',
                nestClasses: true,
                chatCount: chatMetaByKlasse.size
            }
        };
        base.nodes.push(root);
    } else {
        root.id = SCHUL_ORG_SEC_KLASSENCHATS;
        root.refKind = 'sammelgruppe';
        root.refKey = 'klassenchats';
        root.tone = root.tone || 'lehrer';
        // Nutzer-Umbenennung behalten (wie Eltern/Schüler/Lehrkräfte)
        if (!normStr(root.label)) root.label = 'KlassenChats';
        root.meta = Object.assign({}, root.meta || {}, {
            board: true,
            sammelgruppe: 'klassenchats',
            nestClasses: true,
            chatCount: chatMetaByKlasse.size
        });
    }

    /** @type {Set<string>} */
    const stufen = new Set();
    /** @type {Map<string, object>} */
    const classByCanon = new Map();
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (!code) return;
        classByCanon.set(code, c);
        stufen.add(schulstufeFromClassCode(code));
    });
    base.nodes.forEach(function (n) {
        if (!isKlassenchatClassNode(n)) return;
        const code = canonicalSchulOrgClassCode(n.refKey) || canonicalSchulOrgClassCode(n.label);
        if (code) stufen.add(schulstufeFromClassCode(code));
    });

    const ordered = Array.from(stufen);
    ordered.sort(function (a, b) {
        if (a === SCHUL_ORG_JAHRGANG_SONST) return 1;
        if (b === SCHUL_ORG_JAHRGANG_SONST) return -1;
        return Number(a) - Number(b) || String(a).localeCompare(String(b), 'de');
    });

    ordered.forEach(function (stufe) {
        const jid = klassenchatJahrgangAreaId(stufe);
        let node = base.nodes.find(function (n) {
            return n.id === jid || klassenchatJahrgangKeyForArea(n) === stufe;
        });
        if (!node) {
            maxOrder += 1;
            node = {
                id: jid,
                type: NODE.SECTION,
                label: jahrgangAreaLabel(stufe),
                parentId: SCHUL_ORG_SEC_KLASSENCHATS,
                order: maxOrder,
                tone: 'lehrer',
                refKind: 'klassenchatJahrgang',
                refKey: stufe,
                meta: { jahrgang: stufe, nestClasses: true, audience: 'klassenchats' }
            };
            base.nodes.push(node);
        } else {
            node.id = jid;
            node.type = NODE.SECTION;
            node.parentId = SCHUL_ORG_SEC_KLASSENCHATS;
            node.refKind = 'klassenchatJahrgang';
            node.refKey = stufe;
            node.tone = node.tone || 'lehrer';
            node.label = jahrgangAreaLabel(stufe);
            delete node.x;
            delete node.y;
            node.meta = Object.assign({}, node.meta || {}, {
                jahrgang: stufe,
                nestClasses: true,
                audience: 'klassenchats'
            });
        }
    });

    /** @type {Set<string>} */
    const wantClassCodes = new Set(classByCanon.keys());
    classByCanon.forEach(function (c, code) {
        const kid = klassenchatClassNodeId(code);
        const stufe = schulstufeFromClassCode(code);
        const wantParent = klassenchatJahrgangAreaId(stufe);
        const chatMeta = chatMetaByKlasse.get(code) || null;
        let node = base.nodes.find(function (n) {
            return (
                n.id === kid ||
                (isKlassenchatClassNode(n) && canonicalSchulOrgClassCode(n.refKey) === code)
            );
        });
        if (!node) {
            maxOrder += 1;
            base.nodes.push({
                id: kid,
                type: NODE.SECTION,
                label: normStr(c && c.name) || 'Klasse ' + code,
                parentId: wantParent,
                order: maxOrder,
                tone: 'lehrer',
                refKind: 'klassenchatClass',
                refKey: code,
                meta: {
                    year: normStr(c && c.year),
                    jahrgang: stufe,
                    audience: 'klassenchats',
                    nestPeople: true,
                    chatId: chatMeta ? chatMeta.chatId : '',
                    chatTopic: chatMeta ? chatMeta.topic : '',
                    chatWebUrl: chatMeta ? chatMeta.webUrl : ''
                },
                m365: {
                    groupMode: 'none',
                    groupKind: 'groupChat',
                    groupId: null,
                    groupLabel: null,
                    groupMail: null,
                    mailboxMode: 'none',
                    mailboxUpn: null,
                    mailboxLabel: null,
                    mailboxId: null
                }
            });
            return;
        }
        node.id = kid;
        node.type = NODE.SECTION;
        node.refKind = 'klassenchatClass';
        node.refKey = code;
        node.parentId = wantParent;
        node.tone = node.tone || 'lehrer';
        if (!normStr(node.label)) node.label = normStr(c && c.name) || 'Klasse ' + code;
        node.meta = Object.assign({}, node.meta || {}, {
            year: normStr(c && c.year) || (node.meta && node.meta.year) || '',
            jahrgang: stufe,
            audience: 'klassenchats',
            nestPeople: true,
            chatId: chatMeta ? chatMeta.chatId : (node.meta && node.meta.chatId) || '',
            chatTopic: chatMeta ? chatMeta.topic : (node.meta && node.meta.chatTopic) || '',
            chatWebUrl: chatMeta ? chatMeta.webUrl : (node.meta && node.meta.chatWebUrl) || ''
        });
        // Default-Intent Gruppenchat, wenn noch keine M365-Bindung gesetzt ist
        if (!node.m365 || typeof node.m365 !== 'object') {
            node.m365 = {
                groupMode: 'none',
                groupKind: 'groupChat',
                groupId: null,
                groupLabel: null,
                groupMail: null,
                mailboxMode: 'none',
                mailboxUpn: null,
                mailboxLabel: null,
                mailboxId: null
            };
        }
    });

    base.nodes = base.nodes.filter(function (n) {
        if (!isKlassenchatClassNode(n)) return true;
        const code = canonicalSchulOrgClassCode(n.refKey);
        return !!(code && wantClassCodes.has(code));
    });

    base.nodes = base.nodes.filter(function (n) {
        if (!isKlassenchatJahrgangArea(n)) return true;
        return base.nodes.some(function (ch) {
            return ch.parentId === n.id;
        });
    });

    return normalizeSchulOrganigramm(base);
}

/**
 * Zufalls-/Müll-Keys aus der alten Speichernormalisierung (z. B. OG_KLJ43FJC1A).
 * @param {unknown} v
 */
export function isGarbageSchulOrgRef(v) {
    const s = normStr(v);
    if (!s) return false;
    if (/^(OG|VW)[_-]/i.test(s)) return true;
    if (/^(og|vw|n)_[a-z0-9]+$/i.test(s)) return true;
    return false;
}

function cleanRefToken(v) {
    const s = normStr(v);
    if (!s || isGarbageSchulOrgRef(s)) return '';
    return s;
}

/**
 * Nach fehlerhafter Verwaltung-Normalisierung (refKind/refKey weg, type→group) heilen.
 * @param {object} n
 */
function repairSchulOrganigrammNode(n) {
    if (!n || typeof n !== 'object') return n;
    const id = normStr(n.id);
    let type = normStr(n.type).toLowerCase() || NODE.CUSTOM;
    let refKind = normStr(n.refKind) || 'none';
    let refKey = cleanRefToken(n.refKey || n.roleCode || n.code);
    const tone = normStr(n.tone);

    const CANON_SECTIONS = {
        sec_schulleitung: {
            label: 'Schulleitung',
            tone: 'lead',
            refKind: 'audience',
            refKey: 'schulleitung'
        },
        sec_verwaltung: {
            label: 'Verwaltung (Personal)',
            tone: 'staff',
            refKind: 'audience',
            refKey: 'verwaltung'
        },
        sec_lehrer: {
            label: 'Lehrkräfte',
            tone: 'lehrer',
            refKind: 'sammelgruppe',
            refKey: 'lehrer'
        },
        sec_schueler: {
            label: 'Schüler',
            tone: 'schueler',
            refKind: 'sammelgruppe',
            refKey: 'schueler'
        },
        sec_eltern: {
            label: 'Eltern',
            tone: 'council',
            refKind: 'sammelgruppe',
            refKey: 'eltern'
        },
        sec_klassenchats: {
            label: 'KlassenChats',
            tone: 'lehrer',
            refKind: 'sammelgruppe',
            refKey: 'klassenchats'
        },
        sec_arge: { label: 'ARGEs / Fachschaften', tone: 'arge', refKind: 'none', refKey: '' },
        sec_arge_leiter: {
            label: 'ARGE-Leiter',
            tone: 'arge',
            refKind: 'argeLeiter',
            refKey: 'argeLeiter',
            parentId: SCHUL_ORG_SEC_ARGE
        },
        sec_klassenvorstand: {
            label: 'Klassenvorstände',
            tone: 'lehrer',
            refKind: 'klassenvorstand',
            refKey: 'klassenvorstand',
            parentId: null
        },
        sec_classfn_klassensprecher: {
            label: 'Klassensprecher',
            tone: 'schueler',
            refKind: 'classFunction',
            refKey: 'klassensprecher',
            parentId: null
        },
        sec_klassen: {
            label: 'Klassen',
            tone: 'class',
            refKind: 'none',
            refKey: '',
            parentId: SCHUL_ORG_SEC_SCHUELER
        },
        sga_board: {
            label: 'SGA',
            tone: 'sga',
            refKind: 'sga',
            refKey: SCHUL_ORG_BOARD_REF,
            parentId: null
        }
    };
    if (/^sec_jahrgang_/i.test(id) || refKind === 'jahrgang') {
        const stufe =
            cleanRefToken(normStr(refKey).toLowerCase()) ||
            jahrgangKeyForArea({ id: id, type: NODE.AREA, refKind: 'jahrgang', refKey: refKey }) ||
            SCHUL_ORG_JAHRGANG_SONST;
        return Object.assign({}, n, {
            id: jahrgangAreaId(stufe),
            type: NODE.SECTION,
            parentId: SCHUL_ORG_SEC_SCHUELER,
            tone: tone || 'class',
            refKind: 'jahrgang',
            refKey: stufe,
            label: normStr(n.label) || jahrgangAreaLabel(stufe)
        });
    }
    if (id.indexOf('sec_aud_') === 0 && (type === NODE.SECTION || type === 'group' || !type)) {
        const audId = id.slice('sec_aud_'.length);
        // parentId behalten – Custom-Verwaltungsbereiche dürfen verschachtelt werden
        const keepParent =
            n.parentId == null || n.parentId === '' ? null : String(n.parentId);
        return Object.assign({}, n, {
            type: NODE.SECTION,
            parentId: keepParent,
            tone: tone || 'staff',
            refKind: 'audience',
            refKey: audId.indexOf('va-') === 0 ? audId : cleanRefToken(refKey) || audId
        });
    }
    if (
        CANON_SECTIONS[id] ||
        type === NODE.SECTION ||
        (type === 'group' &&
            !n.parentId &&
            (tone === 'lead' ||
                tone === 'staff' ||
                /schulleitung|verwaltung|arge|klasse|vertretung/i.test(n.label)))
    ) {
        if (CANON_SECTIONS[id]) {
            const canon = CANON_SECTIONS[id];
            const canonParent =
                canon.parentId !== undefined
                    ? canon.parentId
                    : n.parentId == null || n.parentId === ''
                      ? null
                      : n.parentId;
            return Object.assign({}, n, {
                type: NODE.SECTION,
                parentId: canonParent == null || canonParent === '' ? null : String(canonParent),
                tone: canon.tone,
                refKind: canon.refKind || 'none',
                refKey: canon.refKey || ''
            });
        }
        if (type === 'group' && !n.parentId) {
            return Object.assign({}, n, {
                type: NODE.SECTION,
                parentId: null,
                refKind: refKind === 'audience' ? 'audience' : 'none',
                refKey: refKind === 'audience' ? refKey : ''
            });
        }
    }

    const roleCodeFromId = /^role_/i.test(id) ? cleanRefToken(normCode(id.replace(/^role_/i, ''))) : '';
    const roleCode =
        cleanRefToken(normCode(n.roleCode)) ||
        (refKind === 'adminRole' ? cleanRefToken(normCode(refKey)) : '') ||
        roleCodeFromId;
    if (type === NODE.ROLE || refKind === 'adminRole' || roleCodeFromId || cleanRefToken(n.roleCode)) {
        if (roleCode) {
            return Object.assign({}, n, {
                type: NODE.ROLE,
                refKind: 'adminRole',
                refKey: roleCode
            });
        }
        // Rolle ohne gültigen Code → als verworfen markieren (kein Müll-refKey)
        return Object.assign({}, n, {
            type: NODE.ROLE,
            refKind: 'adminRole',
            refKey: ''
        });
    }

    if (
        (type === NODE.ARGE || refKind === 'arge' || /^arge_/i.test(id)) &&
        refKind !== 'argeleiter' &&
        id !== SCHUL_ORG_SEC_ARGE_LEITER
    ) {
        const code =
            cleanRefToken(normCode(refKey)) ||
            cleanRefToken(normCode(id.replace(/^arge_/i, ''))) ||
            '';
        // Legacy-Leaf → Bereichskachel (KlassenChat-Muster: Leitung + Mitglieder)
        return Object.assign({}, n, {
            id: code ? argeNodeId(code) : id,
            type: NODE.SECTION,
            refKind: 'arge',
            refKey: code,
            tone: n.tone || 'arge',
            meta: Object.assign({}, n.meta && typeof n.meta === 'object' ? n.meta : {}, {
                nestPeople: true,
                audience: 'arge'
            })
        });
    }
    if (
        (type === NODE.CLASS || refKind === 'class' || /^class_/i.test(id)) &&
        refKind !== 'elternclass' &&
        refKind !== 'klassenchatclass' &&
        !/^eltern_class_/i.test(id) &&
        !/^klassenchat_class_/i.test(id)
    ) {
        const code =
            canonicalSchulOrgClassCode(refKey) ||
            canonicalSchulOrgClassCode(id.replace(/^class_/i, '')) ||
            canonicalSchulOrgClassCode(n.label) ||
            cleanRefToken(normCode(refKey)) ||
            '';
        const nextId = code ? 'class_' + code : id;
        // Legacy-Leaf → Bereichskachel (KV + Schüler, aufklappbar)
        return Object.assign({}, n, {
            id: nextId,
            type: NODE.SECTION,
            refKind: 'class',
            refKey: code,
            tone: n.tone || 'class',
            meta: Object.assign({}, n.meta && typeof n.meta === 'object' ? n.meta : {}, {
                nestPeople: true,
                audience: 'schueler'
            })
        });
    }
    if (type === NODE.KURSTEAM || refKind === 'kursteam' || /^kursteam_/i.test(id)) {
        const rawKey = cleanRefToken(refKey) || cleanRefToken(id.replace(/^kursteam_/i, ''));
        const at = rawKey.indexOf('@');
        const placementClass =
            at > 0
                ? canonicalSchulOrgClassCode(rawKey.slice(at + 1)) || normCode(rawKey.slice(at + 1))
                : canonicalSchulOrgClassCode((n.meta && n.meta.placementClass) || '') ||
                  normCode((n.meta && n.meta.placementClass) || '');
        const stable =
            at > 0
                ? rawKey.slice(0, at)
                : normStr((n.meta && n.meta.stableKey) || rawKey) || 'kurs';
        const nextKey = placementClass ? stable + '@' + placementClass : stable;
        return Object.assign({}, n, {
            type: NODE.KURSTEAM,
            refKind: 'kursteam',
            refKey: nextKey,
            parentId: placementClass ? 'class_' + placementClass : n.parentId,
            tone: tone || 'kursteam'
        });
    }
    if (type === NODE.COUNCIL || refKind === 'studentCouncil' || /^council_/i.test(id) || /^sv_/i.test(id)) {
        const seat = svSeatForNode({ id: id, refKey: refKey, type: NODE.COUNCIL });
        if (seat) {
            return Object.assign({}, n, {
                id: seat.id,
                type: NODE.COUNCIL,
                parentId: SCHUL_ORG_SGA_SV_AREA_ID,
                refKind: 'studentCouncil',
                refKey: seat.refKey,
                label: seat.label,
                tone: 'council'
            });
        }
        if (id === 'council_board' || refKey === SCHUL_ORG_BOARD_REF) {
            // Legacy-Aggregat → wird in ensureSgaStructure entfernt
            return Object.assign({}, n, {
                id: 'council_board',
                type: NODE.COUNCIL,
                refKind: 'studentCouncil',
                refKey: SCHUL_ORG_BOARD_REF,
                label: normStr(n.label) || 'Schülervertretung'
            });
        }
        return Object.assign({}, n, {
            type: NODE.COUNCIL,
            refKind: 'studentCouncil',
            refKey: cleanRefToken(refKey)
        });
    }
    if (id.indexOf('sga_gruppe_') === 0 || refKind === 'sgaScope') {
        const scopeHit = SCHUL_ORG_SGA_SCOPE_AREAS.find(function (s) {
            return s.id === id || s.scope === normStr(refKey).toLowerCase();
        });
        return Object.assign({}, n, {
            id: scopeHit ? scopeHit.id : id,
            type: NODE.AREA,
            parentId: SCHUL_ORG_SGA_BOARD_ID,
            refKind: 'sgaScope',
            refKey: scopeHit ? scopeHit.scope : cleanRefToken(refKey),
            label: normStr(n.label) || (scopeHit && scopeHit.label) || 'SGA-Gruppe',
            tone: (scopeHit && scopeHit.tone) || tone || 'sga'
        });
    }
    if (type === NODE.SGA || refKind === 'sga' || id === SCHUL_ORG_SGA_BOARD_ID || /^sga_/i.test(id)) {
        if (id === SCHUL_ORG_SGA_BOARD_ID || refKey === SCHUL_ORG_BOARD_REF) {
            return Object.assign({}, n, {
                id: SCHUL_ORG_SGA_BOARD_ID,
                type: NODE.AREA,
                refKind: 'sga',
                refKey: SCHUL_ORG_BOARD_REF,
                label: normStr(n.label) || 'SGA',
                tone: tone || 'sga'
            });
        }
        return Object.assign({}, n, {
            type: NODE.SGA,
            refKind: 'sga',
            refKey: cleanRefToken(refKey)
        });
    }

    return Object.assign({}, n, { type: type, refKind: refKind, refKey: refKey });
}

/**
 * @param {unknown} layout
 * @param {object} settings
 */
export function schulOrganigrammNeedsHeal(layout, settings) {
    const nodes = layout && Array.isArray(layout.nodes) ? layout.nodes : [];
    if (!nodes.length) return false;
    if (
        nodes.some(function (n) {
            const t = normStr(n && n.type).toLowerCase();
            if (t === NODE.SECTION) return false;
            return isGarbageSchulOrgRef(n && n.refKey) || isGarbageSchulOrgRef(n && n.id);
        })
    ) {
        return true;
    }
    const seeded = seedSchulOrganigrammFromStammdaten(settings);
    const seedContent = seeded.nodes.filter(function (n) {
        return n.type !== NODE.SECTION;
    }).length;
    const layoutContent = nodes.filter(function (n) {
        return normStr(n && n.type).toLowerCase() !== NODE.SECTION;
    }).length;
    if (layoutContent > Math.max(seedContent + 8, Math.ceil(seedContent * 1.35))) return true;

    /** @type {Record<string, number>} */
    const labelCounts = {};
    /** @type {Record<string, number>} */
    const classCanonCounts = {};
    /** @type {Set<string>} */
    const knownFromSettings = new Set();
    (Array.isArray(settings && settings.classes) ? settings.classes : []).forEach(function (c) {
        const code = classCodeFromSettingsRow(c);
        if (code) knownFromSettings.add(code);
    });
    let phantomClassAreas = 0;
    nodes.forEach(function (n) {
        const t = normStr(n && n.type).toLowerCase();
        if (isPhantomSchulOrgClassNode(n, knownFromSettings)) {
            phantomClassAreas += 1;
        }
        if (t !== NODE.ROLE && t !== 'group' && t !== NODE.ARGE && t !== NODE.CLASS) return;
        const L = normStr(n && n.label).toLowerCase();
        if (L) labelCounts[L] = (labelCounts[L] || 0) + 1;
        if (t === NODE.CLASS || normStr(n && n.refKind).toLowerCase() === 'class') {
            const canon =
                canonicalSchulOrgClassCode(n.refKey) ||
                canonicalSchulOrgClassCode(String(n.id || '').replace(/^class_/i, '')) ||
                canonicalSchulOrgClassCode(n.label);
            if (canon) {
                classCanonCounts[canon] = (classCanonCounts[canon] || 0) + 1;
            }
        }
    });
    // Geister-Areas unter Sonstige / doppelte Klassen → hart neu aufbauen
    if (phantomClassAreas > 0) return true;
    if (Object.keys(classCanonCounts).some(function (k) {
        return classCanonCounts[k] >= 2;
    })) {
        return true;
    }
    // Gleiches Klassen-Label zweimal (auch Area+Class)
    const classLikeLabels = {};
    nodes.forEach(function (n) {
        const L = normStr(n && n.label).toLowerCase();
        if (!L || !/^klasse\s+/i.test(L)) return;
        const canon = canonicalSchulOrgClassCode(n.label);
        if (!canon || !/^\d+/.test(canon)) return;
        classLikeLabels[canon] = (classLikeLabels[canon] || 0) + 1;
    });
    if (Object.keys(classLikeLabels).some(function (k) {
        return classLikeLabels[k] >= 2;
    })) {
        return true;
    }
    return Object.keys(labelCounts).some(function (k) {
        return labelCounts[k] >= 3;
    });
}

/**
 * Layout hart aus Stammdaten neu aufbauen; gültige Platzierung (parent/order) behalten.
 * @param {unknown} layout
 * @param {object} settings
 */
export function rebuildSchulOrganigrammFromStammdaten(layout, settings) {
    const seeded = seedSchulOrganigrammFromStammdaten(settings);
    const prevNodes = Array.isArray(layout && layout.nodes) ? layout.nodes : [];
    /** @type {Map<string, { parentId: string|null, order: number }>} */
    const placeByRef = new Map();
    /** @type {Map<string, { parentId: string|null, order: number }>} */
    const placeById = new Map();
    prevNodes.forEach(function (raw) {
        const n = repairSchulOrganigrammNode(raw);
        if (!n) return;
        if (n.id && isAreaNode(n)) {
            placeById.set(String(n.id), {
                parentId: n.parentId == null || n.parentId === '' ? null : normStr(n.parentId),
                order: typeof n.order === 'number' ? n.order : 0
            });
        }
        const rk = normStr(n.refKind);
        const key = rk && rk !== 'none' && cleanRefToken(n.refKey) ? rk + ':' + String(n.refKey).toUpperCase() : '';
        if (!key) return;
        if (placeByRef.has(key)) return;
        placeByRef.set(key, {
            parentId: n.parentId == null || n.parentId === '' ? null : normStr(n.parentId),
            order: typeof n.order === 'number' ? n.order : 0
        });
    });
    const removedAreaIds = new Set(schulOrganigrammRemovedAreaIds(layout));
    const nodes = seeded.nodes
        .filter(function (n) {
            return !(n && removedAreaIds.has(String(n.id)));
        })
        .map(function (n) {
        // Bereiche: verschachtelte parentId aus dem vorherigen Layout behalten
        if (n.type === NODE.SECTION || isAreaNode(n)) {
            const place = placeById.get(String(n.id)) || (
                n.refKind && n.refKind !== 'none' && n.refKey
                    ? placeByRef.get(n.refKind + ':' + String(n.refKey).toUpperCase())
                    : null
            );
            if (!place) return n;
            return Object.assign({}, n, {
                parentId: place.parentId,
                order: place.order
            });
        }
        // Kursteams immer an der Klasse aus der Belegung halten
        if (n.type === NODE.KURSTEAM || n.refKind === 'kursteam') return n;
        // Klassen: Jahrgang aus Seed behalten (nicht alte „Sonstige“-Platzierung zurückholen)
        if (n.type === NODE.CLASS || n.refKind === 'class') {
            const placeCls =
                n.refKind && n.refKey
                    ? placeByRef.get(n.refKind + ':' + String(n.refKey).toUpperCase())
                    : null;
            if (placeCls && typeof placeCls.order === 'number') {
                return Object.assign({}, n, { order: placeCls.order });
            }
            return n;
        }
        const key =
            n.refKind && n.refKind !== 'none' && n.refKey
                ? n.refKind + ':' + String(n.refKey).toUpperCase()
                : '';
        const place = key ? placeByRef.get(key) : null;
        if (!place) return n;
        return Object.assign({}, n, {
            parentId: place.parentId != null ? place.parentId : n.parentId,
            order: place.order
        });
    });
    // Sammel-/Custom und verschachtelte Nutzer-Bereiche (z. B. IT-Team) aus Altbestand behalten
    const seen = new Set(
        nodes.map(function (n) {
            return n.refKind && n.refKey
                ? n.refKind + ':' + String(n.refKey).toUpperCase()
                : 'id:' + n.id;
        })
    );
    prevNodes.forEach(function (raw) {
        const n = repairSchulOrganigrammNode(raw);
        if (!n) return;
        const key =
            n.refKind && n.refKind !== 'none' && cleanRefToken(n.refKey)
                ? n.refKind + ':' + String(n.refKey).toUpperCase()
                : 'id:' + n.id;
        if (seen.has(key) || seen.has('id:' + n.id)) return;
        if (isGarbageSchulOrgRef(n.id) && !cleanRefToken(n.refKey)) return;
        // Custom mit Meta (Zielgruppen o. Ä.)
        if (n.type === NODE.CUSTOM && n.meta && typeof n.meta === 'object') {
            seen.add(key);
            nodes.push(n);
            return;
        }
        // Freie Areas, die der Seed nicht kennt (auch Top-Level: IT-Team, Sekretariat, …)
        if (isAreaNode(n) && !BUILTIN_AREA_IDS[String(n.id)]) {
            const rk = normStr(n.refKind).toLowerCase();
            const idStr = String(n.id || '');
            const isFree =
                !rk ||
                rk === 'none' ||
                idStr.indexOf('area_') === 0 ||
                // Custom-Verwaltungsgruppe, falls Gruppenliste kurzzeitig fehlte
                (rk === 'audience' && idStr.indexOf('sec_aud_') === 0);
            const isNested = n.parentId != null && n.parentId !== '';
            if (isFree || isNested) {
                seen.add(key);
                seen.add('id:' + n.id);
                nodes.push(n);
            }
        }
    });
    const prevView = layout && typeof layout === 'object' ? layout.view : null;
    /** Area-Positionen / Nesting aus Altbestand behalten */
    const prevById = new Map();
    prevNodes.forEach(function (raw) {
        const n = repairSchulOrganigrammNode(raw);
        if (!n || !isAreaNode(n)) return;
        prevById.set(n.id, n);
    });
    const withPlaces = nodes.map(function (n) {
        if (!isAreaNode(n)) return n;
        const prev = prevById.get(n.id);
        if (!prev) return n;
        const px = parseSchulOrgCanvasCoord(prev.x);
        const py = parseSchulOrgCanvasCoord(prev.y);
        const pw = parseSchulOrgCanvasCoord(prev.w);
        const ph = parseSchulOrgCanvasCoord(prev.h);
        const next = Object.assign({}, n, {
            parentId: prev.parentId != null ? prev.parentId : n.parentId,
            order: typeof prev.order === 'number' ? prev.order : n.order,
            x: px != null ? px : n.x,
            y: py != null ? py : n.y,
            w: pw != null ? pw : n.w,
            h: ph != null ? ph : n.h,
            m365: prev.m365 || n.m365
        });
        // Builtin-Bereiche: umbenannte Bezeichnung aus Altbestand behalten
        if (BUILTIN_AREA_IDS[String(n.id)] && normStr(prev.label)) {
            next.label = prev.label;
        }
        if (prev.collapsed === true || n.collapsed === true) next.collapsed = true;
        return next;
    });
    const rebuilt = normalizeSchulOrganigramm({
        version: SCHUL_ORG_VERSION,
        view: prevView,
        nodes: withPlaces,
        removedAreaIds: schulOrganigrammRemovedAreaIds(layout)
    });
    // ARGEs: ungültige Parent-Bereiche (gelöscht / falscher Ast) → zurück zu sec_arge
    let argeFixed = false;
    rebuilt.nodes.forEach(function (n) {
        if (!isArgeNode(n)) return;
        const pid = n.parentId == null || n.parentId === '' ? null : String(n.parentId);
        if (pid && canPlaceSchulOrgNodeInSection(n, pid, rebuilt)) return;
        if (pid !== 'sec_arge') {
            n.parentId = 'sec_arge';
            argeFixed = true;
        }
    });
    return argeFixed ? normalizeSchulOrganigramm(rebuilt) : rebuilt;
}

/**
 * Viewport (Pan/Zoom) setzen – mit IT-Bib/Core persistent.
 * @param {unknown} layout
 * @param {{ x?: number, y?: number, k?: number }} view
 */
export function setSchulOrganigrammView(layout, view) {
    const base = normalizeSchulOrganigramm(layout);
    const v = view && typeof view === 'object' ? view : {};
    base.view = {
        x: Number.isFinite(Number(v.x)) ? Number(v.x) : base.view.x,
        y: Number.isFinite(Number(v.y)) ? Number(v.y) : base.view.y,
        k: Number.isFinite(Number(v.k)) && Number(v.k) > 0.15 && Number(v.k) < 4 ? Number(v.k) : base.view.k
    };
    return base;
}

/**
 * Canvas-Position eines Knotens setzen.
 * @param {unknown} layout
 * @param {string} nodeId
 * @param {number} x
 * @param {number} y
 */
export function setSchulOrganigrammNodePosition(layout, nodeId, x, y) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(nodeId);
    const node = base.nodes.find(function (n) {
        return n.id === id;
    });
    if (!node) return base;
    const cx = parseSchulOrgCanvasCoord(x);
    const cy = parseSchulOrgCanvasCoord(y);
    if (cx != null) node.x = cx;
    if (cy != null) node.y = cy;
    return base;
}

/** Mindest-/Maximalgröße für frei skalierbare Root-Bereiche (Landkarte). */
export const SCHUL_ORG_AREA_MIN_W = 280;
export const SCHUL_ORG_AREA_MAX_W = 2400;
export const SCHUL_ORG_AREA_MIN_H = 110;
export const SCHUL_ORG_AREA_MAX_H = 4000;

/**
 * @param {unknown} w
 * @param {unknown} h
 * @returns {{ w: number|null, h: number|null }}
 */
export function clampSchulOrganigrammAreaSize(w, h) {
    const cw = parseSchulOrgCanvasCoord(w);
    const ch = parseSchulOrgCanvasCoord(h);
    return {
        w:
            cw == null
                ? null
                : Math.max(SCHUL_ORG_AREA_MIN_W, Math.min(SCHUL_ORG_AREA_MAX_W, cw)),
        h:
            ch == null
                ? null
                : Math.max(SCHUL_ORG_AREA_MIN_H, Math.min(SCHUL_ORG_AREA_MAX_H, ch))
    };
}

/**
 * Größe eines Root-Bereichs setzen (Fenster-Resize auf der Landkarte).
 * @param {unknown} layout
 * @param {string} nodeId
 * @param {number} w
 * @param {number} h
 */
export function setSchulOrganigrammNodeSize(layout, nodeId, w, h) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(nodeId);
    const node = base.nodes.find(function (n) {
        return n.id === id && isAreaNode(n);
    });
    if (!node) return base;
    // Größe nur für Top-Level-Areas (freie Fenster auf der Karte)
    if (node.parentId != null && node.parentId !== '') return base;
    const clamped = clampSchulOrganigrammAreaSize(w, h);
    if (clamped.w != null) node.w = clamped.w;
    if (clamped.h != null) node.h = clamped.h;
    return base;
}

/**
 * Bereich auf der Landkarte ein-/ausklappen (kompakte Kopfzeile vs. volle Inhalte).
 * @param {unknown} layout
 * @param {string} areaId
 * @param {boolean} [collapsed] – weglassen = umschalten
 */
export function setSchulOrganigrammAreaCollapsed(layout, areaId, collapsed) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(areaId);
    const node = base.nodes.find(function (n) {
        return n.id === id && isAreaNode(n);
    });
    if (!node) return base;
    const next =
        typeof collapsed === 'boolean' ? collapsed === true : !(node.collapsed === true);
    if (next) node.collapsed = true;
    else delete node.collapsed;
    return normalizeSchulOrganigramm(base);
}

/**
 * Absolute Root-Area-Positionen aus gemessenen Landkarten-Boxen zurückschreiben.
 * Jahrgänge/Klassen werden nicht persistiert (Auto-Pack).
 * @param {unknown} layout
 * @param {object[]} boxes – Ergebnis von computeLandkarteLayout().boxes
 */
export function applyLandkarteRootPositionsFromBoxes(layout, boxes) {
    return applyLandkartePositionsFromBoxes(layout, boxes, { rootsOnly: true });
}

/**
 * Landkarten-Box-Positionen ins Layout schreiben (Speichern / Drag-Ende).
 * Standard: nur Root-Areas (freie Landkarten-Orte). Optional nested Custom-Areas.
 * @param {unknown} layout
 * @param {object[]} boxes
 * @param {{ rootsOnly?: boolean }} [opts]
 */
export function applyLandkartePositionsFromBoxes(layout, boxes, opts) {
    const rootsOnly = !(opts && opts.rootsOnly === false);
    const base = normalizeSchulOrganigramm(layout);
    const list = Array.isArray(boxes) ? boxes : [];
    let changed = false;
    list.forEach(function (b) {
        if (!b || b.kind !== 'area' || !b.node || !b.node.id) return;
        if (b.parentAreaId != null && b.parentAreaId !== '') {
            if (rootsOnly) return;
            // Jahrgänge nie als freie Relativ-Orte speichern
            if (
                normStr(b.node.refKind).toLowerCase() === 'jahrgang' ||
                /^sec_jahrgang_/i.test(normStr(b.node.id))
            ) {
                return;
            }
        }
        const node = base.nodes.find(function (n) {
            return n.id === b.node.id && isAreaNode(n);
        });
        if (!node) return;
        const nx = Math.round(Number(b.x) || 0);
        const ny = Math.round(Number(b.y) || 0);
        if (b.parentAreaId == null || b.parentAreaId === '') {
            if (node.x !== nx || node.y !== ny) {
                node.x = nx;
                node.y = ny;
                changed = true;
            }
            // Fenstergröße nur fortschreiben, wenn schon gesetzt (Resize) –
            // sonst bleibt Auto-Pack und wächst mit dem Inhalt.
            if (node.w != null || node.h != null) {
            // Größe nur fortschreiben, wenn schon gesetzt (Fenster-Resize)
            if (node.w != null || node.h != null) {
                const bw = parseSchulOrgCanvasCoord(b.w);
                const bh = parseSchulOrgCanvasCoord(b.h);
                if (bw != null && bh != null) {
                    const clamped = clampSchulOrganigrammAreaSize(bw, bh);
                    if (clamped.w != null && node.w !== clamped.w) {
                        node.w = clamped.w;
                        changed = true;
                    }
                    if (clamped.h != null && node.h !== clamped.h) {
                        node.h = clamped.h;
                        changed = true;
                    }
                }
            }
            }
        }
    });
    return changed ? normalizeSchulOrganigramm(base) : base;
}

/**
 * Live-Karten-Layout (Drag) mit aufgelöstem Settings-Layout mergen –
 * damit Speichern/Remount Root-Orte nicht verwirft.
 * @param {unknown} resolvedLayout
 * @param {unknown} liveLayout
 */
export function mergeSchulOrganigrammLivePositions(resolvedLayout, liveLayout) {
    const base = normalizeSchulOrganigramm(resolvedLayout);
    const live = normalizeSchulOrganigramm(liveLayout);
    if (!live.nodes.length) return base;
    const liveById = new Map(
        live.nodes.map(function (n) {
            return [n.id, n];
        })
    );
    base.nodes = base.nodes.map(function (n) {
        const prev = liveById.get(n.id);
        if (!prev) return n;
        const out = Object.assign({}, n);
        // Nur Root-Area-Orte/Größe aus der Live-Karte übernehmen (Jahrgänge = Auto-Pack)
        if (isAreaNode(n) && (n.parentId == null || n.parentId === '')) {
            const px = parseSchulOrgCanvasCoord(prev.x);
            const py = parseSchulOrgCanvasCoord(prev.y);
            if (px != null) out.x = px;
            if (py != null) out.y = py;
            const pw = parseSchulOrgCanvasCoord(prev.w);
            const ph = parseSchulOrgCanvasCoord(prev.h);
            if (pw != null && ph != null) {
                const clamped = clampSchulOrganigrammAreaSize(pw, ph);
                if (clamped.w != null) out.w = clamped.w;
                if (clamped.h != null) out.h = clamped.h;
            } else {
                if (pw != null) out.w = clampSchulOrganigrammAreaSize(pw, SCHUL_ORG_AREA_MIN_H).w;
                if (ph != null) out.h = clampSchulOrganigrammAreaSize(SCHUL_ORG_AREA_MIN_W, ph).h;
            }
        }
        if (typeof prev.order === 'number') out.order = prev.order;
        if (prev.parentId !== undefined && isAreaNode(n)) {
            out.parentId =
                prev.parentId == null || prev.parentId === '' ? null : String(prev.parentId);
        }
        if (isAreaNode(n)) {
            // Kein m365/label aus der Live-Karte: sonst überschreibt Remount nach
            // „Gruppe lösen“ die gespeicherte Lösung wieder mit dem alten Match.
            if (prev.collapsed === true) out.collapsed = true;
            else if ('collapsed' in prev) delete out.collapsed;
        }
        return out;
    });
    if (live.view) base.view = live.view;
    if (Array.isArray(live.removedAreaIds) && live.removedAreaIds.length) {
        base.removedAreaIds = schulOrganigrammRemovedAreaIds({
            removedAreaIds: [].concat(base.removedAreaIds || [], live.removedAreaIds)
        });
    }
    return normalizeSchulOrganigramm(base);
}

/**
 * Doppelte / Müll-Knoten entfernen.
 * @param {unknown} layout
 * @returns {{ layout: { version: number, nodes: object[], view?: object }, removed: number, repaired: number }}
 */
export function dedupeSchulOrganigrammLayout(layout) {
    const raw = layout && typeof layout === 'object' ? layout : {};
    const nodesIn = Array.isArray(raw.nodes) ? raw.nodes : [];
    let repaired = 0;
    let removed = 0;
    const repairedNodes = nodesIn.map(function (n) {
        const next = repairSchulOrganigrammNode(n);
        if (
            next &&
            n &&
            (next.type !== n.type ||
                normStr(next.refKind) !== normStr(n.refKind) ||
                normStr(next.refKey) !== normStr(n.refKey))
        ) {
            repaired += 1;
        }
        return next;
    });
    const sorted = repairedNodes.slice().sort(function (a, b) {
        return (a.order || 0) - (b.order || 0);
    });
    const seenRef = new Set();
    const seenId = new Set();
    const kept = [];
    sorted.forEach(function (n) {
        if (!n) return;
        const id = normStr(n.id);
        if (id && seenId.has(id)) {
            removed += 1;
            return;
        }
        if (n.type === NODE.SECTION) {
            if (id) seenId.add(id);
            kept.push(n);
            return;
        }
        if (n.type === NODE.CUSTOM && n.meta && typeof n.meta === 'object') {
            if (id) seenId.add(id);
            kept.push(n);
            return;
        }
        const rk = normStr(n.refKind);
        let ref = cleanRefToken(n.refKey);
        if (rk === 'class') {
            ref = canonicalSchulOrgClassCode(ref) || canonicalSchulOrgClassCode(n.label) || ref;
        }
        if (!(rk && rk !== 'none' && ref)) {
            removed += 1;
            return;
        }
        const key = rk + ':' + String(ref).toUpperCase();
        if (seenRef.has(key)) {
            removed += 1;
            return;
        }
        seenRef.add(key);
        if (id) seenId.add(id);
        kept.push(Object.assign({}, n, { refKey: ref }));
    });
    return {
        layout: normalizeSchulOrganigramm({
            version: SCHUL_ORG_VERSION,
            view: raw.view,
            nodes: kept,
            removedAreaIds: schulOrganigrammRemovedAreaIds(raw)
        }),
        removed: removed,
        repaired: repaired
    };
}

/**
 * Nächsten bekannten Verwaltungs-/Leitungs-Abschnitt für einen Knoten finden.
 * @param {object} node
 * @param {Map<string, object>} byId
 * @returns {string}
 */
function adminSectionAncestorId(node, byId) {
    let cur = node;
    let guard = 0;
    while (cur && guard < 48) {
        guard += 1;
        const id = normStr(cur.id);
        if (
            id === 'sec_schulleitung' ||
            id === 'sec_verwaltung' ||
            id.indexOf('sec_aud_') === 0
        ) {
            return id;
        }
        const pid = cur.parentId == null || cur.parentId === '' ? '' : String(cur.parentId);
        if (!pid) return '';
        cur = byId.get(pid) || null;
    }
    return '';
}

/**
 * Aktuelles Organigramm in die Stammdaten-Listen zurückschreiben
 * (Rollen-Tier/Namen, ARGE-/Klassen-Namen, Verwaltungsgruppen-Labels) + Layout speichern.
 *
 * @param {object} settings
 * @param {unknown} [layout]
 * @returns {{ settings: object, layout: object, changed: { roles: number, arges: number, classes: number, audiences: number } }}
 */
export function applySchulOrganigrammToStammdaten(settings, layout) {
    const base = settings && typeof settings === 'object' ? settings : {};
    const org = normalizeSchulOrganigramm(
        layout != null ? layout : base.schulOrganigramm
    );
    /** @type {Map<string, object>} */
    const byId = new Map();
    org.nodes.forEach(function (n) {
        if (n && n.id) byId.set(String(n.id), n);
    });

    const changed = { roles: 0, arges: 0, classes: 0, audiences: 0 };

    const roles = (Array.isArray(base.adminRoles) ? base.adminRoles : []).map(function (r) {
        return Object.assign({}, r);
    });
    org.nodes.forEach(function (n) {
        if (!n || n.type !== NODE.ROLE || normStr(n.refKind) !== 'adminRole') return;
        const code = normCode(n.refKey);
        if (!code) return;
        const sec = adminSectionAncestorId(n, byId);
        const tier = sec === 'sec_schulleitung' ? 'schulleitung' : 'verwaltung';
        const label = normStr(n.label) || code;
        let idx = roles.findIndex(function (r) {
            return normCode(r && r.code) === code;
        });
        if (idx < 0) {
            idx = roles.findIndex(function (r) {
                return normStr(r && r.name).toLowerCase() === label.toLowerCase();
            });
        }
        if (idx < 0) return;
        const prev = roles[idx];
        const nextTier = tier || normStr(prev.tier) || 'verwaltung';
        const nextName = label || normStr(prev.name) || code;
        if (normStr(prev.tier).toLowerCase() !== nextTier || normStr(prev.name) !== nextName) {
            changed.roles += 1;
        }
        roles[idx] = Object.assign({}, prev, {
            code: normCode(prev.code) || code,
            name: nextName,
            tier: nextTier
        });
    });

    const arges = (Array.isArray(base.arges) ? base.arges : []).map(function (a) {
        return Object.assign({}, a);
    });
    org.nodes.forEach(function (n) {
        if (!isArgeNode(n) || normStr(n.refKind) !== 'arge') return;
        const code = normCode(n.refKey);
        if (!code) return;
        const idx = arges.findIndex(function (a) {
            return normCode(a && a.code) === code;
        });
        if (idx < 0) return;
        const label = normStr(n.label);
        if (label && normStr(arges[idx].name) !== label) {
            arges[idx] = Object.assign({}, arges[idx], { name: label });
            changed.arges += 1;
        }
    });

    const classes = (Array.isArray(base.classes) ? base.classes : []).map(function (c) {
        return Object.assign({}, c);
    });
    org.nodes.forEach(function (n) {
        if (!isClassNode(n) || normStr(n.refKind) !== 'class') return;
        const code = normCode(n.refKey);
        if (!code) return;
        const canon = canonicalSchulOrgClassCode(code) || code;
        const idx = classes.findIndex(function (c) {
            const cc = classCodeFromSettingsRow(c);
            return (
                cc === canon ||
                normCode(c && c.code) === code ||
                normCode(c && c.name) === code
            );
        });
        if (idx < 0) return;
        const label = normStr(n.label);
        if (label && normStr(classes[idx].name) !== label) {
            classes[idx] = Object.assign({}, classes[idx], { name: label });
            changed.classes += 1;
        }
    });

    let audiences = normalizeVerwaltungAudienceGroups(base.verwaltungAudienceGroups);
    org.nodes.forEach(function (n) {
        if (!n || !isAreaNode(n) || normStr(n.refKind) !== 'audience') return;
        const audId = normStr(n.refKey) || audienceIdFromSectionId(n.id);
        if (!audId || audId.indexOf('va-') !== 0) return;
        const label = normStr(n.label);
        if (!label) return;
        const idx = audiences.findIndex(function (g) {
            return normStr(g && g.id) === audId;
        });
        if (idx >= 0) {
            if (normStr(audiences[idx].label) !== label) {
                audiences[idx] = Object.assign({}, audiences[idx], { label: label });
                changed.audiences += 1;
            }
        } else {
            audiences = audiences.concat([
                { id: audId, label: label, builtin: false, order: 100 + audiences.length }
            ]);
            changed.audiences += 1;
        }
    });
    audiences = normalizeVerwaltungAudienceGroups(audiences);

    /** @type {object} */
    const next = Object.assign({}, base, {
        adminRoles: roles,
        arges: arges,
        classes: classes,
        verwaltungAudienceGroups: audiences,
        schulOrganigramm: org
    });

    if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
        next.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(
            roles,
            Array.isArray(base.admin) ? base.admin : []
        );
    }

    return { settings: next, layout: org, changed: changed };
}

/**
 * Gespeichertes Layout mit fehlenden Stammdaten-Knoten ergänzen.
 * Bei Aufblähung/Müll-Refs: hart aus Stammdaten neu aufbauen.
 * @param {unknown} layout
 * @param {object} settings
 * @returns {{ layout: { version: number, nodes: object[] }, added: number, labelsUpdated: number, removed: number, repaired: number, healed: boolean }}
 */
export function syncSchulOrganigrammWithStammdaten(layout, settings) {
    const beforeCount = layout && Array.isArray(layout.nodes) ? layout.nodes.length : 0;
    if (schulOrganigrammNeedsHeal(layout, settings)) {
        const rebuilt = rebuildSchulOrganigrammFromStammdaten(layout, settings);
        const cleaned = stripPhantomSchulOrgClassAreas(rebuilt, settings);
        const withJahrgang = ensureSchulOrganigrammJahrgangStructure(cleaned.layout, settings);
        const withEltern = ensureSchulOrganigrammElternStructure(withJahrgang, settings);
        const withChats = ensureSchulOrganigrammKlassenchatsStructure(withEltern, settings);
        const withArges = ensureSchulOrganigrammArgeStructure(withChats, settings);
        const withFnAgg = ensureSchulOrganigrammFunctionAggregateStructure(withArges, settings);
        return {
            layout: withFnAgg,
            added: 0,
            labelsUpdated: 0,
            removed: Math.max(0, beforeCount - withJahrgang.nodes.length),
            repaired: beforeCount,
            healed: true
        };
    }

    const phantoms = stripPhantomSchulOrgClassAreas(layout, settings);
    const classMerged = mergeDuplicateSchulOrganigrammClasses(phantoms.layout);
    const deduped = dedupeSchulOrganigrammLayout(classMerged.layout);
    const base = deduped.layout;
    if (!base.nodes.length) {
        const seeded = seedSchulOrganigrammFromStammdaten(settings);
        return {
            layout: seeded,
            added: seeded.nodes.length,
            labelsUpdated: 0,
            removed: deduped.removed + classMerged.merged + phantoms.stripped,
            repaired: deduped.repaired,
            healed: deduped.removed > 0 || classMerged.merged > 0 || phantoms.stripped > 0
        };
    }
    const seeded = seedSchulOrganigrammFromStammdaten(settings);
    const byRef = new Map();
    base.nodes.forEach(function (n) {
        if (n.refKind && n.refKind !== 'none' && cleanRefToken(n.refKey)) {
            byRef.set(n.refKind + ':' + String(n.refKey).toUpperCase(), n);
        }
    });
    const byId = new Map(
        base.nodes.map(function (n) {
            return [n.id, n];
        })
    );
    let added = 0;
    let labelsUpdated = 0;
    const removedAreaIds = new Set(schulOrganigrammRemovedAreaIds(base));
    seeded.nodes.forEach(function (n) {
        if (n.type === NODE.SECTION) {
            const existingSec = byId.get(n.id);
            if (!existingSec) {
                // Vom Nutzer entfernte Standard-Bereiche (z. B. Verwaltung) nicht erneut anlegen
                if (removedAreaIds.has(String(n.id))) return;
                base.nodes.push(n);
                byId.set(n.id, n);
                if (n.refKind && n.refKind !== 'none' && cleanRefToken(n.refKey)) {
                    byRef.set(n.refKind + ':' + String(n.refKey).toUpperCase(), n);
                }
                added += 1;
                return;
            }
            const nextLabel = normStr(n.label);
            // Feste Bereiche (z. B. ARGEs): Nutzer-Umbenennung behalten – nicht Seed-Label zurücksetzen
            if (
                nextLabel &&
                nextLabel !== existingSec.label &&
                !BUILTIN_AREA_IDS[String(existingSec.id || n.id)]
            ) {
                existingSec.label = nextLabel;
                labelsUpdated += 1;
            }
            // Root-Positionen der Landkarte nie aus dem Seed überschreiben
            if (isAreaNode(existingSec) && (existingSec.parentId == null || existingSec.parentId === '')) {
                const ex = parseSchulOrgCanvasCoord(existingSec.x);
                const ey = parseSchulOrgCanvasCoord(existingSec.y);
                if (ex != null) existingSec.x = ex;
                if (ey != null) existingSec.y = ey;
            }
            if (n.refKind && n.refKind !== 'none' && n.refKey) {
                if (!existingSec.refKind || existingSec.refKind === 'none') {
                    existingSec.refKind = n.refKind;
                    existingSec.refKey = n.refKey;
                }
            }
            if (n.tone && !existingSec.tone) existingSec.tone = n.tone;
            if (n.meta && typeof n.meta === 'object') {
                existingSec.meta = Object.assign({}, existingSec.meta || {}, n.meta);
            }
            // M365: explizit gelöste Verknüpfung (groupMode=none) nie durch Seed überschreiben
            if (n.m365 && typeof n.m365 === 'object') {
                const exMode = existingSec.m365
                    ? normStr(existingSec.m365.groupMode).toLowerCase()
                    : '';
                if (exMode === 'none') {
                    existingSec.m365 = Object.assign({}, existingSec.m365, {
                        groupMode: 'none',
                        groupId: null,
                        groupLabel: null,
                        groupMail: null
                    });
                } else if (!existingSec.m365) {
                    existingSec.m365 = Object.assign({}, n.m365);
                }
            }
            // Jahrgangs-Bereiche unter „Schüler“ halten – Position immer Auto-Pack (keine freien Orte)
            if (isJahrgangArea(n) || isJahrgangArea(existingSec)) {
                const stufe = jahrgangKeyForArea(n) || jahrgangKeyForArea(existingSec);
                if (stufe) {
                    existingSec.parentId = SCHUL_ORG_SEC_SCHUELER;
                    existingSec.refKind = 'jahrgang';
                    existingSec.refKey = stufe;
                    const wantLabel = jahrgangAreaLabel(stufe);
                    if (existingSec.label !== wantLabel) {
                        existingSec.label = wantLabel;
                        labelsUpdated += 1;
                    }
                    delete existingSec.x;
                    delete existingSec.y;
                }
            }
            return;
        }
        const key = n.refKind + ':' + String(n.refKey || '').toUpperCase();
        if (!(n.refKind && n.refKind !== 'none' && n.refKey)) return;
        let existing = byRef.get(key);
        // Gleiche Rolle ggf. nur per Id vorhanden (andere Parent-Platzierung behalten)
        if (!existing && n.id && byId.has(String(n.id))) {
            existing = byId.get(String(n.id));
            byRef.set(key, existing);
        }
        if (!existing) {
            base.nodes.push(n);
            byRef.set(key, n);
            byId.set(String(n.id), n);
            added += 1;
            return;
        }
        const nextLabel = normStr(n.label);
        if (nextLabel && nextLabel !== existing.label) {
            existing.label = nextLabel;
            labelsUpdated += 1;
        }
        if (n.meta && typeof n.meta === 'object') {
            existing.meta = Object.assign({}, existing.meta || {}, n.meta);
        }
        // Kursteam bleibt unter der Platzierungs-Klasse
        if (n.type === NODE.KURSTEAM && n.parentId) {
            existing.parentId = n.parentId;
        }
        if (n.m365 && typeof n.m365 === 'object') {
            const exMode = existing.m365
                ? normStr(existing.m365.groupMode).toLowerCase()
                : '';
            // Explizit gelöste Gruppenverknüpfung nicht durch Seed/Sync zurücksetzen
            if (exMode === 'none') {
                existing.m365 = Object.assign({}, existing.m365, {
                    groupMode: 'none',
                    groupId: null,
                    groupLabel: null,
                    groupMail: null
                });
            } else {
                existing.m365 = Object.assign({}, existing.m365 || {}, n.m365);
            }
        }
    });

    const seededKursteamKeys = new Set(
        seeded.nodes
            .filter(function (n) {
                return n.type === NODE.KURSTEAM || n.refKind === 'kursteam';
            })
            .map(function (n) {
                return 'kursteam:' + String(n.refKey || '').toUpperCase();
            })
    );
    const seededClassKeys = new Set(
        seeded.nodes
            .filter(function (n) {
                return n.type === NODE.CLASS || n.refKind === 'class';
            })
            .map(function (n) {
                const c = canonicalSchulOrgClassCode(n.refKey) || String(n.refKey || '').toUpperCase();
                return 'class:' + c;
            })
    );

    // Alte SV-/SGA-Personen-Kacheln entfernen; SV-Sitze (Schulsprecher/STV) behalten
    // Veraltete Kursteams (nicht mehr in Belegung) und verwaiste Klassen entfernen
    const beforeStrip = base.nodes.length;
    base.nodes = base.nodes.filter(function (n) {
        if (n.type === NODE.KURSTEAM || n.refKind === 'kursteam') {
            const key = 'kursteam:' + String(n.refKey || '').toUpperCase();
            return seededKursteamKeys.has(key);
        }
        if (n.type === NODE.CLASS || n.refKind === 'class') {
            const c = canonicalSchulOrgClassCode(n.refKey) || String(n.refKey || '').toUpperCase();
            if (!c) return false;
            return seededClassKeys.has('class:' + c);
        }
        if (n.type === NODE.COUNCIL) {
            if (isSvSeatNode(n)) return true;
            // Legacy council_board / freie SV-Personen-Kacheln entfallen
            return false;
        }
        if (n.type === NODE.SGA) {
            return n.id === SCHUL_ORG_SGA_BOARD_ID || n.refKey === SCHUL_ORG_BOARD_REF;
        }
        return true;
    });
    const stripped = beforeStrip - base.nodes.length;
    const withSga = ensureSchulOrganigrammSgaStructure(base, settings);
    const phantomsAfter = stripPhantomSchulOrgClassAreas(withSga, settings);
    const withJahrgang = ensureSchulOrganigrammJahrgangStructure(phantomsAfter.layout, settings);
    const withEltern = ensureSchulOrganigrammElternStructure(withJahrgang, settings);
    const withChats = ensureSchulOrganigrammKlassenchatsStructure(withEltern, settings);
    const withArges = ensureSchulOrganigrammArgeStructure(withChats, settings);
    const withFnAgg = ensureSchulOrganigrammFunctionAggregateStructure(withArges, settings);

    return {
        layout: withFnAgg,
        added: added,
        labelsUpdated: labelsUpdated,
        removed:
            deduped.removed +
            classMerged.merged +
            stripped +
            phantoms.stripped +
            phantomsAfter.stripped,
        repaired: deduped.repaired,
        healed:
            stripped > 0 ||
            classMerged.merged > 0 ||
            phantoms.stripped > 0 ||
            phantomsAfter.stripped > 0
    };
}

/**
 * Roh-Layout des Legacy-Verwaltungs-Organigramms (group/role/sammel) leicht normalisieren.
 * @param {unknown} raw
 */
function normalizeVerwaltungLayoutLoose(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const nodesIn = Array.isArray(o.nodes) ? o.nodes : Array.isArray(o.groups) ? o.groups : [];
    const seen = new Set();
    const nodes = [];
    nodesIn.forEach(function (n, idx) {
        if (!n || typeof n !== 'object') return;
        let id = normStr(n.id);
        if (!id || seen.has(id)) id = newId('vw');
        seen.add(id);
        const type = normStr(n.type).toLowerCase();
        const kind =
            type === 'role' || type === 'sammel' || type === 'group' ? type : 'group';
        nodes.push({
            id: id,
            type: kind,
            label: normStr(n.label || n.name) || (kind === 'role' ? 'Rolle' : 'Gruppe'),
            parentId: n.parentId == null || n.parentId === '' ? null : normStr(n.parentId),
            order: typeof n.order === 'number' ? n.order : idx,
            tone: normStr(n.tone),
            roleCode: normCode(n.roleCode || n.code),
            audienceGroupId: normStr(n.audienceGroupId),
            m365GroupId: normStr(n.m365GroupId),
            m365GroupLabel: normStr(n.m365GroupLabel),
            m365GroupMail: normStr(n.m365GroupMail).toLowerCase()
        });
    });
    nodes.sort(function (a, b) {
        return (a.order || 0) - (b.order || 0);
    });
    return { version: 1, nodes: nodes };
}

/**
 * @param {object} n – Verwaltung-Knoten (group)
 * @returns {string|null} kanonische Section-Id oder null
 */
function canonicalSectionIdFromVerwaltungGroup(n) {
    const tone = normStr(n && n.tone).toLowerCase();
    const label = normStr(n && n.label);
    if (tone === 'lead' || /schulleitung/i.test(label)) return 'sec_schulleitung';
    if (tone === 'staff' || /verwaltung/i.test(label)) return 'sec_verwaltung';
    return null;
}

/**
 * Legacy `verwaltungOrganigramm` → Schul-Organigramm-Knoten (ohne Stammdaten-Sync).
 * @param {unknown} verwaltungRaw
 * @returns {{ layout: { version: number, nodes: object[] }, stats: { sections: number, roles: number, sammel: number, custom: number } }}
 */
export function convertVerwaltungOrganigrammToSchul(verwaltungRaw) {
    const src = normalizeVerwaltungLayoutLoose(verwaltungRaw);
    /** @type {Map<string, string>} */
    const idMap = new Map();
    /** @type {Set<string>} */
    const usedIds = new Set();
    const stats = { sections: 0, roles: 0, sammel: 0, custom: 0 };

    function allocId(preferred, fallbackPrefix) {
        let id = preferred;
        if (!id || usedIds.has(id)) id = newId(fallbackPrefix || 'n');
        let i = 2;
        while (usedIds.has(id)) {
            id = (preferred || fallbackPrefix || 'n') + '_' + i;
            i += 1;
        }
        usedIds.add(id);
        return id;
    }

    src.nodes.forEach(function (n) {
        if (n.type === 'group' && !n.parentId) {
            const canonical = canonicalSectionIdFromVerwaltungGroup(n);
            const nextId = allocId(canonical || 'sec_migrated_' + n.id.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 36), 'sec');
            idMap.set(n.id, nextId);
            return;
        }
        if (n.type === 'group') {
            idMap.set(n.id, allocId('grp_' + n.id.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 36), 'grp'));
            return;
        }
        if (n.type === 'role') {
            const code = n.roleCode || normCode(n.label);
            idMap.set(n.id, allocId(code ? 'role_' + code : '', 'role'));
            return;
        }
        if (n.type === 'sammel') {
            idMap.set(n.id, allocId('sammel_' + (n.audienceGroupId || n.m365GroupId || n.id).slice(0, 40), 'sammel'));
            return;
        }
        idMap.set(n.id, allocId('custom_' + n.id.slice(0, 36), 'custom'));
    });

    const outNodes = [];
    src.nodes.forEach(function (n) {
        const id = idMap.get(n.id);
        const parentId = n.parentId ? idMap.get(n.parentId) || null : null;
        if (n.type === 'group') {
            if (!n.parentId) {
                const tone =
                    id === 'sec_schulleitung'
                        ? 'lead'
                        : id === 'sec_verwaltung'
                          ? 'staff'
                          : n.tone || 'custom';
                outNodes.push({
                    id: id,
                    type: NODE.SECTION,
                    label: n.label,
                    parentId: null,
                    order: n.order,
                    tone: tone,
                    refKind: 'none',
                    refKey: ''
                });
                stats.sections += 1;
            } else {
                outNodes.push({
                    id: id,
                    type: NODE.CUSTOM,
                    label: n.label,
                    parentId: parentId,
                    order: n.order,
                    tone: n.tone || 'custom',
                    refKind: 'none',
                    refKey: ''
                });
                stats.custom += 1;
            }
            return;
        }
        if (n.type === 'role') {
            const code = n.roleCode || normCode(n.label);
            outNodes.push({
                id: id,
                type: NODE.ROLE,
                label: n.label,
                parentId: parentId || 'sec_verwaltung',
                order: n.order,
                tone: n.tone || (parentId === 'sec_schulleitung' ? 'lead' : 'staff'),
                refKind: 'adminRole',
                refKey: code || n.label
            });
            stats.roles += 1;
            return;
        }
        if (n.type === 'sammel') {
            outNodes.push({
                id: id,
                type: NODE.CUSTOM,
                label: n.label,
                parentId: parentId,
                order: n.order,
                tone: 'sammel',
                refKind: n.audienceGroupId ? 'audience' : 'none',
                refKey: n.audienceGroupId || n.m365GroupId || '',
                meta: {
                    audienceGroupId: n.audienceGroupId,
                    m365GroupId: n.m365GroupId,
                    m365GroupLabel: n.m365GroupLabel,
                    m365GroupMail: n.m365GroupMail
                }
            });
            stats.sammel += 1;
            return;
        }
        outNodes.push({
            id: id,
            type: NODE.CUSTOM,
            label: n.label,
            parentId: parentId,
            order: n.order,
            tone: n.tone || 'custom',
            refKind: 'none',
            refKey: ''
        });
        stats.custom += 1;
    });

    return {
        layout: normalizeSchulOrganigramm({ version: SCHUL_ORG_VERSION, nodes: outNodes }),
        stats: stats
    };
}

/**
 * @param {object} settings
 */
export function hasVerwaltungOrganigrammLayout(settings) {
    const v = settings && settings.verwaltungOrganigramm;
    return !!(v && Array.isArray(v.nodes) && v.nodes.length);
}

/**
 * Verwaltung-Layout übernehmen und mit Stammdaten synchronisieren.
 * @param {object} settings
 * @returns {{ layout: { version: number, nodes: object[] }, stats: object, syncedAdded: number, syncedLabels: number }|null}
 */
export function migrateVerwaltungOrganigrammToSchul(settings) {
    if (!hasVerwaltungOrganigrammLayout(settings)) return null;
    const converted = convertVerwaltungOrganigrammToSchul(settings.verwaltungOrganigramm);
    const synced = syncSchulOrganigrammWithStammdaten(converted.layout, settings);
    return {
        layout: synced.layout,
        stats: converted.stats,
        syncedAdded: synced.added,
        syncedLabels: synced.labelsUpdated
    };
}

/**
 * Layout laden: Seed, Migration aus Verwaltung, oder Sync; Persistenz-Hinweis in dirty.
 * @param {unknown} raw
 * @param {object} settings
 */
export function resolveSchulOrganigrammLayout(raw, settings) {
    if (raw && Array.isArray(raw.nodes) && raw.nodes.length) {
        const synced = syncSchulOrganigrammWithStammdaten(raw, settings);
        const cleaned =
            (synced.removed || 0) + (synced.repaired || 0) > 0 || !!synced.healed;
        const prevVer = Number(raw.version) || 1;
        const needsV2Migrate =
            prevVer < SCHUL_ORG_VERSION ||
            !raw.view ||
            (raw.nodes || []).some(function (n) {
                return normStr(n && n.type).toLowerCase() === 'section';
            });
        return {
            layout: synced.layout,
            seeded: false,
            dirty:
                synced.added > 0 ||
                synced.labelsUpdated > 0 ||
                synced.removed > 0 ||
                synced.repaired > 0 ||
                !!synced.healed ||
                needsV2Migrate,
            migratedFromVerwaltung: false,
            deduped: cleaned,
            healed: !!synced.healed,
            removedDuplicates: synced.removed || 0
        };
    }
    const migrated = migrateVerwaltungOrganigrammToSchul(settings || {});
    if (migrated) {
        return {
            layout: migrated.layout,
            seeded: false,
            dirty: true,
            migratedFromVerwaltung: true,
            migrationStats: migrated.stats
        };
    }
    const seeded = seedSchulOrganigrammFromStammdaten(settings);
    return {
        layout: seeded,
        seeded: true,
        dirty: true,
        migratedFromVerwaltung: false
    };
}

/**
 * Kurzstatistik für Meta-Zeile / IT-Abgleich.
 * @param {{ nodes?: object[] }} layout
 */
export function schulOrganigrammStats(layout) {
    const nodes = normalizeSchulOrganigramm(layout).nodes;
    const counts = {
        sections: 0,
        roles: 0,
        arges: 0,
        classes: 0,
        kursteams: 0,
        other: 0,
        total: nodes.length,
        content: 0
    };
    nodes.forEach(function (n) {
        if (isArgeAreaNode(n) || n.type === NODE.ARGE) counts.arges += 1;
        else if (isClassAreaNode(n) || n.type === NODE.CLASS) counts.classes += 1;
        else if (n.type === NODE.SECTION) counts.sections += 1;
        else if (n.type === NODE.ROLE) counts.roles += 1;
        else if (n.type === NODE.KURSTEAM) counts.kursteams += 1;
        else counts.other += 1;
    });
    counts.content =
        counts.roles + counts.arges + counts.classes + counts.kursteams + counts.other;
    return counts;
}

/**
 * Bereiche, die Funktionen/Rollen aufnehmen dürfen (für Auswahl beim Anlegen / Verschieben).
 * @param {unknown} layout
 * @returns {object[]}
 */
export function listSchulOrgRoleParentAreas(layout) {
    const base = normalizeSchulOrganigramm(layout);
    return base.nodes
        .filter(function (n) {
            return (
                isAreaNode(n) &&
                canPlaceSchulOrgNodeInSection({ type: NODE.ROLE, refKind: 'adminRole' }, n.id, base)
            );
        })
        .sort(function (a, b) {
            return String(a.label || a.id).localeCompare(String(b.label || b.id), 'de');
        });
}

/**
 * @param {object} node
 * @param {string|null} sectionId
 */
export function canPlaceSchulOrgNodeInSection(node, sectionId, layout) {
    const sid = sectionId == null || sectionId === '' ? null : String(sectionId);
    if (!node || !sid) return false;
    const t = normStr(node.type).toLowerCase();
    if (t === NODE.ROLE) {
        if (isRoleCapableSection(sid)) return true;
        const base = layout ? normalizeSchulOrganigramm(layout) : null;
        const sec = base
            ? base.nodes.find(function (n) {
                  return n.id === sid;
              })
            : null;
        if (isRoleCapableSection(sid, sec || undefined)) return true;
        // Verschachtelte Areas unter Schulleitung/Verwaltung/Audience dürfen Rollen aufnehmen
        if (base && sec && isAreaNode(sec)) {
            let walk = sec;
            let guard = 0;
            while (walk && guard < 40) {
                guard += 1;
                if (isRoleCapableSection(walk.id, walk)) return true;
                const pid =
                    walk.parentId == null || walk.parentId === '' ? '' : String(walk.parentId);
                if (!pid) break;
                walk =
                    base.nodes.find(function (n) {
                        return n.id === pid;
                    }) || null;
            }
            // Freie / Custom-Areas: Rollen erlaubt – nicht in ARGE/Lehrer/Schüler/SGA/Jahrgang/Klasse
            // und nicht in Funktions-Sammelbereichen (Klassensprecher, KV, ARGE-Leiter, …)
            const rk = normStr(sec.refKind).toLowerCase();
            const forbiddenId =
                sid === 'sec_arge' ||
                sid === SCHUL_ORG_SEC_LEHRER ||
                sid === SCHUL_ORG_SEC_SCHUELER ||
                sid === SCHUL_ORG_SEC_ELTERN ||
                sid === SCHUL_ORG_SEC_KLASSENCHATS ||
                sid === 'sec_vertretung' ||
                sid === 'sec_klassen' ||
                sid === SCHUL_ORG_SGA_BOARD_ID ||
                sid === SCHUL_ORG_SEC_KLASSENVORSTAND ||
                /^sga_/i.test(sid) ||
                /^sec_jahrgang_/i.test(sid) ||
                /^sec_eltern_jahrgang_/i.test(sid) ||
                /^sec_klassenchat_jahrgang_/i.test(sid) ||
                /^class_/i.test(sid) ||
                /^eltern_class_/i.test(sid) ||
                /^klassenchat_class_/i.test(sid) ||
                /^sec_classfn_/i.test(sid) ||
                /^sec_argefn_/i.test(sid);
            if (
                !forbiddenId &&
                !isSammelgruppeArea(sec) &&
                !isFunctionAggregateArea(sec) &&
                !isJahrgangArea(sec) &&
                !isElternJahrgangArea(sec) &&
                !isElternClassNode(sec) &&
                !isKlassenvorstandArea(sec) &&
                !isKlassenchatJahrgangArea(sec) &&
                !isKlassenchatClassNode(sec) &&
                !isSgaBoardArea(sec) &&
                !isSgaScopeArea(sec) &&
                rk !== 'class' &&
                rk !== 'jahrgang' &&
                rk !== 'elternjahrgang' &&
                rk !== 'elternclass' &&
                rk !== 'klassenchatjahrgang' &&
                rk !== 'klassenchatclass' &&
                rk !== 'sammelgruppe' &&
                rk !== 'arge' &&
                rk !== 'sga' &&
                rk !== 'classfunction' &&
                rk !== 'argefunction' &&
                rk !== 'klassenvorstand' &&
                rk !== 'argeleiter' &&
                normStr(sec.type).toLowerCase() !== NODE.CLASS &&
                normStr(sec.type).toLowerCase() !== NODE.ARGE
            ) {
                return true;
            }
        }
        return false;
    }
    if (t === NODE.ARGE || isArgeAreaNode(node)) {
        if (sid === 'sec_arge') return true;
        const base = layout ? normalizeSchulOrganigramm(layout) : null;
        const sec = base
            ? base.nodes.find(function (n) {
                  return n.id === sid;
              })
            : null;
        if (!base || !sec || !isAreaNode(sec)) return false;
        // Nicht in Lehrer/Schüler/SGA/Jahrgang/Klassen/Leitung-Personal
        const forbidden =
            sid === SCHUL_ORG_SEC_LEHRER ||
            sid === SCHUL_ORG_SEC_SCHUELER ||
            sid === 'sec_vertretung' ||
            sid === 'sec_klassen' ||
            sid === 'sec_schulleitung' ||
            sid === 'sec_verwaltung' ||
            sid === SCHUL_ORG_SGA_BOARD_ID ||
            sid === SCHUL_ORG_SEC_ARGE_LEITER ||
            /^sga_/i.test(sid) ||
            /^sec_jahrgang_/i.test(sid) ||
            /^class_/i.test(sid) ||
            isSammelgruppeArea(sec) ||
            isJahrgangArea(sec) ||
            isSgaBoardArea(sec) ||
            isSgaScopeArea(sec) ||
            isArgeAreaNode(sec) ||
            isArgeLeiterArea(sec) ||
            normStr(sec.refKind).toLowerCase() === 'class' ||
            normStr(sec.refKind).toLowerCase() === 'jahrgang' ||
            normStr(sec.refKind).toLowerCase() === 'sammelgruppe' ||
            normStr(sec.refKind).toLowerCase() === 'sga' ||
            normStr(sec.type).toLowerCase() === NODE.CLASS;
        if (forbidden) return false;
        // Unterbereiche von „ARGEs / Fachschaften“ (z. B. Fachleitungen GS/Gym) + tone arge
        let walk = sec;
        let guard = 0;
        while (walk && guard < 40) {
            guard += 1;
            if (String(walk.id) === 'sec_arge') return true;
            if (normStr(walk.tone).toLowerCase() === 'arge') return true;
            const pid =
                walk.parentId == null || walk.parentId === '' ? '' : String(walk.parentId);
            if (!pid) break;
            walk =
                base.nodes.find(function (n) {
                    return n.id === pid;
                }) || null;
        }
        return false;
    }
    if (t === NODE.CLASS || isClassAreaNode(node)) {
        if (sid === 'sec_klassen' || sid === SCHUL_ORG_SEC_SCHUELER) return true;
        if (/^sec_jahrgang_/i.test(sid)) return true;
        const base = layout ? normalizeSchulOrganigramm(layout) : null;
        const sec = base
            ? base.nodes.find(function (n) {
                  return n.id === sid;
              })
            : null;
        return isJahrgangArea(sec);
    }
    if (t === NODE.KURSTEAM) {
        if (/^class_/i.test(sid)) return true;
        const base = layout ? normalizeSchulOrganigramm(layout) : null;
        const parent = base
            ? base.nodes.find(function (n) {
                  return n.id === sid;
              })
            : null;
        return !!(parent && isClassNode(parent));
    }
    if (t === NODE.COUNCIL || t === NODE.SGA) {
        return sid === 'sec_vertretung' || sid === SCHUL_ORG_SGA_BOARD_ID || sid.indexOf('sga_gruppe_') === 0;
    }
    if (t === NODE.PERSON) return !!sid;
    if (t === NODE.CUSTOM) return true;
    return false;
}

/**
 * Karte verschieben / neu ordnen (nicht für Sections).
 * @param {unknown} layout
 * @param {string} nodeId
 * @param {string|null} newParentId
 * @param {string} [beforeNodeId] – vor diesem Geschwister einfügen; leer = ans Ende
 */
export function moveSchulOrganigrammNode(layout, nodeId, newParentId, beforeNodeId) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(nodeId);
    if (!id) return base;
    const node = base.nodes.find(function (n) {
        return n.id === id;
    });
    if (!node || isAreaNode(node)) return base;
    const parentId = newParentId == null || newParentId === '' ? null : String(newParentId);
    const parent = base.nodes.find(function (n) {
        return n.id === parentId;
    });
    if (!parent || !isAreaNode(parent)) return base;
    if (!canPlaceSchulOrgNodeInSection(node, parentId, base)) return base;

    node.parentId = parentId;
    // Alte Koordinaten verwerfen – Drop setzt Relativ-Position neu; sonst Auto-Pack
    delete node.x;
    delete node.y;
    const beforeId = normStr(beforeNodeId);
    const siblings = base.nodes
        .filter(function (n) {
            return n.id !== id && String(n.parentId || '') === String(parentId);
        })
        .sort(function (a, b) {
            return (a.order || 0) - (b.order || 0);
        });
    let insertAt = siblings.length;
    if (beforeId) {
        const bi = siblings.findIndex(function (n) {
            return n.id === beforeId;
        });
        if (bi >= 0) insertAt = bi;
    }
    siblings.splice(insertAt, 0, node);
    siblings.forEach(function (n, i) {
        n.order = i;
    });
    return normalizeSchulOrganigramm(base);
}

/**
 * Root-Bereiche (Sections) neu ordnen.
 * @param {unknown} layout
 * @param {string} sectionId
 * @param {string} [beforeSectionId]
 */
export function reorderSchulOrganigrammSections(layout, sectionId, beforeSectionId) {
    const base = normalizeSchulOrganigramm(layout);
    const id = normStr(sectionId);
    if (!id) return base;
    const section = base.nodes.find(function (n) {
        return n.id === id && isAreaNode(n);
    });
    if (!section) return base;
    const beforeId = normStr(beforeSectionId);
    const roots = base.nodes
        .filter(function (n) {
            return isAreaNode(n) && n.id !== id && (n.parentId == null || n.parentId === '');
        })
        .sort(function (a, b) {
            return (a.order || 0) - (b.order || 0);
        });
    let insertAt = roots.length;
    if (beforeId) {
        const bi = roots.findIndex(function (n) {
            return n.id === beforeId;
        });
        if (bi >= 0) insertAt = bi;
    }
    roots.splice(insertAt, 0, section);
    roots.forEach(function (n, i) {
        n.order = i;
        n.parentId = null;
    });
    return normalizeSchulOrganigramm(base);
}

/**
 * Textfilter über Label, Kürzel und Meta-Text.
 * @param {object} node
 * @param {string} query
 * @param {string} [metaText]
 */
export function nodeMatchesSchulOrgFilter(node, query, metaText) {
    const q = String(query || '')
        .trim()
        .toLowerCase();
    if (!q) return true;
    const blob = [node && node.label, node && node.refKey, node && node.type, metaText]
        .map(function (x) {
            return String(x || '');
        })
        .join(' ')
        .toLowerCase();
    return blob.indexOf(q) !== -1;
}

const api = {
    SCHUL_ORG_VERSION: SCHUL_ORG_VERSION,
    SCHUL_ORG_BOARD_REF: SCHUL_ORG_BOARD_REF,
    SCHUL_ORG_SEC_LEHRER: SCHUL_ORG_SEC_LEHRER,
    SCHUL_ORG_SEC_SCHUELER: SCHUL_ORG_SEC_SCHUELER,
    SCHUL_ORG_SEC_ELTERN: SCHUL_ORG_SEC_ELTERN,
    SCHUL_ORG_SEC_KLASSENCHATS: SCHUL_ORG_SEC_KLASSENCHATS,
    SCHUL_ORG_SEC_ARGE: SCHUL_ORG_SEC_ARGE,
    SCHUL_ORG_SEC_ARGE_LEITER: SCHUL_ORG_SEC_ARGE_LEITER,
    SCHUL_ORG_SEC_KLASSENVORSTAND: SCHUL_ORG_SEC_KLASSENVORSTAND,
    SCHUL_ORG_SGA_BOARD_ID: SCHUL_ORG_SGA_BOARD_ID,
    SCHUL_ORG_SGA_SV_AREA_ID: SCHUL_ORG_SGA_SV_AREA_ID,
    SCHUL_ORG_SGA_SCOPE_AREAS: SCHUL_ORG_SGA_SCOPE_AREAS,
    SCHUL_ORG_SV_SEATS: SCHUL_ORG_SV_SEATS,
    NODE: NODE,
    kursteamStableKey: kursteamStableKey,
    kursteamLabelFromRow: kursteamLabelFromRow,
    kursteamPlacementsFromBelegung: kursteamPlacementsFromBelegung,
    isAreaNode: isAreaNode,
    isSammelgruppeArea: isSammelgruppeArea,
    sammelgruppeKeyForArea: sammelgruppeKeyForArea,
    sgaScopeForArea: sgaScopeForArea,
    isSgaScopeArea: isSgaScopeArea,
    isSgaBoardArea: isSgaBoardArea,
    isSvSeatNode: isSvSeatNode,
    svSeatForNode: svSeatForNode,
    personForSvSeat: personForSvSeat,
    ensureSchulOrganigrammSgaStructure: ensureSchulOrganigrammSgaStructure,
    ensureSchulOrganigrammJahrgangStructure: ensureSchulOrganigrammJahrgangStructure,
    ensureSchulOrganigrammElternStructure: ensureSchulOrganigrammElternStructure,
    ensureSchulOrganigrammKlassenchatsStructure: ensureSchulOrganigrammKlassenchatsStructure,
    ensureSchulOrganigrammArgeLeiterStructure: ensureSchulOrganigrammArgeLeiterStructure,
    ensureSchulOrganigrammKlassenvorstandStructure: ensureSchulOrganigrammKlassenvorstandStructure,
    ensureSchulOrganigrammClassFunctionStructure: ensureSchulOrganigrammClassFunctionStructure,
    ensureSchulOrganigrammArgeFunctionStructure: ensureSchulOrganigrammArgeFunctionStructure,
    ensureSchulOrganigrammFunctionAggregateStructure: ensureSchulOrganigrammFunctionAggregateStructure,
    SCHUL_ORG_CLASS_FN_SEC_PREFIX: SCHUL_ORG_CLASS_FN_SEC_PREFIX,
    SCHUL_ORG_ARGE_FN_SEC_PREFIX: SCHUL_ORG_ARGE_FN_SEC_PREFIX,
    SCHUL_ORG_DEFAULT_CLASS_FUNCTIONS: SCHUL_ORG_DEFAULT_CLASS_FUNCTIONS,
    SCHUL_ORG_DEFAULT_ARGE_FUNCTIONS: SCHUL_ORG_DEFAULT_ARGE_FUNCTIONS,
    cleanSchulOrgClassFunctionId: cleanSchulOrgClassFunctionId,
    classFunctionBoardId: classFunctionBoardId,
    argeFunctionBoardId: argeFunctionBoardId,
    normalizeClassFunctionsMap: normalizeClassFunctionsMap,
    listSchulOrgClassFunctions: listSchulOrgClassFunctions,
    listSchulOrgArgeFunctions: listSchulOrgArgeFunctions,
    listSchulOrgFunctionAggregates: listSchulOrgFunctionAggregates,
    classFunctionIdForArea: classFunctionIdForArea,
    argeFunctionIdForArea: argeFunctionIdForArea,
    isClassFunctionArea: isClassFunctionArea,
    isArgeFunctionArea: isArgeFunctionArea,
    isFunctionAggregateArea: isFunctionAggregateArea,
    functionAggregateDefForArea: functionAggregateDefForArea,
    collectClassFunctionPeopleFromSettings: collectClassFunctionPeopleFromSettings,
    collectArgeFunctionPeopleFromSettings: collectArgeFunctionPeopleFromSettings,
    collectFunctionAggregatePeople: collectFunctionAggregatePeople,
    ensureSchulOrganigrammArgeStructure: ensureSchulOrganigrammArgeStructure,
    isArgeLeiterArea: isArgeLeiterArea,
    isKlassenvorstandArea: isKlassenvorstandArea,
    isArgeAreaNode: isArgeAreaNode,
    isArgeNode: isArgeNode,
    argeNodeId: argeNodeId,
    collectArgeLeiterPeopleFromSettings: collectArgeLeiterPeopleFromSettings,
    collectKlassenvorstandPeopleFromSettings: collectKlassenvorstandPeopleFromSettings,
    listArgeCatalogGroupRefs: listArgeCatalogGroupRefs,
    schulstufeFromClassCode: schulstufeFromClassCode,
    jahrgangAreaId: jahrgangAreaId,
    jahrgangAreaLabel: jahrgangAreaLabel,
    jahrgangKeyForArea: jahrgangKeyForArea,
    isJahrgangArea: isJahrgangArea,
    elternJahrgangAreaId: elternJahrgangAreaId,
    elternJahrgangKeyForArea: elternJahrgangKeyForArea,
    isElternJahrgangArea: isElternJahrgangArea,
    elternClassNodeId: elternClassNodeId,
    isElternClassNode: isElternClassNode,
    countUniqueGuardiansInSettings: countUniqueGuardiansInSettings,
    klassenchatJahrgangAreaId: klassenchatJahrgangAreaId,
    klassenchatJahrgangKeyForArea: klassenchatJahrgangKeyForArea,
    isKlassenchatJahrgangArea: isKlassenchatJahrgangArea,
    klassenchatClassNodeId: klassenchatClassNodeId,
    isKlassenchatClassNode: isKlassenchatClassNode,
    classChatMetaByKlasse: classChatMetaByKlasse,
    isClassAreaNode: isClassAreaNode,
    isClassNode: isClassNode,
    SCHUL_ORG_JAHRGANG_SONST: SCHUL_ORG_JAHRGANG_SONST,
    isBuiltinSchulOrganigrammArea: isBuiltinSchulOrganigrammArea,
    isOptionalDeletableBuiltinArea: isOptionalDeletableBuiltinArea,
    canDeleteSchulOrganigrammArea: canDeleteSchulOrganigrammArea,
    schulOrganigrammRemovedAreaIds: schulOrganigrammRemovedAreaIds,
    withRemovedSchulOrganigrammArea: withRemovedSchulOrganigrammArea,
    defaultCanvasSlot: defaultCanvasSlot,
    parseSchulOrgCanvasCoord: parseSchulOrgCanvasCoord,
    applyLandkarteRootPositionsFromBoxes: applyLandkarteRootPositionsFromBoxes,
    applyLandkartePositionsFromBoxes: applyLandkartePositionsFromBoxes,
    mergeSchulOrganigrammLivePositions: mergeSchulOrganigrammLivePositions,
    normalizeSchulOrganigramm: normalizeSchulOrganigramm,
    childrenOf: childrenOf,
    seedSchulOrganigrammFromStammdaten: seedSchulOrganigrammFromStammdaten,
    syncSchulOrganigrammWithStammdaten: syncSchulOrganigrammWithStammdaten,
    applySchulOrganigrammToStammdaten: applySchulOrganigrammToStammdaten,
    dedupeSchulOrganigrammLayout: dedupeSchulOrganigrammLayout,
    schulOrganigrammNeedsHeal: schulOrganigrammNeedsHeal,
    rebuildSchulOrganigrammFromStammdaten: rebuildSchulOrganigrammFromStammdaten,
    setSchulOrganigrammView: setSchulOrganigrammView,
    setSchulOrganigrammNodePosition: setSchulOrganigrammNodePosition,
    setSchulOrganigrammNodeSize: setSchulOrganigrammNodeSize,
    clampSchulOrganigrammAreaSize: clampSchulOrganigrammAreaSize,
    SCHUL_ORG_AREA_MIN_W: SCHUL_ORG_AREA_MIN_W,
    SCHUL_ORG_AREA_MAX_W: SCHUL_ORG_AREA_MAX_W,
    SCHUL_ORG_AREA_MIN_H: SCHUL_ORG_AREA_MIN_H,
    SCHUL_ORG_AREA_MAX_H: SCHUL_ORG_AREA_MAX_H,
    setSchulOrganigrammAreaCollapsed: setSchulOrganigrammAreaCollapsed,
    setSchulOrganigrammAreaM365: setSchulOrganigrammAreaM365,
    findSchulOrgNodesUsingM365Group: findSchulOrgNodesUsingM365Group,
    findSchulOrgNodesUsingM365Mailbox: findSchulOrgNodesUsingM365Mailbox,
    schulOrgM365ConflictLabel: schulOrgM365ConflictLabel,
    applyClassGroupLinksToSchulOrganigramm: applyClassGroupLinksToSchulOrganigramm,
    isSchulOrgLocalPlaceholderEmail: isSchulOrgLocalPlaceholderEmail,
    personHasM365UserLink: personHasM365UserLink,
    nodeM365LinkFlags: nodeM365LinkFlags,
    resolveClassChatIdForNode: resolveClassChatIdForNode,
    normalizeSchulOrgGroupKind: normalizeSchulOrgGroupKind,
    schulOrgGroupKindLabel: schulOrgGroupKindLabel,
    schulOrgGroupKindIcon: schulOrgGroupKindIcon,
    schulOrgMailboxIcon: schulOrgMailboxIcon,
    schulOrgGroupKindIsEntraGroup: schulOrgGroupKindIsEntraGroup,
    SCHUL_ORG_M365_OBJECT_KINDS: SCHUL_ORG_M365_OBJECT_KINDS,
    normalizeSchulOrgM365ObjectKind: normalizeSchulOrgM365ObjectKind,
    schulOrgM365ObjectKindLabel: schulOrgM365ObjectKindLabel,
    schulOrgM365ObjectKindIcon: schulOrgM365ObjectKindIcon,
    emptySchulOrgM365Bindings: emptySchulOrgM365Bindings,
    normalizeSchulOrgM365Binding: normalizeSchulOrgM365Binding,
    normalizeSchulOrgM365: normalizeSchulOrgM365,
    getSchulOrgM365Binding: getSchulOrgM365Binding,
    setSchulOrgM365Binding: setSchulOrgM365Binding,
    inferSchulOrgGroupKind: inferSchulOrgGroupKind,
    suggestSchulOrgMailNickname: suggestSchulOrgMailNickname,
    upsertSchulOrganigrammAreaPerson: upsertSchulOrganigrammAreaPerson,
    removeSchulOrganigrammAreaPerson: removeSchulOrganigrammAreaPerson,
    audienceIdForArea: audienceIdForArea,
    areaAcceptsDirectPeople: areaAcceptsDirectPeople,
    personNodeIdFromEmail: personNodeIdFromEmail,
    moveSchulOrganigrammArea: moveSchulOrganigrammArea,
    isGarbageSchulOrgRef: isGarbageSchulOrgRef,
    convertVerwaltungOrganigrammToSchul: convertVerwaltungOrganigrammToSchul,
    migrateVerwaltungOrganigrammToSchul: migrateVerwaltungOrganigrammToSchul,
    hasVerwaltungOrganigrammLayout: hasVerwaltungOrganigrammLayout,
    resolveSchulOrganigrammLayout: resolveSchulOrganigrammLayout,
    schulOrganigrammStats: schulOrganigrammStats,
    canPlaceSchulOrgNodeInSection: canPlaceSchulOrgNodeInSection,
    listSchulOrgRoleParentAreas: listSchulOrgRoleParentAreas,
    isRoleCapableSection: isRoleCapableSection,
    audienceSectionId: audienceSectionId,
    audienceIdFromSectionId: audienceIdFromSectionId,
    moveSchulOrganigrammNode: moveSchulOrganigrammNode,
    reorderSchulOrganigrammSections: reorderSchulOrganigrammSections,
    nodeMatchesSchulOrgFilter: nodeMatchesSchulOrgFilter
};

if (typeof window !== 'undefined') {
    window.ms365SchulOrganigramm = api;
}

export default api;
