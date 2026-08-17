/**
 * Mapping-/Match-Helfer für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` ausgelagert (Phase 2 Schnitt 13).
 *
 * Stellt die **pure Logik** für die Persistenz und UI-Synchronisation des
 * Match-Modus bereit. Die globalen Globals (`loadMatchState`,
 * `saveMatchState`, `window.dispatchEvent`, `__ms365MatchLinks`-Spiegel)
 * leben weiterhin im Hauptfile – dort werden die Pure-Versionen mit einem
 * dünnen Wrapper an den State angebunden.
 *
 * Bereitgestellte Funktionen:
 *  - {@link parseMatchSelectValue}: trennt einen UI-Select-Wert in
 *    `{ tenantGroupId, tenantUserId }` (`g:` / `u:` Präfix).
 *  - {@link applyMatchLinkUpdate}: liefert die nächste Links-Map nach einem
 *    Save oder Clear (immutable).
 *  - {@link persistedMatchSelectValuePure}: leitet aus einem gespeicherten
 *    Link den korrekten Select-Wert ab (`u:<id>` / `g:<id>` / `''`).
 *  - {@link computeMatchDraftDirty}: bewertet, ob das UI vom gespeicherten
 *    Stand abweicht (Wert oder Notiz).
 *
 * Die Funktionen sind **window-frei** und vollständig unit-testbar.
 */

import { normStr } from '../../shared/utils/strings.js';
import {
    STRUCT_TREE_ROOT_ADMIN,
    STRUCT_TREE_ROOT_STUDENTS,
    STRUCT_TREE_ROOT_TEACHERS,
    isStructureTreeRootId,
    isTopVerwaltungStructureNode
} from './schulstruktur-sync-tree.js';

/**
 * Form, in der ein gespeicherter Match-Link in der Links-Map abgelegt ist.
 * @typedef {{
 *   tenantGroupId?: string,
 *   tenantUserId?: string,
 *   note?: string,
 *   updatedAt?: string
 * }} MatchLink
 */

/**
 * Zerlegt einen UI-Select-Wert (`'g:<id>'`, `'u:<id>'` oder
 * `'<id>'`/`''`) in IDs.
 *  - Kein Präfix → wird als Gruppen-ID interpretiert.
 *  - Leerer/Whitespace-Eingang liefert leere IDs.
 *
 * @param {unknown} selectValue
 * @returns {{ tenantGroupId: string, tenantUserId: string }}
 */
export function parseMatchSelectValue(selectValue) {
    const raw = String(selectValue || '').trim();
    if (!raw) return { tenantGroupId: '', tenantUserId: '' };
    if (raw.startsWith('u:')) return { tenantGroupId: '', tenantUserId: raw.slice(2).trim() };
    if (raw.startsWith('g:')) return { tenantGroupId: raw.slice(2).trim(), tenantUserId: '' };
    return { tenantGroupId: raw, tenantUserId: '' };
}

/**
 * Erzeugt eine neue Links-Map mit dem Update für eine Struktur-ID.
 *  - Sind sowohl `tenantGroupId` als auch `tenantUserId` leer, wird der Eintrag
 *    entfernt.
 *  - Sonst wird ein neuer Eintrag mit aktuellem `updatedAt`-Zeitstempel
 *    gesetzt.
 *  - `structureId` wird auf eine Stringform normalisiert; leere ID liefert
 *    die Eingabe-Map unverändert zurück.
 *
 * @param {Record<string, MatchLink>} currentLinks Aktuelle Links-Map (wird
 *        nicht mutiert).
 * @param {string | number | null | undefined} structureId
 * @param {{
 *   tenantGroupId?: string | null,
 *   tenantUserId?: string | null,
 *   note?: string | null,
 *   updatedAt?: string
 * }} update
 * @returns {Record<string, MatchLink>} Neue Links-Map.
 */
export function applyMatchLinkUpdate(currentLinks, structureId, update) {
    const cur = currentLinks && typeof currentLinks === 'object' ? currentLinks : {};
    const id = String(structureId == null ? '' : structureId).trim();
    if (!id) return cur;
    const gid = normStr((update && update.tenantGroupId) || '');
    const uid = normStr((update && update.tenantUserId) || '');
    const next = Object.assign({}, cur);
    if (!gid && !uid) {
        if (Object.prototype.hasOwnProperty.call(next, id)) delete next[id];
        return next;
    }
    next[id] = {
        tenantGroupId: gid,
        tenantUserId: uid,
        note: String((update && update.note) || ''),
        updatedAt: (update && update.updatedAt) || new Date().toISOString()
    };
    return next;
}

/**
 * Liefert den Select-Wert (`u:<id>` / `g:<id>` / `''`), der zum gespeicherten
 * Match passt.
 *
 *  - Hat der Link eine `tenantUserId` → `u:<id>`.
 *  - Hat er nur eine `tenantGroupId` → `g:<id>`, **außer** der Callback
 *    `isUserId(id)` markiert sie als User-ID (Übergangsfall: alte Mappings
 *    haben User-IDs im Group-Slot).
 *  - Kein Eintrag bzw. leere IDs → `''`.
 *
 * @param {string | number | null | undefined} structureId
 * @param {Record<string, MatchLink> | null | undefined} links
 * @param {(id: string) => boolean} [isUserId] Optional, default `() => false`.
 * @returns {string}
 */
export function persistedMatchSelectValuePure(structureId, links, isUserId) {
    if (!links || typeof links !== 'object') return '';
    const id = String(structureId == null ? '' : structureId);
    if (!id) return '';
    const saved = links[id];
    if (!saved) return '';
    const uid = normStr(saved.tenantUserId);
    if (uid) return 'u:' + uid;
    const gid = normStr(saved.tenantGroupId);
    if (!gid) return '';
    const treatAsUser = typeof isUserId === 'function' ? !!isUserId(gid) : false;
    return (treatAsUser ? 'u:' : 'g:') + gid;
}

/**
 * Prüft, ob die UI-Eingaben von einem gespeicherten Match-Link abweichen.
 *
 *  - Berücksichtigt sowohl Select-Wert (`u:`/`g:`) als auch Notiz.
 *  - `savedLink` darf `null/undefined` sein.
 *  - Whitespace wird beim Vergleich getrimmt.
 *
 * @param {MatchLink | null | undefined} savedLink
 * @param {string} currentSelectValue Aktueller `select.value`.
 * @param {string} currentNote Aktueller Notiz-Text.
 * @param {(id: string) => boolean} [isUserId] s. {@link persistedMatchSelectValuePure}.
 * @returns {boolean}
 */
export function computeMatchDraftDirty(savedLink, currentSelectValue, currentNote, isUserId) {
    const curVal = String(currentSelectValue || '').trim();
    const curNote = String(currentNote || '').trim();
    let savedVal = '';
    if (savedLink) {
        const uid = normStr(savedLink.tenantUserId);
        if (uid) savedVal = 'u:' + uid;
        else {
            const gid = normStr(savedLink.tenantGroupId);
            if (gid) {
                const treatAsUser = typeof isUserId === 'function' ? !!isUserId(gid) : false;
                savedVal = (treatAsUser ? 'u:' : 'g:') + gid;
            }
        }
    }
    const savedNote = savedLink && savedLink.note ? String(savedLink.note).trim() : '';
    return curVal !== savedVal || curNote !== savedNote;
}

function normCatalogCode(v) {
    return String(v || '')
        .trim()
        .toUpperCase();
}

function sammelgruppeCodeFromRootId(id) {
    const s = String(id || '');
    if (s === STRUCT_TREE_ROOT_STUDENTS) return 'schueler';
    if (s === STRUCT_TREE_ROOT_TEACHERS) return 'lehrer';
    if (s === STRUCT_TREE_ROOT_ADMIN) return 'verwaltung';
    return '';
}

/**
 * Katalog-Bezug einer SOLL-Zeile (Sammelgruppe / ARGE / Fachschaft).
 * Sonstige Knoten (Jahrgang, Klasse, Kursteam, freie Gruppe, Person) → `null`.
 *
 * @param {object | null | undefined} row
 * @returns {{ kind: 'sammelgruppe'|'arge'|'subject', code: string } | null}
 */
export function catalogRefForStructureRow(row) {
    if (!row || typeof row !== 'object') return null;
    const id = String(row.id || '');
    if (isStructureTreeRootId(id)) {
        const code = sammelgruppeCodeFromRootId(id);
        return code ? { kind: 'sammelgruppe', code: code } : null;
    }
    if (isTopVerwaltungStructureNode(row)) return { kind: 'sammelgruppe', code: 'verwaltung' };
    const typ = String(row.typ || '');
    if (typ === 'Arbeitsgemeinschaft') {
        const code = normCatalogCode(row.argeCode || row.bezeichnung);
        return code ? { kind: 'arge', code: code } : null;
    }
    if (typ === 'Gruppe' && row.fachschaftFach) {
        const code = normCatalogCode(row.ktFach || row.bezeichnung);
        return code ? { kind: 'subject', code: code } : null;
    }
    return null;
}

function catalogGraphId(ref, setup) {
    if (!ref) return '';
    const links = setup && Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
    const matched = setup && setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
    if (ref.kind === 'sammelgruppe') {
        const field =
            ref.code === 'schueler'
                ? 'schuelerGroupId'
                : ref.code === 'lehrer'
                  ? 'lehrerGroupId'
                  : ref.code === 'verwaltung'
                    ? 'verwaltungGroupId'
                    : '';
        const fromMatched = field && matched[field] ? String(matched[field]).trim() : '';
        const link = links.find(function (x) {
            return x && x.kind === 'sammelgruppe' && String(x.code || '') === ref.code;
        });
        const fromCat = link && link.graphGroupId ? String(link.graphGroupId).trim() : '';
        return fromCat || fromMatched;
    }
    const link = links.find(function (x) {
        return x && x.kind === ref.kind && normCatalogCode(x.code) === ref.code;
    });
    return link && link.graphGroupId ? String(link.graphGroupId).trim() : '';
}

/**
 * SOLL-Match-Links mit Katalog/Sammelgruppen und Entra-Personen überlagern.
 * Für katalog-gemappte Knoten gilt der Katalog (auch leer = unverknüpft).
 *
 * @param {Record<string, MatchLink>} links
 * @param {object[]} rows
 * @param {object | null | undefined} setup
 * @returns {Record<string, MatchLink>}
 */
export function overlayCatalogOnMatchLinks(links, rows, setup) {
    const next = links && typeof links === 'object' ? Object.assign({}, links) : {};
    const dir =
        setup && setup.directoryMatchByEmail && typeof setup.directoryMatchByEmail === 'object'
            ? setup.directoryMatchByEmail
            : {};
    (Array.isArray(rows) ? rows : []).forEach(function (row) {
        if (!row || row.id == null) return;
        const id = String(row.id);
        const ref = catalogRefForStructureRow(row);
        if (ref) {
            const gid = catalogGraphId(ref, setup);
            const prev = next[id] && typeof next[id] === 'object' ? next[id] : {};
            if (gid) {
                next[id] = Object.assign({}, prev, { tenantGroupId: gid });
            } else if (prev.tenantGroupId) {
                const copy = Object.assign({}, prev, { tenantGroupId: '' });
                if (!normStr(copy.tenantUserId) && !normStr(copy.note)) delete next[id];
                else next[id] = copy;
            }
        }
        if (String(row.typ || '') !== 'Person') return;
        const em = String(row.personEmail || '')
            .trim()
            .toLowerCase();
        if (!em || em.indexOf('@') === -1) return;
        const hit = dir[em];
        const uid = hit && hit.graphUserId ? String(hit.graphUserId).trim() : '';
        if (!uid) return;
        const prev = next[id] && typeof next[id] === 'object' ? next[id] : {};
        if (!normStr(prev.tenantUserId)) next[id] = Object.assign({}, prev, { tenantUserId: uid });
    });
    return next;
}

/**
 * Match-Links, die noch nicht im Katalog stehen, als Upserts vorschlagen
 * (einmalige Übernahme aus dem alten SOLL-Speicher).
 *
 * @param {object[]} rows
 * @param {Record<string, MatchLink>} links
 * @param {object | null | undefined} setup
 * @returns {Array<{ kind: string, code: string, graphGroupId: string, mode: string }>}
 */
export function catalogUpsertsFromOrphanMatchLinks(rows, links, setup) {
    const out = [];
    const seen = new Set();
    const map = links && typeof links === 'object' ? links : {};
    (Array.isArray(rows) ? rows : []).forEach(function (row) {
        const ref = catalogRefForStructureRow(row);
        if (!ref) return;
        if (catalogGraphId(ref, setup)) return;
        const saved = map[String(row.id || '')];
        const gid = saved && saved.tenantGroupId ? String(saved.tenantGroupId).trim() : '';
        if (!gid) return;
        const k = ref.kind + ':' + ref.code;
        if (seen.has(k)) return;
        seen.add(k);
        out.push({ kind: ref.kind, code: ref.code, graphGroupId: gid, mode: 'matched' });
    });
    return out;
}
