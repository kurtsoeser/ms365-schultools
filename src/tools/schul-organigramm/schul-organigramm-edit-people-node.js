/**
 * Node-scoped people listing for Schul-Organigramm edit.
 */
import { normStr, normEmail, isBlankPerson } from './schul-organigramm-edit-utils.js';
import {
    personChipKey,
    SCHUL_ORG_PEOPLE_PREVIEW_LIMIT,
    teachersForArge,
    teachersForClassChat,
    studentsForClass,
    guardiansForClass,
    peopleForAdminRole,
    findAdminRole
} from './schul-organigramm-edit-people.js';
import { findArge, findClass, normalizeSgaScope, findStudentCouncil, findSga } from './schul-organigramm-edit-find.js';
import { normalizeAdminAudienceMemberships } from '../../shared/administration-audience-groups.js';
import {
    isArgeNode,
    isArgeAreaNode,
    isClassNode,
    isClassAreaNode,
    isElternClassNode,
    isKlassenchatClassNode,
    isSammelgruppeArea,
    isFunctionAggregateArea,
    functionAggregateDefForArea,
    listSchulOrgClassFunctions,
    listSchulOrgArgeFunctions,
    collectFunctionAggregatePeople,
    normalizeClassFunctionsMap,
    sgaScopeForArea,
    isSgaScopeArea,
    isSgaBoardArea,
    svSeatForNode,
    personForSvSeat,
    SCHUL_ORG_SV_SEATS,
    SCHUL_ORG_SGA_SV_AREA_ID,
    SCHUL_ORG_BOARD_REF,
    NODE,
    sammelgruppeKeyForArea,
    areaAcceptsDirectPeople,
    childrenOf,
    normalizeSchulOrganigramm,
    isAreaNode,
    audienceIdForArea,
    isSchulOrgLocalPlaceholderEmail
} from './schul-organigramm-logic.js';

export function listPeopleForSchulOrgNode(settings, node) {
    if (!node || !node.type) return [];
    const out = [];
    const push = function (row) {
        const name = normStr(row.name);
        const email = normEmail(row.email);
        const vacant = !!row.vacant || !(name || email);
        if (!name && !email && !row.vacant && !row.seatLabel) return;
        const kind = row.kind || 'person';
        const seatLabel = normStr(row.seatLabel);
        const label = name || email || (seatLabel ? seatLabel + ' (offen)' : '');
        /** @type {{ kind: string, key: string, name: string, email: string, label: string, movable: boolean, scope?: string, klasse?: string, seatLabel?: string, vacant?: boolean, functionId?: string, classCode?: string, className?: string }} */
        const item = {
            kind: kind,
            key: personChipKey({ kind: kind, name: name || seatLabel, email: email }),
            name: name,
            email: email,
            label: label,
            movable: row.movable !== false
        };
        if (seatLabel) item.seatLabel = seatLabel;
        if (vacant) item.vacant = true;
        if (row.scope) item.scope = String(row.scope);
        if (row.klasse) item.klasse = normStr(row.klasse);
        if (row.functionId) item.functionId = normStr(row.functionId);
        if (row.classCode) item.classCode = normStr(row.classCode);
        if (row.className) item.className = normStr(row.className);
        out.push(item);
    };

    if (node.type === 'role' && node.refKind === 'adminRole' && node.refKey) {
        const role = findAdminRole(settings, node.refKey);
        peopleForAdminRole(settings, role)
            .filter(function (p) {
                return !isBlankPerson(p);
            })
            .forEach(function (p) {
                push({ kind: 'admin', name: p.name, email: p.email });
            });
        return out;
    }

    if (isArgeNode(node) && node.refKey) {
        const arge = findArge(settings, node.refKey);
        const members = teachersForArge(settings, node.refKey);
        if (!members.length) {
            push({
                kind: 'head',
                seatLabel: 'ARGE-Leitung',
                name: '',
                email: '',
                label: 'ARGE-Leitung (offen)',
                vacant: true,
                movable: true
            });
        } else {
            members.forEach(function (t) {
                push({
                    kind: t.role === 'head' ? 'head' : 'teacher',
                    name: t.name,
                    email: t.email,
                    label:
                        t.role === 'head'
                            ? t.name || t.email || 'ARGE-Leitung (offen)'
                            : t.name || t.email,
                    vacant: t.role === 'head' && !(t.name || t.email),
                    movable: t.role === 'head',
                    scope: 'teacher',
                    seatLabel: t.role === 'head' ? 'ARGE-Leitung' : 'ARGE-Mitglieder'
                });
            });
        }
        const argeFnSeats = normalizeClassFunctionsMap(arge && arge.functions);
        /** @type {Set<string>} */
        const argeFnEmails = new Set();
        listSchulOrgArgeFunctions(settings).forEach(function (def) {
            const seat = argeFnSeats[def.id] || { name: '', email: '' };
            const seatEmail = normEmail(seat.email);
            if (seatEmail) argeFnEmails.add(seatEmail);
            push({
                kind: 'argeFunction',
                seatLabel: def.label,
                name: normStr(seat.name),
                email: seatEmail,
                label: normStr(seat.name) || seatEmail || def.label + ' (offen)',
                vacant: !(normStr(seat.name) || seatEmail),
                movable: true,
                functionId: def.id,
                scope: 'teacher'
            });
        });
        // Zusätzliche Bereichspersonen (Landkarte / Abgleich mit M365-Gruppe)
        if (isArgeAreaNode(node)) {
            const seen = new Set(
                out
                    .map(function (p) {
                        return normEmail(p.email);
                    })
                    .filter(Boolean)
            );
            argeFnEmails.forEach(function (em) {
                seen.add(em);
            });
            const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
            childrenOf(layout, node.id).forEach(function (ch) {
                if (ch.type !== NODE.PERSON || ch.refKind !== 'adminPerson') return;
                const em = normEmail(ch.refKey);
                if (!em || seen.has(em)) return;
                seen.add(em);
                push({
                    kind: 'areaPerson',
                    name: ch.label,
                    email: em,
                    movable: true
                });
            });
        }
        return out;
    }

    if (isClassNode(node) && node.refKey) {
        const cls = findClass(settings, node.refKey);
        const headName = cls ? normStr(cls.headName) : '';
        const headEmail = cls ? normEmail(cls.headEmail) : '';
        push({
            kind: 'head',
            seatLabel: 'Klassenvorstand',
            name: headName,
            email: headEmail,
            label: headName || headEmail || 'Klassenvorstand (offen)',
            vacant: !(headName || headEmail),
            movable: true
        });
        const fnSeats = normalizeClassFunctionsMap(cls && cls.functions);
        /** @type {Set<string>} */
        const functionEmails = new Set();
        listSchulOrgClassFunctions(settings).forEach(function (def) {
            const seat = fnSeats[def.id] || { name: '', email: '' };
            const seatEmail = normEmail(seat.email);
            if (seatEmail) functionEmails.add(seatEmail);
            push({
                kind: 'classFunction',
                seatLabel: def.label,
                name: normStr(seat.name),
                email: seatEmail,
                label: normStr(seat.name) || seatEmail || def.label + ' (offen)',
                vacant: !(normStr(seat.name) || seatEmail),
                movable: true,
                functionId: def.id
            });
        });
        studentsForClass(settings, cls)
            .sort(function (a, b) {
                return normStr(a.name).localeCompare(normStr(b.name), 'de');
            })
            .forEach(function (s) {
                const em = normEmail(s.email);
                if (em && functionEmails.has(em)) return;
                push({
                    kind: 'student',
                    name: s.name,
                    email: s.email,
                    seatLabel: 'Schüler'
                });
            });
        return out;
    }

    if (isElternClassNode(node) && node.refKey) {
        guardiansForClass(settings, node.refKey).forEach(function (g) {
            push({
                kind: 'guardian',
                name: g.name,
                email: g.email,
                label: g.name || g.email,
                movable: false,
                scope: 'external'
            });
        });
        return out;
    }

    if (isKlassenchatClassNode(node) && node.refKey) {
        teachersForClassChat(settings, node.refKey).forEach(function (t) {
            push({
                kind: t.role === 'kv' ? 'head' : 'teacher',
                name: t.name,
                email: t.email,
                label: t.role === 'kv' ? (t.name || t.email) + ' (KV)' : t.name || t.email,
                movable: false,
                scope: 'teacher',
                seatLabel: t.role === 'kv' ? 'Klassenvorstand' : ''
            });
        });
        return out;
    }

    if (node.type === NODE.KURSTEAM || node.type === 'kursteam') {
        const meta = node.meta && typeof node.meta === 'object' ? node.meta : {};
        let headName = normStr(meta.lehrerName);
        let headEmail = normEmail(meta.lehrerEmail);
        if (!headEmail && meta.lehrerCode) {
            const code = normStr(meta.lehrerCode).toUpperCase();
            const t = (Array.isArray(settings.teachers) ? settings.teachers : []).find(function (row) {
                return normStr(row && row.code).toUpperCase() === code;
            });
            if (t) {
                headName = headName || normStr(t.name);
                headEmail = normEmail(t.email);
            }
        }
        push({
            kind: 'head',
            seatLabel: 'Besitzer',
            name: headName,
            email: headEmail,
            label: headName || headEmail || 'Besitzer (offen)',
            vacant: !(headName || headEmail),
            movable: false,
            scope: 'teacher'
        });
        return out;
    }

    if (node.type === 'council' && node.refKind === 'studentCouncil') {
        const seat = svSeatForNode(node);
        if (seat) {
            const person = personForSvSeat(settings, seat.seat);
            if (person) {
                push({
                    kind: 'council',
                    name: person.name,
                    email: person.email,
                    label: person.name || person.email,
                    movable: false,
                    scope: 'student',
                    klasse: person.klasse || '',
                    seatLabel: seat.label
                });
            }
            return out;
        }
        if (node.refKey === SCHUL_ORG_BOARD_REF || node.id === 'council_board') {
            SCHUL_ORG_SV_SEATS.forEach(function (s) {
                const person = personForSvSeat(settings, s.seat);
                if (!person) return;
                push({
                    kind: 'council',
                    name: person.name,
                    email: person.email,
                    label: s.label + ': ' + (person.name || person.email),
                    movable: false,
                    scope: 'student',
                    klasse: person.klasse || '',
                    seatLabel: s.label
                });
            });
            return out;
        }
        const row = findStudentCouncil(settings, node.refKey);
        if (row) push({ kind: 'council', name: row.name, email: row.email, movable: false });
        else if (node.meta) push({ kind: 'council', name: node.label, email: node.meta.email, movable: false });
        return out;
    }

    if (node.type === 'sga' && node.refKind === 'sga') {
        if (node.refKey === SCHUL_ORG_BOARD_REF || node.id === 'sga_board') {
            // Personen liegen in den Scope-Unterbereichen
            return out;
        }
        const row = findSga(settings, node.refKey);
        if (row) push({ kind: 'sga', name: row.name, email: row.email, movable: false });
        else if (node.meta) push({ kind: 'sga', name: node.label, email: node.meta.email, movable: false });
        return out;
    }

    if (isAreaNode(node) || node.type === NODE.PERSON) {
        if (node.type === NODE.PERSON && node.refKind === 'adminPerson') {
            push({
                kind: 'areaPerson',
                name: node.label,
                email: node.refKey,
                movable: true
            });
            return out;
        }
        if (isSgaBoardArea(node)) {
            return out;
        }
        if (isFunctionAggregateArea(node)) {
            const agg = functionAggregateDefForArea(node, settings);
            const seatLabel = (agg && agg.seatLabel) || 'Funktion';
            collectFunctionAggregatePeople(settings, agg || node.id).forEach(function (p) {
                const groupHint = p.className || p.argeName || '';
                push({
                    kind: (agg && agg.personKind) === 'student' ? 'student' : 'teacher',
                    seatLabel: p.seatLabel || seatLabel,
                    name: p.name,
                    email: p.email,
                    label: p.label + (groupHint ? ' · ' + groupHint : ''),
                    vacant: false,
                    movable: false,
                    scope: (agg && agg.personKind) === 'student' ? 'student' : 'teacher',
                    classCode: p.classCode || '',
                    className: groupHint,
                    functionId: p.functionId || (agg && agg.seatKey) || (agg && agg.id) || ''
                });
            });
            return out;
        }
        const sgaScope = sgaScopeForArea(node);
        if (sgaScope) {
            if (sgaScope === 'student' || node.id === SCHUL_ORG_SGA_SV_AREA_ID) {
                // Schülervertretung im SGA = drei Funktionen
                SCHUL_ORG_SV_SEATS.forEach(function (s) {
                    const person = personForSvSeat(settings, s.seat);
                    if (!person) return;
                    push({
                        kind: 'council',
                        name: person.name,
                        email: person.email,
                        label: s.label + ': ' + (person.name || person.email),
                        movable: false,
                        scope: 'student',
                        klasse: person.klasse || '',
                        seatLabel: s.label
                    });
                });
                return out;
            }
            (Array.isArray(settings.sga) ? settings.sga : [])
                .filter(function (row) {
                    if (!(normStr(row && row.name) || normEmail(row && row.email))) return false;
                    return (normalizeSgaScope(row.scope) || 'student') === sgaScope;
                })
                .sort(function (a, b) {
                    return normStr(a.name).localeCompare(normStr(b.name), 'de');
                })
                .forEach(function (row) {
                    push({
                        kind: 'sga',
                        name: row.name,
                        email: row.email,
                        label: normStr(row.name) || normEmail(row.email),
                        movable: false,
                        scope: sgaScope
                    });
                });
            return out;
        }
        const sammelKey = sammelgruppeKeyForArea(node);
        if (sammelKey === 'lehrer') {
            (Array.isArray(settings.teachers) ? settings.teachers : [])
                .filter(function (t) {
                    return normStr(t && t.name) || normEmail(t && t.email);
                })
                .sort(function (a, b) {
                    return normStr(a.name).localeCompare(normStr(b.name), 'de');
                })
                .forEach(function (t) {
                    push({
                        kind: 'teacher',
                        name: t.name,
                        email: t.email,
                        label: normStr(t.name) || normEmail(t.email),
                        movable: true,
                        scope: 'teacher'
                    });
                });
            return out;
        }
        if (sammelKey === 'schueler') {
            // Keine flache Schülerliste – Klassen liegen in den Jahrgangs-Bereichen darunter
            return out;
        }
        if (sammelKey === 'eltern') {
            // Keine flache Elternliste – Klassen liegen in den Jahrgangs-Bereichen darunter
            return out;
        }
        if (sammelKey === 'klassenchats') {
            // Keine flache Chat-Liste – Klassen liegen in den Jahrgangs-Bereichen darunter
            return out;
        }
        const seen = new Set();
        const audId = audienceIdForArea(node);
        if (audId) {
            normalizeAdminAudienceMemberships(settings.adminAudienceMemberships).forEach(function (m) {
                if (String(m.groupId) !== audId) return;
                if (seen.has(m.email)) return;
                seen.add(m.email);
                const adminHit = (Array.isArray(settings.admin) ? settings.admin : []).find(function (a) {
                    return normEmail(a && a.email) === m.email;
                });
                push({
                    kind: 'areaPerson',
                    name: m.name || (adminHit && adminHit.name) || '',
                    email: m.email,
                    movable: true
                });
            });
        }
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        childrenOf(layout, node.id).forEach(function (ch) {
            if (ch.type !== NODE.PERSON || ch.refKind !== 'adminPerson') return;
            const em = normEmail(ch.refKey);
            if (!em || seen.has(em)) return;
            seen.add(em);
            push({
                kind: 'areaPerson',
                name: ch.label,
                email: em,
                movable: true
            });
        });
        out.sort(function (a, b) {
            return normStr(a.name).localeCompare(normStr(b.name), 'de');
        });
        return out;
    }

    return out;
}

/**
 * Personen für M365-Verknüpfung eines Bereichs (Gruppe/Postfach):
 * direkte Bereichspersonen, Rollen darunter sowie rekursiv alle Unterbereiche
 * (flach als Mitglieder/Berechtigte – keine verschachtelten M365-Gruppen).
 * @param {object} settings
 * @param {object} areaNode
 * @returns {Array<{ email: string, name: string, label: string, source: string }>}
 */
export function listM365PeopleForArea(settings, areaNode) {
    if (!areaNode || !isAreaNode(areaNode)) return [];
    /** @type {Map<string, { email: string, name: string, label: string, source: string }>} */
    const byMail = new Map();
    const pushPerson = function (p, source) {
        const email = normEmail(p && p.email);
        if (!email || isSchulOrgLocalPlaceholderEmail(email)) return;
        if (p && p.vacant) return;
        if (byMail.has(email)) return;
        const name = normStr(p && p.name);
        byMail.set(email, {
            email: email,
            name: name,
            label: name || email,
            source: source || 'area'
        });
    };

    const layout = normalizeSchulOrganigramm(settings && settings.schulOrganigramm);
    /** @type {Set<string>} */
    const visitedAreas = new Set();

    /**
     * @param {object} area
     * @param {boolean} nested
     */
    function walkArea(area, nested) {
        if (!area || !isAreaNode(area) || !area.id) return;
        const aid = String(area.id);
        if (visitedAreas.has(aid)) return;
        visitedAreas.add(aid);

        listPeopleForSchulOrgNode(settings, area).forEach(function (p) {
            pushPerson(p, nested ? 'nested' : 'area');
        });

        childrenOf(layout, aid).forEach(function (ch) {
            if (!ch) return;
            if (isAreaNode(ch)) {
                walkArea(ch, true);
                return;
            }
            const ct = String(ch.type || '').toLowerCase();
            if (ct !== NODE.ROLE && ct !== 'role') return;
            listPeopleForSchulOrgNode(settings, ch).forEach(function (p) {
                pushPerson(p, nested ? 'nested' : 'role');
            });
        });
    }

    walkArea(areaNode, false);

    return Array.from(byMail.values()).sort(function (a, b) {
        return normStr(a.name || a.email).localeCompare(normStr(b.name || b.email), 'de');
    });
}

/** @deprecated Alias – gleiche Liste wie listM365PeopleForArea */
export function listMailboxDelegatePeopleForArea(settings, areaNode) {
    return listM365PeopleForArea(settings, areaNode);
}

/**
 * @param {object} settings
 * @param {object} node
 * @param {number} [limit]
 */
export function previewPeopleForSchulOrgNode(settings, node, limit) {
    const max =
        typeof limit === 'number' && limit >= 0 ? limit : SCHUL_ORG_PEOPLE_PREVIEW_LIMIT;
    const all = listPeopleForSchulOrgNode(settings, node);
    return {
        all: all,
        shown: all.slice(0, max),
        more: Math.max(0, all.length - max),
        total: all.length
    };
}
