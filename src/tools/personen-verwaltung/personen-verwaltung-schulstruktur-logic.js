/**
 * Stammdaten-/Organigramm-Zuordnungen einer Person (Funktionen, Bereiche, Leitungen, …).
 */

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

/**
 * Settings aus Tenant + optional App-Data (Belegung).
 * @param {object|null|undefined} [raw]
 */
export function loadPersonSchulstrukturSettings(raw) {
    let s =
        raw && typeof raw === 'object'
            ? raw
            : typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function'
              ? window.ms365TenantSettingsLoad() || {}
              : {};
    if (!s.unterrichtsbelegung) {
        try {
            const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
            if (api && typeof api.getUnterrichtsbelegung === 'function') {
                const snap = api.getUnterrichtsbelegung();
                if (snap && Array.isArray(snap.rows) && snap.rows.length) {
                    s = Object.assign({}, s, { unterrichtsbelegung: snap });
                }
            }
        } catch {
            /* ignore */
        }
    }
    return s;
}

/**
 * @param {object} settings
 * @param {string} email
 * @param {{ name?: string, upn?: string }} [identity]
 * @returns {{
 *   email: string,
 *   categories: Array<{ id: string, title: string, items: object[] }>,
 *   total: number
 * }}
 */
export function collectPersonSchulstrukturAssignments(settings, email, identity) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const upn = normEmail(identity && identity.upn);
    const em = normEmail(email) || upn;
    const nameHint = normStr(identity && identity.name).toLowerCase();
    const matchEmail = function (raw) {
        const x = normEmail(raw);
        return !!x && (x === em || (upn && x === upn));
    };

    /** @type {Map<string, { id: string, title: string, items: object[] }>} */
    const cats = new Map([
        ['functions', { id: 'functions', title: 'Funktionen / Rollen', items: [] }],
        ['areas', { id: 'areas', title: 'Bereiche (Verwaltung / Organigramm)', items: [] }],
        ['leadership', { id: 'leadership', title: 'Leitungen', items: [] }],
        ['classes', { id: 'classes', title: 'Klassen & Unterricht', items: [] }],
        ['other', { id: 'other', title: 'Weitere Zuordnungen', items: [] }]
    ]);

    function push(catId, item) {
        const cat = cats.get(catId) || cats.get('other');
        const key = [item.kind, item.label, item.detail].join('|');
        if (cat.items.some(function (x) {
            return [x.kind, x.label, x.detail].join('|') === key;
        })) {
            return;
        }
        cat.items.push(item);
    }

    // —— Verwaltungs-Rollen (admin[]) ——
    (Array.isArray(s.admin) ? s.admin : []).forEach(function (p) {
        if (!matchEmail(p && p.email)) return;
        const role = normStr(p.role) || normStr(p.defaultKey) || 'Rolle';
        push('functions', {
            kind: 'adminRole',
            label: role,
            detail: 'Stammdaten → Verwaltung (Personal)',
            source: 'admin'
        });
    });

    // —— Audience / Verwaltungsgruppen ——
    const groups = Array.isArray(s.verwaltungAudienceGroups) ? s.verwaltungAudienceGroups : [];
    const groupLabel = function (id) {
        const g = groups.find(function (x) {
            return normStr(x && x.id) === normStr(id);
        });
        return (g && normStr(g.label)) || normStr(id) || 'Bereich';
    };
    (Array.isArray(s.adminAudienceMemberships) ? s.adminAudienceMemberships : []).forEach(
        function (m) {
            if (!matchEmail(m && m.email)) return;
            push('areas', {
                kind: 'audience',
                label: groupLabel(m.groupId),
                detail: 'Verwaltungsgruppe / Zielgruppe',
                source: 'adminAudienceMemberships',
                ref: normStr(m.groupId)
            });
        }
    );

    // —— Organigramm: Person-Slots & Pfad ——
    const org =
        s.schulOrganigramm && typeof s.schulOrganigramm === 'object' ? s.schulOrganigramm : null;
    const nodes = org && Array.isArray(org.nodes) ? org.nodes : [];
    const byId = new Map();
    nodes.forEach(function (n) {
        if (n && n.id) byId.set(String(n.id), n);
    });
    function areaPath(nodeId) {
        const parts = [];
        let walk = byId.get(String(nodeId || ''));
        let guard = 0;
        while (walk && guard < 24) {
            guard += 1;
            parts.unshift(normStr(walk.label) || walk.id);
            const pid = walk.parentId == null || walk.parentId === '' ? '' : String(walk.parentId);
            walk = pid ? byId.get(pid) : null;
        }
        return parts.join(' → ');
    }
    nodes.forEach(function (n) {
        if (!n) return;
        const t = normStr(n.type).toLowerCase();
        if (t === 'person' && normStr(n.refKind) === 'adminPerson' && matchEmail(n.refKey)) {
            const parent = byId.get(String(n.parentId || ''));
            const parentType = normStr(parent && parent.type).toLowerCase();
            const path = areaPath(n.parentId) || normStr(n.label) || 'Organigramm';
            if (parentType === 'role') {
                push('functions', {
                    kind: 'orgRoleSeat',
                    label: normStr(parent.label) || path,
                    detail: 'Funktionssitz im Schul-Organigramm' + (path ? ' · ' + path : ''),
                    source: 'schulOrganigramm',
                    nodeId: n.id,
                    parentId: n.parentId || null
                });
            } else {
                push('areas', {
                    kind: 'orgPerson',
                    label: path,
                    detail: 'Direkt im Schul-Organigramm zugeordnet',
                    source: 'schulOrganigramm',
                    nodeId: n.id,
                    parentId: n.parentId || null
                });
            }
        }
    });

    // —— ARGE-Leitung ——
    (Array.isArray(s.arges) ? s.arges : []).forEach(function (a) {
        if (!matchEmail(a && a.headEmail)) return;
        push('leadership', {
            kind: 'argeHead',
            label: normStr(a.name) || normStr(a.code) || 'ARGE',
            detail: 'ARGE-Leitung',
            source: 'arges',
            ref: normStr(a.code)
        });
    });

    // —— Klassenvorstand ——
    (Array.isArray(s.classes) ? s.classes : []).forEach(function (c) {
        if (!matchEmail(c && c.headEmail)) return;
        push('leadership', {
            kind: 'classHead',
            label: normStr(c.name) || normStr(c.code) || 'Klasse',
            detail: 'Klassenvorstand',
            source: 'classes',
            ref: normStr(c.code)
        });
    });

    // —— Lehrkraft ——
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (t) {
        if (!matchEmail(t && t.email)) return;
        push('classes', {
            kind: 'teacher',
            label: normStr(t.name) || normStr(t.code) || 'Lehrkraft',
            detail:
                'Lehrerliste' +
                (normStr(t.code) ? ' · Kürzel ' + normStr(t.code) : ''),
            source: 'teachers',
            ref: normStr(t.code)
        });
    });

    // —— Schüler ——
    (Array.isArray(s.students) ? s.students : []).forEach(function (st) {
        if (!matchEmail(st && st.email)) return;
        push('classes', {
            kind: 'student',
            label: normStr(st.klasse) || 'Klasse',
            detail: 'Schüler:in in Stammdaten',
            source: 'students',
            ref: normStr(st.klasse)
        });
    });

    // —— Unterrichtsteams / Belegung ——
    const belegung = s.unterrichtsbelegung;
    const rows =
        belegung && typeof belegung === 'object' && Array.isArray(belegung.rows)
            ? belegung.rows
            : Array.isArray(belegung)
              ? belegung
              : [];
    rows.forEach(function (row) {
        if (!row || !matchEmail(row.lehrerEmail)) return;
        const label =
            normStr(row.teamName) ||
            [normStr(row.fach), normStr(row.klasse)].filter(Boolean).join(' · ') ||
            'Unterrichtsteam';
        push('classes', {
            kind: 'kursteam',
            label: label,
            detail: 'Unterrichtsbelegung (Besitzer/Lehrkraft)',
            source: 'unterrichtsbelegung',
            ref: normStr(row.gruppenmail) || normStr(row.graphGroupId)
        });
    });

    // —— SGA / Schülervertretung ——
    (Array.isArray(s.sga) ? s.sga : []).forEach(function (row) {
        if (!matchEmail(row && row.email)) return;
        push('other', {
            kind: 'sga',
            label: normStr(row.name) || 'SGA',
            detail: 'SGA · ' + (normStr(row.scope) || 'Eintrag'),
            source: 'sga'
        });
    });
    (Array.isArray(s.studentCouncil) ? s.studentCouncil : []).forEach(function (row) {
        if (!matchEmail(row && row.email)) return;
        push('other', {
            kind: 'studentCouncil',
            label: normStr(row.name) || 'Schülervertretung',
            detail: 'Schülervertretung' + (normStr(row.klasse) ? ' · ' + normStr(row.klasse) : ''),
            source: 'studentCouncil'
        });
    });

    // Name-only fallback for admin without email (rare)
    if (!em && nameHint) {
        (Array.isArray(s.admin) ? s.admin : []).forEach(function (p) {
            if (normStr(p && p.name).toLowerCase() !== nameHint) return;
            push('functions', {
                kind: 'adminRole',
                label: normStr(p.role) || 'Rolle',
                detail: 'Stammdaten (Namensmatch, ohne E-Mail)',
                source: 'admin'
            });
        });
    }

    const categories = Array.from(cats.values()).filter(function (c) {
        return c.items.length > 0;
    });
    let total = 0;
    categories.forEach(function (c) {
        total += c.items.length;
    });

    return {
        email: em,
        categories: categories,
        total: total,
        empty: total === 0
    };
}

export default {
    loadPersonSchulstrukturSettings,
    collectPersonSchulstrukturAssignments
};
