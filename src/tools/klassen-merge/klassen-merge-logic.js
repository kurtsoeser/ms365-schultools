/**
 * Reine Logik: zwei oder mehr Klassen zu einer zusammenführen (Stammdaten-Plan).
 * Kein DOM, kein Graph.
 */

function normStr(v) {
    return String(v == null ? '' : v).trim();
}

function normCode(v) {
    return normStr(v).toUpperCase().replace(/\s+/g, '');
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

/**
 * @param {{ code?: string, name?: string, year?: string, headName?: string, headEmail?: string, stableMailNickname?: string }} cls
 * @param {object[]} classTeams
 */
export function findTeamForClass(cls, classTeams) {
    const teams = Array.isArray(classTeams) ? classTeams : [];
    const code = normCode(cls && cls.code);
    const year = normStr(cls && cls.year);
    if (code) {
        for (let i = 0; i < teams.length; i++) {
            const t = teams[i];
            if (normCode(t && t.classCode) !== code) continue;
            if (year && t.abschlussJahr && String(t.abschlussJahr) !== year) continue;
            return t;
        }
    }
    const nick = normStr(cls && cls.stableMailNickname)
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
    if (nick) {
        for (let j = 0; j < teams.length; j++) {
            const n = normStr(teams[j] && teams[j].stableMailNickname)
                .toLowerCase()
                .replace(/[^a-z0-9]/g, '');
            if (n && n === nick) return teams[j];
        }
    }
    return null;
}

/**
 * @param {object} opts
 * @param {object} opts.survivor Klasse, deren M365-Alias/Gruppe erhalten bleibt
 * @param {string} [opts.survivorOriginalCode] Original-Kürzel vor Umbenennung (falls survivor.code schon neu ist)
 * @param {object[]} opts.sources Weitere Klassen, die aufgelöst werden
 * @param {object[]} [opts.students]
 * @param {object[]} [opts.classTeams]
 * @param {string} [opts.newCode]
 * @param {string} [opts.newName]
 * @param {string} [opts.newDisplayName]
 * @param {'archive'|'delete'|'keep'} [opts.sourceAction]
 */
export function buildMergePlan(opts) {
    const o = opts || {};
    const survivor = o.survivor && typeof o.survivor === 'object' ? o.survivor : null;
    const sources = Array.isArray(o.sources) ? o.sources.filter(Boolean) : [];
    if (!survivor) {
        return { ok: false, error: 'Survivor-Klasse fehlt.', steps: [], memberEmails: [], warnings: [] };
    }
    if (!sources.length) {
        return { ok: false, error: 'Mindestens eine Quell-Klasse wählen.', steps: [], memberEmails: [], warnings: [] };
    }

    const survivorOriginalCode = normCode(o.survivorOriginalCode || survivor.code);
    const newCode = normCode(o.newCode || survivorOriginalCode);
    const newName = normStr(o.newName || survivor.name || newCode);
    const newDisplayName = normStr(o.newDisplayName || newName);
    const sourceAction = o.sourceAction === 'delete' || o.sourceAction === 'keep' ? o.sourceAction : 'archive';
    const students = Array.isArray(o.students) ? o.students : [];
    const classTeams = Array.isArray(o.classTeams) ? o.classTeams : [];

    const remapCodes = new Set(
        sources
            .map((s) => normCode(s.code))
            .concat([survivorOriginalCode])
            .filter(Boolean)
    );

    const memberEmails = [];
    const seen = new Set();
    students.forEach(function (s) {
        const k = normCode(s && s.klasse);
        if (!k || !remapCodes.has(k)) return;
        const em = normEmail(s && s.email);
        if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
        seen.add(em);
        memberEmails.push(em);
    });

    const survivorForLookup = Object.assign({}, survivor, { code: survivorOriginalCode });
    const survivorTeam = findTeamForClass(survivorForLookup, classTeams);
    const sourceTeams = sources.map(function (src) {
        return { class: src, team: findTeamForClass(src, classTeams) };
    });

    const warnings = [];
    if (!survivorTeam || !survivorTeam.graphGroupId) {
        warnings.push('Survivor hat keine verknüpfte M365-Gruppe – bitte zuerst in Klassengruppen matchen.');
    }
    sourceTeams.forEach(function (st) {
        if (!st.team || !st.team.graphGroupId) {
            warnings.push(
                'Quelle ' + normCode(st.class.code) + ' ohne verknüpfte Gruppe – nur Stammdaten werden zusammengeführt.'
            );
        }
    });
    if (survivorOriginalCode !== newCode) {
        warnings.push('Kürzel ändert sich von ' + survivorOriginalCode + ' auf ' + newCode + '.');
    }

    const steps = [
        {
            id: 'local-students',
            label: 'Schüler:innen auf Klasse „' + newCode + '“ umschreiben',
            count: memberEmails.length
        },
        {
            id: 'local-classes',
            label: 'Stammdaten: Survivor aktualisieren, Quell-Klassen entfernen',
            count: sources.length
        },
        {
            id: 'local-teams',
            label: 'classTeams: Survivor behalten, Quell-Verknüpfungen entfernen',
            count: sourceTeams.filter((x) => x.team).length
        }
    ];
    if (survivorTeam && survivorTeam.graphGroupId) {
        steps.push({
            id: 'graph-members',
            label: 'Mitglieder der Survivor-Gruppe mit Schülerliste abgleichen',
            groupId: survivorTeam.graphGroupId,
            count: memberEmails.length
        });
        if (newDisplayName) {
            steps.push({
                id: 'graph-rename',
                label: 'Anzeigename der Survivor-Gruppe setzen: „' + newDisplayName + '“',
                groupId: survivorTeam.graphGroupId,
                displayName: newDisplayName
            });
        }
    }
    sourceTeams.forEach(function (st) {
        if (!st.team || !st.team.graphGroupId) return;
        steps.push({
            id: 'graph-source-' + sourceAction,
            label:
                (sourceAction === 'archive'
                    ? 'Quell-Team archivieren'
                    : sourceAction === 'delete'
                      ? 'Quell-Gruppe löschen'
                      : 'Quell-Gruppe belassen') +
                ': ' +
                normCode(st.class.code),
            groupId: st.team.graphGroupId,
            action: sourceAction,
            classCode: normCode(st.class.code)
        });
    });

    return {
        ok: true,
        error: '',
        survivorOriginalCode,
        survivor: Object.assign({}, survivor, { code: newCode, name: newName }),
        sources: sources.slice(),
        newCode,
        newName,
        newDisplayName,
        sourceAction,
        survivorTeam,
        sourceTeams,
        memberEmails,
        remapCodes: Array.from(remapCodes),
        steps,
        warnings
    };
}

/**
 * Lokaler Teil eines Merge-Plans (Stammdaten + classTeams), ohne Graph.
 * @param {ReturnType<typeof buildMergePlan>} plan
 * @param {{ classes?: object[], students?: object[] }} settings
 * @param {object[]} classTeams
 */
export function applyLocalMerge(plan, settings, classTeams) {
    if (!plan || !plan.ok) {
        return {
            settings: settings || {},
            classTeams: Array.isArray(classTeams) ? classTeams : [],
            error: (plan && plan.error) || 'Ungültiger Plan'
        };
    }

    const newCode = plan.newCode;
    const remap = new Set(plan.remapCodes || []);
    const dropCodes = new Set((plan.sources || []).map((s) => normCode(s.code)).filter(Boolean));
    const survivorOrig = normCode(plan.survivorOriginalCode);

    const studentsOut = (Array.isArray(settings && settings.students) ? settings.students : []).map(function (s) {
        const k = normCode(s && s.klasse);
        if (k && remap.has(k)) return Object.assign({}, s, { klasse: newCode });
        return s;
    });

    const classesIn = Array.isArray(settings && settings.classes) ? settings.classes : [];
    const classesOut = [];
    let survivorWritten = false;
    classesIn.forEach(function (c) {
        const code = normCode(c && c.code);
        if (dropCodes.has(code)) return;
        if (code === survivorOrig || code === newCode) {
            if (survivorWritten) return;
            survivorWritten = true;
            classesOut.push(
                Object.assign({}, c, {
                    code: newCode,
                    name: plan.newName || c.name,
                    headName: plan.survivor.headName != null ? plan.survivor.headName : c.headName,
                    headEmail: plan.survivor.headEmail != null ? plan.survivor.headEmail : c.headEmail,
                    year: plan.survivor.year != null ? plan.survivor.year : c.year,
                    stableMailNickname:
                        (plan.survivorTeam && plan.survivorTeam.stableMailNickname) || c.stableMailNickname || ''
                })
            );
            return;
        }
        classesOut.push(c);
    });
    if (!survivorWritten) {
        classesOut.push({
            code: newCode,
            name: plan.newName,
            year: plan.survivor.year || '',
            headName: plan.survivor.headName || '',
            headEmail: plan.survivor.headEmail || '',
            stableMailNickname: (plan.survivorTeam && plan.survivorTeam.stableMailNickname) || ''
        });
    }

    const dropNicks = new Set();
    const dropIds = new Set();
    (plan.sourceTeams || []).forEach(function (st) {
        if (st.team && st.team.stableMailNickname) {
            dropNicks.add(String(st.team.stableMailNickname).toLowerCase());
        }
        if (st.team && st.team.graphGroupId) dropIds.add(String(st.team.graphGroupId));
    });

    const survivorNick = plan.survivorTeam && plan.survivorTeam.stableMailNickname
        ? String(plan.survivorTeam.stableMailNickname).toLowerCase()
        : '';
    const survivorGid = plan.survivorTeam && plan.survivorTeam.graphGroupId ? String(plan.survivorTeam.graphGroupId) : '';

    const teamsOut = (Array.isArray(classTeams) ? classTeams : [])
        .filter(function (t) {
            if (t.graphGroupId && dropIds.has(String(t.graphGroupId))) return false;
            if (t.stableMailNickname && dropNicks.has(String(t.stableMailNickname).toLowerCase())) return false;
            return true;
        })
        .map(function (t) {
            const same =
                (survivorGid && t.graphGroupId && String(t.graphGroupId) === survivorGid) ||
                (survivorNick &&
                    t.stableMailNickname &&
                    String(t.stableMailNickname).toLowerCase() === survivorNick);
            if (same) {
                return Object.assign({}, t, {
                    classCode: newCode,
                    displayNameHint: plan.newDisplayName || t.displayNameHint
                });
            }
            return t;
        });

    return {
        settings: Object.assign({}, settings || {}, { classes: classesOut, students: studentsOut }),
        classTeams: teamsOut,
        error: ''
    };
}

export default { buildMergePlan, applyLocalMerge, findTeamForClass };
