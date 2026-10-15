/**
 * Selektiver Import aus IT-Bibliothek / Browser-Backup:
 * Bereiche vergleichen, Vorschau, nur gewählte Teile übernehmen.
 */

const SCHOOLTOOL_KEY = 'ms365-schooltool-data-v2';

/** @typedef {'same'|'differs'|'local-only'|'remote-only'|'missing'} SectionStatus */

/**
 * Stammdaten-Bereiche innerhalb von ms365-schooltool-data-v2.
 * @type {Array<{
 *   id: string,
 *   label: string,
 *   hint: string,
 *   kind: 'year-fields'|'core-fields'|'root-field',
 *   fields: string[],
 *   group: string
 * }>}
 */
export const SCHOOLTOOL_SELECTIVE_SECTIONS = [
    {
        id: 'students',
        label: 'Schüler:innen',
        hint: 'Schülerlisten aller Schuljahre im Backup',
        kind: 'year-fields',
        fields: ['students'],
        group: 'stammdaten'
    },
    {
        id: 'classes',
        label: 'Klassen',
        hint: 'Klassenlisten inkl. Klassenvorstand/KV (headName/headEmail) aller Schuljahre',
        kind: 'year-fields',
        fields: ['classes'],
        group: 'stammdaten'
    },
    {
        id: 'guardians',
        label: 'Eltern / Erziehungsberechtigte',
        hint: 'Guardians und Elternlisten (oft zusammen mit Schüler:innen)',
        kind: 'year-fields',
        fields: ['guardians', 'parentLists'],
        group: 'stammdaten'
    },
    {
        id: 'studentCouncil',
        label: 'Schülervertretung',
        hint: 'Schulsprecher:innen / Student Council',
        kind: 'year-fields',
        fields: ['studentCouncil'],
        group: 'stammdaten'
    },
    {
        id: 'unterricht',
        label: 'Unterrichtsbelegung & Klassenchats',
        hint: 'Belegung Klasse×Lehrkraft×Fach (Lehrer je Klasse) sowie Klassenchats',
        kind: 'year-fields',
        fields: ['unterrichtsbelegung', 'classChats'],
        group: 'stammdaten'
    },
    {
        id: 'teachers',
        label: 'Lehrkräfte',
        hint: 'Lehrerstammdaten (core.teachers)',
        kind: 'core-fields',
        fields: ['teachers'],
        group: 'stammdaten'
    },
    {
        id: 'subjects',
        label: 'Fächer',
        hint: 'Fächerkatalog (core.subjects)',
        kind: 'core-fields',
        fields: ['subjects'],
        group: 'stammdaten'
    },
    {
        id: 'schoolMeta',
        label: 'Schule & Profil',
        hint: 'Schulname, Domain, Schullogo/Profil',
        kind: 'core-fields',
        fields: ['schoolName', 'domain', 'schoolProfile'],
        group: 'stammdaten'
    },
    {
        id: 'schulOrganigramm',
        label: 'Schul-Organigramm',
        hint: 'Knoten und Einträge im Schul-Organigramm',
        kind: 'core-fields',
        fields: ['schulOrganigramm'],
        group: 'stammdaten'
    },
    {
        id: 'verwaltung',
        label: 'Verwaltung & Zielgruppen',
        hint: 'Rollen, Personen, Verwaltungs-Organigramm, Zielgruppen',
        kind: 'core-fields',
        fields: [
            'verwaltungAudienceGroups',
            'adminAudienceMemberships',
            'adminRoles',
            'admin',
            'administration',
            'verwaltungOrganigramm'
        ],
        group: 'stammdaten'
    },
    {
        id: 'setup',
        label: 'Setup / Gruppenverknüpfungen',
        hint: 'matched Groups, Katalog-Links, IT-Bibliothek-Meta',
        kind: 'root-field',
        fields: ['setup'],
        group: 'stammdaten'
    }
];

/** Weitere localStorage-Schlüssel als wählbare Pakete. */
export const STORAGE_SELECTIVE_BUNDLES = [
    {
        id: 'bundle-dashboard',
        label: 'Dashboard & Werkzeugzugriff',
        keys: [
            'ms365-dashboard-tool-access-v1',
            'ms365-dashboard-audience-groups-v1',
            'ms365-stammdaten-listen-perms-v1'
        ]
    },
    {
        id: 'bundle-freistellung',
        label: 'Freistellung',
        keys: ['ms365-freistellung-setup-v1', 'ms365-freistellung-perms-v1']
    },
    {
        id: 'bundle-planer',
        label: 'Schularbeiten / LFR',
        keys: [
            'ms365-schularbeiten-perms-v1',
            'ms365-sa-settings-v1',
            'ms365-lfr-setup-v1',
            'ms365-lfr-perms-v1'
        ]
    },
    {
        id: 'bundle-schulstruktur',
        label: 'Schulstruktur / SLG',
        keys: [
            'ms365-schueler-lehrer-gruppen-v2',
            'ms365-schulstruktur-sync-v1',
            'ms365-schulstruktur-match-v1'
        ]
    },
    {
        id: 'bundle-webuntis',
        label: 'WebUntis-Zustand',
        keys: ['webuntis-teams-creator-state-v1']
    }
];

function deepClone(value) {
    if (value == null) return value;
    try {
        return JSON.parse(JSON.stringify(value));
    } catch {
        return value;
    }
}

function stableValue(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v;
    try {
        return JSON.stringify(v);
    } catch {
        return String(v);
    }
}

function storeFromPayload(payload) {
    const p = payload && typeof payload === 'object' ? payload : {};
    return p.localStorage && typeof p.localStorage === 'object' ? p.localStorage : {};
}

function parseSchooltool(store) {
    const raw = store && store[SCHOOLTOOL_KEY];
    if (raw == null || raw === '') return null;
    if (typeof raw === 'object' && !Array.isArray(raw)) return deepClone(raw);
    try {
        return JSON.parse(String(raw));
    } catch {
        return null;
    }
}

function emptyYearBucket() {
    return {
        students: [],
        studentCouncil: [],
        classes: [],
        guardians: [],
        parentLists: [],
        unterrichtsbelegung: null,
        classChats: null
    };
}

function ensureContainerShape(raw) {
    const c = raw && typeof raw === 'object' ? deepClone(raw) : {};
    if (!c.core || typeof c.core !== 'object') c.core = {};
    if (!c.years || typeof c.years !== 'object') c.years = { current: '', byLabel: {} };
    if (!c.years.byLabel || typeof c.years.byLabel !== 'object') c.years.byLabel = {};
    if (!c.setup || typeof c.setup !== 'object') c.setup = {};
    return c;
}

function yearLabels(container) {
    const by = container && container.years && container.years.byLabel ? container.years.byLabel : {};
    return Object.keys(by).sort();
}

function pickYearFieldSummary(container, fields) {
    const labels = yearLabels(container);
    const byYear = [];
    let total = 0;
    const samples = [];
    labels.forEach(function (lab) {
        const bucket = container.years.byLabel[lab] || emptyYearBucket();
        const parts = {};
        let yearCount = 0;
        fields.forEach(function (f) {
            const val = bucket[f];
            if (Array.isArray(val)) {
                parts[f] = val.length;
                yearCount += val.length;
                if (samples.length < 8 && (f === 'students' || f === 'classes' || f === 'guardians')) {
                    val.slice(0, 4).forEach(function (row) {
                        if (samples.length >= 8) return;
                        const name =
                            (row && (row.name || row.displayName || row.className || row.code)) || '';
                        const s = String(name).trim();
                        if (s) samples.push(s);
                    });
                }
            } else if (val && typeof val === 'object') {
                parts[f] = 'gesetzt';
                yearCount += 1;
            } else {
                parts[f] = val == null || val === '' ? 0 : 1;
            }
        });
        total += yearCount;
        byYear.push({ year: lab, count: yearCount, parts: parts });
    });
    return {
        present: labels.length > 0 || !!(container && container.core),
        total: total,
        years: byYear,
        samples: samples,
        summaryText: labels.length
            ? total + ' Einträge · ' + labels.length + ' Schuljahr/e'
            : 'keine Schuljahre'
    };
}

function pickCoreFieldSummary(container, fields) {
    const core = container && container.core && typeof container.core === 'object' ? container.core : {};
    const parts = {};
    let total = 0;
    const samples = [];
    fields.forEach(function (f) {
        const val = core[f];
        if (Array.isArray(val)) {
            parts[f] = val.length;
            total += val.length;
            val.slice(0, 4).forEach(function (row) {
                if (samples.length >= 8) return;
                const name =
                    (row && (row.name || row.displayName || row.code || row.title || row.id)) || '';
                const s = String(name).trim();
                if (s) samples.push(s);
            });
        } else if (val && typeof val === 'object') {
            if (Array.isArray(val.nodes)) {
                parts[f] = val.nodes.length + ' Knoten';
                total += val.nodes.length;
            } else {
                parts[f] = 'gesetzt';
                total += 1;
            }
        } else if (typeof val === 'string') {
            parts[f] = val.trim() ? val.trim() : '–';
            if (val.trim()) total += 1;
        } else {
            parts[f] = val == null || val === '' ? '–' : String(val);
        }
    });
    return {
        present: true,
        total: total,
        parts: parts,
        samples: samples,
        summaryText: fields
            .map(function (f) {
                return f + ': ' + (parts[f] != null ? parts[f] : '–');
            })
            .join(' · ')
    };
}

function pickRootFieldSummary(container, fields) {
    const parts = {};
    fields.forEach(function (f) {
        const val = container ? container[f] : null;
        if (val && typeof val === 'object') {
            const keys = Object.keys(val);
            parts[f] = keys.length + ' Felder';
        } else {
            parts[f] = val == null ? 'fehlt' : 'gesetzt';
        }
    });
    return {
        present: !!(container && fields.some(function (f) {
            return container[f] != null;
        })),
        total: fields.length,
        parts: parts,
        samples: [],
        summaryText: Object.keys(parts)
            .map(function (k) {
                return k + ': ' + parts[k];
            })
            .join(' · ')
    };
}

function extractSectionValue(container, section) {
    if (!container) return null;
    if (section.kind === 'year-fields') {
        const out = {};
        yearLabels(container).forEach(function (lab) {
            const bucket = container.years.byLabel[lab] || {};
            const slice = {};
            section.fields.forEach(function (f) {
                slice[f] = deepClone(bucket[f]);
            });
            out[lab] = slice;
        });
        return out;
    }
    if (section.kind === 'core-fields') {
        const core = container.core || {};
        const out = {};
        section.fields.forEach(function (f) {
            out[f] = deepClone(core[f]);
        });
        return out;
    }
    if (section.kind === 'root-field') {
        const out = {};
        section.fields.forEach(function (f) {
            out[f] = deepClone(container[f]);
        });
        return out;
    }
    return null;
}

function summarizeSection(container, section) {
    if (!container) {
        return {
            present: false,
            total: 0,
            samples: [],
            summaryText: 'fehlt'
        };
    }
    if (section.kind === 'year-fields') return pickYearFieldSummary(container, section.fields);
    if (section.kind === 'core-fields') return pickCoreFieldSummary(container, section.fields);
    return pickRootFieldSummary(container, section.fields);
}

/**
 * @param {SectionStatus} status
 * @param {{ present: boolean, total?: number }} localSum
 * @param {{ present: boolean, total?: number }} remoteSum
 * @param {'local'|'remote'|'equal'|'unknown'} [newerSide]
 */
function recommendImport(status, localSum, remoteSum, newerSide) {
    if (status === 'same' || status === 'missing' || status === 'local-only') return false;
    if (status === 'remote-only') return true;
    if (status === 'differs') {
        const remoteHas =
            remoteSum &&
            remoteSum.present !== false &&
            (remoteSum.total == null || remoteSum.total > 0 || remoteSum.summaryText !== 'fehlt');
        if (!remoteHas) return false;
        // Lokal leer, SharePoint hat Daten → immer empfehlen (auch wenn Gesamt-Backup lokal neuer wirkt)
        if ((localSum.total || 0) === 0 && (remoteSum.total || 0) > 0) return true;
        if (newerSide === 'remote') return true;
        if (newerSide === 'local') return false;
        return true;
    }
    return false;
}

function statusFromValues(localVal, remoteVal, hasLocalContainer, hasRemoteContainer) {
    const locEmpty = localVal == null || stableValue(localVal) === '{}' || stableValue(localVal) === '';
    const remEmpty = remoteVal == null || stableValue(remoteVal) === '{}' || stableValue(remoteVal) === '';
    if (!hasLocalContainer && !hasRemoteContainer) return 'missing';
    if (!hasRemoteContainer || remEmpty) {
        if (!hasLocalContainer || locEmpty) return 'missing';
        return 'local-only';
    }
    if (!hasLocalContainer || locEmpty) return 'remote-only';
    if (stableValue(localVal) === stableValue(remoteVal)) return 'same';
    return 'differs';
}

/**
 * @param {object|null|undefined} localPayload
 * @param {object|null|undefined} remotePayload
 * @param {{ newerSide?: string }} [ctx]
 */
export function buildSelectiveImportPlan(localPayload, remotePayload, ctx) {
    const options = ctx || {};
    const newerSide = options.newerSide || 'unknown';
    const locStore = storeFromPayload(localPayload);
    const remStore = storeFromPayload(remotePayload);
    const locV2 = parseSchooltool(locStore);
    const remV2 = parseSchooltool(remStore);
    const hasLocal = !!locV2;
    const hasRemote = !!remV2;

    /** @type {Array<object>} */
    const sections = [];

    SCHOOLTOOL_SELECTIVE_SECTIONS.forEach(function (def) {
        const localSum = summarizeSection(locV2, def);
        const remoteSum = summarizeSection(remV2, def);
        const localVal = extractSectionValue(locV2, def);
        const remoteVal = extractSectionValue(remV2, def);
        const status = statusFromValues(localVal, remoteVal, hasLocal, hasRemote);
        const recommended = recommendImport(status, localSum, remoteSum, newerSide);
        sections.push({
            id: def.id,
            label: def.label,
            hint: def.hint,
            group: def.group,
            kind: 'schooltool',
            status: status,
            recommended: recommended,
            local: localSum,
            remote: remoteSum
        });
    });

    STORAGE_SELECTIVE_BUNDLES.forEach(function (bundle) {
        const localParts = {};
        const remoteParts = {};
        let localPresent = false;
        let remotePresent = false;
        let differs = false;
        let onlyLocal = false;
        let onlyRemote = false;
        bundle.keys.forEach(function (key) {
            const hasL = Object.prototype.hasOwnProperty.call(locStore, key);
            const hasR = Object.prototype.hasOwnProperty.call(remStore, key);
            if (hasL) localPresent = true;
            if (hasR) remotePresent = true;
            localParts[key] = hasL ? 'vorhanden' : 'fehlt';
            remoteParts[key] = hasR ? 'vorhanden' : 'fehlt';
            if (hasL && hasR) {
                if (stableValue(locStore[key]) !== stableValue(remStore[key])) differs = true;
            } else if (hasL && !hasR) onlyLocal = true;
            else if (!hasL && hasR) onlyRemote = true;
        });
        let status = 'missing';
        if (!localPresent && !remotePresent) status = 'missing';
        else if (differs || (onlyLocal && onlyRemote)) status = 'differs';
        else if (onlyRemote && !localPresent) status = 'remote-only';
        else if (onlyLocal && !remotePresent) status = 'local-only';
        else if (onlyLocal) status = 'differs';
        else if (onlyRemote) status = 'differs';
        else status = 'same';

        const localSum = {
            present: localPresent,
            total: bundle.keys.filter(function (k) {
                return Object.prototype.hasOwnProperty.call(locStore, k);
            }).length,
            parts: localParts,
            samples: [],
            summaryText: localPresent ? 'Schlüssel lokal vorhanden' : 'fehlt lokal'
        };
        const remoteSum = {
            present: remotePresent,
            total: bundle.keys.filter(function (k) {
                return Object.prototype.hasOwnProperty.call(remStore, k);
            }).length,
            parts: remoteParts,
            samples: [],
            summaryText: remotePresent ? 'Schlüssel auf SharePoint' : 'fehlt auf SharePoint'
        };
        sections.push({
            id: bundle.id,
            label: bundle.label,
            hint: bundle.keys.join(', '),
            group: 'werkzeuge',
            kind: 'storage-bundle',
            keys: bundle.keys.slice(),
            status: status,
            recommended: recommendImport(status, localSum, remoteSum, newerSide),
            local: localSum,
            remote: remoteSum
        });
    });

    // Einzelne geänderte Schlüssel, die in keinem Bundle stecken
    const bundled = {};
    STORAGE_SELECTIVE_BUNDLES.forEach(function (b) {
        b.keys.forEach(function (k) {
            bundled[k] = true;
        });
    });
    bundled[SCHOOLTOOL_KEY] = true;

    const remKeys = Object.keys(remStore);
    remKeys.forEach(function (key) {
        if (bundled[key]) return;
        if (!(key.indexOf('ms365-') === 0 || key.indexOf('webuntis-') === 0)) return;
        const hasL = Object.prototype.hasOwnProperty.call(locStore, key);
        const status = !hasL
            ? 'remote-only'
            : stableValue(locStore[key]) === stableValue(remStore[key])
              ? 'same'
              : 'differs';
        if (status === 'same') return;
        const localSum = {
            present: hasL,
            total: hasL ? 1 : 0,
            samples: [],
            summaryText: hasL ? 'vorhanden' : 'fehlt'
        };
        const remoteSum = {
            present: true,
            total: 1,
            samples: [],
            summaryText: 'vorhanden'
        };
        sections.push({
            id: 'key:' + key,
            label: key,
            hint: 'Weiterer App-Schlüssel',
            group: 'weitere',
            kind: 'storage-key',
            keys: [key],
            status: status,
            recommended: status === 'remote-only' || (status === 'differs' && newerSide === 'remote'),
            local: localSum,
            remote: remoteSum
        });
    });

    const importable = sections.filter(function (s) {
        return s.status === 'differs' || s.status === 'remote-only';
    });
    const recommendedIds = sections.filter(function (s) {
        return s.recommended;
    }).map(function (s) {
        return s.id;
    });

    return {
        newerSide: newerSide,
        localExportedAt: localPayload && localPayload.exportedAt ? localPayload.exportedAt : '',
        remoteExportedAt: remotePayload && remotePayload.exportedAt ? remotePayload.exportedAt : '',
        sections: sections,
        importableCount: importable.length,
        recommendedIds: recommendedIds
    };
}

/**
 * Merged Schultool-Container: lokal als Basis, gewählte Bereiche aus remote.
 * @param {object|null} localV2
 * @param {object|null} remoteV2
 * @param {string[]} selectedIds
 */
export function mergeSchooltoolSelective(localV2, remoteV2, selectedIds) {
    const selected = {};
    (selectedIds || []).forEach(function (id) {
        selected[id] = true;
    });
    const defs = SCHOOLTOOL_SELECTIVE_SECTIONS.filter(function (d) {
        return selected[d.id];
    });
    if (!defs.length) return localV2 ? ensureContainerShape(localV2) : null;
    if (!remoteV2) throw new Error('SharePoint-Stammdaten fehlen für den selektiven Import.');

    const base = ensureContainerShape(localV2 || { core: {}, years: { current: remoteV2.years && remoteV2.years.current, byLabel: {} }, setup: {} });
    const remote = ensureContainerShape(remoteV2);

    // Schuljahr-Labels aus Remote anlegen, wenn nötig
    yearLabels(remote).forEach(function (lab) {
        if (!base.years.byLabel[lab]) base.years.byLabel[lab] = emptyYearBucket();
    });
    if (!base.years.current && remote.years.current) {
        base.years.current = remote.years.current;
    }

    defs.forEach(function (def) {
        if (def.kind === 'year-fields') {
            yearLabels(remote).forEach(function (lab) {
                if (!base.years.byLabel[lab]) base.years.byLabel[lab] = emptyYearBucket();
                const remBucket = remote.years.byLabel[lab] || emptyYearBucket();
                def.fields.forEach(function (f) {
                    base.years.byLabel[lab][f] = deepClone(remBucket[f]);
                });
            });
        } else if (def.kind === 'core-fields') {
            def.fields.forEach(function (f) {
                base.core[f] = deepClone(remote.core[f]);
            });
        } else if (def.kind === 'root-field') {
            def.fields.forEach(function (f) {
                base[f] = deepClone(remote[f]);
            });
        }
    });

    return base;
}

/**
 * Baut localStorage-Patch (nur zu schreibende Schlüssel).
 * @param {object|null|undefined} localPayload
 * @param {object|null|undefined} remotePayload
 * @param {string[]} selectedIds
 */
export function buildSelectiveLocalStoragePatch(localPayload, remotePayload, selectedIds) {
    const ids = Array.isArray(selectedIds) ? selectedIds.filter(Boolean) : [];
    if (!ids.length) throw new Error('Keine Bereiche ausgewählt.');

    const locStore = storeFromPayload(localPayload);
    const remStore = storeFromPayload(remotePayload);
    const patch = {};

    const schoolIds = ids.filter(function (id) {
        return SCHOOLTOOL_SELECTIVE_SECTIONS.some(function (s) {
            return s.id === id;
        });
    });
    if (schoolIds.length) {
        const merged = mergeSchooltoolSelective(parseSchooltool(locStore), parseSchooltool(remStore), schoolIds);
        patch[SCHOOLTOOL_KEY] = merged;
    }

    const keySet = {};
    ids.forEach(function (id) {
        if (id.indexOf('key:') === 0) {
            keySet[id.slice(4)] = true;
            return;
        }
        const bundle = STORAGE_SELECTIVE_BUNDLES.find(function (b) {
            return b.id === id;
        });
        if (bundle) {
            bundle.keys.forEach(function (k) {
                keySet[k] = true;
            });
        }
    });

    Object.keys(keySet).forEach(function (key) {
        if (!Object.prototype.hasOwnProperty.call(remStore, key)) return;
        patch[key] = deepClone(remStore[key]);
    });

    if (!Object.keys(patch).length) {
        throw new Error('Für die Auswahl gibt es keine übernehmbaren SharePoint-Daten.');
    }
    return patch;
}

/**
 * @param {object|null|undefined} localPayload
 * @param {object|null|undefined} remotePayload
 * @param {string[]} selectedIds
 * @returns {{ writtenKeys: string[], sectionIds: string[], previewLines: string[] }}
 */
export function describeSelectiveImport(localPayload, remotePayload, selectedIds) {
    const plan = buildSelectiveImportPlan(localPayload, remotePayload, {});
    const selected = {};
    (selectedIds || []).forEach(function (id) {
        selected[id] = true;
    });
    const lines = [];
    const sectionIds = [];
    plan.sections.forEach(function (s) {
        if (!selected[s.id]) return;
        sectionIds.push(s.id);
        lines.push(
            s.label +
                ': lokal ' +
                (s.local && s.local.summaryText ? s.local.summaryText : '–') +
                ' → SharePoint ' +
                (s.remote && s.remote.summaryText ? s.remote.summaryText : '–')
        );
    });
    const patch = buildSelectiveLocalStoragePatch(localPayload, remotePayload, selectedIds);
    return {
        writtenKeys: Object.keys(patch),
        sectionIds: sectionIds,
        previewLines: lines
    };
}

/**
 * Baut Patch und schreibt ihn lokal (ohne Sync-Meta / Reload).
 * @param {object} remotePayload
 * @param {string[]} selectedIds
 * @param {{ localPayload?: object }} [opts]
 */
export function applySelectiveLocalStoragePatch(remotePayload, selectedIds, opts) {
    const options = opts || {};
    const bb = typeof window !== 'undefined' ? window.ms365BrowserBackup : null;
    if (!bb || typeof bb.mergeLocalStoragePatch !== 'function') {
        throw new Error('Backup-Modul fehlt (mergeLocalStoragePatch).');
    }

    let localPayload = options.localPayload || null;
    if (!localPayload && typeof bb.buildBackup === 'function') {
        try {
            localPayload = bb.buildBackup();
        } catch {
            localPayload = { localStorage: {} };
        }
    }

    const patch = buildSelectiveLocalStoragePatch(localPayload, remotePayload, selectedIds);
    const result = bb.mergeLocalStoragePatch(patch);

    try {
        window.dispatchEvent(
            new CustomEvent('ms365-tenant-settings-changed', {
                detail: { source: 'spo-selective-import', keys: Object.keys(patch) }
            })
        );
        if (typeof bb.notifyLocalDataChanged === 'function') {
            bb.notifyLocalDataChanged('spo-selective-import', { keys: Object.keys(patch) });
        }
    } catch {
        /* ignore */
    }

    return {
        result: result,
        writtenKeys: Object.keys(patch),
        selectedIds: (selectedIds || []).slice(),
        patch: patch
    };
}

export function statusLabelDe(status) {
    const s = String(status || '');
    if (s === 'same') return 'Gleich';
    if (s === 'differs') return 'Unterschiedlich';
    if (s === 'local-only') return 'Nur lokal';
    if (s === 'remote-only') return 'Nur SharePoint';
    return 'Fehlt';
}

export default {
    SCHOOLTOOL_SELECTIVE_SECTIONS,
    STORAGE_SELECTIVE_BUNDLES,
    buildSelectiveImportPlan,
    mergeSchooltoolSelective,
    buildSelectiveLocalStoragePatch,
    describeSelectiveImport,
    applySelectiveLocalStoragePatch,
    statusLabelDe
};
