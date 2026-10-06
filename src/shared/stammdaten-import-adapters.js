/**
 * Konkrete Stammdaten-Import-Adapter (WebUntis, SIS, EdTech-Stub).
 */
import {
    registerStammdatenImportAdapter,
    ADAPTER_WEBUNTIS_HANDOFF,
    ADAPTER_SIS_FILE
} from './stammdaten-import-pipeline.js';
import { emptyStammdatenImportBundle, bundleFromWebuntisHandoff } from './stammdaten-import-bundle.js';
import { mergeWebuntisImportWithExisting } from './webuntis-stammdaten-wizard-logic.js';

export const ADAPTER_EDTECH_GENERIC = 'edtech-generic';

function emptyMerge(existingLines) {
    const ex = existingLines || {};
    return {
        lines: {
            teachersLines: String(ex.teachersLines || ''),
            studentsLines: String(ex.studentsLines || ''),
            subjectsLines: String(ex.subjectsLines || ''),
            classesLines: String(ex.classesLines || '')
        },
        stats: {},
        studentDiff: null
    };
}

/** WebUntis-Handoff (bereits in Pipeline registriert – hier für Re-Export). */
export { ADAPTER_WEBUNTIS_HANDOFF, ADAPTER_SIS_FILE };

/**
 * SIS / CSV-Excel-Schülerlisten: Merge über school-sis-import.
 */
export function registerSisFileAdapter() {
    registerStammdatenImportAdapter({
        id: ADAPTER_SIS_FILE,
        label: 'SIS / CSV / Excel (Schüler)',
        toBundle: function (input) {
            const r = input && typeof input === 'object' ? input : {};
            const records = Array.isArray(r.records) ? r.records : [];
            const b = emptyStammdatenImportBundle(ADAPTER_SIS_FILE);
            b.provenance = { source: String(r.source || 'sis'), file: true };
            b.counts.students = records.length;
            b.lines.studentsLines = String(r.lines || '');
            b.meta = r.meta && typeof r.meta === 'object' ? Object.assign({}, r.meta) : {};
            b.rawPayload = r;
            return b;
        },
        merge: function (existingLines, input, deps) {
            const sis = deps && deps.schoolSis;
            const parseStudents =
                deps && typeof deps.parseStudentsLines === 'function'
                    ? deps.parseStudentsLines
                    : function () {
                          return [];
                      };
            const ex = existingLines || {};
            const payload = input && typeof input === 'object' ? input : {};
            const records = Array.isArray(payload.records) ? payload.records : [];
            const mode = payload.mode === 'replace' ? 'replace' : 'merge';
            const existing = parseStudents(ex.studentsLines);

            if (!sis || typeof sis.applySisImport !== 'function') {
                const base = emptyMerge(ex);
                if (payload.lines) base.lines.studentsLines = String(payload.lines);
                return base;
            }

            const next = sis.applySisImport(existing, records, { mode: mode });
            const diff =
                typeof sis.diffSisImport === 'function' ? sis.diffSisImport(existing, records) : null;
            const lines =
                typeof sis.recordsToSemicolonLines === 'function'
                    ? sis.recordsToSemicolonLines(next)
                    : String(payload.lines || '');

            return {
                lines: {
                    teachersLines: String(ex.teachersLines || ''),
                    studentsLines: lines,
                    subjectsLines: String(ex.subjectsLines || ''),
                    classesLines: String(ex.classesLines || '')
                },
                stats: { students: (diff && diff.counts) || { added: 0, updated: 0, removed: 0 } },
                studentDiff: diff,
                records: next,
                mode: mode
            };
        }
    });
}

/**
 * EdTech-Stub: Interface für zukünftige Schul-Softwares (Sokrates, IServ, …).
 * Erwartet später: Spalten-Mapping → StammdatenImportBundle.
 */
export function registerEdtechGenericAdapter() {
    registerStammdatenImportAdapter({
        id: ADAPTER_EDTECH_GENERIC,
        label: 'EdTech (generisch, Stub)',
        toBundle: function (input) {
            const b = emptyStammdatenImportBundle(ADAPTER_EDTECH_GENERIC);
            b.provenance = {
                source: 'edtech',
                stub: true,
                note: 'Noch kein Produkt-Mapping – Spalten-Mapping folgt.'
            };
            if (input && typeof input === 'object') {
                b.meta = { receivedKeys: Object.keys(input) };
            }
            return b;
        },
        merge: function (existingLines) {
            const out = emptyMerge(existingLines);
            out.stats = { stub: true };
            out.message =
                'EdTech-Adapter ist vorbereitet, aber noch ohne Produkt-Mapping. Bitte WebUntis oder SIS/CSV nutzen.';
            return out;
        }
    });
}

/** Einmalige Registrierung aller konkreten Adapter (idempotent über Map-set). */
export function registerAllStammdatenImportAdapters() {
    registerStammdatenImportAdapter({
        id: ADAPTER_WEBUNTIS_HANDOFF,
        label: 'WebUntis (Übergabe)',
        toBundle: bundleFromWebuntisHandoff,
        merge: function (existingLines, payload, deps) {
            return mergeWebuntisImportWithExisting(existingLines, payload, deps);
        }
    });
    registerSisFileAdapter();
    registerEdtechGenericAdapter();
}

registerAllStammdatenImportAdapters();

if (typeof window !== 'undefined') {
    window.ms365StammdatenImportAdapters = {
        ADAPTER_EDTECH_GENERIC,
        ADAPTER_SIS_FILE,
        ADAPTER_WEBUNTIS_HANDOFF,
        registerAllStammdatenImportAdapters,
        registerSisFileAdapter,
        registerEdtechGenericAdapter
    };
}
