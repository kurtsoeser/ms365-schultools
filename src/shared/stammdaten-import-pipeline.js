/**
 * Einheitliche Import-Pipeline: Adapter → Merge → Review → Apply.
 */
import { mergeWebuntisImportWithExisting } from './webuntis-stammdaten-wizard-logic.js';
import { emptyStammdatenImportBundle, bundleFromWebuntisHandoff } from './stammdaten-import-bundle.js';

export const ADAPTER_WEBUNTIS_HANDOFF = 'webuntis-handoff';
export const ADAPTER_SIS_FILE = 'sis-file';
export const ADAPTER_EDTECH_GENERIC = 'edtech-generic';

/** @type {Map<string, StammdatenImportAdapter>} */
const registry = new Map();

/**
 * @typedef {object} StammdatenImportAdapter
 * @property {string} id
 * @property {string} label
 * @property {(input: unknown, ctx?: object) => object} [toBundle]
 * @property {(existingLines: object, input: unknown, deps: object) => object} merge
 */

export function registerStammdatenImportAdapter(adapter) {
    if (!adapter || !adapter.id) return;
    registry.set(String(adapter.id), adapter);
}

export function getStammdatenImportAdapter(id) {
    return registry.get(String(id || '')) || null;
}

export function listStammdatenImportAdapterIds() {
    return Array.from(registry.keys());
}

registerStammdatenImportAdapter({
    id: ADAPTER_WEBUNTIS_HANDOFF,
    label: 'WebUntis (Übergabe)',
    toBundle: function (payload) {
        return bundleFromWebuntisHandoff(payload);
    },
    merge: function (existingLines, payload, deps) {
        return mergeWebuntisImportWithExisting(existingLines, payload, deps);
    }
});

/**
 * @param {string} adapterId
 * @param {object} existingLines
 * @param {unknown} input
 * @param {object} deps
 */
export function mergeStammdatenImport(adapterId, existingLines, input, deps) {
    const adapter = getStammdatenImportAdapter(adapterId);
    if (!adapter || typeof adapter.merge !== 'function') {
        throw new Error('Unbekannter Stammdaten-Import-Adapter: ' + String(adapterId));
    }
    return adapter.merge(existingLines, input, deps);
}

/**
 * @param {string} adapterId
 * @param {unknown} input
 * @param {object} [ctx]
 */
export function toStammdatenImportBundle(adapterId, input, ctx) {
    const adapter = getStammdatenImportAdapter(adapterId);
    if (adapter && typeof adapter.toBundle === 'function') {
        return adapter.toBundle(input, ctx);
    }
    return emptyStammdatenImportBundle(adapterId);
}

/**
 * @param {Array<{ name?: string, aoa?: unknown[][] }>} sheets
 * @returns {{ adapterId: string|null, confidence: 'high'|'low'|'none', classified?: object }}
 */
export function detectStammdatenImportAdapter(sheets) {
    const list = Array.isArray(sheets) ? sheets : [];
    const wu = typeof globalThis !== 'undefined' ? globalThis.ms365WebuntisExportImport : null;
    if (wu && typeof wu.classifySheets === 'function') {
        const classified = wu.classifySheets(list);
        if (
            classified &&
            (classified.studentAoa ||
                classified.guardianAoa ||
                classified.teacherAoa ||
                classified.subjectAoa ||
                classified.classAoa)
        ) {
            const multi =
                !!(classified.studentAoa || classified.guardianAoa) &&
                !!(classified.teacherAoa || classified.subjectAoa || classified.classAoa);
            return {
                adapterId: ADAPTER_SIS_FILE,
                confidence: 'high',
                classified: classified,
                hint: multi ? 'webuntis-or-sis' : 'sis'
            };
        }
    }
    const first = list[0];
    if (first && Array.isArray(first.aoa) && first.aoa.length > 1) {
        const headers = (first.aoa[0] || [])
            .map(function (h) {
                return String(h || '').toLowerCase();
            })
            .join(' ');
        if (/sokrates|iserv|edupage|untis|webuntis|schueler|klasse|guardian/.test(headers)) {
            if (/sokrates|iserv|edupage/.test(headers)) {
                return { adapterId: ADAPTER_EDTECH_GENERIC, confidence: 'low' };
            }
            return { adapterId: ADAPTER_SIS_FILE, confidence: 'low' };
        }
        return { adapterId: ADAPTER_SIS_FILE, confidence: 'low' };
    }
    return { adapterId: null, confidence: 'none' };
}

if (typeof window !== 'undefined') {
    window.ms365StammdatenImportPipeline = {
        ADAPTER_WEBUNTIS_HANDOFF,
        ADAPTER_SIS_FILE,
        ADAPTER_EDTECH_GENERIC,
        registerStammdatenImportAdapter,
        getStammdatenImportAdapter,
        listStammdatenImportAdapterIds,
        mergeStammdatenImport,
        toStammdatenImportBundle,
        detectStammdatenImportAdapter
    };
}
