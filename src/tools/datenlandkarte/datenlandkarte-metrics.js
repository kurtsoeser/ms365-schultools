/**
 * Live-Zähler für Datenblöcke (Browser-Stammdaten, Setup, lokale Planer-URLs).
 */
import { SPO_LIST_PROBES } from './datenlandkarte-spo-metrics.js';
import { collectRegisterLayerChips } from '../../shared/stammdaten-register-status.js';

function loadSettings() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            return window.ms365TenantSettingsLoad();
        }
    } catch {
        /* ignore */
    }
    return null;
}

function loadSetup() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            return window.ms365AppDataV2.getSetup() || {};
        }
    } catch {
        /* ignore */
    }
    return {};
}

function loadYearBucket() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getYearBucket === 'function') {
            const yb = window.ms365AppDataV2.getYearBucket();
            return yb && yb.bucket ? yb.bucket : yb || {};
        }
    } catch {
        /* ignore */
    }
    return {};
}

function localStorageTrim(key) {
    try {
        return String(localStorage.getItem(key) || '').trim();
    } catch {
        return '';
    }
}

/**
 * @returns {Record<string, { value: number|string, hint?: string }>}
 */
export function collectDatenMetrics() {
    const settings = loadSettings() || {};
    const setup = loadSetup();
    const bucket = loadYearBucket();
    const classes = Array.isArray(settings.classes) ? settings.classes.length : 0;
    const teachers = Array.isArray(settings.teachers) ? settings.teachers.length : 0;
    const subjects = Array.isArray(settings.subjects) ? settings.subjects.length : 0;
    const arges = Array.isArray(settings.arges) ? settings.arges.length : 0;
    const students = Array.isArray(bucket.students) ? bucket.students.length : 0;
    const guardians = Array.isArray(bucket.guardians) ? bucket.guardians.length : 0;
    const ub = bucket.unterrichtsbelegung && Array.isArray(bucket.unterrichtsbelegung.rows)
        ? bucket.unterrichtsbelegung.rows.length
        : 0;
    const links = Array.isArray(setup.catalogLinks) ? setup.catalogLinks.length : 0;
    const intranet = String(setup.intranetSiteUrl || '').trim();
    const saSite = localStorageTrim('ms365-sa-site-url') || intranet;
    const pwSite = localStorageTrim('ms365-pw-site-url') || intranet;
    const frSite = localStorageTrim('ms365-freistellung-planer-site-v1') || intranet;
    const aktSite = localStorageTrim('ms365-akt-planer-site-v1') || intranet;
    const sisImports = Array.isArray(setup.sisImportHistory) ? setup.sisImportHistory.length : 0;

    const stammSum = classes + teachers + subjects + arges;

    const base = {
        stammHub: {
            value: stammSum,
            hint: stammSum ? 'Summe Klassen, Lehrkräfte, Fächer, ARGE' : 'Noch leer – Stammdaten'
        },
        classes: { value: classes },
        teachers: { value: teachers },
        subjects: { value: subjects },
        arges: { value: arges },
        students: { value: students },
        guardians: { value: guardians },
        unterrichtRows: { value: ub, hint: ub ? 'Kursteams-Belegung' : 'Noch keine Belegung gespeichert' },
        sisImports: {
            value: sisImports,
            hint: sisImports ? 'Einträge in Import-Historie' : 'Noch kein WebUntis/SIS-Import'
        },
        schularbeitenSite: {
            value: saSite ? '✓' : '—',
            hint: saSite ? 'Planer-Site konfiguriert' : 'Site-URL im Planer setzen'
        },
        projektwochenSite: {
            value: pwSite ? '✓' : '—',
            hint: pwSite ? 'Planer-Site konfiguriert' : 'Site-URL im Planer setzen'
        },
        freistellungSite: {
            value: frSite ? '✓' : '—',
            hint: frSite ? 'Freistellungen-Site' : 'Site im Planer/Setup setzen'
        },
        aktivitaetenSite: {
            value: aktSite ? '✓' : '—',
            hint: aktSite ? 'Schulaktivitäten-Site' : 'Site im Planer setzen'
        },
        catalogLinks: { value: links, hint: links ? 'Verknüpfungen Einrichtung' : 'Noch keine Gruppen-Links' }
    };

    SPO_LIST_PROBES.forEach((p) => {
        if (!base[p.countKey]) {
            base[p.countKey] = { value: '…', hint: 'SharePoint-Zähler: Aktualisieren oder anmelden' };
        }
    });

    try {
        if (typeof window !== 'undefined') {
            const chips = collectRegisterLayerChips();
            const bad = chips.filter(function (c) {
                return c.kind === 'error' || c.kind === 'warn';
            });
            base.registerAmpel = {
                value: bad.length ? '!' : '✓',
                hint: chips.map(function (c) {
                    return c.label;
                }).join(' · ')
            };
        }
    } catch {
        /* ignore */
    }

    return base;
}

/** @param {Record<string, object>} local @param {Record<string, object>} spo */
export function mergeDatenMetrics(local, spo) {
    return { ...(local || {}), ...(spo || {}) };
}

export function spoLoadingMetrics() {
    /** @type {Record<string, { value: string, hint: string }>} */
    const o = {};
    SPO_LIST_PROBES.forEach((p) => {
        o[p.countKey] = { value: '…', hint: 'Lade SharePoint-Zähler …' };
    });
    return o;
}

/**
 * @param {string} countKey
 * @param {Record<string, { value: number|string, hint?: string }>} metrics
 */
export function formatBlockCount(countKey, metrics) {
    const m = metrics && metrics[countKey];
    if (!m) return { text: '—', hint: '', unit: '' };
    const v = m.value;
    const text = typeof v === 'number' ? (v > 9999 ? '9999+' : String(v)) : String(v);
    const hint = m.hint || '';
    const unit =
        typeof v === 'number'
            ? hint.indexOf('SharePoint') !== -1 ||
              hint.indexOf('Intranet-Sync') !== -1 ||
              String(countKey).startsWith('spoList')
                ? ' Zeilen'
                : ' Einträge'
            : '';
    return { text, hint, unit };
}
