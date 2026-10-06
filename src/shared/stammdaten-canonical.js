/**
 * Eine Quelle für Stammdaten-Klassen (und Schüler) in allen Planern.
 *
 * Kanonisch: localStorage `ms365-schooltool-data-v2`
 *   → years.byLabel[current].classes / .students
 *
 * Spiegel (nur Lesen/Backup, nicht führend): `ms365-tenant-settings-v1`
 *
 * Aufräumen: Klassen aus allen Schuljahren + Legacy in das aktuelle Schuljahr mergen,
 * wenn dort noch nichts liegt – und den v1-Spiegel an v2 angleichen.
 */
(function (global) {
    'use strict';

    const KEY_V2 = 'ms365-schooltool-data-v2';
    const KEY_TENANT_V1 = 'ms365-tenant-settings-v1';

    function safeParse(raw) {
        try {
            return JSON.parse(String(raw || ''));
        } catch {
            return null;
        }
    }

    function normCode(v) {
        return String(v || '').trim();
    }

    function isLegacyAhw(code) {
        return /^[1-5]AHW$/i.test(normCode(code));
    }

    function normalizeClassRow(r) {
        if (!r || typeof r !== 'object') return null;
        const code = normCode(r.code || r.name);
        if (!code || isLegacyAhw(code)) return null;
        return {
            code,
            name: String(r.name || code).trim(),
            year: String(r.year || r.abschlussJahr || '').trim(),
            headName: String(r.headName || r.klassenvorstandName || '').trim(),
            headEmail: String(r.headEmail || r.klassenvorstandEmail || r.kvEmail || '')
                .trim()
                .toLowerCase(),
            stableMailNickname: String(r.stableMailNickname || '').trim()
        };
    }

    function mergeClassLists() {
        const out = [];
        const seen = new Set();
        const pushMany = (arr) => {
            (Array.isArray(arr) ? arr : []).forEach((raw) => {
                const row = normalizeClassRow(raw);
                if (!row || seen.has(row.code.toLowerCase())) return;
                seen.add(row.code.toLowerCase());
                out.push(row);
            });
        };

        try {
            const api = global.ms365AppDataV2;
            if (api && typeof api.getContainer === 'function') {
                const c = api.getContainer() || {};
                const by = (c.years && c.years.byLabel) || {};
                const cur = String((c.years && c.years.current) || '');
                const labels = cur
                    ? [cur].concat(Object.keys(by).filter((k) => k !== cur))
                    : Object.keys(by);
                labels.forEach((lab) => pushMany(by[lab] && by[lab].classes));
            }
        } catch {
            /* ignore */
        }

        try {
            const v2 = safeParse(global.localStorage && global.localStorage.getItem(KEY_V2));
            if (v2 && v2.years && v2.years.byLabel) {
                Object.keys(v2.years.byLabel).forEach((lab) => {
                    pushMany(v2.years.byLabel[lab] && v2.years.byLabel[lab].classes);
                });
            }
        } catch {
            /* ignore */
        }

        try {
            const t1 = safeParse(global.localStorage && global.localStorage.getItem(KEY_TENANT_V1));
            pushMany(t1 && t1.classes);
        } catch {
            /* ignore */
        }

        out.sort((a, b) => a.code.localeCompare(b.code, 'de', { numeric: true }));
        return out;
    }

    function currentYearLabel() {
        try {
            const api = global.ms365AppDataV2;
            if (api && typeof api.getContainer === 'function') {
                const c = api.getContainer();
                if (c && c.years && c.years.current) return String(c.years.current);
            }
        } catch {
            /* ignore */
        }
        try {
            const v2 = safeParse(global.localStorage && global.localStorage.getItem(KEY_V2));
            if (v2 && v2.years && v2.years.current) return String(v2.years.current);
        } catch {
            /* ignore */
        }
        return '';
    }

    function listClassesForCurrentYear() {
        const year = currentYearLabel();
        try {
            const api = global.ms365AppDataV2;
            if (api && typeof api.getYearBucket === 'function' && year) {
                const { bucket } = api.getYearBucket(year);
                return (bucket.classes || [])
                    .map(normalizeClassRow)
                    .filter(Boolean)
                    .sort((a, b) => a.code.localeCompare(b.code, 'de', { numeric: true }));
            }
        } catch {
            /* ignore */
        }
        return mergeClassLists();
    }

    /**
     * Speicher bereinigen / angleichen.
     * @returns {{ ok: boolean, year: string, classCount: number, mergedIntoCurrent: boolean, mirroredV1: boolean, message: string }}
     */
    function reconcileStammdatenStorage() {
        const all = mergeClassLists();
        const year = currentYearLabel();
        let mergedIntoCurrent = false;
        let mirroredV1 = false;

        const api = global.ms365AppDataV2;
        if (!api || typeof api.getYearBucket !== 'function' || typeof api.saveYearBucket !== 'function') {
            return {
                ok: all.length > 0,
                year,
                classCount: all.length,
                mergedIntoCurrent: false,
                mirroredV1: false,
                message:
                    all.length > 0
                        ? all.length + ' Klassen im Speicher gefunden (App-Daten-Modul fehlt noch).'
                        : 'Keine Klassen gefunden. Bitte Stammdaten → Klassen speichern.'
            };
        }

        const { year: y, bucket } = api.getYearBucket(year);
        const curClasses = (bucket.classes || []).map(normalizeClassRow).filter(Boolean);
        if (!curClasses.length && all.length) {
            bucket.classes = all.map((r) => ({
                code: r.code,
                name: r.name,
                year: r.year,
                headName: r.headName,
                headEmail: r.headEmail,
                stableMailNickname: r.stableMailNickname
            }));
            api.saveYearBucket(y, bucket);
            mergedIntoCurrent = true;
            if (typeof api.reconcileClassTeamsFromYearClasses === 'function') {
                try {
                    const c = api.getContainer();
                    api.reconcileClassTeamsFromYearClasses(c, y, bucket.classes);
                    if (typeof api.setContainer === 'function') api.setContainer(c);
                } catch {
                    /* ignore */
                }
            }
        }

        // v1-Spiegel: nicht mehr schreiben (kanonisch v2) – siehe tenant-storage-mirror.js
        try {
            const mirrorOn =
                global.ms365TenantStorageMirror &&
                typeof global.ms365TenantStorageMirror.isV1MirrorWriteEnabled === 'function' &&
                global.ms365TenantStorageMirror.isV1MirrorWriteEnabled();
            if (mirrorOn) {
                let snapshot = null;
                if (typeof global.ms365TenantSettingsLoad === 'function') {
                    snapshot = global.ms365TenantSettingsLoad();
                }
                if (snapshot && typeof snapshot === 'object') {
                    const classesNow = listClassesForCurrentYear();
                    if (classesNow.length) snapshot.classes = classesNow;
                    global.localStorage.setItem(KEY_TENANT_V1, JSON.stringify(snapshot));
                    mirroredV1 = true;
                }
            }
        } catch {
            /* ignore */
        }

        const finalCount = listClassesForCurrentYear().length || all.length;
        return {
            ok: finalCount > 0,
            year: y || year,
            classCount: finalCount,
            mergedIntoCurrent,
            mirroredV1,
            message:
                'Stammdaten-Quelle: App-Daten v2' +
                (y ? ' (Schuljahr ' + y + ')' : '') +
                ' · ' +
                finalCount +
                ' Klassen' +
                (mergedIntoCurrent ? ' · in aktuelles Schuljahr übernommen' : '') +
                (mirroredV1 ? ' · v1-Spiegel aktualisiert' : '')
        };
    }

    const api = {
        KEY_V2,
        KEY_TENANT_V1,
        listAllClasses: mergeClassLists,
        listClassesForCurrentYear,
        currentYearLabel,
        reconcileStammdatenStorage,
        normalizeClassRow
    };

    global.ms365StammdatenCanonical = api;

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = api;
    }
})(typeof window !== 'undefined' ? window : globalThis);
