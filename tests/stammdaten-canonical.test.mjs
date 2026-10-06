import { describe, it, expect, beforeEach } from 'vitest';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);

describe('stammdaten-canonical', () => {
    let store;
    let api;

    beforeEach(() => {
        store = {};
        globalThis.localStorage = {
            getItem(k) {
                return store[k] ?? null;
            },
            setItem(k, v) {
                store[k] = String(v);
            },
            removeItem(k) {
                delete store[k];
            }
        };
        globalThis.window = globalThis;
        // Reload module into fresh global
        delete require.cache[require.resolve('../src/shared/stammdaten-canonical.js')];
        require('../src/shared/stammdaten-canonical.js');
        api = globalThis.ms365StammdatenCanonical;
    });

    it('listAllClasses liest aus app-data-v2 Jahrgangs-Bucket', () => {
        store['ms365-schooltool-data-v2'] = JSON.stringify({
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        classes: [
                            { code: '1A', name: 'DEMO Klasse 1A', year: '2032' },
                            { code: '1B', name: 'DEMO Klasse 1B', year: '2032' },
                            { code: '1AHW', name: 'Legacy' }
                        ]
                    }
                }
            }
        });
        expect(api.listAllClasses().map((c) => c.code)).toEqual(['1A', '1B']);
    });

    it('reconcile übernimmt Klassen ins aktuelle Schuljahr wenn leer', () => {
        store['ms365-schooltool-data-v2'] = JSON.stringify({
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': { classes: [] },
                    '2024/25': {
                        classes: [{ code: '2A', name: 'DEMO 2A', year: '2031' }]
                    }
                }
            }
        });
        const container = JSON.parse(store['ms365-schooltool-data-v2']);
        globalThis.ms365AppDataV2 = {
            getContainer() {
                return container;
            },
            setContainer(c) {
                Object.assign(container, c);
            },
            getYearBucket(label) {
                const y = label || container.years.current;
                if (!container.years.byLabel[y]) container.years.byLabel[y] = { classes: [] };
                return { year: y, bucket: container.years.byLabel[y] };
            },
            saveYearBucket(label, bucket) {
                container.years.byLabel[label] = bucket;
                store['ms365-schooltool-data-v2'] = JSON.stringify(container);
            },
            reconcileClassTeamsFromYearClasses() {}
        };
        globalThis.ms365TenantSettingsLoad = () => ({
            schoolName: 'Demo',
            classes: [],
            teachers: [],
            students: []
        });

        const result = api.reconcileStammdatenStorage();
        expect(result.ok).toBe(true);
        expect(result.mergedIntoCurrent).toBe(true);
        expect(result.classCount).toBe(1);
        expect(container.years.byLabel['2025/26'].classes[0].code).toBe('2A');
    });
});
