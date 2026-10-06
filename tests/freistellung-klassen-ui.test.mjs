import { describe, it, expect, beforeEach } from 'vitest';
import { createRequire } from 'module';
import { listStammdatenKlassenForFreistellung } from '../src/tools/freistellung-planer/freistellung-klassen-ui.js';

const require = createRequire(import.meta.url);

describe('listStammdatenKlassenForFreistellung', () => {
    let store;

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
        delete require.cache[require.resolve('../src/shared/stammdaten-canonical.js')];
        require('../src/shared/stammdaten-canonical.js');
        store['ms365-schooltool-data-v2'] = JSON.stringify({
            years: {
                current: '2025/26',
                byLabel: {
                    '2025/26': {
                        classes: [
                            { code: '1B', name: 'DEMO Klasse 1B', year: '2032' },
                            { code: '1A', name: 'DEMO Klasse 1A', year: '2032' },
                            { code: '1AHW', name: 'Legacy' }
                        ]
                    }
                }
            }
        });
    });

    it('liefert Stammdaten-Klassen über kanonische Quelle', () => {
        const rows = listStammdatenKlassenForFreistellung();
        expect(rows.map((r) => r.code)).toEqual(['1A', '1B']);
    });
});
