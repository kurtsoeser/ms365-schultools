import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import { searchEntraUsers, formatUserOptionLabel } from '../src/shared/entra-user-picker.js';

describe('searchEntraUsers', () => {
    const orig = globalThis.window;

    beforeEach(() => {
        globalThis.window = {
            ms365GraphUnifiedGroups: {
                getGraphToken: vi.fn().mockResolvedValue('tok'),
                graphJson: vi.fn()
            },
            ms365SpoGraph: undefined
        };
    });

    afterEach(() => {
        globalThis.window = orig;
    });

    it('formatUserOptionLabel baut lesbare Labels', () => {
        expect(formatUserOptionLabel({ displayName: 'Max Zisser', mail: 'max.zisser@schule.at' })).toBe(
            'Max Zisser (max.zisser@schule.at)'
        );
        expect(formatUserOptionLabel({ displayName: 'Ohne Mail' })).toBe('Ohne Mail');
    });

    it('leere Query liefert [] ohne Graph-Call', async () => {
        const list = await searchEntraUsers('  ');
        expect(list).toEqual([]);
        expect(globalThis.window.ms365GraphUnifiedGroups.graphJson).not.toHaveBeenCalled();
    });

    it('findet Nachnamen-Fragmente über $search (nicht nur startswith)', async () => {
        const json = globalThis.window.ms365GraphUnifiedGroups.graphJson;
        json.mockImplementation(async (_m, path) => {
            if (String(path).indexOf('$search=') !== -1) {
                return {
                    value: [
                        {
                            id: 'u1',
                            displayName: 'Max Zisser',
                            mail: 'max.zisser@schule.at',
                            userPrincipalName: 'max.zisser@schule.at',
                            surname: 'Zisser'
                        }
                    ]
                };
            }
            return { value: [] };
        });

        const list = await searchEntraUsers('ziss');
        expect(list).toHaveLength(1);
        expect(list[0].displayName).toBe('Max Zisser');
        expect(json).toHaveBeenCalled();
        const searchCall = json.mock.calls.find((c) => String(c[1]).indexOf('$search=') !== -1);
        expect(searchCall).toBeTruthy();
        expect(searchCall[4]).toEqual({ ConsistencyLevel: 'eventual' });
    });

    it('fällt auf contains/startswith zurück wenn $search fehlschlägt', async () => {
        const json = globalThis.window.ms365GraphUnifiedGroups.graphJson;
        json.mockImplementation(async (_m, path) => {
            const p = String(path);
            if (p.indexOf('$search=') !== -1) {
                throw new Error('$search unavailable');
            }
            if (p.indexOf('contains(') !== -1) {
                return {
                    value: [
                        {
                            id: 'u2',
                            displayName: 'Anna Zisser',
                            mail: 'anna.zisser@schule.at',
                            userPrincipalName: 'anna.zisser@schule.at'
                        }
                    ]
                };
            }
            return { value: [] };
        });

        const list = await searchEntraUsers('ziss');
        expect(list).toHaveLength(1);
        expect(list[0].id).toBe('u2');
    });

    it('löst E-Mail exakt auf', async () => {
        const json = globalThis.window.ms365GraphUnifiedGroups.graphJson;
        json.mockImplementation(async (_m, path) => {
            const decoded = decodeURIComponent(String(path));
            if (decoded.indexOf("mail eq 'kurt@schule.at'") !== -1) {
                return {
                    value: [
                        {
                            id: 'u3',
                            displayName: 'Kurt',
                            mail: 'kurt@schule.at',
                            userPrincipalName: 'kurt@schule.at'
                        }
                    ]
                };
            }
            return { value: [] };
        });

        const list = await searchEntraUsers('kurt@schule.at');
        expect(list).toHaveLength(1);
        expect(list[0].mail).toBe('kurt@schule.at');
    });
});
