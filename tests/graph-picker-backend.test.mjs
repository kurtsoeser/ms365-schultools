import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import { getGraphPickerApi } from '../src/shared/graph-picker-backend.js';

describe('graph-picker-backend', () => {
    const orig = globalThis.window;

    beforeEach(() => {
        globalThis.window = {};
    });

    afterEach(() => {
        globalThis.window = orig;
    });

    it('nutzt ms365GraphUnifiedGroups wenn spo-graph fehlt', async () => {
        const token = vi.fn().mockResolvedValue('tok');
        const json = vi.fn().mockResolvedValue({ value: [] });
        globalThis.window.ms365GraphUnifiedGroups = {
            getGraphToken: token,
            graphJson: json
        };
        const api = getGraphPickerApi();
        await api.getGraphToken(['scope']);
        expect(token).toHaveBeenCalled();
        await api.graphJson('GET', '/groups', 'tok', undefined, 'v1.0');
        expect(json).toHaveBeenCalledWith('GET', '/groups', 'tok', undefined, undefined);
    });

    it('wirft verständliche Meldung ohne Backend', () => {
        expect(() => getGraphPickerApi()).toThrow(/Microsoft Graph/);
    });
});
