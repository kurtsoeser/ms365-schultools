import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

vi.mock('../src/shared/entra-user-picker.js', () => ({
    pickEntraUser: vi.fn(async () => ({
        mail: 'direktor@ms365.schule',
        displayName: 'Direktor'
    }))
}));

import { pickEntraUser } from '../src/shared/entra-user-picker.js';
import { wireFlowSetupEmailPickers } from '../src/shared/flow-setup-email-pickers.js';

describe('wireFlowSetupEmailPickers', () => {
    const prevDoc = globalThis.document;
    let input;
    let btn;
    let listeners;

    beforeEach(() => {
        listeners = [];
        input = { value: '', dispatchEvent: vi.fn() };
        btn = {
            dataset: {},
            addEventListener: vi.fn((type, fn) => {
                listeners.push({ type, fn });
            })
        };
        globalThis.document = {
            getElementById(id) {
                if (id === 'aktEmailDirektion') return input;
                if (id === 'aktEmailDirektionPick') return btn;
                return null;
            }
        };
        pickEntraUser.mockClear();
    });

    afterEach(() => {
        if (prevDoc === undefined) delete globalThis.document;
        else globalThis.document = prevDoc;
    });

    it('verdrahtet Suchen und setzt E-Mail aus dem Picker', async () => {
        const onPicked = vi.fn();
        wireFlowSetupEmailPickers(
            [{ inputId: 'aktEmailDirektion', title: 'Direktion', onPicked }],
            { wiredFlag: 'testWired' }
        );
        expect(btn.dataset.testWired).toBe('1');
        expect(listeners).toHaveLength(1);
        await listeners[0].fn();
        expect(pickEntraUser).toHaveBeenCalled();
        expect(input.value).toBe('direktor@ms365.schule');
        expect(onPicked).toHaveBeenCalledWith('direktor@ms365.schule', expect.any(Object));
    });
});
