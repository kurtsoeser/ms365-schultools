import { describe, expect, it } from 'vitest';
import {
    classCodesFromStammdaten,
    ensureWzView,
    navSectionsForWz
} from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-state.js';

describe('wahlfach-zuteilung-state', () => {
    it('baut Nav ohne System standardmäßig und mit System bei Flag', () => {
        const plain = navSectionsForWz({ showSystem: false });
        expect(plain.every((s) => s.id !== 'system')).toBe(true);
        expect(plain.flatMap((s) => s.items).some((v) => v.id === 'zuteilung')).toBe(true);

        const withSys = navSectionsForWz({ showSystem: true });
        expect(withSys.some((s) => s.id === 'system')).toBe(true);
        expect(ensureWzView('system', { showSystem: false })).toBe('uebersicht');
        expect(ensureWzView('system', { showSystem: true })).toBe('system');
    });

    it('extrahiert Klassen-Codes aus Stammdaten', () => {
        const codes = classCodesFromStammdaten({
            classes: [{ code: '3AK' }, { name: '3BK' }, { code: '3AK' }]
        });
        expect(codes).toEqual(['3AK', '3BK']);
    });
});
