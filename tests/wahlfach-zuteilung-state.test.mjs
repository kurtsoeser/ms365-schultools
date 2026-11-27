import { describe, expect, it } from 'vitest';
import { buildWahlfachDemoState } from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-demo.js';
import {
    canonicalKlasse,
    classCodesFromStammdaten,
    classResultRows,
    ensureWzView,
    findOwnAnmeldung,
    kvKlassenForEmail,
    matchAnmeldungIndex,
    navSectionsForWz,
    shouldShowWahlfachZuteilungMenu
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

    it('erkennt den Klassenvorstand nur an der hinterlegten Klasse', () => {
        const stammdaten = {
            classes: [
                { code: '3AK', headEmail: 'kv@schule.at' },
                { code: '3BK', kvEmail: 'andere@schule.at' }
            ]
        };
        expect(kvKlassenForEmail(stammdaten, 'kv@schule.at')).toEqual(['3AK']);
        expect(kvKlassenForEmail(stammdaten, 'lehrer@schule.at')).toEqual([]);
        expect(kvKlassenForEmail(stammdaten, '')).toEqual([]);
    });

    it('zeigt einem Klassenvorstand nur die eigene Klasse', () => {
        const demo = buildWahlfachDemoState();
        const rows = classResultRows(demo, ['3AK'], [
            { name: 'Otto Neu', klasse: '3AK', email: 'otto.neu@demo.schule' },
            { name: 'Lara Fuchs', klasse: '1BK', email: 'lara.fuchs@demo.schule' }
        ]);
        const names = rows.map((row) => row.name);
        expect(names).toContain('Anna Berger');
        expect(names).toContain('Felix Moser');
        expect(names).toContain('Otto Neu');
        expect(names).not.toContain('Lara Fuchs');
        expect(names).not.toContain('Lukas Hofer');
        expect(rows.find((row) => row.name === 'Otto Neu').anmeldung).toBeNull();
    });

    it('führt Klassenschreibweisen zusammen und aktualisiert dieselbe Anmeldung', () => {
        expect(canonicalKlasse('3 ak', ['3AK'])).toBe('3AK');
        expect(canonicalKlasse('3ak', [])).toBe('3AK');
        const list = [
            { id: 'a1', schuelerEmail: 'ada@schule.at', status: 'eingereicht' },
            { id: 'a2', schuelerEmail: 'ada@schule.at', status: 'storniert' }
        ];
        expect(matchAnmeldungIndex(list, { schuelerEmail: 'Ada@schule.at' })).toBe(0);
        expect(findOwnAnmeldung({ anmeldungen: list }, 'ada@schule.at').id).toBe('a1');
    });

    it('zeigt die Zuteilung nur Direktion, Klassenvorstand und Schul-IT', () => {
        expect(shouldShowWahlfachZuteilungMenu({ email: 'lehrer@schule.at' })).toBe(false);
        expect(
            shouldShowWahlfachZuteilungMenu({ email: 'kv@schule.at', kvKlassen: ['3AK'] })
        ).toBe(true);
        expect(shouldShowWahlfachZuteilungMenu({ email: 'slg@schule.at', localAdmin: true })).toBe(
            true
        );
        expect(shouldShowWahlfachZuteilungMenu({ isIt: true, email: 'it@schule.at' })).toBe(true);
        expect(shouldShowWahlfachZuteilungMenu({ demoRole: 'lehrer', email: 'it@schule.at' })).toBe(
            false
        );
        expect(shouldShowWahlfachZuteilungMenu({})).toBe(true);
    });
});
