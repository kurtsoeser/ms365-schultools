import { describe, it, expect } from 'vitest';
import {
    buildSeminarShowFormula,
    buildTitleHideFormula,
    buildTitleColumnFormatJson,
    buildPruefplanHeaderExpression,
    buildPillColumnFormatJson,
    buildWahlfachColumnFormatJson,
    buildClientFormCustomFormatterObject,
    buildClientFormCustomFormatterString,
    buildFieldFormatSpecs
} from '../src/tools/srdp-anmeldung/srdp-anmeldung-formatting.js';

describe('srdp-anmeldung-formatting', () => {
    it('Seminar-Show-Formel erkennt Seminar-Präfix', () => {
        const f = buildSeminarShowFormula();
        expect(f).toMatch(/^=/);
        expect(f).toContain('WahlfachMuendlich');
        expect(f).toContain('Seminar');
        expect(f).toContain('indexOf');
    });

    it('Title wird im Formular ausgeblendet', () => {
        expect(buildTitleHideFormula()).toBe('=false');
    });

    it('Title-Column-Format nutzt Nachname/Vorname', () => {
        const j = buildTitleColumnFormatJson();
        expect(j.txtContent).toContain('Nachname');
        expect(j.txtContent).toContain('Vorname');
    });

    it('Prüfplan-Header deckt Profil-Varianten ab', () => {
        const expr = buildPruefplanHeaderExpression('hak');
        expect(expr).toContain('Variante 1');
        expect(expr).toContain('Variante 2');
        expect(expr).toContain('Variante 3');
        expect(expr).toContain('BFK');
        expect(buildPruefplanHeaderExpression('htl')).toContain('SWP');
        expect(buildPruefplanHeaderExpression('ahs')).toContain('mündlich 1');
    });

    it('Variante-Pills und Wahlfach-Format', () => {
        const v = buildPillColumnFormatJson('variante', 'hak');
        expect(v.children.length).toBe(3);
        expect(buildPillColumnFormatJson('variante', 'hlw').children.length).toBe(2);
        const w = buildWahlfachColumnFormatJson();
        expect(w.style['background-color']).toContain('Seminar');
    });

    it('ClientFormCustomFormatter hat Sektionen ohne Title', () => {
        const obj = buildClientFormCustomFormatterObject('hak');
        const allFields = obj.bodyJSONFormatter.sections.flatMap((s) => s.fields);
        expect(allFields).toContain('Nachname');
        expect(allFields).toContain('Variante');
        expect(allFields).not.toContain('Title');
        expect(allFields).toContain('Seminar');
        const str = buildClientFormCustomFormatterString('htl');
        expect(() => JSON.parse(str)).not.toThrow();
        expect(buildClientFormCustomFormatterObject('ahs').bodyJSONFormatter.sections[1].displayname).toMatch(/Abschließende/);
    });

    it('buildFieldFormatSpecs für Terminjahr', () => {
        const specs = buildFieldFormatSpecs(2026, 'hak');
        const title = specs.find((s) => s.internalName === 'Title');
        expect(title?.required).toBe(false);
        expect(title?.conditionalShowFormula).toBe('=false');
        const seminar = specs.find((s) => s.internalName === 'Seminar');
        expect(seminar?.conditionalShowFormula).toContain('Seminar');
        const jahr = specs.find((s) => s.internalName === 'TerminJahr');
        expect(jahr?.defaultValue).toBe('2026');
        expect(buildFieldFormatSpecs(2026, 'hlw').some((s) => s.internalName === 'KlausurKombi')).toBe(true);
    });
});
