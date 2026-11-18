import { describe, it, expect } from 'vitest';
import {
    parseCodeList,
    joinCodeList,
    filterFreistellungen,
    resolveFrOwnership,
    assertFrOwnership
} from '../src/tools/freistellung-planer/freistellung-planer-logic.js';
import {
    createInitialState,
    navSectionsForRole,
    viewsForRole
} from '../src/tools/freistellung-planer/freistellung-planer-state.js';
import { DEMO_KV_PERSONA, DEMO_SEED_TAG } from '../src/tools/freistellung-planer/freistellung-planer-demo.js';

describe('freistellung upgrade – multi filter & ownership', () => {
    it('parseCodeList / joinCodeList', () => {
        expect(parseCodeList('1AK; 2AK')).toEqual(['1AK', '2AK']);
        expect(joinCodeList(['Ausstehend', 'Genehmigt'])).toBe('Ausstehend; Genehmigt');
    });

    it('filterFreistellungen akzeptiert Mehrfach Status/Klasse', () => {
        const items = [
            { klasse: '1AK', status: 'Ausstehend', kategorie: 'Sonstiges' },
            { klasse: '2AK', status: 'Genehmigt', kategorie: 'Ärztlicher Termin' },
            { klasse: '1AK', status: 'Genehmigt', kategorie: 'Sonstiges' }
        ];
        const hit = filterFreistellungen(items, { klasse: '1AK;2AK', status: 'Genehmigt' }, {});
        expect(hit).toHaveLength(2);
        expect(hit.every((i) => i.status === 'Genehmigt')).toBe(true);
    });

    it('resolveFrOwnership setzt beantragtVon/authorEmail', () => {
        const owned = resolveFrOwnership({ klasse: '1AK' }, { accountEmail: 'a@schule.at', accountName: 'Ada' });
        expect(owned.beantragtVon).toBe('a@schule.at');
        expect(owned.authorEmail).toBe('a@schule.at');
        expect(owned.authorName).toBe('Ada');
        expect(assertFrOwnership({}, { accountEmail: '' }).ok).toBe(false);
        expect(assertFrOwnership({}, { accountEmail: 'x@y.at' }).ok).toBe(true);
    });
});

describe('freistellung upgrade – state defaults & nav', () => {
    it('createInitialState enthält StammdatenLive, Filter-Multi und Kalender-Modus', () => {
        const s = createInitialState();
        expect(s.stammdatenLive).toBeTruthy();
        expect(s.filterMultiOpen).toEqual({ klasse: false, status: false, kategorie: false });
        expect(s.calMode).toBe('month');
        expect(s.calWeekAnchor).toBeTruthy();
        expect(s.demoPersonaEmail).toBe('');
    });

    it('Schüler-Nav ist flach ohne Sektionslabel', () => {
        const sections = navSectionsForRole('schueler');
        expect(sections).toHaveLength(1);
        expect(sections[0].label).toBe('');
        expect(viewsForRole('schueler').map((v) => v.id)).toContain('kalender');
    });

    it('Demo-Persona Konstanten vorhanden', () => {
        expect(DEMO_KV_PERSONA.email).toContain('@');
        expect(DEMO_SEED_TAG).toMatch(/^fr-demo-/);
    });
});
