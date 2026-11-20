import { describe, it, expect } from 'vitest';
import {
    navSectionsForRole,
    viewsForRole,
    countOpenAusstehend,
    scopeFromState,
    visibleItemsForState
} from '../src/tools/lehrer-freistellung-planer/lfr-state.js';

describe('lfr navSectionsForRole', () => {
    it('Lehrkraft: flache Arbeit ohne System', () => {
        const secs = navSectionsForRole('lehrer');
        expect(secs).toHaveLength(1);
        expect(secs[0].label).toBe('');
        const ids = viewsForRole('lehrer').map((v) => v.id);
        expect(ids).toEqual(['dashboard', 'meine', 'kalender', 'antrag']);
        expect(ids).not.toContain('einrichtung');
        expect(ids).not.toContain('freigabe');
    });

    it('Direktion: Arbeit · Auswertung · System', () => {
        const secs = navSectionsForRole('direktion');
        expect(secs.map((s) => s.id)).toEqual(['arbeit', 'auswertung', 'system']);
        const ids = viewsForRole('direktion').map((v) => v.id);
        expect(ids).toContain('freigabe');
        expect(ids).toContain('liste');
        expect(ids).toContain('export');
        expect(ids).toContain('einrichtung');
        expect(ids).not.toContain('meine');
    });

    it('zählt offene Ausstehend', () => {
        expect(
            countOpenAusstehend([
                { status: 'Ausstehend' },
                { status: 'Genehmigt' },
                { status: 'ausstehend' }
            ])
        ).toBe(2);
    });

    it('Lehrkraft-Scope: Kalender/Listen nur eigene Anträge', () => {
        const state = {
            role: 'lehrer',
            accountEmail: 'lehrer01.demo@ms365.schule',
            items: [
                { lehrerEmail: 'kurt@schule.at', status: 'Genehmigt', titel: 'Fremd' },
                {
                    lehrerEmail: 'lehrer01.demo@ms365.schule',
                    status: 'Ausstehend',
                    titel: 'Eigen'
                },
                { lehrerEmail: '', status: 'Genehmigt', titel: 'OhneMail' }
            ]
        };
        expect(scopeFromState(state).scopeAll).toBe(false);
        expect(scopeFromState(state, { scopeAll: false }).scopeAll).toBe(false);
        const mine = visibleItemsForState(state, {}, { scopeAll: false });
        expect(mine.map((x) => x.titel)).toEqual(['Eigen']);
    });

    it('Direktion sieht alle, explizit scopeAll false nur eigene', () => {
        const state = {
            role: 'direktion',
            accountEmail: 'direktor@ms365.schule',
            items: [
                { lehrerEmail: 'a@schule.at', titel: 'A' },
                { lehrerEmail: 'direktor@ms365.schule', titel: 'Dir' }
            ]
        };
        expect(visibleItemsForState(state, {}).map((x) => x.titel)).toEqual(['A', 'Dir']);
        expect(
            visibleItemsForState(state, {}, { scopeAll: false }).map((x) => x.titel)
        ).toEqual(['Dir']);
    });
});
