import { describe, it, expect } from 'vitest';
import {
    directoryStatusView,
    groupLinkStatusView,
    presentText,
    presentValue
} from '../src/shared/tenant-workbench-chrome.js';

describe('tenant-workbench-chrome', () => {
    it('zeigt Verknüpft statt der Object-ID', () => {
        const view = directoryStatusView('a@schule.at', {
            graphUserId: '15aff94c-c2f0-4b2a-9c1d-000000000001',
            displayName: 'Ada'
        });
        expect(view.label).toBe('Verknüpft');
        expect(view.kind).toBe('linked');
        expect(view.title).toContain('15aff94c-c2f0-4b2a-9c1d-000000000001');
        expect(view.label).not.toContain('15aff94c');
    });

    it('unterscheidet offen, nicht verknüpft und Fehler', () => {
        expect(directoryStatusView('a@schule.at', null).label).toBe('Abgleich offen');
        expect(directoryStatusView('a@schule.at', { notFound: true }).label).toBe('Nicht verknüpft');
        expect(directoryStatusView('', null).label).toBe('Nicht verknüpft');
        expect(directoryStatusView('a@schule.at', { error: 'timeout' }).label).toBe('Fehler');
        expect(directoryStatusView('a@schule.at', { diff: true }).label).toBe('Abweichung');
    });

    it('blendet undefined aus Namen aus', () => {
        expect(presentValue('Michael undefined')).toBe('Michael');
        expect(presentValue('undefined undefined')).toBe('');
        expect(presentText(undefined)).toBe('–');
        expect(presentText(null)).toBe('–');
        expect(presentText('')).toBe('–');
        expect(presentValue('Anna Null')).toBe('Anna Null');
    });

    it('Gruppenstatus ohne GUID in der Beschriftung', () => {
        const linked = groupLinkStatusView({ graphGroupId: 'abc-123', mailNickname: 'fach-d' });
        expect(linked.label).toBe('Verknüpft');
        expect(linked.title).toContain('abc-123');
        expect(groupLinkStatusView({ notFound: true }).label).toBe('Nicht verknüpft');
        expect(groupLinkStatusView(null).label).toBe('Abgleich offen');
    });
});
