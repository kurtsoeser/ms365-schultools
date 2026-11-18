import { describe, expect, it } from 'vitest';
import { normalizeSettings } from '../src/tools/teams-antrag/teams-antrag-settings.js';

describe('teams-antrag-settings', () => {
    it('normalisiert Freigeber-Gruppe und Extra-Mails', () => {
        const s = normalizeSettings({
            freigeberGroupId: 'a7f3c91e-2b4d-4e8a-9c1f-6d5e8b0a4f27',
            freigeberGroupName: 'GruppenErsteller',
            freigeberMails: 'it@schule.at\nfoo',
            schuelerDuertenBeantragen: false
        });
        expect(s.freigeberGroupId).toBe('a7f3c91e-2b4d-4e8a-9c1f-6d5e8b0a4f27');
        expect(s.freigeberGroupName).toBe('GruppenErsteller');
        expect(s.freigeberMails).toEqual(['it@schule.at']);
        expect(s.schuelerDuertenBeantragen).toBe(false);
    });

    it('verwirft ungültige Gruppen-IDs', () => {
        const s = normalizeSettings({ freigeberGroupId: 'nicht-guid' });
        expect(s.freigeberGroupId).toBe('');
    });
});
