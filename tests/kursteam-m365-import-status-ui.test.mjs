import { describe, it, expect } from 'vitest';
import {
    renderKursteamM365ImportStatusHtml,
    summarizeKursteamM365ImportHint
} from '../src/shared/kursteam-m365-import-status-ui.js';

describe('kursteam-m365-import-status-ui', () => {
    it('zeigt leeren Hinweis ohne Report', () => {
        const html = renderKursteamM365ImportStatusHtml(null, { emptyHint: 'Bitte laden' });
        expect(html).toContain('Bitte laden');
    });

    it('rendert Counts und Passend-Sektion', () => {
        const html = renderKursteamM365ImportStatusHtml(
            {
                counts: {
                    planned: 2,
                    m365: 2,
                    matched: 1,
                    missingInM365: 1,
                    onlyInM365: 0,
                    mailConflict: 0
                },
                matched: [{ klasse: '1A', fach: 'D', lehrerCode: 'MAY', matchBy: 'gruppenmail' }],
                missingInM365: [{ klasse: '1B', fach: 'M', lehrerCode: 'FOO', gruppenmail: '' }],
                onlyInM365: [],
                mailConflict: []
            },
            { planLabel: 'Katalog' }
        );
        expect(html).toContain('Katalog');
        expect(html).toContain('1 passend');
        expect(html).toContain('1 fehlen in M365');
        expect(html).toContain('Passend verknüpft');
        expect(html).toContain('Fehlt in M365');
    });

    it('summarize liefert Warnung bei fehlenden Teams', () => {
        expect(
            summarizeKursteamM365ImportHint({
                counts: { matched: 5, missingInM365: 2, onlyInM365: 0 }
            })
        ).toEqual({ text: '2 fehlen', tone: 'warn' });
        expect(
            summarizeKursteamM365ImportHint({
                counts: { matched: 5, missingInM365: 0, onlyInM365: 0 }
            })
        ).toEqual({ text: '5 ok', tone: 'ok' });
    });
});
