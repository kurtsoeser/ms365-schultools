import { describe, expect, it } from 'vitest';
import {
    buildProvisionPlan,
    filterAntraege,
    isFreigeber,
    suggestMailNickname,
    validateAntrag
} from '../src/tools/teams-antrag/teams-antrag-logic.js';
import { ART, STATUS } from '../src/tools/teams-antrag/teams-antrag-model.js';

describe('teams-antrag-logic', () => {
    it('validiert Pflichtfelder und blockiert Diplom/VWA', () => {
        const bad = validateAntrag(
            { title: 'X', zweck: 'kurz', art: ART.AG, antragstellerEmail: 'a@b.at' },
            { requireApplicant: true }
        );
        expect(bad.ok).toBe(false);

        const diplom = validateAntrag(
            {
                title: 'DA',
                zweck: 'Ausführliche Begründung hier',
                art: ART.DIPLOM_VWA,
                antragstellerEmail: 'a@b.at'
            },
            { requireApplicant: true }
        );
        expect(diplom.ok).toBe(false);

        const ok = validateAntrag(
            {
                title: 'Robotik-AG',
                zweck: 'Wöchentliches Treffen der AG',
                art: ART.AG,
                antragstellerEmail: 'schueler@schule.at'
            },
            { requireApplicant: true }
        );
        expect(ok.ok).toBe(true);
    });

    it('schlägt Alias und Provision-Plan vor', () => {
        const nick = suggestMailNickname({
            title: 'Robotik AG',
            art: ART.AG,
            schuljahr: '2026/27'
        });
        expect(nick).toMatch(/^ag-2026-27-/);

        const plan = buildProvisionPlan({
            title: 'Robotik-AG',
            namensvorschlag: 'AG Robotik',
            art: ART.AG,
            schuljahr: '2026/27',
            antragstellerEmail: 'max@schule.at',
            besitzerMails: 'lehrer@schule.at',
            zweck: 'Test'
        });
        expect(plan.ok).toBe(true);
        expect(plan.ownerEmails).toContain('max@schule.at');
        expect(plan.ownerEmails).toContain('lehrer@schule.at');
    });

    it('erkennt Freigeber und filtert Warteschlange', () => {
        expect(isFreigeber('it@schule.at', ['it@schule.at'])).toBe(true);
        expect(isFreigeber('x@schule.at', ['it@schule.at'])).toBe(false);

        const rows = filterAntraege(
            [
                {
                    id: '1',
                    title: 'A',
                    status: STATUS.BEANTRAGT,
                    antragstellerEmail: 'a@b.at',
                    art: ART.AG,
                    updatedAt: '2026-01-02'
                },
                {
                    id: '2',
                    title: 'B',
                    status: STATUS.ANGELEGT,
                    antragstellerEmail: 'a@b.at',
                    art: ART.AG,
                    updatedAt: '2026-01-01'
                }
            ],
            { status: STATUS.BEANTRAGT }
        );
        expect(rows).toHaveLength(1);
        expect(rows[0].id).toBe('1');
    });
});
