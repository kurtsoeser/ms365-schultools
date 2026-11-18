import { describe, it, expect } from 'vitest';
import {
    applyManuelleZuteilung,
    assertWzOwnership,
    countBelegung,
    countUeberbuchung,
    ergebnisListenProAngebot,
    isAnmeldefensterOffen,
    isWzAccessDeniedError,
    kapazitaetStatus,
    resolveWzActionHint,
    resolveWzAdminActor,
    resolveWzOwnership,
    resolveWzPhase,
    validateAnmeldung,
    validateAngebot
} from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-logic.js';

const rundeOffen = {
    schuljahr: '2026/27',
    titel: 'Demo',
    anmeldeVon: '2026-09-01',
    anmeldeBis: '2026-09-30',
    status: 'offen'
};

describe('wahlfach-zuteilung-logic', () => {
    it('erkennt offenes und geschlossenes Anmeldefenster', () => {
        expect(isAnmeldefensterOffen(rundeOffen, new Date('2026-09-15T12:00:00'))).toBe(true);
        expect(isAnmeldefensterOffen(rundeOffen, new Date('2026-08-31T12:00:00'))).toBe(false);
        expect(
            isAnmeldefensterOffen({ ...rundeOffen, status: 'geschlossen' }, new Date('2026-09-15'))
        ).toBe(false);
    });

    it('prüft Kapazität und warnt bei Ausschöpfung', () => {
        const angebot = { id: 'a1', kapazitaetMax: 2, kapazitaetMin: 1 };
        const frei = kapazitaetStatus(angebot, 1);
        expect(frei.frei).toBe(1);
        expect(frei.voll).toBe(false);
        const voll = kapazitaetStatus(angebot, 2);
        expect(voll.voll).toBe(true);
        expect(voll.warnung).toMatch(/ausgeschöpft/i);
        const ueber = kapazitaetStatus(angebot, 3);
        expect(ueber.ueberbuchung).toBe(1);
    });

    it('wendet manuelle Zuteilung an und aktualisiert Status', () => {
        const angebote = [{ id: 'ang-chor', bezeichnung: 'Chor', kapazitaetMax: 10, kapazitaetMin: 5, aktiv: true }];
        const state = {
            anmeldungen: [
                {
                    id: 'anm-1',
                    schuelerName: 'Ada',
                    klasse: '3AK',
                    wunsch1Id: 'ang-chor',
                    status: 'eingereicht'
                }
            ],
            zuteilungen: []
        };
        const res = applyManuelleZuteilung(state, 'anm-1', 'ang-chor', { zugeteiltVon: 'Test' }, angebote);
        expect(res.ok).toBe(true);
        expect(res.anmeldungen[0].status).toBe('zugeordnet');
        expect(res.zuteilungen).toHaveLength(1);
        expect(res.zuteilungen[0].angebotId).toBe('ang-chor');
        expect(countBelegung('ang-chor', res.zuteilungen)).toBe(1);

        const listen = ergebnisListenProAngebot(angebote, res.anmeldungen, res.zuteilungen);
        expect(listen[0].teilnehmende).toHaveLength(1);
        expect(listen[0].teilnehmende[0].anmeldung.schuelerName).toBe('Ada');
    });

    it('blockiert Überbuchung und zählt Überschreitungen', () => {
        const angebote = [
            { id: 'ang-rob', bezeichnung: 'Robotik', kapazitaetMax: 1, kapazitaetMin: 1, aktiv: true }
        ];
        const state = {
            anmeldungen: [
                { id: 'a1', schuelerName: 'A', klasse: '3A', wunsch1Id: 'ang-rob', status: 'zugeordnet' },
                { id: 'a2', schuelerName: 'B', klasse: '3B', wunsch1Id: 'ang-rob', status: 'eingereicht' }
            ],
            zuteilungen: [{ anmeldungId: 'a1', angebotId: 'ang-rob' }]
        };
        const blocked = applyManuelleZuteilung(state, 'a2', 'ang-rob', {}, angebote);
        expect(blocked.ok).toBe(false);
        expect(blocked.errors.some((e) => /Kapazität/i.test(e))).toBe(true);

        const forcedZ = [
            { anmeldungId: 'a1', angebotId: 'ang-rob' },
            { anmeldungId: 'a2', angebotId: 'ang-rob' }
        ];
        expect(countUeberbuchung(angebote, forcedZ)).toBe(1);
    });

    it('validiert Angebot und Anmeldung', () => {
        expect(validateAngebot({ bezeichnung: 'X', kapazitaetMax: 0 }).ok).toBe(false);
        expect(validateAngebot({ bezeichnung: 'Chor', kapazitaetMax: 20, kapazitaetMin: 8 }).ok).toBe(
            true
        );
        const angebote = [{ id: 'a1', bezeichnung: 'Chor', aktiv: true, kapazitaetMax: 10 }];
        expect(
            validateAnmeldung(
                { schuelerName: 'Ada', klasse: '3AK', wunsch1Id: 'a1', wunsch2Id: 'a1' },
                angebote
            ).ok
        ).toBe(false);
        expect(
            validateAnmeldung({ schuelerName: 'Ada', klasse: '3AK', wunsch1Id: 'a1' }, angebote).ok
        ).toBe(true);
    });

    it('löst Phasen und Action-Hints ohne Matching/Genehmigung', () => {
        const offen = resolveWzPhase(rundeOffen, new Date('2026-09-15'));
        expect(offen.phase).toBe('anmeldung');
        expect(offen.anmeldeOffen).toBe(true);

        const zuteil = resolveWzPhase(
            { ...rundeOffen, status: 'zuteilung' },
            new Date('2026-09-15')
        );
        expect(zuteil.phase).toBe('zuteilung');

        const hint = resolveWzActionHint({
            runde: { ...rundeOffen, status: 'geschlossen' },
            angebote: [{ id: 'a1', kapazitaetMax: 10, kapazitaetMin: 1 }],
            anmeldungen: [
                {
                    id: 'anm-1',
                    schuelerName: 'Ada',
                    klasse: '3AK',
                    wunsch1Id: 'a1',
                    status: 'eingereicht'
                }
            ],
            zuteilungen: []
        });
        expect(hint).toBeTruthy();
        expect(hint.ctaView).toBe('zuteilung');
        expect(hint.text).toMatch(/manuell/i);
    });

    it('setzt Ownership und erkennt soft 403', () => {
        const own = resolveWzOwnership(
            { schuelerName: '', klasse: '' },
            {
                accountEmail: 'ada@schule.at',
                accountName: 'Ada Lovelace',
                students: [
                    {
                        email: 'ada@schule.at',
                        vorname: 'Ada',
                        nachname: 'Lovelace',
                        klasse: '3AK'
                    }
                ]
            }
        );
        expect(own.schuelerEmail).toBe('ada@schule.at');
        expect(own.klasse).toBe('3AK');
        expect(own.schuelerName).toMatch(/Ada/);
        expect(() => assertWzOwnership(own)).not.toThrow();
        expect(() => assertWzOwnership({ schuelerEmail: '' })).toThrow(/E-Mail|Konto/i);
        expect(isWzAccessDeniedError(new Error('403 Forbidden'))).toBe(true);
        expect(resolveWzAdminActor({ accountEmail: 'slg@schule.at', accountName: 'SLG' })).toMatch(
            /slg@schule.at/
        );
    });
});
