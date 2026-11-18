import { describe, expect, it } from 'vitest';
import { classifyElectionError, softElectionMessage } from '../src/tools/schulsprecherwahl/schulsprecherwahl-errors.js';
import {
    electionPhaseSteps,
    electionQuorumInfo,
    nextAdminAction,
    wantsElectionDemo
} from '../src/tools/schulsprecherwahl/schulsprecherwahl-phase.js';
import { createElection, ELECTION_STATUS } from '../src/tools/schulsprecherwahl/schulsprecherwahl-model.js';
import { ELECTION_KIND_SGA_LEHRER } from '../src/tools/schulsprecherwahl/schulsprecherwahl-kinds.js';
import {
    upsertElection,
    publishResults,
    getElection
} from '../src/tools/schulsprecherwahl/schulsprecherwahl-store.js';

describe('schulsprecherwahl soft errors', () => {
    it('klassifiziert 401 als auth_required und keepLocal', () => {
        const err = Object.assign(new Error('Unauthorized'), { status: 401 });
        const soft = classifyElectionError(err);
        expect(soft.code).toBe('auth_required');
        expect(soft.keepLocal).toBe(true);
        expect(softElectionMessage(err)).toMatch(/anmelden/i);
    });

    it('klassifiziert 403 als forbidden ohne Wipe-Signal', () => {
        const err = Object.assign(new Error('Keine Berechtigung'), { status: 403, code: 'forbidden' });
        const soft = classifyElectionError(err);
        expect(soft.code).toBe('forbidden');
        expect(soft.keepLocal).toBe(true);
    });

    it('erkennt Netzwerkfehler', () => {
        const soft = classifyElectionError(new Error('Failed to fetch'));
        expect(soft.code).toBe('network');
        expect(soft.keepLocal).toBe(true);
    });
});

describe('schulsprecherwahl demo gate', () => {
    it('wantsElectionDemo nur bei ?demo=1', () => {
        expect(wantsElectionDemo('?demo=1')).toBe(true);
        expect(wantsElectionDemo('?foo=1')).toBe(false);
        expect(wantsElectionDemo('')).toBe(false);
    });
});

describe('schulsprecherwahl phasen / next action', () => {
    it('zeigt Entwurf → Offen → … für Schulsprecher', () => {
        const e = createElection({ status: ELECTION_STATUS.draft, candidates: [] });
        const steps = electionPhaseSteps(e);
        expect(steps[0].state).toBe('current');
        expect(steps.map((s) => s.id)).toContain('runoff');
        const action = nextAdminAction(e);
        expect(action && action.code).toBe('candidates');
    });

    it('fordert Eröffnen wenn genug Kandidaturen inkl. CTA', () => {
        const e = createElection({
            status: ELECTION_STATUS.draft,
            candidates: [
                { id: 'a', name: 'A', className: '1A', email: 'a@x.at' },
                { id: 'b', name: 'B', className: '1A', email: 'b@x.at' },
                { id: 'c', name: 'C', className: '1A', email: 'c@x.at' }
            ]
        });
        const action = nextAdminAction(e);
        expect(action?.code).toBe('open');
        expect(action?.ctaAction).toBe('open');
    });

    it('SGA Lehrer ohne Stichwahl-Phase', () => {
        const e = createElection({
            electionKind: ELECTION_KIND_SGA_LEHRER,
            status: ELECTION_STATUS.open
        });
        const steps = electionPhaseSteps(e);
        expect(steps.map((s) => s.id)).not.toContain('runoff');
        expect(steps.find((s) => s.state === 'current')?.id).toBe('open');
        expect(nextAdminAction(e)?.code).toBe('count');
        expect(nextAdminAction(e)?.ctaAction).toBe('count');
    });

    it('Stichwahl-Aktion nur bei closed Schulsprecher', () => {
        const e = createElection({ status: ELECTION_STATUS.closed });
        expect(nextAdminAction(e)?.code).toBe('runoff');
        expect(nextAdminAction(e)?.ctaAction).toBe('runoff');
        const sga = createElection({
            electionKind: ELECTION_KIND_SGA_LEHRER,
            status: ELECTION_STATUS.closed
        });
        expect(nextAdminAction(sga)?.ctaAction).toBe('publish');
    });

    it('Kundmachung als eigener Schritt nach Finalisierung', () => {
        const e = createElection({
            status: ELECTION_STATUS.finalized,
            publishedResults: false,
            offices: { schulsprecherId: 'a', decidedAfter: 'main' }
        });
        const steps = electionPhaseSteps(e);
        expect(steps.find((s) => s.state === 'current')?.id).toBe('publish');
        expect(nextAdminAction(e)?.ctaAction).toBe('publish');
        const published = { ...e, publishedResults: true };
        expect(electionPhaseSteps(published).every((s) => s.state === 'done')).toBe(true);
        expect(nextAdminAction(published)?.ctaAction).toBe('copy_results');
    });
});

describe('schulsprecherwahl quorum', () => {
    it('meldet Beteiligung erreicht ab 50 %', () => {
        const e = createElection({
            status: ELECTION_STATUS.finalized,
            eligibleCount: 100,
            rounds: [
                {
                    id: 'main',
                    kind: 'main',
                    status: 'counted',
                    stats: { cast: 62, valid: 60, invalid: 2, eligible: 100 }
                }
            ]
        });
        const info = electionQuorumInfo(e);
        expect(info?.reached).toBe(true);
        expect(info?.tone).toBe('ok');
    });

    it('warnt unter 50 % ohne Auszählung zu sperren', () => {
        const e = createElection({
            status: ELECTION_STATUS.finalized,
            eligibleCount: 100,
            rounds: [
                {
                    id: 'main',
                    kind: 'main',
                    status: 'counted',
                    stats: { cast: 40, valid: 40, invalid: 0, eligible: 100 }
                }
            ]
        });
        const info = electionQuorumInfo(e);
        expect(info?.reached).toBe(false);
        expect(info?.tone).toBe('warn');
    });
});

describe('schulsprecherwahl kundmachung store', () => {
    /** @type {Map<string, string>} */
    let mem;

    it('setzt publishedResults erst per publishResults', () => {
        mem = new Map();
        globalThis.localStorage = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => {
                mem.set(k, String(v));
            },
            removeItem: (k) => {
                mem.delete(k);
            }
        };

        const saved = upsertElection(
            createElection({
                id: 'wahl_publish_test',
                status: ELECTION_STATUS.finalized,
                publishedResults: false,
                candidates: [
                    { id: 'a', name: 'A', className: '1A', email: 'a@x.at' },
                    { id: 'b', name: 'B', className: '1A', email: 'b@x.at' },
                    { id: 'c', name: 'C', className: '1A', email: 'c@x.at' }
                ]
            })
        );
        expect(saved.publishedResults).toBe(false);
        const published = publishResults(saved.id);
        expect(published.publishedResults).toBe(true);
        expect(getElection(saved.id)?.publishedResults).toBe(true);
    });
});
