import { describe, it, expect } from 'vitest';
import { buildMergePlan, applyLocalMerge, findTeamForClass } from '../src/tools/klassen-merge/klassen-merge-logic.js';

describe('klassen-merge-logic', () => {
    const survivor = { code: '4HMA', name: '4HMA', year: '2027' };
    const source = { code: '4HMB', name: '4HMB', year: '2027' };
    const teams = [
        { classCode: '4HMA', graphGroupId: 'g-surv', stableMailNickname: 'jg2027-hma' },
        { classCode: '4HMB', graphGroupId: 'g-src', stableMailNickname: 'jg2027-hmb' }
    ];

    it('Happy-Path: Plan ok mit Graph-Match', () => {
        const plan = buildMergePlan({
            survivor,
            sources: [source],
            students: [
                { klasse: '4HMA', email: 'a@s.at' },
                { klasse: '4HMB', email: 'b@s.at' }
            ],
            classTeams: teams,
            newCode: '4HM'
        });
        expect(plan.ok).toBe(true);
        expect(plan.memberEmails.sort()).toEqual(['a@s.at', 'b@s.at']);
        expect(plan.steps.some((s) => s.id === 'graph-members')).toBe(true);
    });

    it('ohne Graph-Match: ok=false, localOnlyAllowed', () => {
        const plan = buildMergePlan({
            survivor,
            sources: [source],
            students: [],
            classTeams: [],
            newCode: '4HM'
        });
        expect(plan.ok).toBe(false);
        expect(plan.requiresGraphConfirm).toBe(true);
        expect(plan.localOnlyAllowed).toBe(true);
        expect(plan.error).toMatch(/Graph-Match/i);
    });

    it('applyLocalMerge schreibt Schüler um und entfernt Quell-Klasse', () => {
        const plan = buildMergePlan({
            survivor,
            sources: [source],
            students: [
                { klasse: '4HMA', email: 'a@s.at' },
                { klasse: '4HMB', email: 'b@s.at' }
            ],
            classTeams: teams,
            newCode: '4HM',
            newName: '4HM'
        });
        expect(plan.ok).toBe(true);
        const applied = applyLocalMerge(
            plan,
            { classes: [survivor, source], students: plan.memberEmails.map((e) => ({ klasse: e.startsWith('a') ? '4HMA' : '4HMB', email: e })) },
            teams
        );
        expect(applied.error).toBe('');
        expect(applied.settings.classes.map((c) => c.code)).toEqual(['4HM']);
        expect(applied.settings.students.every((s) => s.klasse === '4HM')).toBe(true);
    });

    it('findTeamForClass per Code und Nickname', () => {
        expect(findTeamForClass(survivor, teams).graphGroupId).toBe('g-surv');
        expect(
            findTeamForClass({ stableMailNickname: 'jg2027-hmb' }, teams).graphGroupId
        ).toBe('g-src');
    });
});
