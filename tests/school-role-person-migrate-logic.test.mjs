import { describe, expect, it } from 'vitest';
import {
    extractRolePerson,
    pickItContactEmail,
    planRolePersonMigration,
    planRolePersonMigrationForSchools
} from '../src/shared/school-role-person-migrate-logic.js';

describe('school-role-person-migrate-logic', () => {
    it('trennt Name und E-Mail aus vermischten Feldern', () => {
        expect(extractRolePerson('Kurt Söser <kurt@schule.at>', '')).toEqual({
            name: 'Kurt Söser',
            email: 'kurt@schule.at'
        });
        expect(extractRolePerson('leitung@schule.at', '')).toEqual({
            name: '',
            email: 'leitung@schule.at'
        });
        expect(extractRolePerson('Direktion: max@bbs.at', '')).toEqual({
            name: '',
            email: 'max@bbs.at'
        });
        expect(extractRolePerson('Anna Muster', 'Anna@Schule.AT')).toEqual({
            name: 'Anna Muster',
            email: 'anna@schule.at'
        });
    });

    it('nimmt IT-Mail aus contactEmails', () => {
        expect(
            pickItContactEmail({
                contactEmails: ['IT@Schule.at'],
                contactEmail: 'alt@x.at'
            })
        ).toBe('it@schule.at');
    });

    it('plant Patch: Direktion trennen, IT-Mail aus Kontakt', () => {
        const plan = planRolePersonMigration({
            id: '1',
            schoolName: 'Demo',
            contactEmails: ['kurt@kurtsoeser.at'],
            extra: {
                Direktion: 'DEMO Direktor',
                ItM365Admin: 'Kurt Söser',
                Schuladministration: 'admin@schule.at'
            }
        });
        expect(plan.changed).toBe(true);
        expect(plan.extra.Direktion).toBeUndefined();
        expect(plan.extra.ItM365AdminEmail).toBe('kurt@kurtsoeser.at');
        expect(plan.extra.Schuladministration).toBe('');
        expect(plan.extra.SchuladministrationEmail).toBe('admin@schule.at');
    });

    it('überspringt Schulen ohne Änderungsbedarf', () => {
        const r = planRolePersonMigrationForSchools([
            {
                id: '1',
                schoolName: 'A',
                extra: {
                    Direktion: 'Max',
                    DirektionEmail: 'max@a.at',
                    ItM365Admin: 'IT',
                    ItM365AdminEmail: 'it@a.at'
                }
            },
            {
                id: '2',
                schoolName: 'B',
                contactEmails: ['it@b.at'],
                extra: { ItM365Admin: 'Admin' }
            }
        ]);
        expect(r.candidates).toBe(1);
        expect(r.plans[0].id).toBe('2');
        expect(r.plans[0].extra.ItM365AdminEmail).toBe('it@b.at');
    });
});
