import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { direktionUsersFromTenantStammdaten } from '../src/tools/freistellung-planer/freistellung-planer-direktion-users.js';
import { accountIsPlannerUserInList } from '../src/tools/freistellung-planer/freistellung-planer-direktion-users.js';

describe('Schulaktivitäten: Direktion aus Stammdaten', () => {
    const prevWindow = globalThis.window;

    beforeEach(() => {
        globalThis.window = {
            ms365TenantSettingsLoad: () => ({
                admin: [
                    {
                        role: 'Direktion',
                        name: 'Direktor',
                        email: 'Direktor@ms365.schule',
                        tier: 'schulleitung'
                    }
                ],
                adminRoles: [{ name: 'Direktion', tier: 'schulleitung' }]
            })
        };
    });

    afterEach(() => {
        globalThis.window = prevWindow;
        vi.restoreAllMocks();
    });

    it('findet Direktor@ms365.schule als Schulleitung', () => {
        const users = direktionUsersFromTenantStammdaten();
        expect(users.some((u) => u.mail === 'direktor@ms365.schule')).toBe(true);
        expect(accountIsPlannerUserInList('Direktor@ms365.schule', users)).toBe(true);
    });
});
