import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { effectiveLfrDirektionUsersForListAccess } from '../src/tools/lehrer-freistellung-planer/lfr-list-permissions.js';

describe('effectiveLfrDirektionUsersForListAccess', () => {
    const prevWin = globalThis.window;

    beforeEach(() => {
        globalThis.window = {
            ms365TenantSettingsLoad: () => ({
                admin: [
                    {
                        role: 'Direktion',
                        name: 'Direktor',
                        email: 'direktor@ms365.schule',
                        tier: 'schulleitung'
                    }
                ],
                adminRoles: [{ name: 'Direktion', tier: 'schulleitung' }]
            })
        };
    });

    afterEach(() => {
        globalThis.window = prevWin;
        vi.restoreAllMocks();
    });

    it('zieht Schulleitung aus Stammdaten und Approver zusammen', () => {
        const users = effectiveLfrDirektionUsersForListAccess({
            direktionUsers: [{ mail: 'verwaltung@schule.at', displayName: 'Verw.' }],
            emailDirektion: 'direktor@ms365.schule'
        });
        const mails = users.map((u) => u.mail).sort();
        expect(mails).toContain('verwaltung@schule.at');
        expect(mails).toContain('direktor@ms365.schule');
    });
});
