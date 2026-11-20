import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
    normalizeAktPermissionsConfig,
    AKT_PERMS_STORAGE_KEY
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-permissions.js';
import {
    readAktPermissionsFromPickers,
    detectMountedAktPermWhich,
    persistAktPickersToStorage
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-permissions-ui.js';

const LEHRER_GID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const ADMIN_GID = '11111111-2222-3333-4444-555555555555';

function memoryStorage() {
    /** @type {Record<string, string>} */
    const data = {};
    return {
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(data, k) ? data[k] : null;
        },
        setItem(k, v) {
            data[k] = String(v);
        },
        removeItem(k) {
            delete data[k];
        },
        clear() {
            Object.keys(data).forEach((k) => delete data[k]);
        }
    };
}

describe('AKT-Berechtigungen: Teil-Speichern ohne Einreicher zu löschen', () => {
    const prevDoc = globalThis.document;
    const prevLs = globalThis.localStorage;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
        localStorage.setItem(
            AKT_PERMS_STORAGE_KEY,
            JSON.stringify(
                normalizeAktPermissionsConfig({
                    groupAdmin: 'Demo Schulleitung',
                    groupAdminId: ADMIN_GID,
                    groupLehrer: 'Demo Lehrkraefte',
                    groupLehrerId: LEHRER_GID,
                    adminUsers: [{ mail: 'direktor@ms365.schule', displayName: 'Direktor' }],
                    lehrerUsers: []
                })
            )
        );
    });

    afterEach(() => {
        globalThis.document = prevDoc;
        if (prevLs === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prevLs;
        vi.restoreAllMocks();
    });

    it('detectMountedAktPermWhich: nur Admin-Karte → admin', () => {
        globalThis.document = {
            getElementById(id) {
                if (id === 'spaktGroupAdmin' || id === 'spaktAdminUsers') return { value: '' };
                return null;
            }
        };
        expect(detectMountedAktPermWhich()).toBe('admin');
    });

    it('readAktPermissionsFromPickers(admin) schreibt groupLehrer nicht leer', () => {
        globalThis.document = {
            getElementById(id) {
                if (id === 'spaktGroupAdmin') return { value: 'Demo Schulleitung' };
                if (id === 'spaktGroupAdminId') return { value: ADMIN_GID };
                return null;
            }
        };
        const patch = readAktPermissionsFromPickers('admin');
        expect(patch.groupAdminId).toBe(ADMIN_GID);
        expect(patch).not.toHaveProperty('groupLehrerId');
        expect(patch).not.toHaveProperty('groupLehrer');
        expect(patch).not.toHaveProperty('lehrerUsers');
    });

    it('persistAktPickersToStorage(admin) behält gespeicherte Einreicher-Gruppe', () => {
        globalThis.document = {
            getElementById(id) {
                if (id === 'spaktGroupAdmin') return { value: 'Demo Schulleitung' };
                if (id === 'spaktGroupAdminId') return { value: ADMIN_GID };
                return null;
            }
        };
        const next = persistAktPickersToStorage(false, 'admin');
        expect(next.groupAdminId).toBe(ADMIN_GID);
        expect(next.groupLehrerId).toBe(LEHRER_GID);
        expect(next.groupLehrer).toBe('Demo Lehrkraefte');
        expect(next.adminUsers.some((u) => u.mail === 'direktor@ms365.schule')).toBe(true);
    });
});
