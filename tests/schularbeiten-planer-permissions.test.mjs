import { describe, it, expect } from 'vitest';
import {
    LIST_PERM_PROFILES,
    LIST_TITLE_TO_PROFILE,
    normalizePermissionsConfig,
    roleDefIdForLevel
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-permissions.js';
import { LIST_TITLES } from '../src/tools/schularbeiten-planer/schularbeiten-planer-schema.js';
import { SPO_ROLE } from '../src/shared/stammdaten-sharepoint-sync-logic.js';

describe('schularbeiten-planer-permissions', () => {
    it('normalisiert leere Gruppen (Picker)', () => {
        const c = normalizePermissionsConfig({});
        expect(c.groupAdmin).toBe('');
        expect(c.groupAdminId).toBe('');
        expect(c.skipPerms).toBe(false);
    });

    it('Schularbeiten-Liste: Lehrer contribute, Schüler read', () => {
        expect(LIST_PERM_PROFILES.schularbeiten.lehrer).toBe('contribute');
        expect(LIST_PERM_PROFILES.schularbeiten.schueler).toBe('read');
        expect(LIST_TITLE_TO_PROFILE[LIST_TITLES.schularbeiten]).toBe('schularbeiten');
    });

    it('Meta-Listen: Schüler ohne Recht', () => {
        expect(LIST_PERM_PROFILES.regelwerk.schueler).toBe(null);
        expect(LIST_PERM_PROFILES.fachMeta.schueler).toBe(null);
    });

    it('mappt SPO-Rollen-IDs', () => {
        expect(roleDefIdForLevel('read')).toBe(SPO_ROLE.read);
        expect(roleDefIdForLevel('contribute')).toBe(SPO_ROLE.contribute);
        expect(roleDefIdForLevel('design')).toBe(1073741828);
        expect(roleDefIdForLevel('edit')).toBe(SPO_ROLE.edit);
        expect(roleDefIdForLevel('fullControl')).toBe(SPO_ROLE.fullControl);
    });
});
