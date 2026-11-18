import { describe, expect, it } from 'vitest';
import {
    ablaufAdminGroupIds,
    ablaufKvGroupIds,
    ablaufLehrerGroupIds,
    isAblaufAdminFromStammdaten
} from '../src/shared/ablauf-admin-entra.js';

describe('ablauf-admin-entra', () => {
    it('liefert Group-ID-Listen als Arrays', () => {
        expect(Array.isArray(ablaufAdminGroupIds())).toBe(true);
        expect(Array.isArray(ablaufLehrerGroupIds())).toBe(true);
        expect(Array.isArray(ablaufKvGroupIds())).toBe(true);
    });

    it('Stammdaten-Admin ohne Window ist false', () => {
        expect(isAblaufAdminFromStammdaten('direktion@schule.at')).toBe(false);
        expect(isAblaufAdminFromStammdaten('')).toBe(false);
    });
});
