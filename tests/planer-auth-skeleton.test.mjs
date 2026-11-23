import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createInitialState as createFrState } from '../src/tools/freistellung-planer/freistellung-planer-state.js';
import { applyPlanerRoleFromEntra as applyFrRoleFromEntra } from '../src/tools/freistellung-planer/freistellung-planer-entra-role.js';
import { renderApp as renderFrApp } from '../src/tools/freistellung-planer/freistellung-planer-ui.js';
import { createInitialState as createSaState } from '../src/tools/schularbeiten-planer/schularbeiten-planer-state.js';
import { applyPlanerRoleFromEntra as applySaRoleFromEntra } from '../src/tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import { renderApp as renderSaApp } from '../src/tools/schularbeiten-planer/schularbeiten-planer-ui.js';

describe('Freistellung & Schularbeiten – Skeleton vor Konto-Read, Sperre erst nach E-Mail-Prüfung', () => {
    const origWindow = globalThis.window;
    const origLocalStorage = globalThis.localStorage;

    beforeEach(() => {
        const store = new Map();
        globalThis.localStorage = {
            getItem: (k) => store.get(k) ?? null,
            setItem: (k, v) => store.set(k, String(v)),
            removeItem: (k) => store.delete(k)
        };
        globalThis.window = {
            location: { search: '', pathname: '/tools/freistellung-planer.html' },
            ms365AuthIsLoggedIn: () => false,
            ms365AuthGetAccountInfo: () => null
        };
    });

    afterEach(() => {
        globalThis.window = origWindow;
        globalThis.localStorage = origLocalStorage;
    });

    it('Freistellung setzt ohne Anmeldung und ohne gelesene E-Mail niemals planerAccessDenied', async () => {
        const s1 = createFrState();
        await applyFrRoleFromEntra(s1, { demoRoleOverride: false });
        expect(s1.planerAccessDenied).toBe(false);
        expect(s1.kontoOffen).toBe(false);

        globalThis.window.ms365AuthIsLoggedIn = () => true;
        const s2 = createFrState();
        s2.accountEmail = '';
        await applyFrRoleFromEntra(s2, { demoRoleOverride: false });
        expect(s2.planerAccessDenied).toBe(false);
        expect(s2.kontoOffen).toBe(true);
    });

    it('Schularbeiten setzt ohne gelesene E-Mail niemals planerAccessDenied', async () => {
        globalThis.window.ms365AuthIsLoggedIn = () => true;
        const state = createSaState();
        state.accountEmail = '';
        await applySaRoleFromEntra(state, { demoRoleOverride: false });
        expect(state.planerAccessDenied).toBe(false);
        expect(state.kontoOffen).toBe(true);
    });

    it('Freistellung rendert während Konto-Read ein Skeleton und keine Sperrmeldung', () => {
        const state = createFrState();
        state.authRolePending = true;
        state.planerAccessDenied = true;
        state.roleHintPublic = 'Ihr Konto ist für den Freistellungs-Planer noch nicht freigeschaltet.';
        const root = { innerHTML: '' };
        renderFrApp(state, root);
        expect(root.innerHTML).toContain('class="fr-skel"');
        expect(root.innerHTML).toContain('aria-busy="true"');
        expect(root.innerHTML).not.toContain('noch nicht freigeschaltet');
        expect(root.innerHTML).not.toContain('Kein Zugriff auf den Freistellungs-Planer');
    });

    it('Freistellung zeigt nach Auth-Boot ohne Konto eine ruhige Anmeldekarte statt Sperrbox', () => {
        const state = createFrState();
        state.authRolePending = false;
        state.authSettled = true;
        state.accountEmail = '';
        const root = { innerHTML: '' };
        renderFrApp(state, root);
        expect(root.innerHTML).toContain('Anmeldung erforderlich');
        expect(root.innerHTML).toContain('data-fr-auth-login');
        expect(root.innerHTML).not.toContain('noch nicht freigeschaltet');
        expect(root.innerHTML).not.toContain('Kein Zugriff auf den Freistellungs-Planer');
    });

    it('Freistellung zeigt bei echter Sperre das geprüfte Konto und Konto-Wechsel an', () => {
        const state = createFrState();
        state.authRolePending = false;
        state.authSettled = true;
        state.accountEmail = 'gast@privat.at';
        state.planerAccessDenied = true;
        state.roleHintPublic =
            'Ihr Konto ist für den Freistellungs-Planer noch nicht freigeschaltet. Bitte wenden Sie sich an Klassenvorstand oder Sekretariat.';
        const root = { innerHTML: '' };
        renderFrApp(state, root);
        expect(root.innerHTML).toContain('Kein Zugriff auf den Freistellungs-Planer');
        expect(root.innerHTML).toContain('gast@privat.at');
        expect(root.innerHTML).toContain('data-fr-auth-switch');
    });

    it('Schularbeiten rendert während Konto-Read ein Skeleton und keine Sperrmeldung', () => {
        const state = createSaState();
        state.authRolePending = true;
        state.planerAccessDenied = true;
        const root = { innerHTML: '' };
        renderSaApp(state, root);
        expect(root.innerHTML).toContain('class="sa-skel"');
        expect(root.innerHTML).toContain('aria-busy="true"');
        expect(root.innerHTML).not.toContain('Kein Zugriff auf den Schularbeiten-Planer');
    });

    it('Schularbeiten zeigt nach Auth-Boot ohne Konto eine ruhige Anmeldekarte und bei echter Sperre das geprüfte Konto ohne doppeltes Banner', () => {
        const signedOut = createSaState();
        signedOut.authRolePending = false;
        signedOut.authSettled = true;
        signedOut.accountEmail = '';
        const rootOut = { innerHTML: '' };
        renderSaApp(signedOut, rootOut);
        expect(rootOut.innerHTML).toContain('Anmeldung erforderlich');
        expect(rootOut.innerHTML).toContain('data-sa-auth-login');
        expect(rootOut.innerHTML).not.toContain('Kein Zugriff auf den Schularbeiten-Planer');

        const denied = createSaState();
        denied.authRolePending = false;
        denied.authSettled = true;
        denied.accountEmail = 'gast@privat.at';
        denied.planerAccessDenied = true;
        denied.roleHint = 'Keine Planer-Berechtigung: Duplikat-Test.';
        const rootDenied = { innerHTML: '' };
        renderSaApp(denied, rootDenied);
        expect(rootDenied.innerHTML).toContain('Kein Zugriff auf den Schularbeiten-Planer');
        expect(rootDenied.innerHTML).toContain('gast@privat.at');
        expect(rootDenied.innerHTML).toContain('data-sa-auth-switch');
        expect(rootDenied.innerHTML).not.toContain('Duplikat-Test');
    });
});
