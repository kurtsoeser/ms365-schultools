/**
 * App-Shell auf allen Seiten mit pin-gate (Sticky-Header + Werkzeug-Subheader).
 */
import { ensureAppHeaderChrome } from './app-global-header.js';
import { bootToolPageChrome } from './app-tool-chrome.js';
import { bootItWorkspaceShell } from './it-workspace-shell.js';
import { bootPortalWorkspaceShell } from './portal-workspace-shell.js';
import { refreshWahlfachMenuRole } from '../tools/wahlfach-zuteilung/wahlfach-portal-menu.js';

function refreshWahlfachRail() {
    void refreshWahlfachMenuRole().then((changed) => {
        if (!changed) return;
        try {
            window.dispatchEvent(new Event('ms365-dashboard-tool-access-changed'));
        } catch {
            /* ignore */
        }
        const rail = document.getElementById('portalWorkspaceRail');
        if (rail) rail.setAttribute('data-portal-view', '');
        bootPortalWorkspaceShell();
    });
}

function run() {
    ensureAppHeaderChrome();
    bootToolPageChrome();
    bootItWorkspaceShell();
    bootPortalWorkspaceShell();
    refreshWahlfachRail();
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', run);
    } else {
        run();
    }
    window.addEventListener('ms365-menu-header-ready', function () {
        bootToolPageChrome();
        bootItWorkspaceShell();
        bootPortalWorkspaceShell();
    });
    window.addEventListener('ms365-auth-widget-ready', function () {
        bootToolPageChrome();
        bootItWorkspaceShell();
        bootPortalWorkspaceShell();
        refreshWahlfachRail();
    });
}
