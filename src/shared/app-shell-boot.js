/**
 * App-Shell auf allen Seiten mit pin-gate (Sticky-Header + Werkzeug-Subheader).
 */
import { ensureAppHeaderChrome } from './app-global-header.js';
import { bootToolPageChrome } from './app-tool-chrome.js';
import { bootItWorkspaceShell } from './it-workspace-shell.js';
import { bootPortalWorkspaceShell } from './portal-workspace-shell.js';

function run() {
    ensureAppHeaderChrome();
    bootToolPageChrome();
    bootItWorkspaceShell();
    bootPortalWorkspaceShell();
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', run);
    } else {
        run();
    }
    window.addEventListener('ms365-menu-header-ready', function () {
        bootToolPageChrome();
        bootPortalWorkspaceShell();
    });
    window.addEventListener('ms365-auth-widget-ready', function () {
        bootToolPageChrome();
        bootPortalWorkspaceShell();
    });
}
