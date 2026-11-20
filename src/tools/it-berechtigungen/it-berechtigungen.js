/**
 * Lesende Übersicht der Schul-App-Berechtigungen in diesem Browser.
 */
import { readAppRoleInventory } from '../../shared/app-role-inventory.js';
import { announceAppRole } from '../../shared/app-role-chrome.js';

function esc(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function render() {
    const root = document.getElementById('itPermApp');
    if (!root) return;
    const data = readAppRoleInventory(localStorage);
    const principals = data.principals.length
        ? data.principals.map((line) => `<li>${esc(line)}</li>`).join('')
        : '<li>In diesem Browser sind noch keine Sammelgruppen gespeichert.</li>';
    const rows = data.apps
        .map((app) => {
            const body = app.lines.length
                ? app.lines.map((line) => `<div>${esc(line)}</div>`).join('')
                : '<span class="muted">Noch nichts in diesem Browser.</span>';
            const note = app.note ? `<div class="muted">${esc(app.note)}</div>` : '';
            return `<tr>
                <td>${esc(app.family)}</td>
                <td><a href="${esc(app.href)}">${esc(app.title)}</a></td>
                <td>${body}${note}</td>
            </tr>`;
        })
        .join('');
    root.innerHTML = `
        <section class="fr-panel" style="margin-bottom:1rem">
            <h2>Schul-Prinzipale</h2>
            <p>Lehrkräfte- und Schüler-Sammelgruppe aus der lokalen Einrichtung. Schulleitung und Verwaltung stehen in den Zeilen der Apps, sobald sie dort gesetzt sind.</p>
            <ul>${principals}</ul>
        </section>
        <section class="fr-panel">
            <h2>Schul-Apps</h2>
            <p>Das ist die Einrichtung, nicht die gerade berechnete App-Rolle. Die Rolle steht im Kopf der geöffneten App: App · Rolle · Quelle.</p>
            <div style="overflow:auto">
                <table class="tm-table">
                    <thead><tr><th>Familie</th><th>App</th><th>Eingerichtet</th></tr></thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
        </section>`;
    let email = '';
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const account = window.ms365AuthGetAccountInfo();
            email = String((account && (account.username || account.upn)) || '');
        }
    } catch {
        /* ignore */
    }
    announceAppRole({ appId: 'it-berechtigungen', role: 'it', source: 'school-it', email });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', render);
} else {
    render();
}
window.addEventListener('ms365-auth-state-changed', render);
