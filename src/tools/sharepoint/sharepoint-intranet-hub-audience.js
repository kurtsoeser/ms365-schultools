/**
 * Intranet-Hub ist das Einrichten für die Schul-IT.
 * Lehrkräfte und Schüler/innen landen auf der hinterlegten SharePoint-Seite.
 */
import { getDashboardPersonas } from '../../shared/dashboard-persona-session.js';
import { intranetHubAudienceDecision } from '../../shared/school-intranet-open.js';

function readSetup() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            return window.ms365AppDataV2.getSetup();
        }
    } catch {
        /* ignore */
    }
    return null;
}

function setAudience(state) {
    document.documentElement.setAttribute('data-ih-audience', state);
}

function statusEl() {
    return document.getElementById('ihAudienceStatus');
}

function showMissing() {
    setAudience('missing');
    const status = statusEl();
    if (!status) return;
    status.replaceChildren();
    const text = document.createElement('span');
    text.textContent =
        'Die Intranet-Seite ist noch nicht hinterlegt. Die Schul-IT speichert die Adresse im Intranet-Hub. Danach öffnet dieser Eintrag die SharePoint-Seite.';
    const back = document.createElement('a');
    back.href = '../index.html';
    back.textContent = 'Zurück zum Portal';
    status.append(text, document.createTextNode(' '), back);
}

async function applyAudience() {
    let personas = null;
    try {
        personas = await getDashboardPersonas();
    } catch {
        return;
    }
    const decision = intranetHubAudienceDecision(personas, readSetup());
    if (decision.action === 'wait') return;
    if (decision.action === 'stay') {
        setAudience('it');
        return;
    }
    if (decision.action === 'open' && decision.url) {
        setAudience('open');
        const status = statusEl();
        if (status) status.textContent = 'Die Intranet-Seite wird geöffnet …';
        window.location.replace(decision.url);
        return;
    }
    showMissing();
}

applyAudience();
window.addEventListener('ms365-auth-state-changed', () => {
    applyAudience();
});
window.addEventListener('ms365-dashboard-persona-ready', () => {
    applyAudience();
});
