import {
    loadMergedReleaseNotes,
    getLastSeenAt,
    getNewReleaseNotes,
    filterNotesForViewer
} from './release-notes-store.js';

let viewer = 'enduser';

function modeIsIt(mode) {
    return mode === 'full';
}

function currentMode() {
    try {
        const api = window.ms365DashboardAudience;
        if (api && typeof api.getLayoutMode === 'function') return String(api.getLayoutMode() || '');
    } catch {
        /* ignore */
    }
    if (document.documentElement.hasAttribute('data-dash-it-shell')) return 'full';
    return 'preview-lehrer';
}

function placeStrip(isIt) {
    const strip = document.getElementById('dashNewsStrip');
    const tasks = document.getElementById('dashboard-tasks');
    const stack = document.getElementById('dashAppOnlyStack');
    if (!strip) return;
    if (isIt && tasks) tasks.insertBefore(strip, tasks.firstChild);
    else if (stack && stack.parentNode) stack.parentNode.insertBefore(strip, stack);
    const kicker = strip.querySelector('.dash-news-strip__kicker');
    if (kicker) kicker.textContent = isIt ? 'Schul-IT' : 'Neuigkeiten';
    document.querySelectorAll('a[href^="neuigkeiten.html"]').forEach((link) => {
        link.setAttribute('href', isIt ? 'neuigkeiten.html?bereich=it' : 'neuigkeiten.html');
    });
}

async function paint() {
    const strip = document.getElementById('dashNewsStrip');
    const text = document.getElementById('dashNewsStripText');
    if (!strip || !text) return;
    let notes = [];
    try {
        notes = await loadMergedReleaseNotes(localStorage);
    } catch {
        notes = [];
    }
    const visible = filterNotesForViewer(notes, viewer);
    const fresh = getNewReleaseNotes({
        notes: visible,
        lastSeenAtIso: getLastSeenAt(localStorage, viewer)
    });
    strip.classList.toggle('dash-news-strip--fresh', fresh.length > 0);
    const link = strip.querySelector('.dash-news-strip__link');
    if (link) link.textContent = fresh.length ? 'Neuigkeiten lesen' : 'Lesen';
    if (fresh.length) {
        const first = fresh[0];
        const more = fresh.length > 1 ? ' · ' + fresh.length + ' neue Einträge' : '';
        text.textContent = (first.title || 'Neuer Eintrag') + more;
        return;
    }
    text.textContent = visible.length
        ? viewer === 'it'
            ? 'Technische Hinweise für die Schul-IT.'
            : 'Hinweise zum Schulalltag.'
        : 'Hier erscheinen Hinweise, sobald welche veröffentlicht sind.';
}

function applyMode(mode) {
    const isIt = modeIsIt(mode);
    viewer = isIt ? 'it' : 'enduser';
    placeStrip(isIt);
    paint();
}

function boot() {
    applyMode(currentMode());
    window.addEventListener('ms365-dash-app-only-changed', function (ev) {
        const mode = ev && ev.detail ? ev.detail.layoutMode : '';
        applyMode(mode || currentMode());
    });
    window.addEventListener('ms365-it-shell-changed', function (ev) {
        const mode = ev && ev.detail ? ev.detail.mode : '';
        applyMode(mode || currentMode());
    });
    window.addEventListener('ms365-dashboard-persona-ready', function () {
        applyMode(currentMode());
    });
    window.addEventListener('pageshow', function () {
        applyMode(currentMode());
    });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
