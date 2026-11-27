import {
    loadMergedReleaseNotes,
    getLastSeenAt,
    getNewReleaseNotes,
    setLastSeenAt,
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

function setStripVisible(strip, visible) {
    strip.hidden = !visible;
    strip.style.display = visible ? '' : 'none';
    if (!visible) strip.classList.remove('dash-news-strip--fresh');
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
    setStripVisible(strip, fresh.length > 0);
    if (!fresh.length) return;
    strip.classList.add('dash-news-strip--fresh');
    const link = strip.querySelector('.dash-news-strip__link');
    const first = fresh[0];
    const more = fresh.length > 1 ? ' · ' + fresh.length + ' neue Einträge' : '';
    text.textContent = (first.title || 'Neuer Eintrag') + more;
    if (link) {
        link.textContent = 'Neuigkeiten lesen';
        link.onclick = function () {
            const latest = fresh.reduce(function (acc, note) {
                const t = new Date(note.at).getTime();
                if (!Number.isFinite(t)) return acc;
                if (!acc || t > new Date(acc).getTime()) return note.at;
                return acc;
            }, '');
            if (latest) setLastSeenAt(latest, localStorage, viewer);
            setStripVisible(strip, false);
        };
    }
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
