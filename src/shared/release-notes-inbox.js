import {
    loadMergedReleaseNotes,
    getLastSeenAt,
    getNewReleaseNotes,
    filterNotesForViewer
} from './release-notes-store.js';

let viewer = 'enduser';

function viewerFromPersonas(personas) {
    if (personas && (personas.isIt || personas.designatedSchoolIt || personas.globalAdmin)) return 'it';
    return 'enduser';
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
        lastSeenAtIso: getLastSeenAt(localStorage)
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
        ? 'Hinweise zum Schulalltag und zur Einrichtung.'
        : 'Hier erscheinen Hinweise, sobald welche veröffentlicht sind.';
}

function boot() {
    paint();
    window.addEventListener('ms365-dashboard-persona-ready', function (ev) {
        const personas = ev && ev.detail ? ev.detail.personas : null;
        viewer = viewerFromPersonas(personas);
        paint();
    });
    window.addEventListener('pageshow', function () {
        paint();
    });
    try {
        if (window.__ms365DashboardPersonasLast) {
            viewer = viewerFromPersonas(window.__ms365DashboardPersonasLast);
            paint();
        }
    } catch {
        /* ignore */
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
