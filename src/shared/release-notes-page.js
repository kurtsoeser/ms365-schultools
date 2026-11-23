import {
    loadMergedReleaseNotes,
    filterNotesForViewer,
    setLastSeenAt
} from './release-notes-store.js';

function viewerForPage() {
    const bereich = new URLSearchParams(location.search).get('bereich');
    return bereich === 'it' ? 'it' : 'enduser';
}

function whenAuthReady() {
    return new Promise(function (resolve) {
        let done = false;
        const finish = function () {
            if (done) return;
            done = true;
            window.removeEventListener('ms365-auth-widget-ready', finish);
            resolve();
        };
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            finish();
            return;
        }
        window.addEventListener('ms365-auth-widget-ready', finish);
        setTimeout(finish, 2500);
    });
}

function latestAt(notes) {
    return notes.reduce(function (acc, note) {
        const t = new Date(note.at).getTime();
        if (!Number.isFinite(t)) return acc;
        if (!acc || t > new Date(acc).getTime()) return note.at;
        return acc;
    }, '');
}

async function boot() {
    const list = document.getElementById('newsPageList');
    const lead = document.getElementById('newsPageLead');
    if (!list) return;
    await whenAuthReady();
    const viewer = viewerForPage();
    let notes = [];
    try {
        notes = await loadMergedReleaseNotes(localStorage);
    } catch {
        notes = [];
    }
    const visible = filterNotesForViewer(notes, viewer);
    if (lead) {
        lead.textContent =
            viewer === 'it'
                ? 'Einträge für alle und für die Schul-IT. Das ist das IT-Dashboard.'
                : 'Einträge für alle und für Endnutzer. Beiträge nur für die Schul-IT stehen im IT-Dashboard.';
    }
    list.replaceChildren();
    if (!visible.length) {
        const empty = document.createElement('p');
        empty.textContent = 'Noch keine Neuigkeiten für diese Ansicht.';
        list.appendChild(empty);
        return;
    }
    const ui = window.ms365ReleaseNotesUi;
    visible.forEach(function (note) {
        if (ui && typeof ui.renderNoteCard === 'function') {
            list.appendChild(ui.renderNoteCard(note));
            return;
        }
        const article = document.createElement('article');
        article.className = 'rn-card';
        const title = document.createElement('h2');
        title.textContent = note.title || '(ohne Titel)';
        article.appendChild(title);
        list.appendChild(article);
    });
    const seen = latestAt(visible);
    if (seen) setLastSeenAt(seen, localStorage, viewer);
}

boot();
