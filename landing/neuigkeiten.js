(function () {
    'use strict';

    var list = document.getElementById('landingNewsList');
    if (!list) return;

    function audienceOf(note) {
        var raw = String((note && note.audience) || 'all')
            .trim()
            .toLowerCase();
        if (raw === 'it' || raw === 'enduser') return raw;
        return 'all';
    }

    function kindLabel(kind) {
        if (kind === 'fix') return 'Bugfix';
        if (kind === 'other') return 'Hinweis';
        return 'Neuigkeit';
    }

    function formatWhen(iso) {
        var d = new Date(iso);
        if (!Number.isFinite(d.getTime())) return '';
        return d.toLocaleDateString('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric' });
    }

    function render(notes) {
        list.replaceChildren();
        var publicNotes = notes.filter(function (note) {
            return audienceOf(note) === 'all' && String(note.title || '').trim();
        });
        if (!publicNotes.length) {
            var empty = document.createElement('p');
            empty.textContent = 'Noch keine öffentlichen Neuigkeiten.';
            list.appendChild(empty);
            return;
        }
        publicNotes.forEach(function (note) {
            var article = document.createElement('article');
            article.className = 'news-card';

            var meta = document.createElement('p');
            meta.className = 'news-card__meta';
            var badge = document.createElement('span');
            badge.className = 'news-card__badge';
            badge.textContent = kindLabel(note.kind);
            meta.appendChild(badge);
            var when = document.createElement('time');
            when.dateTime = note.at || '';
            when.textContent = formatWhen(note.at);
            meta.appendChild(when);
            article.appendChild(meta);

            var title = document.createElement('h2');
            title.textContent = note.title;
            article.appendChild(title);

            if (note.bodyHtml) {
                var body = document.createElement('div');
                body.className = 'news-card__body';
                body.innerHTML = String(note.bodyHtml);
                article.appendChild(body);
            }

            var images = Array.isArray(note.images) ? note.images : [];
            if (images.length) {
                var gallery = document.createElement('div');
                gallery.className = 'news-card__gallery';
                images.forEach(function (img) {
                    if (!img || !img.src) return;
                    var figure = document.createElement('figure');
                    var image = document.createElement('img');
                    image.src = img.src;
                    image.alt = img.alt || note.title || 'Screenshot';
                    var caption = document.createElement('figcaption');
                    caption.textContent = img.alt || '';
                    figure.appendChild(image);
                    if (caption.textContent) figure.appendChild(caption);
                    gallery.appendChild(figure);
                });
                article.appendChild(gallery);
            }

            list.appendChild(article);
        });
    }

    var PUBLIC_API =
        'https://func-ms365-license-dev.azurewebsites.net/api/license/release-notes?view=public';

    function asNotes(data) {
        if (Array.isArray(data)) return data;
        if (data && Array.isArray(data.notes) && !data.missing) return data.notes;
        return null;
    }

    function loadStatic() {
        return fetch('release-notes.json', { cache: 'no-cache' }).then(function (res) {
            if (!res.ok) throw new Error('status ' + res.status);
            return res.json();
        });
    }

    fetch(PUBLIC_API, { cache: 'no-store' })
        .then(function (res) {
            if (!res.ok) throw new Error('status ' + res.status);
            return res.json();
        })
        .then(function (data) {
            var notes = asNotes(data);
            if (!notes) return loadStatic();
            return notes;
        })
        .catch(function () {
            return loadStatic();
        })
        .then(function (data) {
            render(Array.isArray(data) ? data : asNotes(data) || []);
        })
        .catch(function () {
            list.replaceChildren();
            var fail = document.createElement('p');
            fail.textContent = 'Die Neuigkeiten konnten gerade nicht geladen werden.';
            list.appendChild(fail);
        });
})();
